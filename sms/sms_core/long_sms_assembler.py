import copy
from dataclasses import replace
from datetime import datetime
import hashlib
import re
import time

from sms_core.sms_collected_event import CollectedPendingCandidates, pending_from_collected


SMS_CALLBACK_TIMESTAMP_RE = re.compile(
    r"^\s*\S+\s+(?P<timestamp>\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+)"
)
DEFAULT_LONG_SMS_TTL = 180.0
DEFAULT_COMPLETED_SMS_TTL = 30.0
DEFAULT_COMPLETED_SMS_DUPLICATE_GRACE = 5.0
DEFAULT_INCOMPLETE_SMS_TTL = 300.0
DEFAULT_LONG_SMS_SESSION_WINDOW = 60.0
DEFAULT_INCOMPLETE_SMS_PARTIAL_WAIT = 30.0
DEFAULT_MAX_PENDING_LONG_SMS = 512
DEFAULT_MAX_COMPLETED_LONG_SMS = 512
_DEFAULT_TIMEOUT = object()


class LongSmsAssembler:
    def __init__(
        self,
        parse_callback_head,
        timeout=_DEFAULT_TIMEOUT,
        completed_timeout: float = DEFAULT_COMPLETED_SMS_TTL,
        completed_duplicate_grace: float = DEFAULT_COMPLETED_SMS_DUPLICATE_GRACE,
        incomplete_timeout=None,
        multipart_timestamp_tolerance: float = 5.0,
        multipart_part_timestamp_step: float = 1.0,
        session_window: float = DEFAULT_LONG_SMS_SESSION_WINDOW,
        incomplete_emit_wait: float = DEFAULT_INCOMPLETE_SMS_PARTIAL_WAIT,
        max_pending_entries: int = DEFAULT_MAX_PENDING_LONG_SMS,
        max_completed_entries: int = DEFAULT_MAX_COMPLETED_LONG_SMS,
    ):
        self.parse_callback_head = parse_callback_head
        timeout_explicit = timeout is not _DEFAULT_TIMEOUT
        self._timeout_explicit = timeout_explicit
        self.timeout = DEFAULT_LONG_SMS_TTL if not timeout_explicit else float(timeout)
        self.completed_timeout = float(completed_timeout)
        self.completed_duplicate_grace = max(0.0, float(completed_duplicate_grace))
        if incomplete_timeout is not None:
            self.incomplete_timeout = float(incomplete_timeout)
        elif timeout_explicit:
            self.incomplete_timeout = self.timeout
        else:
            self.incomplete_timeout = DEFAULT_INCOMPLETE_SMS_TTL
        self._pending = {}
        self._completed = {}
        self._deferred = {}
        self.multipart_timestamp_tolerance = float(multipart_timestamp_tolerance)
        self.multipart_part_timestamp_step = float(multipart_part_timestamp_step)
        self.session_window = max(1.0, float(session_window or DEFAULT_LONG_SMS_SESSION_WINDOW))
        self.incomplete_emit_wait = max(0.0, float(incomplete_emit_wait or 0.0))
        self.max_pending_entries = max(1, int(max_pending_entries or DEFAULT_MAX_PENDING_LONG_SMS))
        self.max_completed_entries = max(1, int(max_completed_entries or DEFAULT_MAX_COMPLETED_LONG_SMS))

    def reset(self):
        self._pending.clear()
        self._completed.clear()
        self._deferred.clear()

    def add_collected(self, collected, correction_cache=None, now=None, log=None):
        pending = pending_from_collected(
            collected,
            self.parse_callback_head,
            correction_cache=correction_cache,
            now=now,
        )
        if isinstance(pending, CollectedPendingCandidates):
            pending_snapshot = copy.deepcopy(self._pending)
            completed_snapshot = copy.deepcopy(self._completed)
            strict_timestamp = self._candidate_batch_needs_strict_timestamp(pending.segments)
            results = self._add_candidate_segments(
                pending.segments,
                now=now,
                log=log,
                strict_timestamp=strict_timestamp,
            )
            if results and strict_timestamp:
                self._cancel_deferred_for_segments(pending.segments)
                return _pack_pending_results(results)
            result = results[-1] if results else None
            if result is not None and (
                not pending.require_fallback_match
                or _candidate_matches_fallback(result, pending.fallback)
            ):
                return _pack_pending_results(results)
            state_changed = (
                self._pending != pending_snapshot
                or self._completed != completed_snapshot
            )
            self._pending = pending_snapshot
            self._completed = completed_snapshot
            if strict_timestamp and result is None:
                trial_results = self._add_candidate_segments(
                    pending.segments,
                    now=now,
                    log=None,
                    strict_timestamp=False,
                )
                trial_completed = {
                    key: copy.deepcopy(entry)
                    for key, entry in self._completed.items()
                    if key not in completed_snapshot or completed_snapshot.get(key) != entry
                }
                self._pending = pending_snapshot
                self._completed = completed_snapshot
                trial_result = trial_results[-1] if trial_results else None
                if trial_result is not None and (
                    not pending.require_fallback_match
                    or _candidate_matches_fallback(trial_result, pending.fallback)
                ):
                    self._defer_candidate_results(
                        pending.segments,
                        trial_results,
                        trial_completed,
                        now=now,
                        log=log,
                    )
                    return None
            if result is None and not state_changed:
                return None
            return pending.fallback
        if isinstance(pending, list):
            results = self._add_candidate_segments(
                pending,
                now=now,
                log=log,
                strict_timestamp=self._candidate_batch_needs_strict_timestamp(pending),
            )
            if results:
                self._cancel_deferred_for_segments(pending)
            return _pack_pending_results(results)
        return self.add_message(pending, now=now, log=log)

    def drain_ready(self, now=None, log=None):
        current = time.monotonic() if now is None else now
        ready = []
        for key, entry in list(self._deferred.items()):
            if current < float(entry.get("deadline", current)):
                continue
            self._deferred.pop(key, None)
            for completed_key, completed_entry in dict(entry.get("completed_entries") or {}).items():
                target_key = self._unique_state_key(completed_key)
                restored = copy.deepcopy(completed_entry)
                restored["last_update"] = current
                restored["seen"] = current
                self._completed[target_key] = restored
            ready.extend(list(entry.get("results") or []))
            _emit_log(log, "[SMS] CONCAT AMBIGUOUS RELEASE")
        if ready:
            self._enforce_completed_limit(log=log)
        return _pack_pending_results(ready)

    def _add_candidate_segments(self, segments, *, now=None, log=None, strict_timestamp=False):
        results = []
        for item in list(segments or []):
            current = self.add_message(
                item,
                now=now,
                log=log,
                strict_timestamp=strict_timestamp,
            )
            if current is not None:
                results.append(current)
        return results

    def _candidate_batch_needs_strict_timestamp(self, segments):
        profile = self._candidate_batch_profile(segments)
        return bool(
            profile
            and profile["total"] == 2
            and len(profile["timestamps"]) > 1
        )

    def _candidate_batch_profile(self, segments):
        base = None
        timestamps = set()
        signatures = []
        for pending in list(segments or []):
            concat_info = self._concat_info_from_pending(pending)
            sender = self._sender_from_pending(pending)
            reference = getattr(concat_info, "reference", None) if concat_info is not None else None
            total = int(getattr(concat_info, "total", 0) or 0) if concat_info is not None else 0
            index = int(getattr(concat_info, "index", 0) or 0) if concat_info is not None else 0
            reference_bits = int(getattr(concat_info, "reference_bits", None) or 8) if concat_info is not None else 8
            if not sender or reference is None or total <= 1 or index < 1 or index > total:
                return None
            candidate_base = (sender, reference_bits, int(reference), total)
            if base is None:
                base = candidate_base
            elif candidate_base != base:
                return None
            timestamp = self._timestamp_from_pending(pending)
            if timestamp:
                timestamps.add(timestamp)
            signatures.append((index, timestamp, self._part_body_from_pending(pending)))
        if base is None:
            return None
        return {
            "base": base,
            "total": base[3],
            "timestamps": timestamps,
            "signatures": tuple(sorted(signatures)),
        }

    def _defer_candidate_results(
        self,
        segments,
        results,
        completed_entries,
        *,
        now=None,
        log=None,
    ):
        profile = self._candidate_batch_profile(segments)
        if not profile or not results:
            return
        current = time.monotonic() if now is None else now
        key = (profile["base"], profile["signatures"])
        if key in self._deferred:
            return
        self._deferred[key] = {
            "base": profile["base"],
            "deadline": current + self.completed_duplicate_grace,
            "results": tuple(results),
            "completed_entries": copy.deepcopy(completed_entries),
        }
        _emit_log(log, "[SMS] CONCAT AMBIGUOUS DEFER")

    def _cancel_deferred_for_segments(self, segments):
        profile = self._candidate_batch_profile(segments)
        if not profile:
            return
        base = profile["base"]
        self._deferred = {
            key: entry
            for key, entry in self._deferred.items()
            if entry.get("base") != base
        }

    def _unique_state_key(self, key):
        if key not in self._pending and key not in self._completed:
            return key
        sequence = 1
        candidate = tuple(key) + (sequence,)
        while candidate in self._pending or candidate in self._completed:
            sequence += 1
            candidate = tuple(key) + (sequence,)
        return candidate

    def cleanup(self, now=None, log=None):
        current = time.monotonic() if now is None else now
        expired_pending = [
            entry
            for entry in self._pending.values()
            if current - entry.get("last_update", entry.get("seen", current)) > self._pending_ttl(entry)
        ]
        for entry in expired_pending:
            self._log_timeout(log, entry)
        self._pending = {
            key: entry
            for key, entry in self._pending.items()
            if current - entry.get("last_update", entry.get("seen", current)) <= self._pending_ttl(entry)
        }
        self._completed = {
            key: entry
            for key, entry in self._completed.items()
            if current - entry.get("last_update", entry.get("seen", current)) <= self.completed_timeout
        }
        self._enforce_completed_limit(log=log)

    def add_message(self, pending, now=None, log=None, *, strict_timestamp=False):
        if pending is None:
            return None

        current = time.monotonic() if now is None else now
        self.cleanup(current, log=log)

        concat_info = self._concat_info_from_pending(pending)
        if concat_info is None or int(concat_info.total or 0) <= 1:
            return pending

        total = int(concat_info.total or 0)
        index = int(concat_info.index or 0)
        if total <= 1 or index < 1 or index > total:
            return pending

        sender = self._sender_from_pending(pending)
        if not sender:
            return pending
        if getattr(concat_info, "reference", None) is None:
            return pending
        reference = int(concat_info.reference)
        timestamp = self._timestamp_from_pending(pending)
        callback_timestamp = self._callback_timestamp_from_pending(pending)
        reference_bits = int(getattr(concat_info, "reference_bits", None) or 8)
        body = self._part_body_from_pending(pending)
        key = self._resolve_pending_key(
            sender,
            reference_bits,
            reference,
            total,
            current,
            index,
            body,
            timestamp,
            callback_timestamp,
            strict_timestamp,
        )
        if key is None:
            key = self._new_pending_key(sender, reference_bits, reference, total, current)
        trace_id = message_trace_id(
            sender,
            f"bucket:{self._key_bucket(key)}",
            reference_bits,
            reference,
            total,
        )

        if self._is_completed_duplicate(
            sender,
            reference_bits,
            reference,
            total,
            index,
            body,
            timestamp,
            callback_timestamp,
            strict_timestamp,
            current,
        ):
            self._log_concat(log, sender, reference, index, total, "duplicate complete", trace_id)
            return None

        entry = self._pending.get(key)
        if entry is None:
            entry = {
                "sender": sender,
                "reference": reference,
                "reference_bits": reference_bits,
                "timestamp": timestamp,
                "callback_timestamp": callback_timestamp,
                "total": total,
                "parts": {},
                "parts_seen": set(),
                "part_timestamps": {},
                "part_callback_timestamps": {},
                "heads": {},
                "display_headers": {},
                "trace_id": trace_id,
                "bucket": self._key_bucket(key),
                "first_seen": current,
                "last_update": current,
                "seen": current,
            }
            self._pending[key] = entry
            self._enforce_pending_limit(log=log)
        elif index in entry["parts_seen"]:
            self._log_concat(
                log,
                sender,
                reference,
                index,
                total,
                len(entry["parts"]),
                entry.get("trace_id") or trace_id,
            )
            return self._maybe_emit_incomplete(key, entry, pending, current, log)
        else:
            entry["trace_id"] = entry.get("trace_id") or trace_id

        entry["last_update"] = current
        entry["seen"] = current
        entry["total"] = total

        entry["parts_seen"].add(index)
        entry["parts"][index] = body
        entry.setdefault("part_timestamps", {})[index] = timestamp
        entry.setdefault("part_callback_timestamps", {})[index] = callback_timestamp
        entry["heads"][index] = getattr(pending, "callback_head", "")
        display_lines = list(getattr(pending, "display_lines", []) or [])
        entry["display_headers"][index] = display_lines[0] if display_lines else ""

        cache_count = len(entry["parts"])
        self._log_concat(log, sender, reference, index, total, cache_count, trace_id)

        if not self._is_complete(entry):
            return self._maybe_emit_incomplete(key, entry, pending, current, log)

        complete_body = "".join(entry["parts"].get(i, "") for i in range(1, total + 1))
        first_head = self._first_callback_head(entry, pending)
        first_body = entry["parts"].get(1, "")
        callback_head = self._replace_callback_body(first_head, first_body, complete_body)
        display_header = self._first_display_header(entry, display_lines)
        complete = replace(
            pending,
            callback_head=callback_head,
            full_msg=complete_body,
            display_lines=[display_header, callback_head] if display_header else [callback_head],
            concat_info=None,
            concat_body=None,
            concat_sender=None,
            concat_timestamp=None,
            concat_reference=reference,
            concat_reference_bits=reference_bits,
            concat_total=total,
            message_trace_id=entry.get("trace_id") or trace_id,
        )
        self._pending.pop(key, None)
        self._completed[key] = {
            "sender": sender,
            "reference": reference,
            "reference_bits": reference_bits,
            "timestamp": entry.get("timestamp") or timestamp,
            "bucket": entry.get("bucket"),
            "first_seen": entry.get("first_seen", current),
            "last_update": current,
            "seen": current,
            "parts": dict(entry["parts"]),
            "part_timestamps": dict(entry.get("part_timestamps") or {}),
            "part_callback_timestamps": dict(entry.get("part_callback_timestamps") or {}),
            "total": total,
            "trace_id": entry.get("trace_id") or trace_id,
        }
        self._enforce_completed_limit(log=log)
        self._log_complete(log, sender, reference, len(complete_body), entry.get("trace_id") or trace_id)
        return complete

    def _new_pending_key(self, sender, reference_bits, reference, total, current):
        base = (
            sender,
            int(reference_bits),
            int(reference),
            int(total),
            self._session_bucket(current),
        )
        if base not in self._pending and base not in self._completed:
            return base
        sequence = 1
        while base + (sequence,) in self._pending or base + (sequence,) in self._completed:
            sequence += 1
        return base + (sequence,)

    def _resolve_pending_key(
        self,
        sender,
        reference_bits,
        reference,
        total,
        current,
        index,
        body,
        timestamp,
        callback_timestamp,
        strict_timestamp,
    ):
        base = (sender, int(reference_bits), int(reference), int(total))
        open_candidates = []
        duplicate_candidates = []
        for key, entry in self._pending.items():
            if key[:4] != base:
                continue
            first_seen = entry.get("first_seen", entry.get("seen", current))
            if current - first_seen > self.session_window:
                continue
            parts = entry.get("parts") or {}
            if index in (entry.get("parts_seen") or set()):
                if str(parts.get(index, "")) == str(body or ""):
                    affinity = self._pending_candidate_affinity(
                        entry,
                        index,
                        timestamp,
                        callback_timestamp,
                        strict_callback_timestamp=bool(strict_timestamp or int(total) == 2),
                        strict_pdu_timestamp=bool(strict_timestamp),
                    )
                    if affinity is not None:
                        duplicate_candidates.append((key, entry, affinity))
                continue
            affinity = self._pending_candidate_affinity(
                entry,
                index,
                timestamp,
                callback_timestamp,
                strict_callback_timestamp=bool(strict_timestamp or int(total) == 2),
                strict_pdu_timestamp=bool(strict_timestamp),
            )
            if affinity is not None:
                open_candidates.append((key, entry, affinity))
        if open_candidates:
            open_candidates.sort(key=lambda item: item[2])
            best_affinity = open_candidates[0][2]
            if sum(1 for _key, _entry, affinity in open_candidates if affinity == best_affinity) != 1:
                return None
            return open_candidates[0][0]
        if not duplicate_candidates:
            return None
        duplicate_candidates.sort(
            key=lambda item: (
                item[2],
                -float(item[1].get("last_update", item[1].get("seen", 0.0)) or 0.0),
            ),
        )
        return duplicate_candidates[0][0]

    def _pending_candidate_affinity(
        self,
        entry,
        index,
        timestamp,
        callback_timestamp,
        *,
        strict_callback_timestamp=False,
        strict_pdu_timestamp=False,
    ):
        callback_values = dict(entry.get("part_callback_timestamps") or {})
        if not callback_values and entry.get("callback_timestamp"):
            callback_values[0] = entry.get("callback_timestamp")
        pdu_values = dict(entry.get("part_timestamps") or {})
        if not pdu_values and entry.get("timestamp"):
            pdu_values[0] = entry.get("timestamp")

        callback_comparable = bool(callback_timestamp and any(callback_values.values()))
        pdu_comparable = bool(timestamp and any(pdu_values.values()))
        callback_affinity = self._timestamp_affinity(
            callback_timestamp,
            index,
            callback_values,
            exact_only=strict_callback_timestamp,
        )
        pdu_affinity = self._timestamp_affinity(
            timestamp,
            index,
            pdu_values,
            exact_only=strict_pdu_timestamp,
        )

        if callback_comparable and callback_affinity is None:
            return None
        if (
            pdu_comparable
            and pdu_affinity is None
            and (strict_pdu_timestamp or not callback_comparable)
        ):
            return None

        missing = (2, float("inf"), float("inf"))
        return (
            0 if callback_affinity is not None else 1,
            *(callback_affinity or missing),
            0 if pdu_affinity is not None else 1,
            *(pdu_affinity or missing),
        )

    def _timestamp_affinity(self, timestamp, index, values_by_index, *, exact_only=False):
        incoming_text = str(timestamp or "").strip()
        if not incoming_text:
            return None

        incoming_dt = _parse_sms_timestamp(incoming_text)
        matches = []
        for existing_index, existing in dict(values_by_index or {}).items():
            existing_text = str(existing or "").strip()
            if not existing_text:
                continue
            if incoming_text == existing_text:
                matches.append((0, 0.0, 0.0))
                continue
            existing_dt = _parse_sms_timestamp(existing_text)
            if incoming_dt is None or existing_dt is None:
                continue
            signed_delta = (incoming_dt - existing_dt).total_seconds()
            raw_delta = abs(signed_delta)
            if raw_delta == 0:
                matches.append((0, 0.0, 0.0))
                continue
            if exact_only:
                continue
            try:
                expected_delta = (int(index) - int(existing_index)) * self.multipart_part_timestamp_step
            except Exception:
                expected_delta = 0.0
            adjusted_delta = abs(signed_delta - expected_delta)
            distance = min(raw_delta, adjusted_delta)
            if distance <= self.multipart_timestamp_tolerance:
                matches.append((1, distance, raw_delta))
        return min(matches) if matches else None

    def _key_bucket(self, key):
        try:
            return key[4]
        except Exception:
            return 0

    def _session_bucket(self, current):
        try:
            return int(float(current) // self.session_window)
        except Exception:
            return 0

    def _concat_info_from_pending(self, pending):
        concat_info = getattr(pending, "concat_info", None)
        if concat_info is not None:
            return concat_info
        pdu = getattr(pending, "pdu", None)
        if pdu is not None:
            return getattr(pdu, "concat_info", None)
        return None

    def _part_body_from_pending(self, pending):
        body = getattr(pending, "concat_body", None)
        if body is not None:
            return str(body)
        return str(getattr(pending, "full_msg", "") or "")

    def _sender_from_pending(self, pending):
        concat_sender = str(getattr(pending, "concat_sender", "") or "").strip()
        if concat_sender:
            return _normalize_sender_for_key(concat_sender)

        sender, _body = self.parse_callback_head(getattr(pending, "callback_head", ""))
        return _normalize_sender_for_key(sender)

    def _timestamp_from_pending(self, pending):
        timestamp = str(getattr(pending, "concat_timestamp", "") or "").strip()
        if timestamp:
            return timestamp
        match = SMS_CALLBACK_TIMESTAMP_RE.match(str(getattr(pending, "callback_head", "") or ""))
        return match.group("timestamp") if match else ""

    def _callback_timestamp_from_pending(self, pending):
        match = SMS_CALLBACK_TIMESTAMP_RE.match(str(getattr(pending, "callback_head", "") or ""))
        return match.group("timestamp") if match else ""

    def _is_completed_duplicate(
        self,
        sender,
        reference_bits,
        reference,
        total,
        index,
        body,
        timestamp,
        callback_timestamp,
        strict_timestamp,
        current,
    ):
        for key, completed in self._completed.items():
            key_sender, key_bits, key_ref, key_total = key[:4]
            if (
                key_sender != sender
                or int(key_bits) != int(reference_bits)
                or int(key_ref) != int(reference)
                or int(key_total) != int(total)
            ):
                continue
            last_update = completed.get("last_update", completed.get("seen", current))
            if current - last_update > self.completed_duplicate_grace:
                continue
            parts = completed.get("parts") or {}
            if str(parts.get(index, "")) != str(body or ""):
                continue
            callback_values = dict(completed.get("part_callback_timestamps") or {})
            pdu_values = dict(completed.get("part_timestamps") or {})
            timestamp_checks = []
            if callback_timestamp and any(callback_values.values()):
                timestamp_checks.append(
                    self._timestamp_affinity(
                        callback_timestamp,
                        index,
                        callback_values,
                        exact_only=True,
                    ) is not None
                )
            if timestamp and any(pdu_values.values()):
                timestamp_checks.append(
                    self._timestamp_affinity(
                        timestamp,
                        index,
                        pdu_values,
                        exact_only=True,
                    ) is not None
                )
            if strict_timestamp and timestamp and any(pdu_values.values()):
                if not timestamp_checks[-1]:
                    continue
                return True
            if timestamp_checks and not all(timestamp_checks):
                continue
            return True
        return False

    def _pending_ttl(self, entry):
        total = int(entry.get("total") or 0)
        parts = entry.get("parts") or {}
        if total > 1 and len(parts) < total:
            return self.incomplete_timeout
        return self.timeout

    def _is_complete(self, entry):
        total = int(entry.get("total") or 0)
        parts = entry.get("parts") or {}
        return total > 0 and len(parts) == total and all(i in parts for i in range(1, total + 1))

    def _maybe_emit_incomplete(self, key, entry, pending, current, log=None):
        total = int(entry.get("total") or 0)
        parts = entry.get("parts") or {}
        if total <= 1 or len(parts) != total - 1:
            return None
        first_seen = entry.get("first_seen", entry.get("seen", current))
        if current - first_seen <= self.incomplete_emit_wait:
            return None

        missing = [i for i in range(1, total + 1) if i not in parts]
        complete_body = "[INCOMPLETE SMS]\n" + "".join(
            parts.get(i, f"[MISSING PART {i}/{total}]")
            for i in range(1, total + 1)
        )
        first_head = self._first_callback_head(entry, pending)
        first_body = parts.get(1, "")
        callback_head = self._replace_callback_body(first_head, first_body, complete_body)
        display_header = self._first_display_header(
            entry,
            list(getattr(pending, "display_lines", []) or []),
        )
        trace_id = entry.get("trace_id") or ""
        incomplete = replace(
            pending,
            callback_head=callback_head,
            full_msg=complete_body,
            display_lines=[display_header, callback_head] if display_header else [callback_head],
            concat_info=None,
            concat_body=None,
            concat_sender=None,
            concat_timestamp=None,
            concat_reference=entry.get("reference"),
            concat_reference_bits=entry.get("reference_bits"),
            concat_total=total,
            message_trace_id=trace_id,
        )
        self._pending.pop(key, None)
        self._completed[key] = {
            "sender": entry.get("sender"),
            "reference": entry.get("reference"),
            "reference_bits": entry.get("reference_bits"),
            "timestamp": entry.get("timestamp"),
            "bucket": entry.get("bucket"),
            "first_seen": entry.get("first_seen", current),
            "last_update": current,
            "seen": current,
            "parts": dict(parts),
            "part_timestamps": dict(entry.get("part_timestamps") or {}),
            "part_callback_timestamps": dict(entry.get("part_callback_timestamps") or {}),
            "total": total,
            "trace_id": trace_id,
            "incomplete": True,
            "missing": missing,
        }
        self._enforce_completed_limit(log=log)
        self._log_incomplete(log, entry, missing)
        return incomplete

    def _first_callback_head(self, entry, pending):
        heads = entry.get("heads") or {}
        if 1 in heads:
            return heads.get(1) or getattr(pending, "callback_head", "")
        if heads:
            first_index = min(heads)
            return heads.get(first_index) or getattr(pending, "callback_head", "")
        return getattr(pending, "callback_head", "")

    def _first_display_header(self, entry, fallback_lines):
        headers = entry.get("display_headers") or {}
        if 1 in headers:
            return headers.get(1) or (fallback_lines[0] if fallback_lines else "")
        if headers:
            first_index = min(headers)
            return headers.get(first_index) or (fallback_lines[0] if fallback_lines else "")
        return fallback_lines[0] if fallback_lines else ""

    def _enforce_pending_limit(self, log=None):
        while len(self._pending) > self.max_pending_entries:
            oldest_key = min(
                self._pending,
                key=lambda key: self._pending[key].get(
                    "first_seen",
                    self._pending[key].get("seen", 0.0),
                ),
            )
            entry = self._pending.pop(oldest_key, None)
            if entry is not None:
                self._log_evicted(log, entry, "PENDING")

    def _enforce_completed_limit(self, log=None):
        while len(self._completed) > self.max_completed_entries:
            oldest_key = min(
                self._completed,
                key=lambda key: self._completed[key].get(
                    "last_update",
                    self._completed[key].get("seen", 0.0),
                ),
            )
            entry = self._completed.pop(oldest_key, None)
            if entry is not None:
                self._log_evicted(log, entry, "COMPLETED")

    def _log_concat(self, log, sender, reference, index, total, cache_count, trace_id):
        cache_text = str(cache_count)
        _emit_log(
            log,
            (
                "[SMS] SMS CONCAT "
                f"Sender={sender or 'unknown'} Ref=0x{reference:X} "
                f"Part={index}/{total} Cache={cache_text}/{total} Trace={trace_id}"
            ),
        )

    def _log_complete(self, log, sender, reference, length, trace_id):
        _emit_log(
            log,
            (
                "[SMS] SMS CONCAT COMPLETE "
                f"Sender={sender or 'unknown'} Ref=0x{reference:X} Length={length} Trace={trace_id}"
            ),
        )

    def _log_timeout(self, log, entry):
        parts = entry.get("parts") or {}
        total = int(entry.get("total") or 0)
        reference = int(entry.get("reference") or 0)
        trace_id = entry.get("trace_id") or ""
        _emit_log(
            log,
            (
                "[SMS] SMS CONCAT TIMEOUT "
                f"Sender={entry.get('sender') or 'unknown'} Ref=0x{reference:X} "
                f"Parts={len(parts)}/{total} Trace={trace_id}"
            ),
        )

    def _log_incomplete(self, log, entry, missing):
        parts = entry.get("parts") or {}
        total = int(entry.get("total") or 0)
        reference = int(entry.get("reference") or 0)
        trace_id = entry.get("trace_id") or ""
        missing_text = ",".join(str(item) for item in list(missing or []))
        _emit_log(
            log,
            (
                "[SMS] SMS CONCAT INCOMPLETE "
                f"Sender={entry.get('sender') or 'unknown'} Ref=0x{reference:X} "
                f"Parts={len(parts)}/{total} Missing={missing_text} Trace={trace_id}"
            ),
        )

    def _log_evicted(self, log, entry, cache_name):
        parts = entry.get("parts") or {}
        total = int(entry.get("total") or 0)
        reference = int(entry.get("reference") or 0)
        trace_id = entry.get("trace_id") or ""
        _emit_log(
            log,
            (
                "[SMS] SMS CONCAT EVICT "
                f"Cache={cache_name} Sender={entry.get('sender') or 'unknown'} "
                f"Ref=0x{reference:X} Parts={len(parts)}/{total} Trace={trace_id}"
            ),
        )

    def _replace_callback_body(self, callback_head, old_body, new_body):
        head = str(callback_head or "")
        old = str(old_body or "")
        new = str(new_body or "")
        if not head:
            return new
        if old and head.endswith(old):
            return head[:-len(old)] + new

        try:
            sender, parsed_body = self.parse_callback_head(head)
        except Exception:
            sender, parsed_body = "", ""
        parsed_body = str(parsed_body or "")
        if parsed_body and (sender or parsed_body != head) and head.endswith(parsed_body):
            return head[:-len(parsed_body)] + new

        if "\ufffd" in head:
            return new
        return (head + " " + new).strip()


def _emit_log(log, message):
    if log is None:
        return
    try:
        log(message, "normal")
    except TypeError:
        try:
            log(message)
        except Exception:
            pass
    except Exception:
        pass


def _pack_pending_results(results):
    items = list(results or [])
    if not items:
        return None
    if len(items) == 1:
        return items[0]
    return items


def _candidate_matches_fallback(candidate, fallback):
    candidate_body = _normalized_body(getattr(candidate, "full_msg", ""))
    fallback_body = _normalized_body(getattr(fallback, "full_msg", ""))
    if "\ufffd" in fallback_body:
        return _corrupted_fallback_matches_candidate(candidate_body, fallback_body)
    return (
        candidate_body == fallback_body
        or _body_without_line_breaks(candidate_body) == _body_without_line_breaks(fallback_body)
    )


def _corrupted_fallback_matches_candidate(candidate_body, fallback_body):
    fragments = []
    for fragment in str(fallback_body or "").split("\ufffd"):
        cleaned = _body_without_line_breaks(fragment)
        if len(cleaned) >= 4:
            fragments.append(cleaned)
    if len(fragments) < 2:
        return False

    candidate_text = _body_without_line_breaks(candidate_body)
    cursor = 0
    matched = 0
    for fragment in fragments:
        index = candidate_text.find(fragment, cursor)
        if index < 0:
            return False
        cursor = index + len(fragment)
        matched += 1
    return matched >= 2


def _normalized_body(value):
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def _body_without_line_breaks(value):
    return _normalized_body(value).replace("\n", "")


def _normalize_sender_for_key(sender: str) -> str:
    text = str(sender or "").strip()
    if text.startswith("+86") and len(text) > 3:
        return text[3:]
    if text.startswith("86") and len(text) > 2:
        return text[2:]
    return text


def _parse_sms_timestamp(value: str):
    text = str(value or "").strip()
    try:
        return datetime.strptime(text[:17], "%y/%m/%d,%H:%M:%S")
    except Exception:
        return None


def message_trace_id(sender, timestamp, reference_bits, reference, total):
    raw = "\n".join([
        str(sender or ""),
        str(timestamp or ""),
        str(reference_bits or ""),
        str(reference or ""),
        str(total or ""),
    ])
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]

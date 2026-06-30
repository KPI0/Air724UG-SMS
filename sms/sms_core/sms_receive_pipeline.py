"""
SMS receive pipeline contract:
- observe raw serial lines for side caches;
- forward collected SMS frames to LongSmsAssembler;
- never inspect SMS type, patch SMS body, or group concat parts.
"""

class SmsReceivePipeline:
    """Blind transport from collected serial SMS frames to LongSmsAssembler."""

    def __init__(self, parse_callback_head, correction_cache, long_sms_assembler):
        self.parse_callback_head = parse_callback_head
        self.correction_cache = correction_cache
        self.long_sms_assembler = long_sms_assembler

    def observe_line(self, line: str, now: float, log=None):
        try:
            self.correction_cache.observe_line(line, now, log=log)
        except TypeError:
            self.correction_cache.observe_line(line, now)

    def add_collected(self, collected, now=None, log=None):
        return self.long_sms_assembler.add_collected(
            collected,
            correction_cache=self.correction_cache,
            now=now,
            log=log,
        )

    def reset(self):
        self.long_sms_assembler.reset()

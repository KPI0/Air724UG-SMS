from functools import wraps


def _mapped_call_args(
    runtime_name,
    args,
    kwargs,
    positional_keywords,
    positional_prefix_count,
):
    if not positional_keywords:
        return args, kwargs

    prefix_args = args[:positional_prefix_count]
    keyword_args = args[positional_prefix_count:]
    keyword_count = min(len(keyword_args), len(positional_keywords))
    mapped_kwargs = {}
    for name, value in zip(positional_keywords[:keyword_count], keyword_args[:keyword_count]):
        if name in kwargs:
            raise TypeError(f"{runtime_name}() got multiple values for argument '{name}'")
        mapped_kwargs[name] = value

    if not mapped_kwargs:
        return args, kwargs

    call_kwargs = dict(mapped_kwargs)
    call_kwargs.update(kwargs)
    return prefix_args + keyword_args[keyword_count:], call_kwargs


def bind_namespace_runtime(
    namespace,
    module_globals,
    runtime_name,
    *,
    positional_keywords=(),
    positional_prefix_count=0,
):
    """Bind a namespace-first runtime while keeping late module-level lookup patchable."""
    runtime = module_globals[runtime_name]
    positional_keywords = tuple(positional_keywords)

    @wraps(runtime)
    def bound_runtime(*args, **kwargs):
        call_args, call_kwargs = _mapped_call_args(
            runtime_name,
            args,
            kwargs,
            positional_keywords,
            positional_prefix_count,
        )
        return module_globals[runtime_name](namespace, *call_args, **call_kwargs)

    return bound_runtime


def bind_async_namespace_runtime(
    namespace,
    module_globals,
    runtime_name,
    *,
    positional_keywords=(),
    positional_prefix_count=0,
):
    """Async variant of bind_namespace_runtime."""
    runtime = module_globals[runtime_name]
    positional_keywords = tuple(positional_keywords)

    @wraps(runtime)
    async def bound_runtime(*args, **kwargs):
        call_args, call_kwargs = _mapped_call_args(
            runtime_name,
            args,
            kwargs,
            positional_keywords,
            positional_prefix_count,
        )
        return await module_globals[runtime_name](namespace, *call_args, **call_kwargs)

    return bound_runtime


def make_namespace_runtime_binder(namespace, module_globals):
    def bind(runtime_name, **binding_options):
        return bind_namespace_runtime(
            namespace,
            module_globals,
            runtime_name,
            **binding_options,
        )

    return bind


def make_async_namespace_runtime_binder(namespace, module_globals):
    def bind(runtime_name, **binding_options):
        return bind_async_namespace_runtime(
            namespace,
            module_globals,
            runtime_name,
            **binding_options,
        )

    return bind

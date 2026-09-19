"""Best-effort, monotonic request spans; metadata only, never document payloads."""
from contextlib import contextmanager
from contextvars import ContextVar
from functools import wraps
import time
import uuid

_current = ContextVar('wps_request_timing', default=None)
_fields = ContextVar('wps_timing_fields', default={})
_parent = ContextVar('wps_timing_parent', default=None)


class RequestTiming:
    def __init__(self, journal, *, clock=time.perf_counter_ns, application="word"):
        self.journal = journal
        self.application = application
        self.clock = clock
        self.request_id = 'request-' + uuid.uuid4().hex
        self.task_id = None

    def emit(self, event, **fields):
        try:
            self.journal.request_event(request_id=self.request_id, event=event,
                                       taskId=self.task_id, application=self.application, **fields)
        except Exception:
            pass


@contextmanager
def fields(**values):
    token = _fields.set(dict(_fields.get(), **values))
    try:
        yield
    finally:
        _fields.reset(token)


def bind_task(task_id):
    recorder = _current.get()
    if recorder is not None:
        recorder.task_id = task_id
        recorder.emit('request.task_bound')


@contextmanager
def span(name, **values):
    recorder = _current.get()
    result = {}
    if recorder is None:
        yield result
        return
    identity = dict(_fields.get(), **values, spanId=uuid.uuid4().hex,
                    parentSpanId=_parent.get(), name=name)
    token = _parent.set(identity['spanId'])
    started = recorder.clock()
    recorder.emit('span.started', **identity, monotonicNs=started)
    try:
        yield result
    except BaseException as exc:
        result.update(status='exception', exceptionType=type(exc).__name__)
        raise
    finally:
        ended = recorder.clock()
        recorder.emit('span.finished', **identity, **dict({'status': 'returned'}, **result),
                      monotonicNs=ended, durationNs=ended-started, durationMs=(ended-started)/1e6)
        _parent.reset(token)


@contextmanager
def submission(enabled=True, *, journal=None, clock=time.perf_counter_ns, application="word"):
    if not enabled or _current.get() is not None:
        yield
        return
    if journal is None:
        from wps_skills.core.trace_journal import JsonlTraceJournal
        journal = JsonlTraceJournal.default()
    recorder = RequestTiming(journal, clock=clock, application=application)
    token = _current.set(recorder)
    try:
        with span('request.total'):
            yield
    finally:
        _current.reset(token)


def timed(name):
    def decorate(function):
        @wraps(function)
        def wrapped(*args, **kwargs):
            with span(name):
                return function(*args, **kwargs)
        return wrapped
    return decorate

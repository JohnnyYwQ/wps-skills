"""First real Word Action handlers behind the Word Adapter seam."""

from typing import Mapping

from wps_skills.core.action_runtime import ActionError, ControllerResult
from wps_skills.windows.bridge_types import BackendActionFailure
from wps_skills.word.backend.operation import WordBackendOperation
from wps_skills.word.persistence.output_names import MAX_OUTPUT_ATTEMPTS, output_candidate


def _failure_result(failure):
    factory = (
        ControllerResult.failed
        if failure.outcome == "failed"
        else ControllerResult.unknown
    )
    return factory(
        error=ActionError(
            code=failure.code,
            message=failure.message,
        ),
        controller_state=(
            "usable"
            if failure.binding_disposition == "unchanged"
            else "broken"
        ),
        binding_disposition=failure.binding_disposition,
    )


def _invoke(backend, document, operation, context):
    try:
        if (operation.name in {"save_as_artifact", "export_pdf_artifact"}
                and operation.arguments.get("overwritePolicy") == "renameIfExists"):
            data = _invoke_with_unique_name(backend, document, operation, context)
        else:
            data = backend.invoke(document, operation, context)
    except BackendActionFailure as failure:
        return _failure_result(failure)
    if not isinstance(data, Mapping):
        raise TypeError("Word Backend Action result must be an object")
    return ControllerResult.succeeded(
        data=data,
        controller_state="usable",
        binding_disposition="unchanged",
    )


def _invoke_with_unique_name(backend, document, operation, context):
    requested = operation.arguments["outputPath"]
    for attempt in range(MAX_OUTPUT_ATTEMPTS):
        candidate = output_candidate(requested, attempt)
        arguments = dict(operation.arguments, outputPath=candidate, overwritePolicy="failIfExists")
        try:
            # Every attempt keeps the exact document and the original deadline.
            # The backend atomically rejects occupied destinations; no probing
            # through the active window and no repeated content Actions.
            data = backend.invoke(document, WordBackendOperation(operation.name, arguments), context)
        except BackendActionFailure as failure:
            if (failure.outcome != "failed" or failure.binding_disposition != "unchanged"
                    or failure.code not in {"OUTPUT_ALREADY_EXISTS", "OUTPUT_IN_USE"}):
                raise
            continue
        if not isinstance(data, Mapping):
            raise TypeError("Word Backend Action result must be an object")
        return dict(data, outputResolution={
            "requestedPath": requested, "attempts": attempt + 1, "renamed": attempt > 0,
        })
    raise BackendActionFailure(
        outcome="failed", code="OUTPUT_NAME_EXHAUSTED", binding_disposition="unchanged",
        message="No available output name after {} attempts in the requested directory: {}".format(
            MAX_OUTPUT_ATTEMPTS, requested,
        ),
    )


def _handler(operation_name):
    def invoke(backend, document, params, context):
        return _invoke(
            backend,
            document,
            WordBackendOperation(
                name=operation_name,
                arguments=params,
            ),
            context,
        )

    return invoke


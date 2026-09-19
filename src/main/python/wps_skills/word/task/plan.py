"""Word plan entry point backed by the shared Task policy."""

from wps_skills.client.task_plan import (
    DOCUMENT_STEP, SAVE_STEP, PDF_STEP, RESERVED_IDS, check_document,
    compile_request as _compile_request,
)


def compile_request(request):
    return _compile_request(request, application="word")

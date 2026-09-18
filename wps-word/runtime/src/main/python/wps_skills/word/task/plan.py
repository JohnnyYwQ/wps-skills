"""Validate three groups of Action requests and fix document/completion order."""

from wps_skills.client.task_request import TaskRequestError, _object, _identifier, validate_value


DOCUMENT_STEP = "task_document"
SAVE_STEP = "task_save"
PDF_STEP = "task_pdf"
RESERVED_IDS = frozenset({DOCUMENT_STEP, SAVE_STEP, PDF_STEP})


def _action(value, allowed):
    _object(value, ("address", "params"), ())
    _object(value["address"], ("app", "action"))
    name = value["address"]["action"]
    if value["address"]["app"] != "word" or not isinstance(name, str) or name not in allowed:
        raise TaskRequestError("Action is not allowed in this Task section: " + str(name))
    if not isinstance(value["params"], dict) or "$ref" in value["params"]:
        raise TaskRequestError("Action params must be an object")
    # Lifecycle targets are explicit intent, not locators inferred from content.
    validate_value(value["params"], set())
    return name


def _step(identifier, action):
    return dict(action, id=identifier)


def compile_request(request):
    """No alias translation: every section carries canonical address + params."""
    _object(request, ("app", "document", "steps", "completion"), ("includeExistingChanges",))
    if request["app"] != "word":
        raise TaskRequestError("Task application must be word")
    from wps_skills.word.actions.registry import WORD_ACTIONS, content_actions
    acquisition = {name for name, entry in WORD_ACTIONS.items() if entry.contract.binding_role == "establish"}
    persistence = {name for name, entry in WORD_ACTIONS.items() if entry.contract.category == "persistence"}
    document = request["document"]
    document_name = _action(document, acquisition)
    completion = request["completion"]
    if not isinstance(completion, list) or len(completion) > 2:
        raise TaskRequestError("completion must be a list with at most one save/saveAs and one exportPdf")
    save = pdf = None
    for action in completion:
        name = _action(action, persistence)
        if name == "exportPdf":
            if pdf is not None:
                raise TaskRequestError("completion may contain exportPdf only once")
            pdf = _step(PDF_STEP, action)
        else:
            if save is not None:
                raise TaskRequestError("completion may contain only one of save or saveAs")
            if document_name == "createDocument" and name == "save":
                raise TaskRequestError("A new document requires saveAs for its first save")
            save = _step(SAVE_STEP, action)
    if "includeExistingChanges" in request:
        if type(request["includeExistingChanges"]) is not bool:
            raise TaskRequestError("includeExistingChanges must be Boolean")
        if document_name != "openDocument" or save is None:
            raise TaskRequestError("includeExistingChanges only applies when saving an existing document")

    body = request["steps"]
    if not isinstance(body, list) or len(body) > 125:
        raise TaskRequestError("Word Task requires 0-125 content steps")
    allowed = content_actions()
    seen = {DOCUMENT_STEP}
    for step in body:
        _object(step, ("id", "address", "params"), ())
        _object(step["address"], ("app", "action"))
        if isinstance(step["id"], str) and step["id"] in RESERVED_IDS:
            raise TaskRequestError("Step ID is reserved for Task lifecycle", step_id=step["id"])
        _identifier(step["id"])
        if step["id"] in seen:
            raise TaskRequestError("Step IDs must be unique", step_id=step["id"])
        if step["address"]["app"] != "word":
            raise TaskRequestError("Every Action must belong to Word", step_id=step["id"])
        if not isinstance(step["params"], dict) or "$ref" in step["params"]:
            raise TaskRequestError("Step params must be an object", step_id=step["id"])
        validate_value(step["params"], seen)
        seen.add(step["id"])
        name = step["address"]["action"]
        if not isinstance(name, str) or name not in allowed:
            raise TaskRequestError(
                "steps must name a registered content Action; use document/completion for acquisition and persistence",
                step_id=step["id"],
            )
    steps = [_step(DOCUMENT_STEP, document), *body]
    # The executor owns tail ordering, independent of the caller's list order.
    steps.extend(action for action in (save, pdf) if action is not None)
    return {"app": "word", "steps": steps}


def check_document(request, response):
    """Never start content edits that would save unapproved existing changes."""
    saving = any(action["address"]["action"] in {"save", "saveAs"} for action in request["completion"])
    if (request["document"]["address"]["action"] == "openDocument"
            and saving and not request.get("includeExistingChanges", False)):
        state = response["data"]["documentState"]["persistenceState"]
        if state != "saved":
            raise TaskRequestError(
                "The document has existing unsaved changes; confirm saving them together before submitting a new request",
                code="TASK_EXISTING_CHANGES_CONFIRMATION_REQUIRED",
            )

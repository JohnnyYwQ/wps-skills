"""Semantic validation of Word Action parameters and results."""

from wps_skills.word.persistence.output_names import output_candidate


def _content_ranges(value):
    if isinstance(value, dict) or hasattr(value, "items"):
        if {"start", "end", "revision"}.issubset(value):
            yield value
        for item in value.values():
            yield from _content_ranges(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _content_ranges(item)


def _range_inside(inner, outer):
    return outer["start"] <= inner["start"] <= inner["end"] <= outer["end"]


def _ranges_are_ordered(ranges):
    return all(
        left["end"] <= right["start"]
        for left, right in zip(ranges, ranges[1:])
    )


def _revision_change_error(result, *, must_change=False):
    before = result["revisionBefore"]
    after = result["revisionAfter"]
    if must_change and before == after:
        return "revisionAfter must differ from revisionBefore"
    return None


def _result_ranges_match_revision(result, revision):
    return all(
        content_range["revision"] == revision
        for content_range in _content_ranges(result)
    )


def _input_revisions(value):
    if isinstance(value, dict) or hasattr(value, "items"):
        revision = value.get("revision")
        if isinstance(revision, str):
            yield revision
        for item in value.values():
            yield from _input_revisions(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _input_revisions(item)


def _selected_section_indexes(selector, selected_count):
    if selector["kind"] == "indexes":
        return list(selector["indexes"])
    return list(range(selected_count))


def _anchor_position(anchor):
    if anchor["kind"] == "documentStart":
        return 0
    if anchor["kind"] == "before":
        return anchor["range"]["start"]
    if anchor["kind"] == "after":
        return anchor["range"]["end"]
    return None


def _utf16_units(value):
    return len(value.encode("utf-16-le")) // 2


def _length_in_points(length):
    factors = {
        "pt": 1,
        "in": 72,
        "cm": 72 / 2.54,
        "mm": 72 / 25.4,
    }
    return length["value"] * factors[length["unit"]]


def _observed_length_matches(observed, requested, tolerance=0.5):
    return abs(observed["value"] - _length_in_points(requested)) <= tolerance


def _text_format_patches(action, params):
    if action == "writeContent":
        blocks = params["blocks"]
    elif action == "replaceContent":
        replacement = params["replacement"]
        if replacement["kind"] == "delete":
            return
        if replacement["kind"] == "text":
            blocks = ({"runs": replacement["runs"]},)
        else:
            blocks = replacement["blocks"]
    else:
        return
    for block in blocks:
        for run in block["runs"]:
            if "format" in run:
                yield run["format"]


def _semantic_params_error(action, params):
    if any(
        "fontFamily" in patch
        and (
            "westernFontFamily" in patch
            or "eastAsiaFontFamily" in patch
        )
        for patch in _text_format_patches(action, params)
    ):
        return (
            "fontFamily cannot be combined with script-specific font families"
        )
    if action == "setHeaderFooter" and any(
        update["operation"]["kind"] == "linkToPrevious"
        for update in params["updates"]
    ):
        selector = params["sections"]
        if selector["kind"] == "all" or 0 in selector["indexes"]:
            return "linkToPrevious cannot target section 0"
    return None


def _semantic_result_error(action, params, result):
    expected_input_revision = None
    if action in {"inspectDocument", "findContent"}:
        expected_input_revision = result["revision"]
    elif "revisionBefore" in result:
        expected_input_revision = result["revisionBefore"]
    if expected_input_revision is not None and any(
        revision != expected_input_revision
        for revision in _input_revisions(params)
    ):
        return "every revision-bound input must equal the observed base revision"

    if action == "openDocument":
        if result["artifact"]["path"] != params["path"]:
            return "artifact path must equal the requested document locator"

    elif action == "writeContent":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        if result["range"]["revision"] != result["revisionAfter"]:
            return "inserted range must use revisionAfter"
        anchor_position = _anchor_position(params["anchor"])
        if (
            anchor_position is not None
            and result["range"]["start"] != anchor_position
        ):
            return "inserted range must begin at the resolved Body Anchor"

    elif action == "inspectDocument":
        revision = result["revision"]
        if not _result_ranges_match_revision(result, revision):
            return "all inspection ranges must use the snapshot revision"
        if (
            params["scope"]["kind"] == "range"
            and result["scopeRange"] != params["scope"]["range"]
        ):
            return "scopeRange must equal the requested range scope"
        limits = params["limits"]
        if len(result["text"]) > limits["maxTextCharacters"]:
            return "inspection text exceeds maxTextCharacters"
        if len(result["paragraphs"]) > limits["maxParagraphs"]:
            return "inspection paragraphs exceed maxParagraphs"
        run_count = sum(
            len(paragraph["runs"])
            for paragraph in result["paragraphs"]
        )
        if run_count > limits["maxRuns"]:
            return "inspection runs exceed maxRuns"
        structure = result["structure"]
        if structure["sectionCount"] != len(structure["sections"]):
            return "sectionCount must equal the section snapshot count"
        if [section["index"] for section in structure["sections"]] != list(
            range(len(structure["sections"]))
        ):
            return "section snapshots must use complete ordered indexes"
        expected_stories = [
            (area, variant)
            for area in ("header", "footer")
            for variant in ("primary", "firstPage", "evenPages")
        ]
        for section in structure["sections"]:
            observed = [
                (story["area"], story["variant"])
                for story in section["headerFooter"]["stories"]
            ]
            if observed != expected_stories:
                return "section stories must use the canonical six-item order"
        scope_range = result["scopeRange"]
        returned_range = result["returnedRange"]
        if not _range_inside(returned_range, scope_range):
            return "returnedRange must be inside scopeRange"
        remaining = result["remainingRange"]
        if result["truncated"] != (remaining is not None):
            return "truncated must exactly reflect remainingRange"
        if remaining is None:
            if returned_range != scope_range:
                return "an untruncated returnedRange must equal scopeRange"
        elif (
            remaining["start"] != returned_range["end"]
            or remaining["end"] != scope_range["end"]
            or remaining["start"] >= remaining["end"]
            or not _range_inside(remaining, scope_range)
        ):
            return "remainingRange must be the contiguous scope suffix"
        if any(
            not _range_inside(paragraph["range"], returned_range)
            for paragraph in result["paragraphs"]
        ):
            return "paragraph ranges must be inside returnedRange"
        paragraph_ranges = [
            paragraph["range"] for paragraph in result["paragraphs"]
        ]
        if not _ranges_are_ordered(paragraph_ranges):
            return "paragraph ranges must be ordered and non-overlapping"
        for paragraph in result["paragraphs"]:
            run_ranges = [run["range"] for run in paragraph["runs"]]
            if not _ranges_are_ordered(run_ranges) or any(
                not _range_inside(run_range, paragraph["range"])
                for run_range in run_ranges
            ):
                return "run ranges must be ordered inside their paragraph"

    elif action == "findContent":
        revision = result["revision"]
        if not _result_ranges_match_revision(result, revision):
            return "all find ranges must use the result revision"
        requested_scope = params["query"]["scope"]
        if (
            requested_scope["kind"] == "range"
            and result["scopeRange"] != requested_scope["range"]
        ):
            return "scopeRange must equal the requested range scope"
        if len(result["matches"]) > params["limit"]:
            return "returned matches exceed the requested limit"
        ranges = [match["range"] for match in result["matches"]]
        if not _ranges_are_ordered(ranges):
            return "matches must be ordered and non-overlapping"
        if any(
            not _range_inside(match_range, result["scopeRange"])
            for match_range in ranges
        ):
            return "match ranges must be inside scopeRange"
        query_text = params["query"]["text"]
        for match in result["matches"]:
            observed_text = match["text"]
            if match["range"]["end"] - match["range"]["start"] != (
                _utf16_units(observed_text)
            ):
                return "each match range must span its observed UTF-16 text"
            if params["query"]["caseSensitive"]:
                text_matches = observed_text == query_text
            else:
                text_matches = observed_text.casefold() == query_text.casefold()
            if not text_matches:
                return "each observed match must equal the literal query"
        remaining = result["remainingRange"]
        if result["truncated"] != (remaining is not None):
            return "truncated must exactly reflect remainingRange"
        if remaining is not None:
            if len(ranges) != params["limit"] or not ranges:
                return "a truncated find must fill the requested match limit"
            if (
                not _range_inside(remaining, result["scopeRange"])
                or remaining["start"] != ranges[-1]["end"]
                or remaining["end"] != result["scopeRange"]["end"]
                or remaining["start"] >= remaining["end"]
            ):
                return "remainingRange must be the contiguous scope suffix"

    elif action == "replaceContent":
        matched = result["matchedCount"]
        must_change = params["replacement"]["kind"] == "delete"
        expected = (
            1
            if params["target"]["kind"] == "range"
            else params["target"]["expectedMatchCount"]
        )
        if matched != expected:
            return "matchedCount must equal the target precondition"
        if len(result["ranges"]) != matched:
            return "ranges length must equal matchedCount"
        if (
            params["target"]["kind"] == "range"
            and result["ranges"][0]["start"]
            != params["target"]["range"]["start"]
        ):
            return "an exact-range replacement must preserve its start position"
        if params["replacement"]["kind"] == "delete":
            if any(
                content_range["start"] != content_range["end"]
                for content_range in result["ranges"]
            ):
                return "deleted targets must return collapsed final ranges"
        elif any(
            content_range["start"] >= content_range["end"]
            for content_range in result["ranges"]
        ):
            return "non-empty replacements must return non-empty final ranges"
        if (
            params["target"]["kind"] == "query"
            and params["replacement"]["kind"] == "text"
        ):
            query = params["target"]["query"]
            replacement_text = "".join(
                run["text"] for run in params["replacement"]["runs"]
            )
            definitely_different = (
                replacement_text != query["text"]
                if query["caseSensitive"]
                else replacement_text.casefold() != query["text"].casefold()
            )
            must_change = definitely_different
            replacement_length = _utf16_units(replacement_text)
            if any(
                content_range["end"] - content_range["start"]
                != replacement_length
                for content_range in result["ranges"]
            ):
                return "query replacement ranges must span the replacement text"
        error = _revision_change_error(result, must_change=must_change)
        if error is not None:
            return error
        if not _result_ranges_match_revision(result, result["revisionAfter"]):
            return "replacement ranges must use revisionAfter"
        if not _ranges_are_ordered(result["ranges"]):
            return "replacement ranges must be ordered and non-overlapping"

    elif action == "insertTable":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        table = result["table"]
        if table["rowCount"] != len(table["data"]):
            return "rowCount must equal the observed data row count"
        if table["columnCount"] != len(table["data"][0]):
            return "columnCount must equal the observed data width"
        observed_data = tuple(tuple(row) for row in table["data"])
        requested_data = tuple(tuple(row) for row in params["data"])
        if observed_data != requested_data:
            return "observed table data must equal the requested matrix"
        if table["headerRow"] != params["headerRow"]:
            return "observed table header state must equal headerRow"
        if table["range"]["revision"] != result["revisionAfter"]:
            return "table range must use revisionAfter"
        anchor_position = _anchor_position(params["anchor"])
        if (
            anchor_position is not None
            and table["range"]["start"] != anchor_position
        ):
            return "table range must begin at the resolved Body Anchor"

    elif action == "insertImage":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        image_range = result["image"].get(
            "range",
            result["image"].get("anchorRange"),
        )
        if image_range["revision"] != result["revisionAfter"]:
            return "image range must use revisionAfter"
        image = result["image"]
        placement = params["placement"]
        if image["kind"] != placement["kind"]:
            return "observed image placement kind must equal the request"
        if image["alternativeText"] != params["alternativeText"]:
            return "observed alternative text must equal the request"
        anchor_position = _anchor_position(params["anchor"])
        if anchor_position is not None and image_range["start"] != anchor_position:
            return "image range must begin at the resolved Body Anchor"
        if placement["kind"] == "floating":
            if image["wrap"] != placement["wrap"]:
                return "observed image wrap must equal the request"
            for axis in ("horizontal", "vertical"):
                if image[axis]["relativeTo"] != placement[axis]["relativeTo"]:
                    return f"observed image {axis} reference must equal the request"
                if not _observed_length_matches(
                    image[axis]["offset"],
                    placement[axis]["offset"],
                ):
                    return f"observed image {axis} offset must equal the request"
        size_request = params["size"]
        observed_size = image["size"]
        if size_request["kind"] == "width" and not _observed_length_matches(
            observed_size["width"],
            size_request["width"],
        ):
            return "observed image width must equal the request"
        if size_request["kind"] == "height" and not _observed_length_matches(
            observed_size["height"],
            size_request["height"],
        ):
            return "observed image height must equal the request"
        if size_request["kind"] == "box":
            width_limit = _length_in_points(size_request["width"])
            height_limit = _length_in_points(size_request["height"])
            if size_request["fit"] == "stretch":
                if (
                    abs(observed_size["width"]["value"] - width_limit) > 0.5
                    or abs(observed_size["height"]["value"] - height_limit)
                    > 0.5
                ):
                    return "stretched image dimensions must equal the requested box"
            elif (
                observed_size["width"]["value"] > width_limit + 0.5
                or observed_size["height"]["value"] > height_limit + 0.5
            ):
                return "contained image dimensions must fit the requested box"

    elif action in {"setHeaderFooter", "setPageLayout"}:
        selected_count = result["selectedSectionCount"]
        selected = params["sections"]
        selected_indexes = _selected_section_indexes(
            selected,
            selected_count,
        )
        if selected["kind"] == "indexes" and selected_count != len(
            selected["indexes"]
        ):
            return "selectedSectionCount must equal the requested index count"
        if action == "setHeaderFooter":
            area_order = {"header": 0, "footer": 1}
            variant_order = {
                "primary": 0,
                "firstPage": 1,
                "evenPages": 2,
            }
            update_pairs = sorted(
                (
                    (update["area"], update["variant"])
                    for update in params["updates"]
                ),
                key=lambda pair: (
                    area_order[pair[0]],
                    variant_order[pair[1]],
                ),
            )
            expected_stories = [
                (section_index, area, variant)
                for section_index in selected_indexes
                for area, variant in update_pairs
            ]
            observed_stories = [
                (
                    story["sectionIndex"],
                    story["area"],
                    story["variant"],
                )
                for story in result["stories"]
            ]
            if observed_stories != expected_stories:
                return (
                    "stories must contain every selected section/update pair "
                    "once in canonical order"
                )
            operations = {
                (update["area"], update["variant"]): update["operation"]
                for update in params["updates"]
            }
            for story in result["stories"]:
                operation = operations[(story["area"], story["variant"])]
                if story["variant"] != "primary" and not story[
                    "variantEnabled"
                ]:
                    return "first-page and even-page updates must enable their variant"
                if operation["kind"] == "replace" and (
                    story["text"] != operation["text"]
                    or story["linkToPrevious"]
                    or not story["exists"]
                ):
                    return "a replaced story must be unlinked and equal the requested text"
                if operation["kind"] == "clear" and (
                    story["text"] != "" or story["linkToPrevious"]
                ):
                    return "a cleared story must be empty and unlinked"
                if operation["kind"] == "linkToPrevious" and (
                    story["sectionIndex"] == 0
                    or not story["linkToPrevious"]
                ):
                    return "linkToPrevious must link only a non-first section"
        else:
            if selected_count != len(result["sections"]):
                return "selectedSectionCount must equal the section result count"
            observed_indexes = [
                section["index"] for section in result["sections"]
            ]
            if observed_indexes != selected_indexes:
                return "result sections must equal the complete selection"
            requested_layout = params["layout"]
            for section in result["sections"]:
                observed_layout = section["layout"]
                if (
                    "orientation" in requested_layout
                    and observed_layout["orientation"]
                    != requested_layout["orientation"]
                ):
                    return "observed orientation must equal the request"
                if "margins" in requested_layout:
                    for side in ("top", "right", "bottom", "left"):
                        if not _observed_length_matches(
                            observed_layout["margins"][side],
                            requested_layout["margins"][side],
                        ):
                            return "observed margins must equal the request"

    elif action == "insertBreak":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        observed = result["break"]
        if observed["type"] != params["type"]:
            return "observed break type must equal the requested type"
        if observed["range"]["revision"] != result["revisionAfter"]:
            return "break range must use revisionAfter"
        anchor_position = _anchor_position(params["anchor"])
        if (
            anchor_position is not None
            and observed["range"]["start"] != anchor_position
        ):
            return "break range must begin at the resolved Body Anchor"
        if observed["type"] == "page":
            if observed["sectionCountBefore"] != observed["sectionCountAfter"]:
                return "page break must preserve section count"
        elif observed["sectionCountAfter"] != observed["sectionCountBefore"] + 1:
            return "section break must add exactly one section"
        elif observed["followingSectionIndex"] >= observed["sectionCountAfter"]:
            return "followingSectionIndex must name a resulting section"

    elif action in {"save", "saveAs", "exportPdf"}:
        if result["revisionBefore"] != result["revisionAfter"]:
            return "persistence must preserve the stable Content Revision"
        if action == "save":
            return None
        if params["overwritePolicy"] == "renameIfExists":
            resolution = result.get("outputResolution")
            if resolution is None or resolution["requestedPath"] != params["outputPath"]:
                return "renameIfExists requires the original requested path in outputResolution"
            attempt = resolution["attempts"] - 1
            if result["artifact"]["path"] != output_candidate(params["outputPath"], attempt):
                return "renamed artifact must be the reported numbered sibling of the requested path"
            if resolution["renamed"] != (attempt > 0):
                return "renamed must agree with the number of attempts"
        elif result["artifact"]["path"] != params["outputPath"]:
            return "artifact path must equal the authorized output locator"
        elif "outputResolution" in result:
            return "outputResolution is only valid with renameIfExists"
        if (
            action == "exportPdf"
            and result["documentStateBefore"] != result["documentStateAfter"]
        ):
            return "PDF export must preserve the locator-free document state"
        if (
            params["overwritePolicy"] in {"failIfExists", "renameIfExists"}
            and result["replacedExisting"]
        ):
            return "non-overwriting policies cannot report replacedExisting"

    return None

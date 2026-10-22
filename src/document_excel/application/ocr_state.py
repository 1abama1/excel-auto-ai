"""Revision-checked OCR transactions and masked, in-memory audit events."""

from __future__ import annotations

import hashlib
import time
from copy import deepcopy
from datetime import UTC, datetime

from document_excel.domain.models import ClientData, DocumentSide, FieldStatus, OcrRun, WorkflowMode
from document_excel.domain.resolution import manual


def image_signature(image):
    """Hash bytes actually used, not a stale import-time digest or filename."""
    try:
        with image.selected_path.open("rb") as stream:
            digest = hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError:
        digest = "unavailable"
    value = (
        f"{image.selected_path}|{digest}|{image.selected_rotation}|{image.role}|{image.processed_confirmed}"
    )
    return hashlib.sha256(value.encode()).hexdigest()


def signatures(case):
    return {image.id: image_signature(image) for image in case.images}


def identity_token(case):
    return hashlib.sha256(
        repr(
            sorted(
                (i.id, image_signature(i.model_copy(update={"role": DocumentSide.UNKNOWN})))
                for i in case.images
            )
        ).encode()
    ).hexdigest()


def event(case, name, run_id="", **metadata):
    # Allowlist excludes values, OCR text, names, paths and exception repr.
    allowed = {
        k: v
        for k, v in metadata.items()
        if k in {"field", "image_id", "count", "reason", "status", "method", "bbox"}
    }
    case.diagnostic_events.append(
        {"event": name, "run_id": run_id, "time": datetime.now(UTC).isoformat(), **allowed}
    )
    case.diagnostic_events = case.diagnostic_events[-1000:]


class OcrState:
    def begin(self, case, read_images=None, fields=(), scope="full"):
        if case.active_run_id and case.active_run_id in case.ocr_runs:
            older = case.ocr_runs[case.active_run_id]
            if older.status == "RUNNING":
                older.status = "SUPERSEDED"
        for image in case.images:
            if image.case_id and image.case_id != case.id:
                raise ValueError("Изображение принадлежит другому клиенту")
            image.case_id = case.id
        run = OcrRun(
            case_id=case.id,
            image_hashes=signatures(case),
            read_images=[
                i.model_copy(deep=True) for i in (read_images if read_images is not None else case.images)
            ],
            fields=sorted(fields),
            scope=scope,
            document_type=case.data.document_type,
            manual_type=case.document_type_manual,
        )
        case.ocr_runs[run.id] = run
        run.read_hashes = {i.id: image_signature(i) for i in run.read_images}
        case.active_run_id = run.id
        case.invalidate()
        event(case, "OCR_STARTED", run.id, count=len(run.read_images))
        return run

    def current(self, case, run_id):
        run = case.ocr_runs.get(run_id)
        return bool(
            run
            and run.status == "RUNNING"
            and case.active_run_id == run_id
            and run.case_id == case.id
            and run.image_hashes == signatures(case)
            and run.manual_type == case.document_type_manual
            and run.read_hashes == {i.id: image_signature(i) for i in run.read_images}
            and (not run.manual_type or run.document_type == case.data.document_type)
        )

    def stop(self, case, run_id, status):
        run = case.ocr_runs.get(run_id)
        if run and run.status == "RUNNING":
            run.status = status
            run.completed_at = datetime.now(UTC)
            event(case, "OCR_COMPLETED", run_id, status=status)

    def commit(self, case, run_id, output):
        # Resolve against a private copy and publish as one UI-thread operation.
        # Parser/validation failure cannot leave half of a new client's data.
        working = case.model_copy(deep=True)
        accepted = self._commit(working, run_id, deepcopy(output))
        for name in type(case).model_fields:
            if name not in {"id", "images"}:
                setattr(case, name, getattr(working, name))
        roles = {i.id: i.role for i in working.images}
        for image in case.images:
            image.role = roles[image.id]
        return accepted

    def _commit(self, case, run_id, output, *, preview=False):
        from document_excel.infrastructure.extraction import apply_candidates
        from document_excel.infrastructure.mrz import mrz_candidates

        if not self.current(case, run_id):
            run = case.ocr_runs.get(run_id)
            if run and run.status == "RUNNING":
                run.status = "STALE"
                run.completed_at = datetime.now(UTC)
            event(case, "STALE_RESULT_IGNORED", run_id, reason="run_or_image_changed")
            return False
        run = case.ocr_runs[run_id]
        image_ids = {image.id for image in run.read_images}
        future = [
            c for e in case.evidence.values() for c in e.candidates if c.source_image_id not in image_ids
        ] + output.candidates
        self.check_identity(case, future)
        if case.identity_status == "review" and not preview:
            case.pending_identity = (run_id, output)
            case.diagnostic_blocks.update(getattr(output, "diagnostic_blocks", {}))
            case.processing_metrics.update(getattr(output, "processing_metrics", {}))
            for metric in case.processing_metrics.values():
                metric["publication_status"] = "IDENTITY_REVIEW"
            run.status = "IDENTITY_REVIEW"
            run.completed_at = datetime.now(UTC)
            event(case, "OCR_COMPLETED", run_id, status=run.status, reason="identity_not_established")
            return True
        case.pending_identity = None
        touched = set(run.fields)
        before = ClientData()
        # Replace a run's contribution, including fields which disappeared.
        # Other photos and all manual overrides remain available.
        for name, evidence in list(case.evidence.items()):
            if touched and name not in touched:
                continue
            had = any(c.source_image_id in image_ids for c in evidence.candidates)
            evidence.candidates = [c for c in evidence.candidates if c.source_image_id not in image_ids]
            if had and not manual(evidence):
                if name in type(case.data).model_fields:
                    setattr(case.data, name, getattr(before, name))
                if not evidence.candidates:
                    del case.evidence[name]
        if run.scope != "field":
            for image_id in image_ids:
                case.mrz_reports.pop(image_id, None)
                case.diagnostic_blocks.pop(image_id, None)
            case.mrz_reports.update(output.mrz_reports)
        # Side assignments are proposals until the caller confirms disagreement.
        for image in case.images:
            side = output.sides.get(image.id)
            if image.id in image_ids and side:
                if image.role == DocumentSide.UNKNOWN:
                    image.role = side
                elif side in {DocumentSide.FRONT, DocumentSide.BACK} and side != image.role:
                    case.side_proposals[image.id] = side
        incoming = []
        for candidate in output.candidates:
            if candidate.source_image_id not in image_ids or touched and candidate.field_name not in touched:
                continue
            candidate = candidate.model_copy(
                update={"run_id": run.id, "image_hash": run.image_hashes[candidate.source_image_id]}
            )
            incoming.append(candidate)
            event(
                case,
                "FIELD_CANDIDATE_REJECTED"
                if candidate.validation_status == FieldStatus.INVALID
                else "FIELD_CANDIDATE_FOUND",
                run.id,
                field=candidate.field_name,
                image_id=candidate.source_image_id,
                method=candidate.extraction_method,
                bbox=candidate.source_bbox,
            )
        # MRZ comparisons must be regenerated against the current visual data,
        # not retain a conflict from an earlier image/run.
        visual = [c for e in case.evidence.values() for c in e.candidates if c.source_method != "MRZ"]
        visual += [c for c in incoming if c.source_method != "MRZ"]
        for evidence in case.evidence.values():
            evidence.candidates = [c for c in evidence.candidates if c.source_method != "MRZ"]
        incoming = [c for c in incoming if c.source_method != "MRZ"]
        for report in case.mrz_reports.values():
            report.conflicts.clear()
            incoming.extend(mrz_candidates(report, visual))
            event(
                case,
                "MRZ_VALIDATED",
                run.id,
                image_id=report.image_id,
                status="VALID" if report.valid else "INVALID",
            )
        # Re-resolve unchanged sources too, so removal cannot retain stale FIO.
        all_candidates = [c for e in case.evidence.values() for c in e.candidates] + incoming
        run.status = "PREVIEW" if preview else "COMPLETED"
        run.completed_at = datetime.now(UTC)
        case.field_regions.update(getattr(output, "field_regions", {}))
        resolution_started = time.perf_counter()
        apply_candidates(case, all_candidates, output.document_type, output.label_seen)
        resolution_seconds = time.perf_counter() - resolution_started
        case.diagnostic_blocks.update(getattr(output, "diagnostic_blocks", {}))
        case.processing_metrics.update(getattr(output, "processing_metrics", {}))
        from document_excel.infrastructure.ocr_diagnostics import safe_resolution_summary

        for image_id, metric in case.processing_metrics.items():
            metric.update(safe_resolution_summary(case, image_id))
            metric["case_resolution_seconds"] = resolution_seconds
            metric["publication_status"] = "PREVIEW" if preview else "COMMITTED"
        case.field_regions.update(getattr(output, "field_regions", {}))
        self.check_identity(case, all_candidates)
        for image_id, blocks in getattr(output, "diagnostic_blocks", {}).items():
            event(case, "IMAGE_PROCESSED", run.id, image_id=image_id)
            event(case, "DOCUMENT_CLASSIFIED", run.id, image_id=image_id, status=output.document_type)
            event(case, "TEXT_BLOCKS_EXTRACTED", run.id, image_id=image_id, count=len(blocks))
        for name, evidence in case.evidence.items():
            if evidence.status == FieldStatus.CONFLICT:
                event(case, "FIELD_CONFLICT_DETECTED", run.id, field=name)
        event(case, "OCR_COMPLETED", run.id, status=run.status)
        return True

    @staticmethod
    def check_identity(case, candidates):
        primary = [i for i in case.images if i.role != DocumentSide.ADDITIONAL]
        if len(primary) < 2:
            case.identity_status = "not_checked"
            return
        identity: dict[str, set[tuple[str, str]]] = {i.id: set() for i in primary}
        for c in candidates:
            if (
                c.source_image_id in identity
                and c.field_name in {"passport_number", "personal_identifier"}
                and c.validation_status != FieldStatus.INVALID
            ):
                identity[c.source_image_id].add((c.field_name, c.normalized_value))
        connected = {primary[0].id}
        while True:
            add = {i for i in identity if any(identity[i] & identity[j] for j in connected)}
            new = connected | add
            if new == connected:
                break
            connected = new
        conflicting = False
        if case.workflow_mode == WorkflowMode.BEST_EFFORT_UPDATE:
            # No shared identifier is ordinary for opposite sides. Contradictory
            # identifiers or names still quarantine results from different people.
            for field in (
                "passport_number",
                "personal_identifier",
                "full_name",
                "surname",
                "first_name",
                "birth_date",
            ):
                by_image: dict[str, set[str]] = {}
                for candidate in candidates:
                    if (
                        candidate.source_image_id in identity
                        and candidate.field_name == field
                        and candidate.validation_status != FieldStatus.INVALID
                        and not candidate.validation_errors
                    ):
                        by_image.setdefault(candidate.source_image_id, set()).add(
                            candidate.normalized_value.casefold()
                        )
                values = list(by_image.values())
                if any(
                    not left & right for index, left in enumerate(values) for right in values[index + 1 :]
                ):
                    conflicting = True
                    break
        token = identity_token(case)
        if case.identity_approved == token:
            case.identity_status = "user_confirmed"
        elif conflicting:
            case.identity_status = "review"
        elif len(connected) == len(primary):
            case.identity_status = "matched"
        elif case.workflow_mode == WorkflowMode.BEST_EFFORT_UPDATE:
            case.identity_status = "unverified"
        else:
            case.identity_status = "review"
        case.identity_message = (
            "Проверьте, что все фотографии принадлежат одному человеку."
            if case.identity_status == "review"
            else "Стороны не удалось сопоставить по общему номеру; проверьте их в итоговой сводке."
            if case.identity_status == "unverified"
            else "Принадлежность сторон сопоставлена; сверьте документ."
        )

    @staticmethod
    def approve_identity(case):
        case.identity_approved = identity_token(case)
        case.identity_status = "user_confirmed"
        case.invalidate()

    def accept_identity(self, case):
        pending = case.pending_identity
        if not pending:
            self.approve_identity(case)
            return True
        run_id, output = pending
        run = case.ocr_runs.get(run_id)
        if not self.pending_current(case):
            case.pending_identity = None
            event(case, "STALE_RESULT_IGNORED", run_id, reason="identity_images_changed")
            return False
        self.approve_identity(case)
        run.status = "RUNNING"
        return self.commit(case, run_id, output)

    @staticmethod
    def pending_current(case):
        if not case.pending_identity:
            return False
        run_id, _ = case.pending_identity
        run = case.ocr_runs.get(run_id)
        return bool(
            run
            and run.status == "IDENTITY_REVIEW"
            and case.active_run_id == run_id
            and run.image_hashes == signatures(case)
            and run.read_hashes == {i.id: image_signature(i) for i in run.read_images}
            and run.manual_type == case.document_type_manual
            and (not run.manual_type or run.document_type == case.data.document_type)
        )

    def preview(self, case):
        """Read-only staging: no approval, mutation or export of the live case."""
        if not self.pending_current(case):
            return None
        working = case.model_copy(deep=True)
        run_id, output = working.pending_identity
        working.ocr_runs[run_id].status = "RUNNING"
        if not self._commit(working, run_id, deepcopy(output), preview=True):
            return None
        working.confirmed_digest = None
        return working

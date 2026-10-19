from document_excel.application.benchmark import evaluate
from document_excel.domain.models import ClientCase, FieldEvidence, FieldStatus


def test_benchmark_counts_plausible_wrong_values_without_hiding_conflicts():
    case = ClientCase()
    case.data.personal_identifier = "00123456789012"
    case.data.passport_type, case.data.passport_number = "ID", "0000001"
    result = evaluate(case, {"personal_identifier": "00123456789013", "document_number": "ID0000001"})
    assert result["fields"] == {"personal_identifier": False, "document_number": True}
    assert not result["fully_correct"] and result["dangerous_errors"] == 1
    case.evidence["personal_identifier"] = FieldEvidence(status=FieldStatus.CONFLICT)
    result = evaluate(case, {"personal_identifier": "00123456789012"})
    assert not result["fully_correct"] and result["dangerous_errors"] == 0

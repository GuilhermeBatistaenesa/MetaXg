from outcomes import (
    OUTCOME_VERIFIED_SUCCESS,
    OUTCOME_SAVED_NOT_VERIFIED,
    OUTCOME_FAILED_ACTION,
    OUTCOME_FAILED_VERIFICATION,
    OUTCOME_SKIPPED_ALREADY_EXISTS,
    compute_totals,
)


def test_compute_totals_by_outcome():
    people = [
        {"outcome": OUTCOME_VERIFIED_SUCCESS, "no_photo": False},
        {"outcome": OUTCOME_SAVED_NOT_VERIFIED, "no_photo": True},
        {"outcome": OUTCOME_FAILED_ACTION, "no_photo": False},
        {"outcome": OUTCOME_FAILED_VERIFICATION, "no_photo": False},
        {"outcome": OUTCOME_SKIPPED_ALREADY_EXISTS, "no_photo": False},
    ]

    totals = compute_totals(people, detected=7)
    assert totals["detected"] == 7
    assert totals["people_total"] == 5
    assert totals["no_photo"] == 1
    assert totals["by_outcome"][OUTCOME_VERIFIED_SUCCESS] == 1
    assert totals["by_outcome"][OUTCOME_SAVED_NOT_VERIFIED] == 1
    assert totals["by_outcome"][OUTCOME_FAILED_ACTION] == 1
    assert totals["by_outcome"][OUTCOME_FAILED_VERIFICATION] == 1
    assert totals["by_outcome"][OUTCOME_SKIPPED_ALREADY_EXISTS] == 1

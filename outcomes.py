OUTCOME_VERIFIED_SUCCESS = "VERIFIED_SUCCESS"
OUTCOME_SAVED_NOT_VERIFIED = "SAVED_NOT_VERIFIED"
OUTCOME_FAILED_ACTION = "FAILED_ACTION"
OUTCOME_FAILED_VERIFICATION = "FAILED_VERIFICATION"
OUTCOME_SKIPPED_ALREADY_EXISTS = "SKIPPED_ALREADY_EXISTS"
OUTCOME_SKIPPED_DRY_RUN = "SKIPPED_DRY_RUN"
OUTCOME_SKIPPED_NO_RECIPIENT = "SKIPPED_NO_RECIPIENT"
OUTCOME_SKIPPED_EMAIL_DISABLED = "SKIPPED_EMAIL_DISABLED"

OUTCOME_ORDER = [
    OUTCOME_FAILED_ACTION,
    OUTCOME_FAILED_VERIFICATION,
    OUTCOME_SAVED_NOT_VERIFIED,
    OUTCOME_VERIFIED_SUCCESS,
    OUTCOME_SKIPPED_ALREADY_EXISTS,
    OUTCOME_SKIPPED_DRY_RUN,
    OUTCOME_SKIPPED_NO_RECIPIENT,
    OUTCOME_SKIPPED_EMAIL_DISABLED,
]


def compute_totals(people: list[dict], detected: int | None = None) -> dict:
    totals_by_outcome = {k: 0 for k in OUTCOME_ORDER}
    unknown_outcome = 0
    no_photo = 0

    for person in people:
        outcome = person.get("outcome")
        if outcome in totals_by_outcome:
            totals_by_outcome[outcome] += 1
        else:
            unknown_outcome += 1
        if person.get("no_photo"):
            no_photo += 1

    return {
        "detected": detected if detected is not None else len(people),
        "people_total": len(people),
        "by_outcome": totals_by_outcome,
        "no_photo": no_photo,
        "unknown_outcome": unknown_outcome,
    }

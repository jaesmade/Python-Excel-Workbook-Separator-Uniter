from collections import OrderedDict

import pytest

from excel_splitter.spelling_review import KEEP_SEPARATE, SpellingReview


def make_review():
    groups = OrderedDict([
        (("JOB ORDER",), [(1,), (2,), (3,)]),
        (("JO ORDER",), [(4,)]),
        (("Job Order",), [(5,), (6,)]),
        (("PAYROLL",), [(7,), (8,)]),
        (("PAYR0LL",), [(9,)]),
    ])
    return SpellingReview(groups, ["Type"], {}, {("JOB ORDER",), ("Job Order",)})


def test_review_lists_suggestions_with_record_counts_and_undoes_decisions():
    review = make_review()
    assert review.suggestions == [
        ("Type", "Job Order", "JOB ORDER", 2, 3),
        ("Type", "JO ORDER", "JOB ORDER", 1, 3),
    ]
    review.choose(0, "JOB ORDER")
    assert len(review.groups[("JOB ORDER",)]) == 5
    assert review.selected == {("JOB ORDER",)}
    assert review.mappings == {("Type", "Job Order"): "JOB ORDER"}
    review.choose(1, KEEP_SEPARATE)
    assert review.decisions[1] is KEEP_SEPARATE
    assert review.undo()
    assert 1 not in review.decisions
    assert review.undo()
    assert review.selected == {("JOB ORDER",), ("Job Order",)}
    assert review.mappings == {}
    assert len(review.groups[("JOB ORDER",)]) == 3
    assert not review.can_undo


def test_overlapping_merges_can_choose_the_other_spelling_without_a_cycle():
    review = make_review()
    review.choose(0, "JOB ORDER")
    review.choose(1, "JO ORDER")
    assert len(review.groups[("JO ORDER",)]) == 6
    assert review.mappings == {
        ("Type", "Job Order"): "JOB ORDER",
        ("Type", "JOB ORDER"): "JO ORDER",
    }
    assert review.selected == {("JO ORDER",)}
    review.undo()
    assert len(review.groups[("JOB ORDER",)]) == 5
    assert ("JO ORDER",) in review.groups


def test_invalid_or_repeated_choice_does_not_change_review():
    review = make_review()
    with pytest.raises(ValueError):
        review.choose(0, "unrelated")
    review.choose(0, KEEP_SEPARATE)
    with pytest.raises(ValueError):
        review.choose(0, "JOB ORDER")
    assert review.decisions == {0: KEEP_SEPARATE}


def test_cannot_choose_a_spelling_already_replaced_by_an_earlier_merge():
    review = make_review()
    review.choose(0, "Job Order")
    with pytest.raises(ValueError, match="already replaced"):
        review.choose(1, "JOB ORDER")
    assert review.decisions == {0: "Job Order"}
    assert len(review.groups[("Job Order",)]) == 5


def test_copy_preserves_undo_history_without_mutating_original_session():
    original = make_review()
    original.choose(0, "JOB ORDER")
    original.choose(1, KEEP_SEPARATE)
    copy = original.copy()

    assert copy.decisions[1] is KEEP_SEPARATE
    assert copy.undo()
    copy.choose(1, "JOB ORDER")
    assert len(copy.groups[("JOB ORDER",)]) == 6
    assert len(original.groups[("JOB ORDER",)]) == 5
    assert original.decisions[1] is KEEP_SEPARATE
    assert original.undo()
    assert original.undo()
    assert original.mappings == {}
    assert original.selected == {("JOB ORDER",), ("Job Order",)}
    assert len(copy.groups[("JOB ORDER",)]) == 6

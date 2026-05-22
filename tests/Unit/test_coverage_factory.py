import pytest
from vba_unit.Coverage.coverage_factory import CovFact
from vba_unit.Coverage.coveralls import Coveralls


def test_get_coveralls() -> None:
    co = CovFact.provider("coveralls")
    assert isinstance(co, Coveralls)


def test_get_except() -> None:
    with pytest.raises(Exception) as e:
        CovFact.provider("foo")
    assert str(e.value) == "Unknown coverage provider."

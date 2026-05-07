import pytest
from vba_unit.Coverage.coverage import Coverage


def test_constructor() -> None:
    obj = Coverage()
    with pytest.raises(Exception) as e:
        obj.generate_report()
    assert str(e.value) == "Must be implemented by an extending class"

import pytest
from vba_unit.Coverage.coverage import Coverage


def test_constructor() -> None:
    obj = Coverage()
    with pytest.raises(Exception):
        obj.generate_report()

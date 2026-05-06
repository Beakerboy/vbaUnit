from typing import TypeVar
from vba_unit.Coverage.coverage import Coverage


T = TypeVar('T', bound='TextReport')


class TextReport(Coverage):
    """
    Coverage metrics for text output
    """
    pass

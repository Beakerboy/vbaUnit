from typing import TypeVar
from vba_unit.Coverage.coverage import Coverage


T = TypeVar('T', bound='HtmlReport')


class HtmlReport(Coverage):
    """
    Static webpages of code coverage
    """
    pass

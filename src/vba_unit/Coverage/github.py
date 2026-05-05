from typing import TypeVar
from vba_unit.Coverage.git_repo import GitRepo


T = Typevar('T', bound='Github')


class Coveralls(GitRepo):
    pass

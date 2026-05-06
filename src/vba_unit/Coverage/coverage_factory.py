from .coverage import Coverage
from .git_repo import GitRepo

class CovFact():
    @staticmethod
    def provider(name: str) -> Coverage:
        if name == "coveralls":
            return Coveralls(GitRepo())

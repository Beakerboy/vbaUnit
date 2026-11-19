from .github import Github
from .git_repo import GitRepo


class GitFact():
    @staticmethod
    def provider(name: str) -> GitRepo:
        if name == "github":
            return Github()
        else:
            raise Exception("Unknown git provider.")

from .github import GitHub

class GitFact():
    @staticmethod
    def provider(name: str) -> Coverage:
        if name == "github":
            return Github()

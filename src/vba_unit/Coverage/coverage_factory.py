from .coverage import Coverage


class CovFact():
    @staticmethod
    def provider(name: str) -> Coverage:
        if name == "coveralls":
            return Coveralls()

from .coverage import Coverage
from .coveralls import Coveralls


class CovFact():
    @staticmethod
    def provider(name: str) -> Coverage:
        if name == "coveralls":
            return Coveralls()
        else:
            raise Exception("Unknown coverage provider.")

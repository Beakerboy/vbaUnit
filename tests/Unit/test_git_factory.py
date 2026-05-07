import pytest
from vba_unit.Coverage.git_factory import GitFact
from vba_unit.Coverage.github import Github


def test_get_github() -> None:
    gh = GitFact.provider("github")
    assert isinstance(gh, Github)


def test_get_except() -> None:
    with pytest.raises(Exception) as e:
        GitFact.provider("foo")
    assert str(e) == "Unknown git provider."

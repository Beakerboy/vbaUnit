from vba_unit.Coverage.github import GitRepo


def test_constructor() -> None:
    gh = GitRepo()
    assert isinstance(gh, GitRepo)

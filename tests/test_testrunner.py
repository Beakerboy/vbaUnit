import os
from vba_unit.vba_test_runner import main
from pytest_mock import MockerFixture


def test_main(mocker: MockerFixture) -> None:
    mock_print = mocker.patch("builtins.print")
    mocker.patch(
        "sys.argv",
        [
            "vba_test_runner.py",
            "--project",
            "VbaProject"
        ],
    )
    os.chdir("./tests")
    main()
    mock_print.assert_called_with("-----------------------\nSummary: 2/2 passed.")

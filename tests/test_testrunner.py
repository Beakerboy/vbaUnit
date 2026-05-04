import os
import pytest
from vba_unit.vba_test_runner import main
from pytest_mock import MockerFixture

@patch('builtins.print')
def test_main(mocker: MockerFixture, mock_print: str) -> None:
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
    mock_print.assert_called_with("Test")

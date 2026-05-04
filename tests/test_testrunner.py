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
    assert mock_print.call_count == 4
    assert mock_print.call_args_list[0].args[0] == "test"
    assert mock_print.call_args_list[1].args[0] == "test"
    assert mock_print.call_args_list[2].args[0] == (
        "-----------------------\nSummary: 2/2 passed."
    )

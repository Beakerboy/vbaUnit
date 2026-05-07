import os
import pytest
from vba_unit.cli import main
from pytest_mock import MockerFixture


@pytest.fixture
def change_dir():
    original_dir = os.getcwd()
    os.chdir("./tests")
    yield  # The test runs here
    os.chdir(original_dir)  # Teardown: happens after test ends


def test_main(change_dir, mocker: MockerFixture) -> None:
    mock_print = mocker.patch("builtins.print")
    mocker.patch(
        "sys.argv",
        [
            "vba_test_runner.py"
        ],
    )
    os.chdir("./tests")
    main()

    messages = [
        "\n--- VBA Test Report ---",
        "test_boolean.test_true: PASS",
        "test_boolean.test_false: FAIL: ",
        "test_boolean.test_and: PASS",
        "-----------------------\nSummary: 2/3 passed."
    ]
    assert mock_print.call_count == len(messages)
    i = 0
    for message in messages:
        assert mock_print.call_args_list[i].args[0] == message
        i += 1

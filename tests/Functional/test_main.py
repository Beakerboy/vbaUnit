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
    
    messages = [
        "\n--- VBA Test Report ---",
        "test_boolean.test_true: PASS",
        "test_boolean.test_false: FAIL:",
        "test_boolean.test_and: PASS",
        "-----------------------\nSummary: 2/3 passed."
    ]
    assert mock_print.call_count == len(messages)
    i = 0
    for message in messages:
        assert mock_print.call_args_list[i].args[0] == message
        i += 1

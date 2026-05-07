import os
import pytest
from pytest_mock import MockerFixture
from unittest import mock
from vba_unit.cli import main


@pytest.fixture
def change_dir() -> None:
    original_dir = os.getcwd()
    os.chdir("./tests")
    yield  # The test runs here
    os.chdir(original_dir)  # Teardown: happens after test ends


@mock.patch.dict(os.environ, {
    "COVERALLS_REPO_TOKEN": "secretsecretsecret",
    "GITHUB_RUN_ID": "25396149145",
    "GITHUB_SHA": "036c36dfac1d00cb37b6510fc423641cda7b1f08",
    "GITHUB_REF": "refs/pull/6/merge"
})
def test_main(change_dir: str, mocker: MockerFixture) -> None:
    mock_post = mocker.patch('requests.post')
    mock_check_output = mocker.patch('subprocess.check_output')
    mock_check_output.side_effect = [
        "John Doe\nme@me.com\nGitHub\nnoreply@github.com\ncommit message",
        'https://github.com/Beakerboy/FooProject'
    ]
    mock_response = mock.MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {
        "message": "25504858355.1",
        "url": "https://coveralls.io/builds/79326270"
    }
    mock_post.return_value = mock_response
    mock_print = mocker.patch("builtins.print")
    mocker.patch(
        "sys.argv",
        [
            "vba_test_runner.py",
            "--coverage"
        ],
    )
    main()

    messages = [
        "\n--- VBA Test Report ---",
        "test_boolean.test_true: PASS",
        "test_boolean.test_false: FAIL: ",
        "test_boolean.test_and: PASS",
        "-----------------------\nSummary: 2/3 passed.",
        "Submitting coverage to coveralls.io..."
    ]
    assert mock_print.call_count == len(messages)
    i = 0
    for message in messages:
        assert mock_print.call_args_list[i].args[0] == message
        i += 1

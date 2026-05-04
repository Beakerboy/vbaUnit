import os
from vba_unit.vba_test_runner import run_tests
from pytest_mock import MockerFixture


def test_main(mocker: MockerFixture) -> None:
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

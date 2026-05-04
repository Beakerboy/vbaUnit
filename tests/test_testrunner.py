from vba_unit.vba_test_runner import run_tests

def test_main(mocker: MockerFixture) -> None:
    mocker.patch(
        "sys.argv",
        [
            "vba_test_runner.py",
            "./project",
        ],
    )
    run_tests()

import pytest
from typing import TypeVar
from unittest import mock
from vba_unit.Coverage.coverage import Coverage


T = TypeVar('T', bound='MockCoverage')


def test_constructor() -> None:
    obj = Coverage()
    with pytest.raises(Exception) as e:
        obj.generate_report()
    assert str(e.value) == "Must be implemented by an extending class"


class MockCoverage(Coverage):
    def generate_report(self: T) -> dict:
        return {
            'repo_token': 'secretsecretsecret',
            'service_name': 'manual',
            'service_job_id': '25396149145',
        }


@mock.patch('requests.post')
def test_submit(mock_post) -> None:
    mock_response = mock.MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {
        "message": "25504858355.1,
        "url": "https://coveralls.io/builds/79326270"
    }
    mock_post.return_value = mock_response

    expected_report = {
        'repo_token': 'secretsecretsecret',
        'service_name': 'manual',
        'service_job_id': '25396149145',
    }
    url = "https://www.example.com/api/v1"
    cov = MockCoverage()
    cov.endpoint = url
    result = cov.submit_report()
    assert result = "{}"
    mock_post.assert_called_once_with(url, json=expected_report) 

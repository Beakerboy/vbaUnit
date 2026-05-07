import os
from typing import TypeVar
from unittest import mock
from vba_unit.Coverage.coveralls import Coveralls



T = TypeVar('T', bound='MockTable')


mock_github = mock.Mock()
mock_github.repo.return_value = (
    {
        'head': {
            'id': '036c36dfac1d00cb37b6510fc423641cda7b1f08',
            'author_name': 'John Doe',
            'author_email': 'me@me.com',
            'committer_name': 'GitHub',
            'committer_email': 'noreply@github.com',
            'message': 'commit message'
        },
        'branch': 'PR-6/merge',
        'remotes': [
            {
                'name': 'origin',
                'url': 'https://github.com/Beakerboy/FooProject'
            }
        ]
    }
)
mock_github.job_id.return_value = '25396149145'


class MockTable():
    def __init__(self: T) -> None:
        self.definitions = {
            "vbaproject": {
                "modules": {
                    "roots": {
                        "name": "roots",
                        "cover": True,
                        "coverage": [1, None, None, None, None, None, None,
                                     None, None, None, None, 1, 1, 1],
                        "path": 'src/Modules/Roots.bas'
                    }
                }
            }
        }


def test_constructor() -> None:
    obj = Coveralls()
    assert obj.endpoint == "https://coveralls.io/api/v1/jobs"


@mock.patch.dict(os.environ, {
    "COVERALLS_REPO_TOKEN": "secretsecretsecret",
})
def test_report() -> None:
    fake_content = (
        'Attribute VB_Name = "Roots"\r\n'
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        "'Comment\r\n"
        'Function Determinant(a, b, c)\r\n'
        '    Determinant = b ^ 2 - 4 * a * c'
        'End Function\r\n'
    )
    m = mock.mock_open(read_data=fake_content)
    obj = Coveralls()
    obj.git = mock_github
    obj.table = MockTable()
    expected = {
        'repo_token': 'secretsecretsecret',
        'service_name': 'manual',
        'service_job_id': '25396149145',
        'source_files': [{
            'name': 'src/Modules/Roots.bas',
            'source_digest': '72a03368f06a7905c304c52068d77755',
            'coverage': [1, None, None, None, None, None, None,
                         None, None, None, None, 1, 1, 1]
        }],
        'git': {
            'head': {
                'id': '036c36dfac1d00cb37b6510fc423641cda7b1f08',
                'author_name': 'John Doe',
                'author_email': 'me@me.com',
                'committer_name': 'GitHub',
                'committer_email': 'noreply@github.com',
                'message': 'commit message'
            },
            'branch': 'PR-6/merge',
            'remotes': [
                {
                    'name': 'origin',
                    'url': 'https://github.com/Beakerboy/FooProject'
                }
            ]
        }
    }
    received = obj.generate_report()
    assert received == expected

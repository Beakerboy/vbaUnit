from unittest import mock
from vba_unit.Coverage.coveralls import Coveralls



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

def test_constructor() -> None:
    obj = Coveralls()
    assert obj.endpoint == "https://coveralls.io/api/v1/jobs"


def test_report() -> None:
    obj = Coveralls()
    obj.git = mock_github
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
    received = coveralls.generate_repoet()
    assert received == expected

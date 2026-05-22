from typing import TypeVar


T = TypeVar('T', bound='GitRepo')


class GitRepo():
    def __init__(self: T) -> None:
        self.job_id = ''
        self.git_repo = ''
        self.author_name = ''
        self.author_email = ''
        self.committer_name = ''
        self.committer_email = ''
        self.message = ''
        self.branch = ''
        self.remote_url = ''
        self.commit_sha = ''

    def repo(self: T) -> dict:
        return {
            "head": {
                "id": self.commit_sha,
                "author_name": self.author_name,
                "author_email": self.author_email,
                "committer_name": self.committer_name,
                "committer_email": self.committer_email,
                "message": self.message
            },
            "branch": self.branch,
            "remotes": [
                {
                    "name": "origin",
                    "url": self.remote_url
                }
            ]
        }

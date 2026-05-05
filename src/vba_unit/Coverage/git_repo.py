from typing import TypeVar


T = TypeVar('T', bound='GitRep')


class GitRepo():
    job_id: str
    
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

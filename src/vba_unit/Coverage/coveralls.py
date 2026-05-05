from typing import TypeVar


T = Typevar('T', bound='Coveralls')


class Coveralls(Coverage):
    def __init__(self: T) -> None:
        self.endpoint = "https://coveralls.io/api/v1/jobs"

    def coveralls_report() -> str:
    # with open(file_path, 'r') as f:
    #    line_count = sum(1 for line in f)
    # coverage = [None] * line_count
    # for i in range(line_count):
    #    line_num = i + 1
    #    if line_num in visited_lines:
    #        coverage[i] = 1
    #    else:
    #        coverage[i] = 0
    source_code = """Attribute VB_Name = "Roots"
' Function: Discriminant
' A function to determine if the roots of a quadratic are real of complex
'
' Parameters:
'    a - x² coefficiant
'    b - x  coefficient
'    c - constant term
'
' Returns:
' a real number
Public Function Discriminant(a, b, c)
    Discriminant = b ^ 2 - (4 * a * c)
End Function
"""
    digest = hashlib.md5(source_code.encode('utf-8')).hexdigest()
    commit_sha = os.environ.get('GITHUB_SHA')
    assert commit_sha is not None
    full_ref = os.environ.get('GITHUB_REF', 'master')
    branch = full_ref.replace('refs/heads/', '').replace('refs/pull/', 'PR-')
    fmt = "%an%n%ae%n%cn%n%ce%n%s"
    details = subprocess.check_output(
        ["git", "log", "-1", f"--pretty=format:{fmt}", commit_sha],
        text=True
    ).splitlines()

    # 3. Remote URL from git config
    remote_url = subprocess.check_output(
        ["git", "config", "--get", "remote.origin.url"],
        text=True
    ).strip()
    report = {
        "repo_token": os.environ['COVERALLS_REPO_TOKEN'],
        "service_name": "manual",
        "service_job_id": os.environ['GITHUB_RUN_ID'],
        "source_files": [
            {
                "name": "src/Modules/Roots.bas",
                "source_digest": digest,
                "source": source_code,
                "coverage": [1, None, None, None, None, None, None, None, None,
                             None, None, 1, 1, 1],
            }
        ],
        "git": {
            "head": {
                "id": commit_sha,
                "author_name": details[0],
                "author_email": details[1],
                "committer_name": details[2],
                "committer_email": details[3],
                "message": details[4]
            },
            "branch": branch,
            "remotes": [
                {
                    "name": "origin",
                    "url": remote_url
                }
            ]
        }
    }
    print(report)

"""CI for GCV add-in repositories (standard library only).

python gcv_ci.py check [--root .] [--tag logger-v1.0.1]
    Source checks: VBA files are ASCII with CRLF, VB_Name equals the file
    name, Option Explicit; the ribbon is valid and uses the shared tab;
    GcvAddin.bas holds a valid identity whose version has a CHANGELOG
    section; catalog.json (when present) is valid. With --tag: the tag is
    <package>-v<version of GcvAddin.bas>.

python gcv_ci.py release --tag logger-v1.0.1 [--root .] [--min-loader 2.0.0]
    GitHub only (GITHUB_TOKEN, GITHUB_REPOSITORY). Finds the DRAFT release of
    the tag, checks the uploaded .xlam with make_addin_release.py (same
    version as the tag, add-in type, ribbon, no test modules), uploads
    manifest.json and the add-in under its asset name, sets the notes from
    CHANGELOG.md and publishes. On any problem the release stays a draft
    (the updater ignores drafts) and the command fails.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol
from xml.etree import ElementTree

import make_addin_release as mar

VB_NAME_RE = re.compile(r'^Attribute VB_Name = "([^"]+)"', re.MULTILINE)
CONST_RE = r'^Public Const {name} As String = "([^"]*)"'
ID_RE = re.compile(r"^[a-z0-9-]+/[a-z0-9-]+$")
VERSION_RE = re.compile(r"^\d{1,6}\.\d{1,6}\.\d{1,6}$")
REPO_RE = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")
ADDIN_FILE_RE = re.compile(r"^GCV [A-Za-z0-9 -]+\.xlam$")
LOADER_FILE = "GCV Loader.xlam"
SHARED_TAB = "gcv:GcvToolsTab"
CUSTOMUI_NS = "{http://schemas.microsoft.com/office/2009/07/customui}"
VBA_SUFFIXES = (".bas", ".cls")
API_ROOT = "https://api.github.com"
UPLOAD_ROOT = "https://uploads.github.com"
DEFAULT_MIN_LOADER = "2.0.0"


class CiError(Exception):
    """The repository or the release cannot pass as it is."""


@dataclass(frozen=True, slots=True)
class AddinModule:
    addin_id: str
    file: str
    version: str
    repo: str

    @property
    def tag(self) -> str:
        return f"{self.addin_id.split('/', 1)[1]}-v{self.version}"


# ---------------------------------------------------------------------------
# Source checks
# ---------------------------------------------------------------------------


def _vba_files(root: Path) -> list[Path]:
    files = sorted(p for p in (root / "src").glob("*") if p.suffix in VBA_SUFFIXES)
    files += sorted(p for p in (root / "loader").glob("*") if p.suffix in VBA_SUFFIXES)
    return files


def _check_text_file(path: Path, root: Path) -> list[str]:
    name = path.relative_to(root).as_posix()
    data = path.read_bytes()
    try:
        text = data.decode("ascii")
    except UnicodeDecodeError:
        return [f"{name}: must be ASCII (write Arabic as AR(...) code points)"]
    if "\n" in text.replace("\r\n", ""):
        return [f"{name}: line endings must be CRLF"]
    return []


def _check_vba_file(path: Path, root: Path) -> list[str]:
    problems = _check_text_file(path, root)
    if problems:
        return problems
    name = path.relative_to(root).as_posix()
    text = path.read_text(encoding="ascii")
    match = VB_NAME_RE.search(text)
    if match is None or match.group(1) != path.stem:
        problems.append(f'{name}: Attribute VB_Name must be "{path.stem}"')
    if "Option Explicit" not in text:
        problems.append(f"{name}: Option Explicit is missing")
    return problems


def _check_ribbon(root: Path) -> list[str]:
    path = root / "customUI14.xml"
    if not path.is_file():
        return []
    try:
        tree = ElementTree.fromstring(path.read_bytes())  # noqa: S314 - file from this repository
    except ElementTree.ParseError as exc:
        return [f"customUI14.xml: invalid XML ({exc})"]
    tabs = [tab.get("idQ") for tab in tree.iter(f"{CUSTOMUI_NS}tab")]
    if SHARED_TAB not in tabs:
        return [f'customUI14.xml: the tab must be idQ="{SHARED_TAB}" (the shared GCV tab)']
    return []


def read_addin_module(root: Path) -> AddinModule:
    path = root / "src" / "GcvAddin.bas"
    if not path.is_file():
        raise CiError("src/GcvAddin.bas is missing (identity of the add-in)")
    text = path.read_text(encoding="ascii", errors="replace")
    values: dict[str, str] = {}
    for const in ("ADDIN_ID", "ADDIN_FILE", "ADDIN_VERSION", "ADDIN_REPO"):
        match = re.search(CONST_RE.format(name=const), text, re.MULTILINE)
        if match is None:
            raise CiError(f"src/GcvAddin.bas: {const} is missing")
        values[const] = match.group(1)
    return AddinModule(values["ADDIN_ID"], values["ADDIN_FILE"], values["ADDIN_VERSION"], values["ADDIN_REPO"])


def _check_identity(root: Path) -> list[str]:
    try:
        module = read_addin_module(root)
    except CiError as exc:
        return [str(exc)]
    problems: list[str] = []
    if not ID_RE.match(module.addin_id):
        problems.append(f"GcvAddin.bas: ADDIN_ID '{module.addin_id}' must be vendor/package")
    if not ADDIN_FILE_RE.match(module.file):
        problems.append(f"GcvAddin.bas: ADDIN_FILE '{module.file}' must be 'GCV <name>.xlam'")
    if not REPO_RE.match(module.repo):
        problems.append(f"GcvAddin.bas: ADDIN_REPO '{module.repo}' must be owner/repository")
    if not VERSION_RE.match(module.version):
        problems.append(f"GcvAddin.bas: ADDIN_VERSION '{module.version}' must be MAJOR.MINOR.PATCH")
        return problems
    try:
        mar.changelog_section(root / "CHANGELOG.md", module.version)
    except mar.ReleaseError as exc:
        problems.append(f"CHANGELOG.md: {exc} (version {module.version})")
    return problems


def check_sources(root: Path) -> list[str]:
    problems: list[str] = []
    for path in _vba_files(root):
        problems += _check_vba_file(path, root)
    for path in sorted([*root.glob("*.code"), *(root / "loader").glob("*.code")]):
        problems += _check_text_file(path, root)
    problems += _check_ribbon(root)
    problems += _check_identity(root)
    return problems


def check_tag(tag: str, root: Path) -> list[str]:
    try:
        module = read_addin_module(root)
    except CiError as exc:
        return [str(exc)]
    if tag != module.tag:
        return [f"tag '{tag}' does not match GcvAddin.bas: expected '{module.tag}'"]
    return []


# ---------------------------------------------------------------------------
# Catalog (excel-updater)
# ---------------------------------------------------------------------------


def check_catalog(path: Path) -> list[str]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        return [f"{path.name}: invalid JSON ({exc})"]
    tools = data.get("tools") if isinstance(data, dict) else None
    if not isinstance(tools, list) or not tools:
        return [f"{path.name}: 'tools' must be a non-empty list"]
    problems: list[str] = []
    seen: set[str] = set()
    for index, entry in enumerate(tools, 1):
        where = f"{path.name} entry {index}"
        if not isinstance(entry, dict):
            problems.append(f"{where}: must be an object")
            continue
        tool_id, file, repo = entry.get("id"), entry.get("file"), entry.get("repo")
        if not isinstance(tool_id, str) or not ID_RE.match(tool_id):
            problems.append(f"{where}: invalid id {tool_id!r}")
        elif tool_id in seen:
            problems.append(f"{where}: {tool_id} is listed twice")
        else:
            seen.add(tool_id)
        if not isinstance(file, str) or not ADDIN_FILE_RE.match(file) or file == LOADER_FILE:
            problems.append(f"{where}: invalid file {file!r}")
        if not isinstance(repo, str) or not REPO_RE.match(repo):
            problems.append(f"{where}: invalid repo {repo!r}")
        if not isinstance(entry.get("required"), bool):
            problems.append(f"{where}: 'required' must be true or false")
    return problems


# ---------------------------------------------------------------------------
# Finishing a draft release on GitHub
# ---------------------------------------------------------------------------


class GitHubApi(Protocol):
    def list_releases(self) -> list[dict[str, object]]: ...

    def download_asset(self, asset_id: int) -> bytes: ...

    def upload_asset(self, release_id: int, name: str, data: bytes) -> None: ...

    def delete_asset(self, asset_id: int) -> None: ...

    def update_release(self, release_id: int, fields: dict[str, object]) -> None: ...


def _assets(release: dict[str, object]) -> list[dict[str, object]]:
    assets = release.get("assets")
    return [a for a in assets if isinstance(a, dict)] if isinstance(assets, list) else []


def _int(value: object) -> int:
    if not isinstance(value, int):
        raise CiError(f"unexpected answer from GitHub: {value!r} is not an id")
    return value


def finish_release(
    api: GitHubApi, repo: str, tag: str, root: Path, work: Path, min_loader: str = DEFAULT_MIN_LOADER
) -> str:
    problems = check_tag(tag, root)
    if problems:
        raise CiError("; ".join(problems))
    release = next((r for r in api.list_releases() if r.get("tag_name") == tag), None)
    if release is None:
        raise CiError(
            f"no draft release for {tag}: build the add-in in Excel, then "
            f'gh release create {tag} "<file>.xlam" --draft --repo {repo}, and re-run this job'
        )
    if release.get("draft") is not True:
        return f"{tag} is already published; nothing to do"
    release_id = _int(release.get("id"))
    addins = [a for a in _assets(release) if str(a.get("name", "")).lower().endswith(".xlam")]
    if len(addins) != 1:
        raise CiError(f"the draft release of {tag} must contain exactly one .xlam file (found {len(addins)})")

    work.mkdir(parents=True, exist_ok=True)
    local = work / "uploaded.xlam"
    local.write_bytes(api.download_asset(_int(addins[0].get("id"))))
    try:
        built = mar.build_release(
            local, root / "CHANGELOG.md", work / "dist", repo=repo, min_loader=min_loader, allow_dev=False
        )
    except mar.ReleaseError as exc:
        raise CiError(str(exc).replace("uploaded.xlam", "the uploaded add-in")) from exc
    if built.tag != tag:
        raise CiError(f"the uploaded add-in is version {built.identity.version} ({built.tag}), not {tag}")

    asset_name = built.identity.asset_name
    for asset in _assets(release):
        if asset.get("name") == mar.MANIFEST_NAME:
            api.delete_asset(_int(asset.get("id")))
    if addins[0].get("name") != asset_name:
        api.upload_asset(release_id, asset_name, (built.folder / asset_name).read_bytes())
        api.delete_asset(_int(addins[0].get("id")))
    api.upload_asset(release_id, mar.MANIFEST_NAME, (built.folder / mar.MANIFEST_NAME).read_bytes())
    api.update_release(
        release_id, {"draft": False, "name": built.identity.title, "body": built.notes_file.read_text(encoding="utf-8")}
    )
    return f"{tag} checked and published ({asset_name} + {mar.MANIFEST_NAME})"


class GitHubRestApi:
    """GitHub REST API with the workflow's GITHUB_TOKEN (urllib, no extra packages)."""

    def __init__(self, repo: str, token: str) -> None:
        self._repo = repo
        self._token = token

    def _request(
        self, method: str, url: str, data: bytes | None = None, content_type: str = "application/json"
    ) -> bytes:
        request = urllib.request.Request(url, data=data, method=method)  # noqa: S310 - fixed https GitHub URLs
        request.add_header("Authorization", f"Bearer {self._token}")
        request.add_header("Accept", "application/vnd.github+json")
        request.add_header("X-GitHub-Api-Version", "2022-11-28")
        if data is not None:
            request.add_header("Content-Type", content_type)
        with urllib.request.urlopen(request, timeout=120) as response:  # noqa: S310 - see above
            body: bytes = response.read()
            return body

    def list_releases(self) -> list[dict[str, object]]:
        answer = json.loads(self._request("GET", f"{API_ROOT}/repos/{self._repo}/releases?per_page=100"))
        if not isinstance(answer, list):
            raise CiError("unexpected answer from GitHub when listing releases")
        return [r for r in answer if isinstance(r, dict)]

    def download_asset(self, asset_id: int) -> bytes:
        # GitHub answers with a redirect to signed storage, which must be
        # fetched WITHOUT the Authorization header.
        request = urllib.request.Request(  # noqa: S310 - fixed https GitHub URL
            f"{API_ROOT}/repos/{self._repo}/releases/assets/{asset_id}"
        )
        request.add_header("Authorization", f"Bearer {self._token}")
        request.add_header("Accept", "application/octet-stream")
        opener = urllib.request.build_opener(_NoRedirect)
        try:
            with opener.open(request, timeout=120) as response:
                direct: bytes = response.read()
                return direct
        except urllib.error.HTTPError as exc:
            location = exc.headers.get("Location") if exc.code in (301, 302, 303, 307, 308) else None
            if location is None:
                raise
        with urllib.request.urlopen(location, timeout=120) as response:  # noqa: S310 - URL given by GitHub
            redirected: bytes = response.read()
            return redirected

    def upload_asset(self, release_id: int, name: str, data: bytes) -> None:
        url = f"{UPLOAD_ROOT}/repos/{self._repo}/releases/{release_id}/assets?name={urllib.parse.quote(name)}"
        self._request("POST", url, data, "application/octet-stream")

    def delete_asset(self, asset_id: int) -> None:
        self._request("DELETE", f"{API_ROOT}/repos/{self._repo}/releases/assets/{asset_id}")

    def update_release(self, release_id: int, fields: dict[str, object]) -> None:
        self._request("PATCH", f"{API_ROOT}/repos/{self._repo}/releases/{release_id}", json.dumps(fields).encode())


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(
        self, req: urllib.request.Request, fp: object, code: int, msg: str, headers: object, newurl: str
    ) -> urllib.request.Request | None:
        return None


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------


def _report(problems: list[str]) -> int:
    for problem in problems:
        print(f"error: {problem}", file=sys.stderr)
    return 1 if problems else 0


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="CI for GCV add-in repositories.")
    commands = parser.add_subparsers(dest="command", required=True)
    check = commands.add_parser("check", help="check the sources (and the tag)")
    check.add_argument("--root", type=Path, default=Path())
    check.add_argument("--tag")
    release = commands.add_parser("release", help="check and publish the draft release of a tag")
    release.add_argument("--root", type=Path, default=Path())
    release.add_argument("--tag", required=True)
    release.add_argument("--min-loader", default=DEFAULT_MIN_LOADER)
    args = parser.parse_args(argv)

    if args.command == "check":
        problems = check_sources(args.root)
        if (args.root / "catalog.json").is_file():
            problems += check_catalog(args.root / "catalog.json")
        if args.tag:
            problems += check_tag(args.tag, args.root)
        if not problems:
            print("All checks passed.")
        return _report(problems)

    token = os.environ.get("GITHUB_TOKEN", "")
    repo = os.environ.get("GITHUB_REPOSITORY", "")
    if not token or not repo:
        return _report(["GITHUB_TOKEN and GITHUB_REPOSITORY are needed (run this in GitHub Actions)"])
    try:
        print(
            finish_release(GitHubRestApi(repo, token), repo, args.tag, args.root, Path(".gcv-release"), args.min_loader)
        )
    except (CiError, urllib.error.URLError) as exc:
        return _report([str(exc)])
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

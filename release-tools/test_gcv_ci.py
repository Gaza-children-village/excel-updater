"""Tests for gcv_ci.py (CI checks and release finishing for GCV add-in repositories)."""

from __future__ import annotations

import email.message
import io
import json
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

import pytest

import gcv_ci

ADDIN_CT = "application/vnd.ms-excel.addin.macroEnabled.main+xml"
REPO = "Gaza-children-village/excel-logger"

GCV_ADDIN = (
    'Attribute VB_Name = "GcvAddin"\r\nOption Explicit\r\n'
    'Public Const ADDIN_ID As String = "gcv/logger"\r\n'
    'Public Const ADDIN_NAME As String = "GCV Logger"\r\n'
    'Public Const ADDIN_FILE As String = "GCV Logger.xlam"\r\n'
    'Public Const ADDIN_VERSION As String = "1.0.1"\r\n'
    f'Public Const ADDIN_REPO As String = "{REPO}"\r\n'
)
RIBBON = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" xmlns:gcv="GCV.Tools.Ribbon">'
    '<ribbon><tabs><tab idQ="gcv:GcvToolsTab" label="Gaza Children Village Tools"/></tabs></ribbon></customUI>'
)
CHANGELOG = "# Changelog\n\n## 1.0.1\n\n- Fixed the report.\n\n## 1.0.0\n\n- First.\n"


def make_repo(root: Path, *, version: str = "1.0.1") -> Path:
    (root / "src").mkdir(parents=True)
    (root / "src" / "GcvAddin.bas").write_bytes(GCV_ADDIN.replace('"1.0.1"', f'"{version}"').encode("ascii"))
    (root / "src" / "LoggerUi.bas").write_bytes(b'Attribute VB_Name = "LoggerUi"\r\nOption Explicit\r\n')
    (root / "src" / "Event.cls").write_bytes(
        b'VERSION 1.0 CLASS\r\nBEGIN\r\nEND\r\nAttribute VB_Name = "Event"\r\nOption Explicit\r\n'
    )
    (root / "ThisWorkbook.code").write_bytes(b"Option Explicit\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n")
    (root / "customUI14.xml").write_text(RIBBON, encoding="utf-8")
    (root / "CHANGELOG.md").write_text(CHANGELOG, encoding="utf-8")
    return root


def addin_bytes(*, version: str = "1.0.1", vba: bytes = b"Module=GcvAddin\r\nModule=LoggerUi\r\n") -> bytes:
    props = "".join(
        f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{i}" name="{name}">'
        f"<vt:lpwstr>{value}</vt:lpwstr></property>"
        for i, (name, value) in enumerate(
            (("GcvAddinId", "gcv/logger"), ("GcvAddinVersion", version), ("GcvAddinFile", "GCV Logger.xlam")), 2
        )
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'<Override PartName="/xl/workbook.xml" ContentType="{ADDIN_CT}"/></Types>',
        )
        zf.writestr(
            "docProps/custom.xml",
            '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            f"{props}</Properties>",
        )
        zf.writestr("xl/vbaProject.bin", vba)
        zf.writestr("customUI/customUI14.xml", RIBBON)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Source checks
# ---------------------------------------------------------------------------


def test_a_correct_repository_has_no_problems(tmp_path: Path) -> None:
    assert gcv_ci.check_sources(make_repo(tmp_path)) == []


def test_reports_lf_line_endings(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "src" / "LoggerUi.bas").write_bytes(b'Attribute VB_Name = "LoggerUi"\nOption Explicit\n')
    assert any("CRLF" in p and "LoggerUi.bas" in p for p in gcv_ci.check_sources(root))


def test_reports_non_ascii_text(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "src" / "LoggerUi.bas").write_bytes(
        'Attribute VB_Name = "LoggerUi"\r\nOption Explicit\r\n\' ملاحظة\r\n'.encode()
    )
    assert any("ASCII" in p for p in gcv_ci.check_sources(root))


def test_reports_a_module_name_that_differs_from_the_file(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "src" / "LoggerUi.bas").write_bytes(b'Attribute VB_Name = "Other"\r\nOption Explicit\r\n')
    assert any("VB_Name" in p for p in gcv_ci.check_sources(root))


def test_reports_missing_option_explicit(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "src" / "LoggerUi.bas").write_bytes(b'Attribute VB_Name = "LoggerUi"\r\n')
    assert any("Option Explicit" in p for p in gcv_ci.check_sources(root))


def test_reports_a_ribbon_without_the_shared_tab(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "customUI14.xml").write_text(RIBBON.replace("gcv:GcvToolsTab", "gcv:Other"), encoding="utf-8")
    assert any("GcvToolsTab" in p for p in gcv_ci.check_sources(root))


def test_reports_invalid_ribbon_xml(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "customUI14.xml").write_text("<customUI>", encoding="utf-8")
    assert any("customUI14.xml" in p for p in gcv_ci.check_sources(root))


def test_reports_a_version_without_changelog_section(tmp_path: Path) -> None:
    root = make_repo(tmp_path, version="1.0.2")
    assert any("CHANGELOG" in p and "1.0.2" in p for p in gcv_ci.check_sources(root))


def test_reports_a_missing_identity_module(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    (root / "src" / "GcvAddin.bas").unlink()
    assert any("GcvAddin.bas" in p for p in gcv_ci.check_sources(root))


def test_reads_the_addin_identity(tmp_path: Path) -> None:
    module = gcv_ci.read_addin_module(make_repo(tmp_path))
    assert module == gcv_ci.AddinModule("gcv/logger", "GCV Logger.xlam", "1.0.1", REPO)
    assert module.tag == "logger-v1.0.1"


def test_tag_must_match_the_version(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    assert gcv_ci.check_tag("logger-v1.0.1", root) == []
    assert any("logger-v1.0.1" in p for p in gcv_ci.check_tag("logger-v1.0.2", root))


# ---------------------------------------------------------------------------
# Catalog
# ---------------------------------------------------------------------------


def write_catalog(root: Path, tools: list[dict[str, object]]) -> Path:
    path = root / "catalog.json"
    path.write_text(json.dumps({"tools": tools}), encoding="utf-8")
    return path


def tool(**overrides: object) -> dict[str, object]:
    base: dict[str, object] = {
        "id": "gcv/logger",
        "name": "GCV Logger",
        "file": "GCV Logger.xlam",
        "repo": REPO,
        "required": True,
        "description": "Logs.",
    }
    base.update(overrides)
    return base


def test_a_correct_catalog_has_no_problems(tmp_path: Path) -> None:
    assert gcv_ci.check_catalog(write_catalog(tmp_path, [tool()])) == []


@pytest.mark.parametrize(
    ("overrides", "word"),
    [
        ({"id": "Logger"}, "id"),
        ({"file": "..\\GCV x.xlam"}, "file"),
        ({"file": "GCV Loader.xlam"}, "file"),
        ({"repo": "no-slash"}, "repo"),
        ({"required": "yes"}, "required"),
    ],
)
def test_reports_bad_catalog_entries(tmp_path: Path, overrides: dict[str, object], word: str) -> None:
    assert any(word in p for p in gcv_ci.check_catalog(write_catalog(tmp_path, [tool(**overrides)])))


def test_reports_duplicate_and_unreadable_catalogs(tmp_path: Path) -> None:
    assert any("twice" in p for p in gcv_ci.check_catalog(write_catalog(tmp_path, [tool(), tool()])))
    (tmp_path / "bad.json").write_text("{", encoding="utf-8")
    assert any("JSON" in p for p in gcv_ci.check_catalog(tmp_path / "bad.json"))
    (tmp_path / "empty.json").write_text("{}", encoding="utf-8")
    assert any("tools" in p for p in gcv_ci.check_catalog(tmp_path / "empty.json"))


# ---------------------------------------------------------------------------
# Finishing a draft release
# ---------------------------------------------------------------------------


class FakeGitHub:
    def __init__(self, releases: list[dict[str, object]], assets: dict[int, bytes]) -> None:
        self.releases = releases
        self.assets = assets
        self.uploaded: dict[str, bytes] = {}
        self.deleted: list[int] = []
        self.updates: dict[str, object] = {}

    def list_releases(self) -> list[dict[str, object]]:
        return self.releases

    def download_asset(self, asset_id: int) -> bytes:
        return self.assets[asset_id]

    def upload_asset(self, release_id: int, name: str, data: bytes) -> None:
        self.uploaded[name] = data

    def delete_asset(self, asset_id: int) -> None:
        self.deleted.append(asset_id)

    def update_release(self, release_id: int, fields: dict[str, object]) -> None:
        self.updates = fields


def draft(tag: str = "logger-v1.0.1", assets: list[dict[str, object]] | None = None) -> dict[str, object]:
    return {
        "id": 7,
        "tag_name": tag,
        "draft": True,
        "assets": assets if assets is not None else [{"id": 11, "name": "GCV.Logger.xlam"}],
    }


def test_release_validates_adds_manifest_and_publishes(tmp_path: Path) -> None:
    data = addin_bytes()
    api = FakeGitHub([draft()], {11: data})
    message = gcv_ci.finish_release(api, REPO, "logger-v1.0.1", make_repo(tmp_path / "repo"), tmp_path / "work")

    assert "published" in message
    assert api.uploaded["GCV-Logger.xlam"] == data
    manifest = json.loads(api.uploaded["manifest.json"])
    assert manifest["id"] == "gcv/logger"
    assert manifest["version"] == "1.0.1"
    assert manifest["file"]["name"] == "GCV-Logger.xlam"
    assert manifest["file"]["size"] == len(data)
    assert api.deleted == [11]  # GitHub's "GCV.Logger.xlam" replaced
    assert api.updates == {"draft": False, "name": "GCV Logger 1.0.1", "body": "- Fixed the report.\n"}


def test_release_keeps_an_asset_that_already_has_the_right_name(tmp_path: Path) -> None:
    api = FakeGitHub([draft(assets=[{"id": 11, "name": "GCV-Logger.xlam"}])], {11: addin_bytes()})
    gcv_ci.finish_release(api, REPO, "logger-v1.0.1", make_repo(tmp_path / "repo"), tmp_path / "work")
    assert "GCV-Logger.xlam" not in api.uploaded
    assert api.deleted == []


def test_release_replaces_an_old_manifest_when_run_again(tmp_path: Path) -> None:
    assets = [{"id": 11, "name": "GCV-Logger.xlam"}, {"id": 12, "name": "manifest.json"}]
    api = FakeGitHub([draft(assets=assets)], {11: addin_bytes()})
    gcv_ci.finish_release(api, REPO, "logger-v1.0.1", make_repo(tmp_path / "repo"), tmp_path / "work")
    assert api.deleted == [12]
    assert "manifest.json" in api.uploaded


def test_release_refuses_a_file_with_another_version(tmp_path: Path) -> None:
    api = FakeGitHub([draft()], {11: addin_bytes(version="1.0.0")})
    with pytest.raises(gcv_ci.CiError, match=r"1\.0\.0"):
        gcv_ci.finish_release(api, REPO, "logger-v1.0.1", make_repo(tmp_path / "repo"), tmp_path / "work")
    assert api.updates == {}


def test_release_refuses_test_modules(tmp_path: Path) -> None:
    api = FakeGitHub([draft()], {11: addin_bytes(vba=b"Module=GcvAddin\r\nModule=TestGcvLog\r\n")})
    with pytest.raises(gcv_ci.CiError, match="development modules"):
        gcv_ci.finish_release(api, REPO, "logger-v1.0.1", make_repo(tmp_path / "repo"), tmp_path / "work")
    assert api.updates == {}


def test_release_explains_a_missing_draft(tmp_path: Path) -> None:
    with pytest.raises(gcv_ci.CiError, match="draft release"):
        gcv_ci.finish_release(FakeGitHub([], {}), REPO, "logger-v1.0.1", make_repo(tmp_path / "r"), tmp_path / "w")


def test_release_already_published_is_left_alone(tmp_path: Path) -> None:
    published = draft()
    published["draft"] = False
    api = FakeGitHub([published], {})
    message = gcv_ci.finish_release(api, REPO, "logger-v1.0.1", make_repo(tmp_path / "r"), tmp_path / "w")
    assert "already published" in message
    assert api.updates == {}


@pytest.mark.parametrize("names", [[], ["a.xlam", "b.xlam"]])
def test_release_needs_exactly_one_addin_file(tmp_path: Path, names: list[str]) -> None:
    assets = [{"id": i, "name": n} for i, n in enumerate(names, 20)]
    with pytest.raises(gcv_ci.CiError, match=r"exactly one \.xlam"):
        gcv_ci.finish_release(
            FakeGitHub([draft(assets=assets)], {}), REPO, "logger-v1.0.1", make_repo(tmp_path / "r"), tmp_path / "w"
        )


def test_release_refuses_a_tag_that_does_not_match_the_sources(tmp_path: Path) -> None:
    with pytest.raises(gcv_ci.CiError, match=r"logger-v1\.0\.1"):
        gcv_ci.finish_release(
            FakeGitHub([draft(tag="logger-v9.0.0")], {}),
            REPO,
            "logger-v9.0.0",
            make_repo(tmp_path / "r"),
            tmp_path / "w",
        )


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------


def test_check_command_passes_and_fails(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    root = make_repo(tmp_path / "ok")
    assert gcv_ci.main(["check", "--root", str(root)]) == 0
    assert gcv_ci.main(["check", "--root", str(root), "--tag", "logger-v2.0.0"]) == 1
    assert "logger-v1.0.1" in capsys.readouterr().err


def test_check_command_reads_the_catalog_when_present(tmp_path: Path) -> None:
    root = make_repo(tmp_path / "ok")
    write_catalog(root, [tool(repo="bad")])
    assert gcv_ci.main(["check", "--root", str(root)]) == 1


def test_release_command_needs_a_token(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("GITHUB_TOKEN", raising=False)
    monkeypatch.setenv("GITHUB_REPOSITORY", REPO)
    assert gcv_ci.main(["release", "--root", str(make_repo(tmp_path / "r")), "--tag", "logger-v1.0.1"]) == 1


def test_release_command_runs_with_the_api(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    api = FakeGitHub([draft()], {11: addin_bytes()})
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setenv("GITHUB_REPOSITORY", REPO)
    monkeypatch.setattr(gcv_ci, "GitHubRestApi", lambda repo, token: api)
    assert gcv_ci.main(["release", "--root", str(make_repo(tmp_path / "r")), "--tag", "logger-v1.0.1"]) == 0
    assert api.updates["draft"] is False


@pytest.mark.parametrize(
    ("old", "new", "word"),
    [
        ('"gcv/logger"', '"Logger"', "ADDIN_ID"),
        ('"GCV Logger.xlam"', '"Logger.xlsm"', "ADDIN_FILE"),
        (f'"{REPO}"', '"no-slash"', "ADDIN_REPO"),
        ('"1.0.1"', '"1.0"', "ADDIN_VERSION"),
    ],
)
def test_reports_an_invalid_identity(tmp_path: Path, old: str, new: str, word: str) -> None:
    root = make_repo(tmp_path)
    path = root / "src" / "GcvAddin.bas"
    path.write_bytes(path.read_bytes().replace(old.encode(), new.encode()))
    assert any(word in p for p in gcv_ci.check_sources(root))


def test_reports_a_missing_constant_and_checks_loader_files(tmp_path: Path) -> None:
    root = make_repo(tmp_path)
    path = root / "src" / "GcvAddin.bas"
    path.write_bytes(path.read_bytes().replace(b"ADDIN_REPO", b"OTHER"))
    (root / "loader").mkdir()
    (root / "loader" / "ThisWorkbook.code").write_bytes(b"Option Explicit\n")
    problems = gcv_ci.check_sources(root)
    assert any("ADDIN_REPO is missing" in p for p in problems)
    assert any("loader/ThisWorkbook.code" in p for p in problems)
    assert any("ADDIN_REPO" in p for p in gcv_ci.check_tag("logger-v1.0.1", root))


def test_reports_catalog_entries_that_are_not_objects(tmp_path: Path) -> None:
    path = tmp_path / "catalog.json"
    path.write_text('{"tools": ["x"]}', encoding="utf-8")
    assert any("object" in p for p in gcv_ci.check_catalog(path))


def test_release_reports_a_bad_id_from_github(tmp_path: Path) -> None:
    bad = draft()
    bad["id"] = "seven"
    with pytest.raises(gcv_ci.CiError, match="not an id"):
        gcv_ci.finish_release(FakeGitHub([bad], {}), REPO, "logger-v1.0.1", make_repo(tmp_path / "r"), tmp_path / "w")


# ---------------------------------------------------------------------------
# GitHub REST client (fake connection, no network)
# ---------------------------------------------------------------------------


def headers(location: str | None) -> email.message.Message:
    message = email.message.Message()
    if location is not None:
        message["Location"] = location
    return message


class FakeResponse(io.BytesIO):
    def __enter__(self) -> FakeResponse:
        return self

    def __exit__(self, *args: object) -> None:
        self.close()


def test_rest_api_sends_the_token_and_reads_releases(monkeypatch: pytest.MonkeyPatch) -> None:
    seen: list[urllib.request.Request] = []

    def fake_urlopen(request: urllib.request.Request, timeout: int) -> FakeResponse:
        seen.append(request)
        return FakeResponse(b'[{"tag_name": "logger-v1.0.1"}, "x"]')

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    api = gcv_ci.GitHubRestApi(REPO, "secret")
    assert api.list_releases() == [{"tag_name": "logger-v1.0.1"}]
    api.upload_asset(7, "GCV-Logger.xlam", b"data")
    api.delete_asset(11)
    api.update_release(7, {"draft": False})
    assert seen[0].get_header("Authorization") == "Bearer secret"
    assert seen[1].full_url.endswith("/releases/7/assets?name=GCV-Logger.xlam")
    assert seen[1].get_header("Content-type") == "application/octet-stream"
    assert seen[2].get_method() == "DELETE"
    body = seen[3].data
    assert isinstance(body, bytes)
    assert json.loads(body) == {"draft": False}


def test_rest_api_rejects_an_unexpected_release_list(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(urllib.request, "urlopen", lambda request, timeout: FakeResponse(b"{}"))
    with pytest.raises(gcv_ci.CiError, match="listing releases"):
        gcv_ci.GitHubRestApi(REPO, "t").list_releases()


def test_rest_api_follows_the_download_redirect_without_the_token(monkeypatch: pytest.MonkeyPatch) -> None:
    redirected: list[str] = []

    class Opener:
        def open(self, request: urllib.request.Request, timeout: int) -> FakeResponse:
            raise urllib.error.HTTPError(
                request.full_url,
                302,
                "Found",
                headers("https://storage/file"),
                None,
            )

    def fake_urlopen(url: str, timeout: int) -> FakeResponse:
        redirected.append(url)
        return FakeResponse(b"xlam bytes")

    monkeypatch.setattr(urllib.request, "build_opener", lambda *handlers: Opener())
    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    assert gcv_ci.GitHubRestApi(REPO, "t").download_asset(11) == b"xlam bytes"
    assert redirected == ["https://storage/file"]


def test_rest_api_download_without_redirect_and_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    class Direct:
        def open(self, request: urllib.request.Request, timeout: int) -> FakeResponse:
            return FakeResponse(b"direct")

    class Missing:
        def open(self, request: urllib.request.Request, timeout: int) -> FakeResponse:
            raise urllib.error.HTTPError(request.full_url, 404, "Not Found", headers(None), None)

    monkeypatch.setattr(urllib.request, "build_opener", lambda *handlers: Direct())
    assert gcv_ci.GitHubRestApi(REPO, "t").download_asset(11) == b"direct"
    monkeypatch.setattr(urllib.request, "build_opener", lambda *handlers: Missing())
    with pytest.raises(urllib.error.HTTPError):
        gcv_ci.GitHubRestApi(REPO, "t").download_asset(11)


def test_no_redirect_handler_stops_redirects() -> None:
    handler = gcv_ci._NoRedirect()
    request = urllib.request.Request("https://api.github.com/x")
    assert handler.redirect_request(request, None, 302, "Found", headers(None), "https://storage/file") is None


def test_release_command_reports_errors(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setenv("GITHUB_REPOSITORY", REPO)
    monkeypatch.setattr(gcv_ci, "GitHubRestApi", lambda repo, token: FakeGitHub([], {}))
    assert gcv_ci.main(["release", "--root", str(make_repo(tmp_path / "r")), "--tag", "logger-v1.0.1"]) == 1

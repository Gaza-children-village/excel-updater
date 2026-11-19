"""Tests for make_addin_release.py (GCV add-in releases)."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

import pytest

import make_addin_release as mtr

ADDIN_CT = "application/vnd.ms-excel.addin.macroEnabled.main+xml"
XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"


def custom_xml(version: str | None, addin_id: str = "gcv/student-lookup", file: str = "GCV Student Lookup.xlam") -> str:
    props = ""
    pid = 2
    for name, value in (("GcvAddinId", addin_id), ("GcvAddinVersion", version), ("GcvAddinFile", file)):
        if value is not None:
            props += (
                f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" name="{name}">'
                f"<vt:lpwstr>{value}</vt:lpwstr></property>"
            )
            pid += 1
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f"{props}</Properties>"
    )


def make_addin(
    path: Path,
    *,
    version: str | None = "0.2.0",
    addin_id: str = "gcv/student-lookup",
    file: str = "GCV Student Lookup.xlam",
    content_type: str = ADDIN_CT,
    vba: bytes | None = b'ID="{x}"\r\nModule=GcvTools\r\nModule=ToolsUpdater\r\nClass=UpdWinHttp\r\n',
    ribbon: bool = True,
) -> Path:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'<Override PartName="/xl/workbook.xml" ContentType="{content_type}"/></Types>',
        )
        zf.writestr("xl/workbook.xml", "<workbook/>")
        zf.writestr("docProps/custom.xml", custom_xml(version, addin_id, file))
        if vba is not None:
            zf.writestr("xl/vbaProject.bin", vba)
        if ribbon:
            zf.writestr("customUI/customUI14.xml", "<customUI/>")
    return path


def write_changelog(
    folder: Path, text: str = "# Changelog\n\n## 0.2.0\n\n- New updater.\n\n## 0.1.0\n\n- First.\n"
) -> Path:
    path = folder / "CHANGELOG.md"
    path.write_text(text, encoding="utf-8")
    return path


def test_reads_the_version_recorded_by_the_add_in(tmp_path: Path) -> None:
    identity = mtr.read_addin_identity(make_addin(tmp_path / "a.xlam"))
    assert identity == mtr.AddinIdentity("gcv/student-lookup", "0.2.0", "GCV Student Lookup.xlam")


def test_rejects_a_file_without_the_version_property(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match="GcvAddinVersion"):
        mtr.read_addin_identity(make_addin(tmp_path / "a.xlam", version=None))


def test_rejects_an_invalid_version(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match=r"MAJOR\.MINOR\.PATCH"):
        mtr.read_addin_identity(make_addin(tmp_path / "a.xlam", version="0.2"))


def test_rejects_a_file_that_is_not_a_zip(tmp_path: Path) -> None:
    path = tmp_path / "a.xlam"
    path.write_bytes(b"not a zip")
    with pytest.raises(mtr.ReleaseError, match="not an Office file"):
        mtr.check_addin(path)


def test_rejects_a_workbook_that_is_not_an_add_in(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match=r"\.xlam add-in"):
        mtr.check_addin(make_addin(tmp_path / "a.xlam", content_type=XLSX_CT))


def test_rejects_an_add_in_without_vba(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match="no VBA project"):
        mtr.check_addin(make_addin(tmp_path / "a.xlam", vba=None))


def test_rejects_an_add_in_without_the_ribbon(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match="ribbon"):
        mtr.check_addin(make_addin(tmp_path / "a.xlam", ribbon=False))


@pytest.mark.parametrize("module", [b"Module=TestToolsUpdater", b"Class=UpdFakeHttp", b"Module=TestGcvLog"])
def test_rejects_test_modules_in_a_release(tmp_path: Path, module: bytes) -> None:
    with pytest.raises(mtr.ReleaseError, match="development modules"):
        mtr.check_addin(make_addin(tmp_path / "a.xlam", vba=b"Module=GcvTools\r\n" + module + b"\r\n"))


def test_dev_modules_are_allowed_when_asked(tmp_path: Path) -> None:
    path = make_addin(tmp_path / "a.xlam", vba=b"Module=GcvTools\r\nModule=TestToolsUpdater\r\n")
    mtr.check_addin(path, allow_dev=True)


def test_builds_release_folder_manifest_and_notes(tmp_path: Path) -> None:
    addin = make_addin(tmp_path / "GCV Tools 0.2.0.xlam")
    changelog = write_changelog(tmp_path)
    release = mtr.build_release(
        addin, changelog, tmp_path / "dist", repo="gcv-org/excel-student-lookup", min_loader="2.0.0", allow_dev=False
    )

    assert release.tag == "student-lookup-v0.2.0"
    asset = release.folder / "GCV-Student-Lookup.xlam"
    assert asset.read_bytes() == addin.read_bytes()
    manifest = json.loads((release.folder / "manifest.json").read_text(encoding="utf-8"))
    assert manifest == {
        "id": "gcv/student-lookup",
        "version": "0.2.0",
        "minLoaderVersion": "2.0.0",
        "file": {
            "name": "GCV-Student-Lookup.xlam",
            "sha256": hashlib.sha256(addin.read_bytes()).hexdigest(),
            "size": addin.stat().st_size,
        },
        "changelog": "- New updater.",
    }
    assert release.notes_file.read_text(encoding="utf-8") == "- New updater.\n"


def test_gh_command_lists_both_assets(tmp_path: Path) -> None:
    release = mtr.build_release(
        make_addin(tmp_path / "a.xlam"),
        write_changelog(tmp_path),
        tmp_path / "dist",
        repo="gcv-org/excel-student-lookup",
        min_loader="2.0.0",
        allow_dev=False,
    )
    command = release.gh_command()
    assert command[:4] == ["gh", "release", "create", "student-lookup-v0.2.0"]
    assert str(release.folder / "GCV-Student-Lookup.xlam") in command
    assert str(release.folder / "manifest.json") in command
    assert command[command.index("--repo") + 1] == "gcv-org/excel-student-lookup"


def test_requires_a_changelog_section_for_the_version(tmp_path: Path) -> None:
    changelog = write_changelog(tmp_path, "# Changelog\n\n## 0.1.0\n\n- First.\n")
    with pytest.raises(mtr.ReleaseError, match=r"no '## 0\.2\.0' section"):
        mtr.build_release(
            make_addin(tmp_path / "a.xlam"),
            changelog,
            tmp_path / "dist",
            repo="gcv-org/excel-student-lookup",
            min_loader="2.0.0",
            allow_dev=False,
        )


@pytest.mark.parametrize(
    ("repo", "min_loader", "message"),
    [
        ("OWNER/REPOSITORY", "1.0.0", "real repository"),
        ("not-a-repo", "1.0.0", "owner/repository"),
        ("gcv-org/excel-student-lookup", "1.0", "--min-loader"),
    ],
)
def test_rejects_bad_options(tmp_path: Path, repo: str, min_loader: str, message: str) -> None:
    with pytest.raises(mtr.ReleaseError, match=message):
        mtr.build_release(
            make_addin(tmp_path / "a.xlam"),
            write_changelog(tmp_path),
            tmp_path / "dist",
            repo=repo,
            min_loader=min_loader,
            allow_dev=False,
        )


def test_main_prints_the_publish_command(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    addin = make_addin(tmp_path / "a.xlam")
    code = mtr.main(
        [
            str(addin),
            "--changelog",
            str(write_changelog(tmp_path)),
            "--out",
            str(tmp_path / "dist"),
            "--repo",
            "gcv-org/excel-student-lookup",
        ]
    )
    assert code == 0
    out = capsys.readouterr().out
    assert "student-lookup-v0.2.0" in out
    assert "gh release create" in out


def test_main_reports_errors(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    addin = make_addin(tmp_path / "a.xlam", ribbon=False)
    code = mtr.main(
        [str(addin), "--changelog", str(write_changelog(tmp_path)), "--repo", "gcv-org/excel-student-lookup"]
    )
    assert code == 1
    assert "error:" in capsys.readouterr().err


def test_rejects_a_file_never_saved_by_the_add_in(tmp_path: Path) -> None:
    path = tmp_path / "a.xlam"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", "<workbook/>")
    with pytest.raises(mtr.ReleaseError, match="save the add-in in Excel first"):
        mtr.read_addin_identity(path)


def test_reports_a_missing_changelog(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match="is missing"):
        mtr.build_release(
            make_addin(tmp_path / "a.xlam"),
            tmp_path / "CHANGELOG.md",
            tmp_path / "dist",
            repo="gcv-org/excel-student-lookup",
            min_loader="2.0.0",
            allow_dev=False,
        )


def test_rebuilding_replaces_the_previous_release_folder(tmp_path: Path) -> None:
    args = (make_addin(tmp_path / "a.xlam"), write_changelog(tmp_path), tmp_path / "dist")
    first = mtr.build_release(*args, repo="gcv-org/excel-student-lookup", min_loader="2.0.0", allow_dev=False)
    (first.folder / "stale.txt").write_text("old", encoding="utf-8")
    second = mtr.build_release(*args, repo="gcv-org/excel-student-lookup", min_loader="2.0.0", allow_dev=False)
    assert not (second.folder / "stale.txt").exists()


def test_rejects_an_id_without_vendor(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match="vendor/package"):
        mtr.read_addin_identity(make_addin(tmp_path / "a.xlam", addin_id="studentlookup"))


def test_rejects_a_file_name_that_is_not_a_gcv_add_in(tmp_path: Path) -> None:
    with pytest.raises(mtr.ReleaseError, match=r"GCV <name>\.xlam"):
        mtr.read_addin_identity(make_addin(tmp_path / "a.xlam", file="Student.xlam"))

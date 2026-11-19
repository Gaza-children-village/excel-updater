"""Build a GitHub release of one GCV Excel add-in.

Input: the release build of the add-in (an .xlam saved in Excel, without test
modules) and its CHANGELOG.md. Every GCV add-in records its id, version and
file name in the document properties GcvAddinId, GcvAddinVersion and
GcvAddinFile when it is saved (GcvAddin.OnAddinSave).

Output: dist/<package>-v<version>/ with the add-in (file name with dashes, e.g.
GCV-Student-Lookup.xlam) and manifest.json (SHA-256 and size), the release
notes, and the `gh release create` command.

Usage:
    python make_addin_release.py "GCV Student Lookup.xlam" --repo owner/repository
        [--changelog CHANGELOG.md] [--out dist] [--min-loader 2.0.0] [--allow-dev]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shlex
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree

MANIFEST_NAME = "manifest.json"
ID_PROPERTY = "GcvAddinId"
VERSION_PROPERTY = "GcvAddinVersion"
FILE_PROPERTY = "GcvAddinFile"
PLACEHOLDER_REPO = "OWNER/REPOSITORY"
DEFAULT_MIN_LOADER = "2.0.0"

ADDIN_CONTENT_TYPE = "application/vnd.ms-excel.addin.macroEnabled.main+xml"
CONTENT_TYPES_PART = "[Content_Types].xml"
CUSTOM_PROPS_PART = "docProps/custom.xml"
VBA_PART = "xl/vbaProject.bin"
RIBBON_PARTS = ("customUI/customUI14.xml", "customUI/customUI.xml")

CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
PROPS_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}"
VT_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"

VERSION_RE = re.compile(r"^\d{1,6}\.\d{1,6}\.\d{1,6}$")
REPO_RE = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")
ID_RE = re.compile(r"^[a-z0-9-]+/[a-z0-9-]+$")
FILE_RE = re.compile(r"^GCV [A-Za-z0-9 -]+\.xlam$")
# The PROJECT stream of vbaProject.bin lists components as plain text lines
# ("Module=Name", "Class=Name"); development components follow the naming
# rule of the package standard (Test*, *Fake*).
DEV_COMPONENT_RE = re.compile(rb"(?:Module|Class)=((?:Test\w*)|(?:\w*Fake\w*))\r?\n")


class ReleaseError(Exception):
    """The add-in or the options cannot be released as they are."""


@dataclass(frozen=True, slots=True)
class AddinIdentity:
    addin_id: str
    version: str
    file: str

    @property
    def tag(self) -> str:
        return f"{self.addin_id.split('/', 1)[1]}-v{self.version}"

    @property
    def asset_name(self) -> str:
        return self.file.replace(" ", "-")

    @property
    def title(self) -> str:
        return f"{self.file.removesuffix('.xlam')} {self.version}"


@dataclass(frozen=True, slots=True)
class AddinRelease:
    identity: AddinIdentity
    repo: str
    folder: Path
    notes_file: Path

    @property
    def tag(self) -> str:
        return self.identity.tag

    def gh_command(self) -> list[str]:
        return [
            "gh",
            "release",
            "create",
            self.tag,
            str(self.folder / self.identity.asset_name),
            str(self.folder / MANIFEST_NAME),
            "--repo",
            self.repo,
            "--title",
            self.identity.title,
            "--notes-file",
            str(self.notes_file),
        ]


def _open_zip(path: Path) -> zipfile.ZipFile:
    try:
        return zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise ReleaseError(f"{path.name} is not an Office file (not a zip package)") from exc


def _properties(path: Path) -> dict[str, str]:
    with _open_zip(path) as package:
        if CUSTOM_PROPS_PART not in package.namelist():
            return {}
        root = ElementTree.fromstring(package.read(CUSTOM_PROPS_PART))  # noqa: S314 - local file built by Ahmed, not untrusted input
    values: dict[str, str] = {}
    for prop in root.iter(f"{PROPS_NS}property"):
        value = prop.find(f"{VT_NS}lpwstr")
        values[prop.get("name", "")] = "" if value is None or value.text is None else value.text.strip()
    return values


def read_addin_identity(path: Path) -> AddinIdentity:
    """Id, version and file name recorded by the add-in in its document properties."""
    props = _properties(path)
    for name in (ID_PROPERTY, VERSION_PROPERTY, FILE_PROPERTY):
        if name not in props:
            raise ReleaseError(f"{path.name}: no {name} property - save the add-in in Excel first")
    identity = AddinIdentity(props[ID_PROPERTY], props[VERSION_PROPERTY], props[FILE_PROPERTY])
    if not ID_RE.match(identity.addin_id):
        raise ReleaseError(f"{path.name}: {ID_PROPERTY} '{identity.addin_id}' must be vendor/package")
    if not VERSION_RE.match(identity.version):
        raise ReleaseError(f"{path.name}: {VERSION_PROPERTY} '{identity.version}' must be MAJOR.MINOR.PATCH")
    if not FILE_RE.match(identity.file):
        raise ReleaseError(f"{path.name}: {FILE_PROPERTY} '{identity.file}' must be 'GCV <name>.xlam'")
    return identity


def check_addin(path: Path, *, allow_dev: bool = False) -> None:
    """Checks that the file is a macro add-in with VBA, the ribbon and no test modules."""
    with _open_zip(path) as package:
        names = set(package.namelist())
        types = ElementTree.fromstring(package.read(CONTENT_TYPES_PART))  # noqa: S314 - local file built by Ahmed
        workbook_types = [
            item.get("ContentType")
            for item in types.iter(f"{CT_NS}Override")
            if item.get("PartName") == "/xl/workbook.xml"
        ]
        if workbook_types != [ADDIN_CONTENT_TYPE]:
            raise ReleaseError(f"{path.name}: must be saved as an .xlam add-in")
        if VBA_PART not in names:
            raise ReleaseError(f"{path.name}: no VBA project - import the modules and save in Excel")
        if not any(part in names for part in RIBBON_PARTS):
            raise ReleaseError(f"{path.name}: the ribbon (customUI14.xml) is missing - start from the shell file")
        vba = package.read(VBA_PART)
    dev = sorted({match.decode("ascii", "replace") for match in DEV_COMPONENT_RE.findall(vba)})
    if dev and not allow_dev:
        raise ReleaseError(f"{path.name}: contains development modules {', '.join(dev)} - remove them for a release")


def changelog_section(changelog: Path, version: str) -> str:
    if not changelog.is_file():
        raise ReleaseError(f"{changelog} is missing")
    lines = changelog.read_text(encoding="utf-8").splitlines()
    heading = re.compile(rf"^##\s+v?{re.escape(version)}\b")
    for start, line in enumerate(lines):
        if heading.match(line):
            body: list[str] = []
            for line_after in lines[start + 1 :]:
                if line_after.startswith("## "):
                    break
                body.append(line_after)
            return "\n".join(body).strip() + "\n"
    raise ReleaseError(f"{changelog.name} has no '## {version}' section")


def build_release(
    addin: Path, changelog: Path, out_dir: Path, *, repo: str, min_loader: str, allow_dev: bool
) -> AddinRelease:
    if repo == PLACEHOLDER_REPO:
        raise ReleaseError("--repo must be the real repository, not OWNER/REPOSITORY")
    if not REPO_RE.match(repo):
        raise ReleaseError("--repo must be owner/repository")
    if not VERSION_RE.match(min_loader):
        raise ReleaseError("--min-loader must be MAJOR.MINOR.PATCH")
    check_addin(addin, allow_dev=allow_dev)
    identity = read_addin_identity(addin)
    notes = changelog_section(changelog, identity.version)

    folder = out_dir / identity.tag
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True)
    data = addin.read_bytes()
    (folder / identity.asset_name).write_bytes(data)
    manifest = {
        "id": identity.addin_id,
        "version": identity.version,
        "minLoaderVersion": min_loader,
        "file": {"name": identity.asset_name, "sha256": hashlib.sha256(data).hexdigest(), "size": len(data)},
        "changelog": notes.strip(),
    }
    (folder / MANIFEST_NAME).write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    notes_file = out_dir / f"{identity.tag}-notes.md"
    notes_file.write_text(notes, encoding="utf-8")
    return AddinRelease(identity=identity, repo=repo, folder=folder, notes_file=notes_file)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Build a GitHub release of one GCV Excel add-in.")
    parser.add_argument("addin", type=Path)
    parser.add_argument("--repo", required=True)
    parser.add_argument("--changelog", type=Path, default=Path("CHANGELOG.md"))
    parser.add_argument("--out", type=Path, default=Path("dist"))
    parser.add_argument("--min-loader", default=DEFAULT_MIN_LOADER)
    parser.add_argument("--allow-dev", action="store_true")
    args = parser.parse_args(argv)
    try:
        release = build_release(
            args.addin, args.changelog, args.out, repo=args.repo, min_loader=args.min_loader, allow_dev=args.allow_dev
        )
    except ReleaseError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(f"Release {release.tag} prepared in {release.folder}")
    print("Publish with:")
    print("  " + shlex.join(release.gh_command()))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

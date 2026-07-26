#!/usr/bin/env python3
"""Read-only structural audit for an Obsidian study vault."""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
import urllib.parse
from collections import defaultdict, deque
from pathlib import Path


WIKILINK_RE = re.compile(r"(!?)\[\[([^\]]+)\]\]")
MARKDOWN_LINK_RE = re.compile(r"!?\[[^\]]*\]\(([^)]+)\)")
FRONTMATTER_RE = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.DOTALL)
TITLE_RE = re.compile(r"(?m)^#\s+(.+?)\s*$")
TYPE_RE = re.compile(r"(?m)^type:\s*[\"']?([^\"'\n]+)")
DEFAULT_NOISE = (
    r"公众号",
    r"扫码关注",
    r"免费分享",
    r"加微信",
    r"微信号",
    r"QQ群",
    r"购买课程",
    r"资料领取",
)
SOURCE_SUFFIXES = {".pdf", ".doc", ".docx", ".ppt", ".pptx", ".png", ".jpg", ".jpeg"}
IGNORED_PARTS = {".git", ".obsidian", ".trash", "__pycache__"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--vault", type=Path, required=True, help="Obsidian vault directory")
    parser.add_argument("--source-root", type=Path, help="Original source directory")
    parser.add_argument("--root-note", help="Root note path or basename")
    parser.add_argument(
        "--noise-pattern",
        action="append",
        default=[],
        help="Additional regular expression; repeat as needed",
    )
    parser.add_argument(
        "--min-concept-chars",
        type=int,
        default=120,
        help="Minimum non-metadata characters for concept notes",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat ambiguous links and unreachable notes as failures",
    )
    return parser.parse_args()


def visible_files(root: Path, suffix: str | None = None) -> list[Path]:
    files = []
    for path in root.rglob("*"):
        if not path.is_file() or any(part in IGNORED_PARTS for part in path.parts):
            continue
        if suffix is None or path.suffix.lower() == suffix:
            files.append(path)
    return sorted(files)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig", errors="replace")


def note_title(text: str, path: Path) -> str:
    frontmatter = FRONTMATTER_RE.match(text)
    if frontmatter:
        match = re.search(r'(?m)^title:\s*["\']?(.+?)["\']?\s*$', frontmatter.group(1))
        if match:
            return match.group(1).strip()
    heading = TITLE_RE.search(text)
    return heading.group(1).strip() if heading else path.stem


def clean_target(raw: str) -> str:
    target = raw.split("|", 1)[0].split("#", 1)[0].strip()
    return urllib.parse.unquote(target).replace("\\", "/")


def build_note_index(vault: Path, notes: list[Path]) -> tuple[dict[str, Path], dict[str, list[Path]]]:
    by_relative = {}
    by_stem: dict[str, list[Path]] = defaultdict(list)
    for note in notes:
        relative = note.relative_to(vault).with_suffix("").as_posix().casefold()
        by_relative[relative] = note
        by_stem[note.stem.casefold()].append(note)
    return by_relative, by_stem


def resolve_note(
    target: str,
    by_relative: dict[str, Path],
    by_stem: dict[str, list[Path]],
) -> tuple[Path | None, bool]:
    key = target.removesuffix(".md").strip("/").casefold()
    if key in by_relative:
        return by_relative[key], False
    matches = by_stem.get(Path(key).name, [])
    if len(matches) == 1:
        return matches[0], False
    return None, len(matches) > 1


def normalized_body(text: str) -> str:
    text = FRONTMATTER_RE.sub("", text, count=1)
    text = WIKILINK_RE.sub("", text)
    text = MARKDOWN_LINK_RE.sub("", text)
    text = re.sub(r"(?m)^#+\s+.*$", "", text)
    text = re.sub(r"[\W_]+", "", text, flags=re.UNICODE)
    return text


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def audit() -> int:
    args = parse_args()
    vault = args.vault.resolve()
    if not vault.is_dir():
        print(f"ERROR vault does not exist: {vault}")
        return 2

    notes = visible_files(vault, ".md")
    texts = {note: read_text(note) for note in notes}
    by_relative, by_stem = build_note_index(vault, notes)
    titles: dict[str, list[Path]] = defaultdict(list)
    graph: dict[Path, set[Path]] = defaultdict(set)
    broken_links: list[tuple[Path, str]] = []
    ambiguous_links: list[tuple[Path, str]] = []
    missing_assets: list[tuple[Path, str]] = []
    empty_concepts: list[Path] = []
    noise_hits: list[tuple[Path, int, str]] = []

    noise_re = re.compile("|".join(f"(?:{p})" for p in (*DEFAULT_NOISE, *args.noise_pattern)), re.I)

    for note, text in texts.items():
        titles[note_title(text, note).casefold()].append(note)
        type_match = TYPE_RE.search(text)
        if type_match and type_match.group(1).strip().casefold() in {"concept", "knowledge-point"}:
            if len(normalized_body(text)) < args.min_concept_chars:
                empty_concepts.append(note)

        for line_number, line in enumerate(text.splitlines(), 1):
            if noise_re.search(line):
                noise_hits.append((note, line_number, line.strip()[:120]))

        for embed, raw_target in WIKILINK_RE.findall(text):
            target = clean_target(raw_target)
            if not target:
                continue
            suffix = Path(target).suffix.lower()
            if suffix and suffix != ".md":
                asset = (vault / target).resolve()
                if not asset.is_file():
                    candidates = [
                        p for p in visible_files(vault) if p.name.casefold() == Path(target).name.casefold()
                    ]
                    if len(candidates) != 1:
                        missing_assets.append((note, target))
                continue
            resolved, ambiguous = resolve_note(target, by_relative, by_stem)
            if resolved:
                graph[note].add(resolved)
            elif ambiguous:
                ambiguous_links.append((note, target))
            else:
                broken_links.append((note, target))

        for raw_target in MARKDOWN_LINK_RE.findall(text):
            target = clean_target(raw_target)
            if not target or re.match(r"^[a-z]+://", target, re.I) or target.startswith("#"):
                continue
            asset = (note.parent / target).resolve()
            if not asset.exists():
                missing_assets.append((note, target))

    duplicate_titles = {title: paths for title, paths in titles.items() if len(paths) > 1}

    unreachable: list[Path] = []
    if args.root_note:
        root_note, ambiguous = resolve_note(args.root_note, by_relative, by_stem)
        if root_note is None:
            print(f"ERROR root note unresolved: {args.root_note} (ambiguous={ambiguous})")
            return 2
        reached = {root_note}
        queue = deque([root_note])
        while queue:
            current = queue.popleft()
            for linked in graph[current]:
                if linked not in reached:
                    reached.add(linked)
                    queue.append(linked)
        unreachable = sorted(set(notes) - reached)

    source_total = source_unique = source_duplicates = uncovered_sources = 0
    if args.source_root:
        source_root = args.source_root.resolve()
        if not source_root.is_dir():
            print(f"ERROR source root does not exist: {source_root}")
            return 2
        sources = [p for p in visible_files(source_root) if p.suffix.lower() in SOURCE_SUFFIXES]
        groups: dict[str, list[Path]] = defaultdict(list)
        for source in sources:
            groups[sha256(source)].append(source)
        corpus = "\n".join(texts.values()).casefold()
        source_total = len(sources)
        source_unique = len(groups)
        source_duplicates = source_total - source_unique
        for group in groups.values():
            names = {path.name.casefold() for path in group} | {path.stem.casefold() for path in group}
            if not any(name in corpus for name in names):
                uncovered_sources += 1

    def rel(path: Path) -> str:
        return path.relative_to(vault).as_posix()

    print(
        "SUMMARY "
        f"notes={len(notes)} broken_links={len(broken_links)} "
        f"ambiguous_links={len(ambiguous_links)} missing_assets={len(missing_assets)} "
        f"duplicate_titles={len(duplicate_titles)} noise_hits={len(noise_hits)} "
        f"empty_concepts={len(empty_concepts)} unreachable={len(unreachable)}"
    )
    if args.source_root:
        print(
            "SOURCES "
            f"physical={source_total} unique={source_unique} "
            f"duplicates={source_duplicates} uncovered_unique={uncovered_sources}"
        )

    details = (
        ("BROKEN_LINK", ((rel(path), target) for path, target in broken_links)),
        ("AMBIGUOUS_LINK", ((rel(path), target) for path, target in ambiguous_links)),
        ("MISSING_ASSET", ((rel(path), target) for path, target in missing_assets)),
        (
            "DUPLICATE_TITLE",
            ((title, ", ".join(rel(p) for p in paths)) for title, paths in duplicate_titles.items()),
        ),
        ("NOISE", ((f"{rel(path)}:{line}", snippet) for path, line, snippet in noise_hits)),
        ("EMPTY_CONCEPT", ((rel(path), "") for path in empty_concepts)),
        ("UNREACHABLE", ((rel(path), "") for path in unreachable)),
    )
    for label, rows in details:
        for left, right in rows:
            print(f"{label} {left}" + (f" -> {right}" if right else ""))

    failures = (
        len(broken_links)
        + len(missing_assets)
        + len(duplicate_titles)
        + len(noise_hits)
        + len(empty_concepts)
        + uncovered_sources
    )
    if args.strict:
        failures += len(ambiguous_links) + len(unreachable)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(audit())

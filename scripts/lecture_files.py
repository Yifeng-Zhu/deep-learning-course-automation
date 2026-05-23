"""Helpers for discovering lecture PowerPoint files."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


LECTURE_FILENAME_PATTERN = re.compile(
    r"^\s*(?P<lecture_number>\d+)\s*-\s*(?P<lecture_title>.+?)\s*\.pptx$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class LectureDeck:
    lecture_number: int
    lecture_title: str
    path: Path


def parse_lecture_deck(path: Path) -> LectureDeck:
    match = LECTURE_FILENAME_PATTERN.match(path.name)
    if match is None:
        raise ValueError(
            f"{path.name} does not match the expected pattern: "
            "{lecture_number} - {lecture_title}.pptx"
        )

    return LectureDeck(
        lecture_number=int(match.group("lecture_number")),
        lecture_title=match.group("lecture_title").strip(),
        path=path,
    )


def discover_lecture_decks(original_slides_folder: Path) -> list[LectureDeck]:
    if not original_slides_folder.is_dir():
        raise FileNotFoundError(f"Original slides folder not found: {original_slides_folder}")

    pptx_files = [
        path
        for path in original_slides_folder.glob("*.pptx")
        if not path.name.startswith("~$")
    ]

    invalid_names: list[str] = []
    decks: list[LectureDeck] = []

    for path in pptx_files:
        try:
            decks.append(parse_lecture_deck(path))
        except ValueError:
            invalid_names.append(path.name)

    if invalid_names:
        invalid_list = "\n".join(f"- {name}" for name in sorted(invalid_names))
        raise ValueError(
            "These PowerPoint filenames do not match "
            "{lecture_number} - {lecture_title}.pptx:\n"
            f"{invalid_list}"
        )

    lecture_numbers: dict[int, list[str]] = {}
    for deck in decks:
        lecture_numbers.setdefault(deck.lecture_number, []).append(deck.path.name)

    duplicates = {
        number: names for number, names in lecture_numbers.items() if len(names) > 1
    }
    if duplicates:
        duplicate_lines = [
            f"- lecture {number}: {', '.join(sorted(names))}"
            for number, names in sorted(duplicates.items())
        ]
        raise ValueError(
            "Multiple PowerPoint files use the same lecture number:\n"
            + "\n".join(duplicate_lines)
        )

    return sorted(
        decks,
        key=lambda deck: (deck.lecture_number, deck.lecture_title.lower(), deck.path.name.lower()),
    )

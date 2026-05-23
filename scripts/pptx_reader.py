"""Read PowerPoint decks, with a read-only fallback for damaged ZIP metadata."""

from __future__ import annotations

import contextlib
import struct
import tempfile
import zipfile
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator


LOCAL_FILE_HEADER = b"PK\x03\x04"


@dataclass(frozen=True)
class RepairInfo:
    used_repair: bool
    entries_written: int = 0
    skipped_entry: str = ""
    skipped_reason: str = ""


class PptxReadError(RuntimeError):
    """Raised when a PowerPoint deck cannot be read or repaired."""


def repair_local_zip_entries(source_path: Path, output_path: Path) -> RepairInfo:
    """Rebuild a PPTX ZIP using complete local entries from a damaged package.

    Some Office files can have intact local ZIP entries but a missing or damaged
    central directory. This creates a temporary valid ZIP package from complete
    entries only. The original file is never modified.
    """

    data = source_path.read_bytes()
    pos = 0
    entries_written = 0
    skipped_entry = ""
    skipped_reason = ""

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_out:
        while pos + 30 <= len(data) and data[pos : pos + 4] == LOCAL_FILE_HEADER:
            (
                _version_needed,
                _flag,
                compression_method,
                _modified_time,
                _modified_date,
                _crc32,
                compressed_size,
                _uncompressed_size,
                filename_length,
                extra_length,
            ) = struct.unpack_from("<HHHHHIIIHH", data, pos + 4)

            filename_start = pos + 30
            filename_end = filename_start + filename_length
            payload_start = filename_end + extra_length
            payload_end = payload_start + compressed_size
            filename = data[filename_start:filename_end].decode("utf-8", errors="replace")

            if payload_end > len(data):
                skipped_entry = filename
                skipped_reason = (
                    f"incomplete payload: expected {compressed_size} bytes, "
                    f"found {max(0, len(data) - payload_start)}"
                )
                break

            payload = data[payload_start:payload_end]
            try:
                if compression_method == 0:
                    content = payload
                elif compression_method == 8:
                    content = zlib.decompress(payload, -15)
                else:
                    skipped_entry = filename
                    skipped_reason = f"unsupported compression method {compression_method}"
                    break
            except zlib.error as exc:
                skipped_entry = filename
                skipped_reason = f"could not decompress payload: {exc}"
                break

            zip_out.writestr(filename, content)
            entries_written += 1
            pos = payload_end

    if entries_written == 0:
        raise PptxReadError(f"No recoverable ZIP entries found in {source_path}")

    with zipfile.ZipFile(output_path, "r") as zip_in:
        names = set(zip_in.namelist())

    required_entries = {"[Content_Types].xml", "ppt/presentation.xml"}
    missing_entries = sorted(required_entries - names)
    if missing_entries:
        raise PptxReadError(
            f"Repaired package is missing required entries: {', '.join(missing_entries)}"
        )

    return RepairInfo(
        used_repair=True,
        entries_written=entries_written,
        skipped_entry=skipped_entry,
        skipped_reason=skipped_reason,
    )


@contextlib.contextmanager
def open_presentation_readonly(source_path: Path) -> Iterator[tuple[object, RepairInfo]]:
    """Open a presentation without modifying the source file.

    If python-pptx cannot read the source directly, try a temporary repaired copy
    built from complete local ZIP entries.
    """

    try:
        from pptx import Presentation
    except ImportError as exc:
        raise PptxReadError(
            "python-pptx is required. Install dependencies with: pip install -r requirements.txt"
        ) from exc

    try:
        presentation = Presentation(str(source_path))
    except Exception as original_error:
        with tempfile.TemporaryDirectory(prefix="pptx_repair_") as temp_dir:
            repaired_path = Path(temp_dir) / source_path.name
            try:
                repair_info = repair_local_zip_entries(source_path, repaired_path)
                presentation = Presentation(str(repaired_path))
            except Exception as repair_error:
                raise PptxReadError(
                    f"Could not read {source_path.name}: {original_error}; "
                    f"temporary repair also failed: {repair_error}"
                ) from repair_error

            yield presentation, repair_info
    else:
        yield presentation, RepairInfo(used_repair=False)

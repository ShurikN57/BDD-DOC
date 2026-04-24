#!/usr/bin/env python3
from pathlib import Path
import argparse
import hashlib
import sys

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
IMP = ROOT / "vba_import"

TEXT_EXTS = {".bas", ".cls", ".frm", ".txt"}
BINARY_EXTS = {".frx"}


def normalize_text(data: bytes) -> str:
    if data.startswith(b"\xef\xbb\xbf"):
        data = data[3:]
    txt = data.decode("utf-8", errors="replace")
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    txt = txt.rstrip("\n")
    return txt


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def rel_files(folder: Path):
    return sorted(p.relative_to(folder) for p in folder.iterdir() if p.is_file())


def main() -> int:
    parser = argparse.ArgumentParser(description="Check sync drift between src/ and vba_import/.")
    parser.add_argument("--strict", action="store_true", help="With --enforce, fail also on BOM/EOL-only text diffs.")
    parser.add_argument("--enforce", action="store_true", help="Return a non-zero code when drift is detected.")
    parser.add_argument(
        "--fix-normalization",
        action="store_true",
        help="Auto-fix normalization-only diffs by copying src/ bytes to vba_import/.",
    )
    args = parser.parse_args()

    if not SRC.exists() or not IMP.exists():
        print("ERROR: src/ or vba_import/ is missing.")
        return 2

    src_files = rel_files(SRC)
    imp_files = rel_files(IMP)

    missing_in_import = [str(p) for p in src_files if p not in imp_files]
    missing_in_src = [str(p) for p in imp_files if p not in src_files]

    if missing_in_import:
        print("ERROR: files present in src/ but missing in vba_import/:")
        for p in missing_in_import:
            print(f"  - {p}")
    if missing_in_src:
        print("ERROR: files present in vba_import/ but missing in src/:")
        for p in missing_in_src:
            print(f"  - {p}")

    had_error = bool(missing_in_import or missing_in_src)
    had_warning = False

    for rel in sorted(set(src_files).intersection(imp_files)):
        a = (SRC / rel).read_bytes()
        b = (IMP / rel).read_bytes()
        ext = rel.suffix.lower()

        if ext in TEXT_EXTS:
            if a == b:
                continue
            if normalize_text(a) == normalize_text(b):
                had_warning = True
                print(f"WARN: normalization-only diff (BOM/EOL) for {rel}")
                if args.fix_normalization:
                    (IMP / rel).write_bytes(a)
                    print(f"FIX: normalized {rel} (vba_import <= src)")
                if args.strict and args.enforce:
                    had_error = True
            else:
                if args.enforce:
                    had_error = True
                print(f"ERROR: content drift for {rel}")
        elif ext in BINARY_EXTS:
            if digest_bytes(a) != digest_bytes(b):
                if args.enforce:
                    had_error = True
                print(f"ERROR: binary drift for {rel}")
        else:
            # Unknown extension: compare raw bytes conservatively.
            if a != b:
                if args.enforce:
                    had_error = True
                print(f"ERROR: drift on unsupported extension for {rel}")

    if had_error:
        return 1
    if args.enforce:
        print("OK: no blocking drift detected.")
        return 0

    print("Report generated (non-blocking mode). Use --enforce to fail on drift.")
    if had_warning:
        print("Note: normalization-only diffs (BOM/EOL) were detected.")
    else:
        print("Note: no normalization-only diffs detected.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

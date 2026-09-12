from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INPUT_DIR = PROJECT_ROOT / "data" / "listings_data"
DEFAULT_OUTPUT_FILE = PROJECT_ROOT / "data" / "listings.csv"


def _normalize_columns(columns: pd.Index) -> list[str]:
    return (
        columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .tolist()
    )


def load_listings_data(
    input_dir: Path = DEFAULT_INPUT_DIR,
    output_file: Path = DEFAULT_OUTPUT_FILE,
) -> pd.DataFrame:
    files = sorted(input_dir.glob("listings-*.csv"))
    if not files:
        raise FileNotFoundError(f"Nenhum CSV encontrado em: {input_dir}")

    frames: list[pd.DataFrame] = []
    for file in files:
        frame = pd.read_csv(file, encoding="utf-8-sig")
        frame = frame.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if frame.empty:
            continue

        frame.columns = _normalize_columns(frame.columns)
        frame["source_file"] = file.name
        frames.append(frame)

    if not frames:
        raise ValueError(f"Nenhum dado encontrado em: {input_dir}")

    listings = pd.concat(frames, ignore_index=True, sort=False)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    listings.to_csv(output_file, index=False, encoding="utf-8-sig")
    return listings


def main() -> None:
    parser = argparse.ArgumentParser(description="Consolida os listings em CSV.")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR)
    parser.add_argument("--output-file", type=Path, default=DEFAULT_OUTPUT_FILE)
    args = parser.parse_args()

    listings = load_listings_data(
        input_dir=args.input_dir,
        output_file=args.output_file,
    )
    print(f"CSV gerado: {args.output_file} ({len(listings):,} linhas, {len(listings.columns)} colunas)")
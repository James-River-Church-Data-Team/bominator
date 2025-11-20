import argparse
import os
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path


BOM = b"\xef\xbb\xbf"

class Args(argparse.Namespace):
	csv_path: Path
	exe_path: Path


def parse_args(argv: Sequence[str]) -> Args:
	parser = argparse.ArgumentParser()
	parser.add_argument(
		"csv_path",
		type=Path,
	)
	parser.add_argument(
		"--exe_path",
		type=Path,
		default=r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
	)
	return parser.parse_args(argv[1:], Args())


def upsert_bom(csv_path: Path) -> None:
	with csv_path.open("rb+") as f:
		if f.read(3) == BOM:
			return
		f.seek(os.SEEK_SET)
		data = f.read()
		f.seek(os.SEEK_SET)
		f.write(BOM)
		f.write(data)


def main() -> None:
	args = parse_args(sys.argv)
	upsert_bom(args.csv_path)
	subprocess.Popen([args.exe_path.absolute(), args.csv_path.absolute()])


if __name__ == "__main__":
	main()

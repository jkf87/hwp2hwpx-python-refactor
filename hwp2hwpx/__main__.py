"""CLI entry point for hwp2hwpx converter."""

import sys
import os
import argparse


def main():
    parser = argparse.ArgumentParser(description="Convert HWP files to HWPX format")
    parser.add_argument("input", help="Input .hwp file path")
    parser.add_argument("-o", "--output", help="Output .hwpx file path (default: same name)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    args = parser.parse_args()

    from .converter import convert_file

    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    try:
        output = convert_file(args.input, args.output)
        print(f"Converted: {args.input} -> {output}")
        if args.verbose:
            print(f"  Output size: {os.path.getsize(output)} bytes")
    except Exception as e:
        print(f"Error converting {args.input}: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

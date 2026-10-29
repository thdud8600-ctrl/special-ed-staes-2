#!/bin/sh
# macOS: 더블클릭으로 실행
cd "$(dirname "$0")"
python3 compare.py --open "$@"

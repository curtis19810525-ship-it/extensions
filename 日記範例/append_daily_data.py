# -*- coding: utf-8 -*-
"""與 append_nba_journal.py 同一支寫入程式。"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from append_nba_journal import main

if __name__ == "__main__":
    main()

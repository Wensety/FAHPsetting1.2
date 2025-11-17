#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""測試模組導入"""

import os
import sys

print(f"Python 版本: {sys.version}")
print(f"當前工作目錄: {os.getcwd()}")
print(f"Python 路徑: {sys.path[:3]}...")

try:
    from fahp_analysis import FAHPAnalyzer, __version__
    print(f"✓ 成功導入 fahp_analysis，版本: {__version__}")
except ImportError as e:
    print(f"✗ 無法導入 fahp_analysis: {e}")
    sys.exit(1)

try:
    import pandas as pd
    print(f"✓ 成功導入 pandas，版本: {pd.__version__}")
except ImportError as e:
    print(f"✗ 無法導入 pandas: {e}")
    sys.exit(1)

try:
    import numpy as np
    print(f"✓ 成功導入 numpy，版本: {np.__version__}")
except ImportError as e:
    print(f"✗ 無法導入 numpy: {e}")
    sys.exit(1)

print("\n所有模組導入成功！")


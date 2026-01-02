import pandas as pd
import re

def extract_grade(val):
    s = '' if pd.isna(val) else str(val)
    m = re.search(r":\s*(\d+)", s)
    if m:
        digits = m.group(1)
        try:
            return int(digits[0])
        except Exception:
            return None

    all_digits = re.findall(r"(\d+)", s)
    if all_digits:
        longest = max(all_digits, key=len)
        try:
            return int(longest[0])
        except Exception:
            return None

    try:
        return int(s[0])
    except Exception:
        return None

samples = ['I25: 13645554','I25: 136455','21008고진우','1234567','X:ABC','']
for s in samples:
    print(s, '->', extract_grade(s))

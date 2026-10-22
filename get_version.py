import re
print(re.search(r'APP_VERSION\s*=\s*"([\d.]+)"', open("invoice_printer.py", encoding="utf-8").read()).group(1))

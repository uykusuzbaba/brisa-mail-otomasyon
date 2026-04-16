import os
from pathlib import Path
Path("brisa_mail.db").unlink(missing_ok=True)
print("DB silindi!")

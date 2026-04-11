from __future__ import annotations

import os

os.environ["APP_MODE"] = "expert"

import app


if __name__ == "__main__":
    app.main("expert")

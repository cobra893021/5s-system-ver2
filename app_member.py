from __future__ import annotations

import os

os.environ["APP_MODE"] = "member"

import app


if __name__ == "__main__":
    app.main("member")

"""Start the HTTP service with ``python -m dfine_la``."""

import os

import uvicorn


if __name__ == "__main__":
    uvicorn.run(
        "dfine_la.api:app",
        host="0.0.0.0",
        port=int(os.getenv("PORT", "8000")),
        workers=1,
        log_level=os.getenv("LOG_LEVEL", "info"),
    )

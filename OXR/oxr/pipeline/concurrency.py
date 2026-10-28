import asyncio
from collections.abc import Awaitable
from typing import Any


async def gather_cancel_on_error(*awaitables: Awaitable[Any]) -> list[Any]:
    """Gather awaitables, cancelling and draining every sibling after a failure."""
    tasks = [asyncio.ensure_future(awaitable) for awaitable in awaitables]
    try:
        return list(await asyncio.gather(*tasks))
    except BaseException:
        for task in tasks:
            if not task.done():
                task.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)
        raise

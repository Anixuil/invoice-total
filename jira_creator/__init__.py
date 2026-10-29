"""Account-scoped Jira creation, isolated from the workbook processors."""

from contextlib import asynccontextmanager


def install(app):
    # Optional dependencies/configuration must not prevent existing tools starting.
    app.state.jira_creator = None
    app.state.jira_creator_error = "Jira 批量建单尚未初始化"
    try:
        from .api import router, start, stop
    except Exception:
        from fastapi.responses import JSONResponse
        app.state.jira_creator_error = "Jira 批量建单模块未加载，请更新 requirements.txt 中的依赖并检查模块文件"

        async def unavailable(path: str):
            return JSONResponse({"detail": app.state.jira_creator_error}, status_code=503,
                                headers={"Cache-Control": "no-store"})

        app.add_api_route("/api/jira-create/{path:path}", unavailable,
                          methods=["GET", "POST", "PUT", "DELETE"], include_in_schema=False)
        return
    app.include_router(router)

    previous_lifespan = app.router.lifespan_context

    @asynccontextmanager
    async def lifespan(application):
        # Preserve existing application/router startup, shutdown and lifespan state.
        async with previous_lifespan(application) as state:
            try:
                await start(application)
                yield state
            finally:
                await stop(application)

    app.router.lifespan_context = lifespan

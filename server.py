"""Alaamri MS Project MCP: a local MCP server that drives the installed Microsoft Project desktop app via COM.

All COM calls run on one dedicated STA thread, because Project's COM objects
must not cross threads.

Developed by Dr Ibrahim M.A. Alaamri. Licensed under the MIT License.
"""

import asyncio
import datetime as dt
import json
from concurrent.futures import ThreadPoolExecutor

import pythoncom
import win32com.client
from mcp.server.mcpserver import MCPServer

__version__ = "0.1.0"
AUTHOR = "Dr Ibrahim M.A. Alaamri"
LINKEDIN_URL = "https://www.linkedin.com/in/dribrahim-alameri/"
REPO_URL = "https://github.com/engibrrahem-wq/alaamri-msproject-mcp"

_executor = ThreadPoolExecutor(max_workers=1, initializer=pythoncom.CoInitialize)
mcp = MCPServer(
    "alaamri-msproject",
    title="Alaamri MS Project MCP",
    website_url=LINKEDIN_URL,
    version=__version__,
    instructions=(
        "Controls Microsoft Project on this PC. Durations are strings like '5d' or '3h'. "
        "Dates are ISO 'YYYY-MM-DD'. Task ids are the visible row IDs. "
        "Always work on a copy of the client's schedule and save_project to a new path. "
        f"Developed by {AUTHOR} — {LINKEDIN_URL}"
    ),
)


async def _run(fn, *args):
    return await asyncio.get_running_loop().run_in_executor(_executor, fn, *args)


def _app(launch=True):
    try:
        return win32com.client.GetActiveObject("MSProject.Application")
    except pythoncom.com_error:
        if not launch:
            return None
        app = win32com.client.Dispatch("MSProject.Application")
        app.Visible = True
        return app


def _proj():
    app = _app()
    if app.Projects.Count == 0:
        raise RuntimeError("No project is open in Microsoft Project. Use open_project or new_project.")
    return app.ActiveProject


def _date(v):
    if v is None or isinstance(v, str):
        return v
    try:
        if getattr(v, "tzinfo", None) is not None:
            v = v.astimezone()
        return dt.datetime(v.year, v.month, v.day, v.hour, v.minute).isoformat(timespec="minutes")
    except Exception:
        return str(v)


def _parse(s):
    # pywin32 treats naive datetimes as UTC; pin to local time so dates don't shift
    return dt.datetime.fromisoformat(s).astimezone()


def _days(proj, minutes):
    return round(minutes / (proj.HoursPerDay * 60), 2)


def _task_dict(proj, t):
    return {
        "id": t.ID,
        "uid": t.UniqueID,
        "name": t.Name,
        "outline_level": t.OutlineLevel,
        "wbs": t.WBS,
        "summary": bool(t.Summary),
        "milestone": bool(t.Milestone),
        "duration_days": _days(proj, t.Duration),
        "start": _date(t.Start),
        "finish": _date(t.Finish),
        "percent_complete": t.PercentComplete,
        "predecessors": t.Predecessors,
        "resources": t.ResourceNames,
        "critical": bool(t.Critical),
        "total_slack_days": _days(proj, t.TotalSlack),
    }


def _find(proj, task_id):
    for t in proj.Tasks:
        if t is not None and t.ID == task_id:
            return t
    raise ValueError(f"Task id {task_id} not found")


def _apply(t, f):
    if "name" in f:
        t.Name = f["name"]
    if "duration" in f:
        t.Duration = f["duration"]
    if "start" in f:
        t.Start = _parse(f["start"])
    if "predecessors" in f:
        t.Predecessors = f["predecessors"]
    if "resources" in f:
        t.ResourceNames = f["resources"]
    if "percent_complete" in f:
        t.PercentComplete = int(f["percent_complete"])
    if "outline_level" in f:
        t.OutlineLevel = int(f["outline_level"])
    if "notes" in f:
        t.Notes = f["notes"]


def _stamp_comments(app, p):
    # Only the Comments property is set; Author/Company stay with the file owner.
    try:
        p.BuiltinDocumentProperties("Comments").Value = STAMP
        return
    except Exception:
        pass
    try:
        app.ProjectSummaryInfo(Comments=STAMP)
    except Exception:
        pass


STAMP = "Created with Alaamri MS Project MCP"


@mcp.tool()
async def about() -> str:
    """Information about this MCP server: name, version, author, license and links."""
    return json.dumps({
        "name": "alaamri-msproject-mcp",
        "version": __version__,
        "author": AUTHOR,
        "linkedin": LINKEDIN_URL,
        "license": "MIT",
        "repository": REPO_URL,
    }, ensure_ascii=False)


@mcp.tool()
async def status() -> str:
    """Check whether Microsoft Project is running and which projects are open."""
    def go():
        app = _app(launch=False)
        if app is None:
            return {"running": False}
        projs = [{"name": p.Name, "path": p.FullName, "tasks": p.Tasks.Count} for p in app.Projects]
        active = app.ActiveProject.Name if app.Projects.Count else None
        return {"running": True, "version": app.Version, "active": active, "open_projects": projs}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def open_project(path: str) -> str:
    """Open an .mpp/.xml/.xlsx schedule file in Microsoft Project (launches Project if needed)."""
    def go():
        app = _app()
        app.FileOpenEx(Name=path)
        p = app.ActiveProject
        return {"opened": p.FullName, "tasks": p.Tasks.Count, "start": _date(p.ProjectStart), "finish": _date(p.ProjectFinish)}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def new_project(start_date: str = "", title: str = "") -> str:
    """Create a new blank project. start_date is ISO YYYY-MM-DD."""
    def go():
        app = _app()
        app.FileNew()
        p = app.ActiveProject
        if start_date:
            p.ProjectStart = _parse(start_date)
        if title:
            p.Title = title
        _stamp_comments(app, p)
        return {"created": p.Name, "start": _date(p.ProjectStart)}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def project_summary() -> str:
    """Summary of the active project: dates, task counts, % complete, critical task count."""
    def go():
        p = _proj()
        tasks = [t for t in p.Tasks if t is not None]
        return {
            "name": p.Name, "path": p.FullName,
            "start": _date(p.ProjectStart), "finish": _date(p.ProjectFinish),
            "status_date": _date(p.StatusDate),
            "hours_per_day": p.HoursPerDay,
            "task_count": len(tasks),
            "summary_tasks": sum(1 for t in tasks if t.Summary),
            "milestones": sum(1 for t in tasks if t.Milestone),
            "critical_tasks": sum(1 for t in tasks if t.Critical and not t.Summary),
            "percent_complete": p.ProjectSummaryTask.PercentComplete,
            "resources": [r.Name for r in p.Resources if r is not None],
        }
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def list_tasks(offset: int = 0, limit: int = 200, critical_only: bool = False) -> str:
    """List tasks of the active project (id, WBS, name, duration, dates, predecessors, resources, %, slack)."""
    def go():
        p = _proj()
        rows = [_task_dict(p, t) for t in p.Tasks if t is not None]
        if critical_only:
            rows = [r for r in rows if r["critical"] and not r["summary"]]
        return {"total": len(rows), "tasks": rows[offset:offset + limit]}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def add_tasks(tasks_json: str) -> str:
    """Append tasks. tasks_json is a JSON array of objects with keys:
    name (required), duration ('5d'), start ('YYYY-MM-DD'), predecessors ('3,4FS+2d'),
    resources ('Mason[2],Crane'), outline_level (1=top), percent_complete, notes.
    Tasks are added in order, so predecessors may reference ids created earlier in the same call."""
    def go():
        p = _proj()
        added = []
        for f in json.loads(tasks_json):
            t = p.Tasks.Add(f["name"])
            _apply(t, {k: v for k, v in f.items() if k != "name"})
            added.append({"id": t.ID, "name": t.Name})
        return {"added": added}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def update_task(task_id: int, fields_json: str) -> str:
    """Update one task by id. fields_json keys: name, duration, start, predecessors, resources,
    outline_level, percent_complete, notes."""
    def go():
        p = _proj()
        t = _find(p, task_id)
        _apply(t, json.loads(fields_json))
        return _task_dict(p, t)
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def delete_task(task_id: int) -> str:
    """Delete one task by id (ids below it shift up)."""
    def go():
        t = _find(_proj(), task_id)
        name = t.Name
        t.Delete()
        return {"deleted": task_id, "name": name}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def set_status_date(date: str) -> str:
    """Set the project status date (ISO YYYY-MM-DD) for progress tracking."""
    def go():
        p = _proj()
        p.StatusDate = _parse(date)
        return {"status_date": _date(p.StatusDate)}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def save_baseline() -> str:
    """Save Baseline (Baseline 0) for the whole active project."""
    def go():
        app = _app()
        app.BaselineSave(All=True)
        return {"baseline_saved": app.ActiveProject.Name}
    return json.dumps(await _run(go), ensure_ascii=False)


@mcp.tool()
async def save_project(path: str = "") -> str:
    """Save the active project. With path, Save As (.mpp, or .xml for MSPDI). Without path, save in place."""
    def go():
        app = _app()
        if path:
            fmt = "MSProject.XML" if path.lower().endswith(".xml") else ""
            app.FileSaveAs(Name=path, FormatID=fmt) if fmt else app.FileSaveAs(Name=path)
        else:
            app.FileSave()
        return {"saved": app.ActiveProject.FullName}
    return json.dumps(await _run(go), ensure_ascii=False)


def main():
    mcp.run()


if __name__ == "__main__":
    main()

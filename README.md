# Alaamri MS Project MCP

An [MCP](https://modelcontextprotocol.io) server that lets Claude (or any MCP client) drive the **Microsoft Project desktop app** on Windows through COM automation: open and create schedules, read and edit tasks, set the status date, save a baseline and save the file.

All COM calls run on a single dedicated STA thread, because Project's COM objects must not cross threads.

## Requirements

- Windows
- Microsoft Project desktop (installed and licensed)
- Python 3.10 or newer

> **بالعربي:** [دليل التنصيب خطوة بخطوة](docs/ar/INSTALL-AR.md) · [البرومبت المستخدم في الفيديو](prompts/villa-schedule-prompt-ar.txt) · تنصيب بنقرة واحدة: `install.bat`

## Install

**Quick (Windows):** download the ZIP (Code → Download ZIP), unzip it, and double-click `install.bat`. It creates the virtual environment, installs the server and prints the Claude Desktop config with your path filled in.

**Manual:**

```powershell
git clone https://github.com/engibrrahem-wq/alaamri-msproject-mcp.git
cd alaamri-msproject-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

## Register the server

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json` (see `examples/claude_desktop_config.json`) and replace the paths with your clone location:

```json
{
  "mcpServers": {
    "alaamri-msproject": {
      "command": "C:\\path\\to\\alaamri-msproject-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\alaamri-msproject-mcp\\server.py"]
    }
  }
}
```

Restart Claude Desktop afterwards.

### Claude Code

Add it to a project's `.mcp.json` (same JSON shape as above), or from the command line:

```powershell
claude mcp add alaamri-msproject -- C:\path\to\alaamri-msproject-mcp\.venv\Scripts\python.exe C:\path\to\alaamri-msproject-mcp\server.py
```

After `pip install -e .` you can also use the console script `C:\path\to\alaamri-msproject-mcp\.venv\Scripts\alaamri-msproject-mcp.exe` as the command with no args.

## Tools

| Tool | What it does |
|---|---|
| `about` | Server name, version, author, license and links |
| `status` | Whether Project is running, its version, and the open projects |
| `open_project` | Open an `.mpp` / `.xml` / `.xlsx` schedule (launches Project if needed) |
| `new_project` | Create a blank project with optional start date and title |
| `project_summary` | Dates, task counts, milestones, critical tasks, % complete, resources |
| `list_tasks` | Tasks with ID, WBS, name, duration, dates, predecessors, resources, %, slack (paged, optional critical-only filter) |
| `add_tasks` | Append tasks from a JSON array (name, duration, start, predecessors, resources, outline level, %, notes) |
| `update_task` | Update one task's fields by ID |
| `delete_task` | Delete one task by ID |
| `set_status_date` | Set the project status date |
| `save_baseline` | Save Baseline 0 for the whole project |
| `save_project` | Save in place, or Save As `.mpp` / `.xml` (MSPDI) |

Conventions: durations are strings like `5d` or `3h`, dates are ISO `YYYY-MM-DD`, task IDs are the visible row IDs. New projects get the Comments property "Created with Alaamri MS Project MCP"; Author and Company are left untouched.

## Safety

The server edits live schedules in Project. **Always work on a copy** of a real schedule and use `save_project` with a new path. There is no undo through MCP, and `delete_task` shifts the IDs below the deleted row.

## Author

Developed by **Dr Ibrahim M.A. Alaamri**: [LinkedIn](https://www.linkedin.com/in/dribrahim-alameri/)

## License

MIT, see [LICENSE](LICENSE).

---

<div dir="rtl">

## بالعربية

**Alaamri MS Project MCP** خادم MCP يتيح لـ Claude (أو أي عميل MCP) التحكم في برنامج **Microsoft Project** المثبّت على ويندوز عبر COM: فتح الجداول الزمنية وإنشاؤها، وقراءة الأنشطة وتعديلها، وتحديد تاريخ الحالة (Status Date)، وحفظ خط الأساس (Baseline)، وحفظ الملف.

### المتطلبات
- نظام Windows
- برنامج Microsoft Project لسطح المكتب (مثبّت ومفعّل)
- Python بإصدار 3.10 أو أحدث

### التثبيت
```powershell
git clone https://github.com/engibrrahem-wq/alaamri-msproject-mcp.git
cd alaamri-msproject-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

### التسجيل في Claude
- **Claude Desktop:** أضف الإعداد الموجود في `examples/claude_desktop_config.json` إلى الملف `%APPDATA%\Claude\claude_desktop_config.json` بعد تعديل المسارات، ثم أعد تشغيل البرنامج.
- **Claude Code:** أضف الإعداد نفسه إلى ملف `.mcp.json` في مشروعك، أو استخدم الأمر `claude mcp add` الموضّح في القسم الإنجليزي.

### الأدوات
`about` (معلومات الخادم)، و`status` (حالة البرنامج والملفات المفتوحة)، و`open_project` (فتح جدول)، و`new_project` (مشروع جديد)، و`project_summary` (ملخص المشروع)، و`list_tasks` (قائمة الأنشطة)، و`add_tasks` (إضافة أنشطة)، و`update_task` (تعديل نشاط)، و`delete_task` (حذف نشاط)، و`set_status_date` (تاريخ الحالة)، و`save_baseline` (حفظ خط الأساس)، و`save_project` (حفظ الملف).

### تنبيه
الخادم يعدّل الجدول المفتوح مباشرة، فاعمل دائماً على **نسخة** من جدول المشروع واحفظها بمسار جديد عبر `save_project`. لا يوجد تراجع (Undo) عبر MCP.

### المطوّر
تطوير **د. إبراهيم م. أ. العامري** (Dr Ibrahim M.A. Alaamri): [LinkedIn](https://www.linkedin.com/in/dribrahim-alameri/)

### الترخيص
MIT، راجع ملف [LICENSE](LICENSE).

</div>

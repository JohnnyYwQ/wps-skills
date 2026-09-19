<div align="center">

# WPS Skills

[简体中文](README.md) · **English**

**Let your Agent get work done in your local WPS applications**

Turn natural-language requests into editable Word documents, Excel spreadsheets, and PPT presentations.

<p>
  <code>Codex</code> &nbsp; <code>Claude Code</code> &nbsp; <code>Windows WPS</code> &nbsp; <code>Python 3.10+</code>
</p>

[Features](#features) · [Quick start](#quick-start) · [Examples](#examples) · [Installation guide (中文)](INSTALL.md)

</div>

## Features

One plugin, three independently usable Skills, and **85 operations (Actions)**. Create, read, edit, format, save, and export documents. Each application Skill can also be installed separately.

| Application | What you can do | Save and export |
| --- | --- | --- |
| 📄 **[Word](src/main/resources/skills/wps-word/SKILL.md)** · 14 Actions | Write content, find and replace text, adjust formatting, insert tables and images, and configure headers, footers, and page layout | DOCX · PDF |
| 📊 **[Excel](src/main/resources/skills/wps-excel/SKILL.md)** · 34 Actions | Manage worksheets, read and write data, calculate formulas, format ranges, adjust rows and columns, sort, and filter | XLSX · PDF |
| 📽️ **[PPT](src/main/resources/skills/wps-ppt/SKILL.md)** · 37 Actions | Organize slides, edit text and shapes, insert images and tables, and adjust layouts and speaker notes | PPTX · PDF · PNG |

The Agent interprets your request and plans the operations. The executor works with real documents through Windows WPS COM and returns the results. The runtime uses the Python standard library and Windows PowerShell, with **no third-party Python dependencies**.

## Quick start

**Requirements:** Windows, Python 3.10+, Windows PowerShell 5.1, and the relevant WPS applications installed and working. Document operations must run in a logged-in interactive desktop session. Your Agent host must support plugins, file writes, and local script execution.

### 1. Build and extract

Run this command from the repository root:

```sh
python scripts/build/plugin.py
```

Extract the generated `build/plugins/wps-skills.zip` to a permanent location, such as `C:/Tools/wps-skills`. Keep all directories, including those whose names start with a dot:

```text
C:/Tools/wps-skills/
├── .agents/           Codex marketplace catalog
├── .claude-plugin/    Claude Code marketplace catalog
├── plugins/
│   └── wps-skills/    Plugin and three complete Skills
└── README.md          Installation instructions
```

### 2. Install in your Agent host

Run the commands for your host in a Windows terminal. Replace the path with your actual extraction directory.

<details open>
<summary><strong>Codex</strong></summary>

```text
codex plugin marketplace add "C:/Tools/wps-skills"
codex plugin add wps-skills@wps-skills-local
```

</details>

<details open>
<summary><strong>Claude Code</strong></summary>

```text
claude plugin marketplace add "C:/Tools/wps-skills"
claude plugin install wps-skills@wps-skills-local --scope user
```

</details>

### 3. Start using it

Start a new conversation, select `wps-word`, `wps-excel`, or `wps-ppt`, and describe your goal, content, and output location. In Claude Code, you can also select a Skill with these commands:

```text
/wps-skills:wps-word
/wps-skills:wps-excel
/wps-skills:wps-ppt
```

For updates, temporary loading, and standalone Skill installation, see the **[installation guide (中文) →](INSTALL.md)**.

## Examples

**📄 Word · Project kickoff notice**

> Create a kickoff notice for the “Galaxy Knowledge Base” project. Include the project goals, participating departments, and a four-week implementation schedule. Format it neatly and save it to `C:/Documents/project-notice.docx`.

**📊 Excel · Project budget**

> Create a project budget with personnel, equipment, and service costs. Calculate totals with formulas, apply currency formatting, and save it to `C:/Documents/project-budget.xlsx`.

**📽️ PPT · Project update**

> Create a four-slide project update covering the background, implementation plan, current progress, and next steps. Save it to `C:/Documents/project-update.pptx`.

Use an actual absolute path on the local machine; its parent directory must already exist. **Documents are saved only when explicitly requested.** Exporting a PDF or PNG does not save the source document.

## How tasks work

**Understand the request → Query the needed Action definitions → Submit a multi-step plan → Operate WPS and verify → Report results**

| Design | Purpose |
| --- | --- |
| **Definitions on demand** | Read the Skill to understand available operations, then query only the parameters and result structures needed for the current plan, reducing unrelated context. |
| **Multi-step Task execution** | Express dependencies through result references, pass intermediate results within a single submission, and reuse the document binding and bridge process. |
| **Exact document binding** | Target a specific document and check relevant content revisions or tokens to reduce the risk of editing the wrong window or using stale locations. |
| **Read-back verification and receipts** | Verify effects against each Action's contract and record step results. Query receipts after a lost response; release the Task's execution resources when it ends. |

Each Task handles one document in one application. The Agent submits separate Tasks for work spanning applications or when it needs to read results before deciding what to do next. See each Skill's instructions and Action definitions for parameters and capability limits.

## Advanced usage

<details>
<summary><strong>Query definitions, submit a Task, and retrieve its receipt manually</strong></summary>

The following example uses the extracted Word Skill. In a Windows terminal, enter its script directory and run the remaining commands from there:

```text
cd "C:/Tools/wps-skills/plugins/wps-skills/skills/wps-word/scripts"
```

**Query definitions** to retrieve input parameters and result structures without starting WPS:

```text
python schema.py createDocument writeContent saveAs
```

**Submit a Task.** Prepare a UTF-8 JSON file following the Skill's instructions, using a new path for each new request:

```text
python word.py --app word --task-file "C:/Tasks/request-001.json"
```

**Retrieve the receipt** using the original request path if command output is lost or the command times out:

```text
python word.py --app word --task-status-file "C:/Tasks/request-001.json"
```

Use the original path even if the input file has already been consumed. Resubmitting an accepted path does not replay operations. A failed or uncertain outcome stops subsequent steps; effects already applied remain, with no automatic rollback or retry.

For Excel and PPT, use `schema.py` and `excel.py` or `ppt.py` in the respective Skill directory, with the corresponding `--app` value. Request formats and complete examples are in each package's `SKILL.md`.

</details>

<details>
<summary><strong>Check the Windows environment on another machine</strong></summary>

On Windows, run doctor from the repository root:

```sh
python scripts/doctor.py --app word
python scripts/doctor.py --app excel
python scripts/doctor.py --app ppt
```

By default, doctor checks prerequisites such as the desktop session, COM registration, PowerShell, and Chinese text transport, without starting WPS or modifying documents. Passing these checks does not guarantee that every document operation is available. Run `python scripts/doctor.py --help` for more options.

</details>

<p align="center">
  <a href="INSTALL.md">Installation guide (中文)</a> ·
  <a href="LICENSE">MIT License</a>
</p>

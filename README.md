# LocalRPA Orchestrator

A lightweight orchestrator that keeps logic in Python and delegates UI work to RPA tools.

---

## Overview

This project is a mini RPA orchestrator written in Python.

It is designed as a lightweight alternative where enterprise orchestrators would be unnecessary overhead,
focusing on clarity, ease of modification, and running on a single machine.

It does NOT replace RPA tools.

Instead, it orchestrates them:
you still need a real RPA tool (Power Automate, UiPath Studio, Blue Prism, etc.)
to perform screen-based automation. Think of this as the control around the RPA — not the robot itself.

---
## Responsibilities

| Step | Orchestrator (this project) | RPA tool |
| ---- |:--------------------------:|:----------:|
| Job intake| ✓ |  |
| Handover | → | ✓ |
| Job execution |  | ✓ |
| Handover | ✓ | ← |
| Job verification | ✓ |  |


---

## Typical examples

### Email-driven job
A user sends an email → Python validates and prepares the job → writes to `handover.json` → RPA executes UI actions → Python verifies and responds.

### Data-driven job
Python polls a data source → detects a valid case → prepares payload → signals RPA → RPA executes → Python verifies the outcome.


Example dashboard during runtime:
<img width="1209" height="635" alt="example_dash" src="https://github.com/user-attachments/assets/dc12a84b-c329-4b91-b402-387128197f9a" />

---

## Key Idea
This project separates responsibilities between the Orchestrator and the RPA tool:

* The **Orchestrator (this project)** handles:
  - how jobs enter the system (e.g. via email)
  - how jobs are discovered autonomously (e.g. queries or data sources)
  - access control and validation
  - job state tracking and audit logging
  - preparing payloads and handing work over to an RPA tool
  - verifying results after execution
  - handling failures and entering a controlled safestop state

* The **RPA tool** handles:
  - UI automation (clicks, keyboard input, ERP/UI interaction)

They communicate through a file-based IPC mechanism (`handover.json`).

---

## What this project is not

You still need a real RPA tool to execute UI automation steps.

This includes tools such as:

* Microsoft Power Automate
* UiPath Studio
* Blue Prism
* Robot Framework
* TagUI
* RPA for Python

These tools perform the actual UI interactions (clicks, keyboard input, screen automation).

---

## Architecture

<img width="1140" height="1709" alt="workflow" src="https://github.com/user-attachments/assets/9e9b1135-76c9-40d6-9f7f-785cfbde715d" />

The diagram shows how:

* The Orchestrator and the RPA tool run independently
* Both operate in their own loops
* State is synchronized via `handover.json`
* Failures transition the system into safestop
* Your RPA tool must follow this model


## Features

* Email-driven job processing (personal inbox)
* Shared inbox support (extensible)
* Data-driven jobs (ERP/data queries)
* File-based IPC (`handover.json`)
* SQLite audit logging (`job_audit.db`)
* Crash-safe mode (`safestop`)
* Remote reboot mechanism (`reboot.flag`)
* Built-in screen recording (ffmpeg)
* Runs without administrator rights
* Cross-platform (Windows and Linux)
* Single-file runtime (`main.py`) for easy sharing and inspection

---


## Job Sources

The system supports multiple job producers:

* Personal inbox (`personal_inbox`)
* Shared mailbox (planned/partial)
* Scheduled jobs (ERP/data queries)

All sources produce standardized **candidates**, processed through a unified flow.

---

## Job Lifecycle

Jobs are tracked in SQLite (`job_audit.db`) with clear states:

* `REJECTED` – invalid request / no access
* `QUEUED` – waiting for RPA
* `RUNNING` – RPA executing
* `VERIFYING` – post-check (Python)
* `DONE` – success
* `FAIL` – error or verification failure

---


## Email Pipeline

* Access controlled via `friends.xlsx`
* Only allowed users can trigger jobs
* Job permissions defined per user
* Unknown senders → silently deleted
* Automatic replies:

  * `DONE`
  * `FAIL`
  * "lifesign" (once per day)

---


## Running the Project

### Requirements

* Python 3.14
* `openpyxl`
* `ffmpeg` (optional, for screen recording)

---

### Start

```bash
python main.py
```

---

### Test setup

Use included dev tools:

* `fake_jobs_generator.py`
* `rpa_tool_simulator.py`

---

## Project Structure (simplified)

```
main.py
personal_inbox/
shared_inbox/
handover.json
job_audit.db
friends.xlsx
recordings/
```

---

## Intended Use Case

* Small internal automation (5–10 users)
* No dedicated RPA infrastructure
* No admin rights required
* Cheap “extra laptop” deployment
* Pilot / proof-of-concept automation

---
## Why not just use X?

#### Why not just use RPA for everything?

You can — but it tends to lead to:

* Business logic spread across visual workflows
* Difficult testing and debugging
* Fragile automations that break on small UI changes

---

#### Why not just use Python for everything?

Python is great for logic and data processing, but:

* It cannot reliably interact with arbitrary GUIs
* Many business systems (ERP, legacy apps) require UI automation

---

#### Why not use an enterprise orchestrator?

Enterprise orchestrators (e.g. UiPath Orchestrator, Control Room):

* Require infrastructure, setup, and licensing
* Are designed for large-scale, multi-bot environments

---

#### Why not use a workflow orchestrator?

Workflow tools (e.g. Airflow, Prefect) are built for:

* Scheduled pipelines
* Data engineering workflows
* Distributed task execution

This project instead focuses on:

* Business-triggered jobs (email, ERP signals)
* Tight coupling to UI automation (via RPA)

---
## What you will likely need to adapt

Most users will need to replace or customize:

- the mail backend
- the ERP/data backend
- the job handlers
- the network health check path
- the screen-recording destination
- the RPA tool implementation

---
## Limitations

* Not designed for large-scale orchestration
* No distributed execution
* File-based IPC only
* Minimal error recovery (by design)


---

## License

MIT (recommended)

---

## Status

Early-stage / experimental, but functional.

---

> Written with help from you-know-who

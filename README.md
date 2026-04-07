# Robot Runtime

A small Python runtime for email- and query-driven RPA

---

## Overview

Robot Runtime is a local Python runtime for small-scale RPA deployments.
'Runtime' is this text refers to job intake, local orchestration, decision logic, logging, and result verification.
The final part, UI automation, is delegated to an external RPA tool such as Power Automate or UiPath Studio.
Together, the Runtime and the RPA tool form the robot.

This project is useful for getting started with automating tasks in a business unit. It runs on a single machine with one single file.
The principle is: **Python owns the logic. The RPA tool owns the screen**

---

## Example dashboard
<img width="1209" height="635" alt="image" src="https://github.com/user-attachments/assets/5f16b39f-99b3-4c82-ad91-0b3092f3b516" />

---

## Job source examples

The runtime supports two types of job sources: emails and queries.

#### Email-driven
A user sends an email → Python validates and prepares the job → writes to `handover.json` → RPA executes UI actions → Python verifies and responds.

#### Query-driven
Python polls a data source → detects a valid case → prepares payload → signals RPA → RPA executes → Python verifies the outcome.


---

## Key Idea
This project separates responsibilities between the Orchestrator and the RPA tool:

* The **Runtime (this project)** handles:
  - how jobs enter the system (e.g. via email)
  - how jobs are discovered autonomously (e.g. queries)
  - preparing payloads and handover to an RPA tool
  - verifying results after execution
  - handling failures

* The **RPA tool** handles:
  - UI automation (clicks, keyboard input, ERP/UI interaction)

They communicate through a file-based IPC mechanism (`handover.json`).

---


---

## Architecture

<img width="1140" height="1709" alt="workflow" src="https://github.com/user-attachments/assets/9e9b1135-76c9-40d6-9f7f-785cfbde715d" />

The diagram shows:

* How the Runtime and the RPA tool run independently
* How your RPA tool must be implemented

---

## Features

* Email-driven job processing (personal inbox)
* Shared inbox support (partially implemented)
* Query-driven jobs (ERP/data polling)
* SQLite audit-style logging (`job_audit.db`)
* Crash-safe mode (`safestop`)
* Built-in screen recording (ffmpeg)
* Final user replies after verification (DONE / FAIL)
* Screen-recording link included in final reply
* Runs without administrator rights
* Single-file runtime (`main.py`) for easy sharing and inspection
* Windows or Linux, with environment-specific setup

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

In this project the RPA tool is used for what it does best: UI interactions (clicks, keyboard input, screen automation).
These tools include Microsoft Power Automate, UiPath Studio, Blue Prism, [Robot Framework](https://github.com/robotframework/robotframework), [TagUI](https://github.com/aisingapore/TagUI), [RPA for Python](https://github.com/tebelorg/RPA-Python)

---

#### Why not just use Python for everything?

Python is great for logic and data processing, but:

* It cannot reliably interact with arbitrary GUIs
* Many business systems (ERP, legacy apps) require UI automation

This project capitalize on the simplicity and large resources avaliable for Python ecosystem.

---

#### Why not use an enterprise orchestrator?

Enterprise orchestrators (e.g. UiPath Orchestrator, Control Room, [orchestrator_rpa](https://github.com/daferferso/orchestrator_rpa), [orchestrator_rpa](https://github.com/daferferso/orchestrator_rpa), [openorchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator)

* Require infrastructure, setup, and licensing
* Are designed for large-scale, multi-bot environments

This project intentionally avoids that scope and runs on a single machine with simple file- and DB-based state.
If you need distributed execution, queues, or centralized control — this project is the wrong tool.

---

## Deployment requirements

- a dedicated machine or “extra laptop”
- a mailbox such as rpa@yourcompany.com
- an external RPA tool
- environment-specific setup for mail backend, ERP/query backend, job handlers, recording path, operating hours, and network health check

---

## License

MIT (recommended)

---

## Status

Early-stage / experimental, but functional.

---

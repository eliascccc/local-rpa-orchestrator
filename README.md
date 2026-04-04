# LocalRPA Orchestrator

A lightweight orchestrator that keeps logic in Python — and uses RPA tools only for UI execution.

---

## Overview

The goal is a simple way to introduce RPA in a team.
It follows the principle: **automation logic belongs in Python — UI execution belongs in RPA tools.**

It is designed as a smaller alternative to enterprise orchestrators,
focusing on clarity, ease of modification, and running on a single machine.

It does NOT replace RPA tools.

Instead, it orchestrates them:
you still need a real RPA tool (Power Automate, UiPath Studio, Blue Prism, etc.)
to perform screen-based automation. Think of this as the shell around the RPA, not the robot itself.

## Responsibilities

| Step | Orchestrator (this project) | RPA tool |
| ---- |:--------------------------:|:----------:|
| Job intake| ✓ |  |
| Handover | → | ✓ |
| Job execution |  | ✓ |
| Handover | ✓ | ← |
| Job verification | ✓ |  |


---

## Example dashboard
<img width="1209" height="635" alt="image" src="https://github.com/user-attachments/assets/5f16b39f-99b3-4c82-ad91-0b3092f3b516" />

---

## Job source examples

The orchestrator supports two types of job sources: emails and queries.

#### Email-driven
A user sends an email → Python validates and prepares the job → writes to `handover.json` → RPA executes UI actions → Python verifies and responds.

#### Query-driven
Python polls a data source → detects a valid case → prepares payload → signals RPA → RPA executes → Python verifies the outcome.


---

## Key Idea
This project separates responsibilities between the Orchestrator and the RPA tool:

* The **Orchestrator (this project)** handles:
  - how jobs enter the system (e.g. via email)
  - how jobs are discovered autonomously (e.g. queries)
  - access control and validation
  - job state tracking and audit logging
  - preparing payloads and handover to an RPA tool
  - verifying results after execution
  - handling failures

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
* [Robot Framework](https://github.com/robotframework/robotframework)
* [TagUI](https://github.com/aisingapore/TagUI)
* [RPA for Python](https://github.com/tebelorg/RPA-Python)

These tools perform the actual UI interactions (clicks, keyboard input, screen automation).

---

## Architecture

<img width="1140" height="1709" alt="workflow" src="https://github.com/user-attachments/assets/9e9b1135-76c9-40d6-9f7f-785cfbde715d" />

The diagram shows how:

* The Orchestrator and the RPA tool run independently
* State is synchronized via `handover.json`
* Failures transition the system into a safestop
* Your RPA tool must follow this model

---

## Two distinct email channels

#### 1. Personal inbox (command channel)

This is a direct communication channel to the robot.

* Users explicitly send requests to the robot (e.g. rpa@company.com)
* The robot may reply to the sender
* Access control is enforced (e.g. via friends.xlsx)
* This behaves like a command interface

#### 2. Shared inbox (operational channel)

This is a passive monitoring channel.

* The robot listens to an existing business inbox (e.g. orders@company.com)
* Senders are typically external and unaware of the robot
* The robot must never reply
* The robot only processes emails that match a defined scope
* All other emails are ignored or returned to the inbox

---

## Features

* Email-driven job processing (personal inbox)
* Shared inbox support (partially implemented)
* Query-driven jobs (ERP/data polling)
* File-based IPC (`handover.json`)
* SQLite audit-style logging (`job_audit.db`)
* Crash-safe mode (`safestop`)
* Controlled restart mechanism (`restart.flag`)
* Stop hook from the RPA side (`stop.flag`)
* Built-in screen recording (ffmpeg)
* Final user replies after verification (DONE / FAIL)
* Screen-recording link included in final reply
* Runs without administrator rights
* Cross-platform (Windows and Linux)
* Single-file runtime (`main.py`) for easy sharing and inspection

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
* `select_all_from_job_audit.py`

---

## Project Structure (simplified)

```
main.py
personal_inbox/
shared_inbox/
handover.json
job_audit.db
friends.xlsx
recordings_destination/
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

This project intentionally avoids that scope and runs on a single machine with simple file- and DB-based state

---

#### Why not use a workflow orchestrator?

Workflow tools (e.g. Airflow, Prefect) are built for:

* Scheduled and data pipelines
* Data engineering workflows
* Distributed task execution

This project is much smaller, local-first, and designed around business-triggered jobs (email, ERP signals) plus screen-based RPA

---
## What you will likely need to adapt

Most users will need to replace or customize:

- the mail backend (e.g. connect a real personal inbox such as rpa@yourcompany.com)
- the query/ERP backend
- the job handlers
- the network health check path
- the screen-recording destination (a shared network drive)
- the operating hours
- the RPA tool implementation

---

## Design Philosophy
* Simplicity over scalability
* Rather crash than guess
* Logic in Python, UI automation in the RPA tool

---

## Limitations

* Not designed for large-scale orchestration
* No distributed execution
* File-based IPC only
* Minimal error recovery (by design)
* Pure Python-only jobs are out of scope by design.

---

## License

MIT (recommended)

---

## Status

Early-stage / experimental, but functional.

---

> Written with help from you-know-who

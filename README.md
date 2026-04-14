# Robot Runtime

A local runtime for email- and query-driven RPA — where Python owns the logic and RPA tools own the screen.

---

## Overview

Robot Runtime is a local Python runtime for small-scale RPA deployments.
It handles everything except UI automation:
job intake, orchestration, decision logic, and result verification.
UI automation is delegated to an external RPA tool such as UiPath or Power Automate.
Together, the runtime and the RPA tool form the robot.

Unlike traditional RPA setups — where users manually trigger predefined automations — this runtime is event-driven. It continuously listens for incoming work (emails or queries) and decides what to do. It runs on a single machine as a single Python file.

The principle is: **UI interaction** → handled by the RPA tool. **The rest (logic and orchestration)** → this Python runtime

---

## Example dashboard
<img width="1209" height="635" alt="image" src="https://github.com/user-attachments/assets/5f16b39f-99b3-4c82-ad91-0b3092f3b516" />

---

## Job intake examples

The runtime supports two types of job sources: emails and queries.

#### Email-driven
A user sends an email → Python validates and prepares the job → writes to `handover.json` → RPA executes UI actions → Python verifies and responds.

#### Query-driven
Python polls a data source → detects a valid case → prepares a payload → signals RPA → RPA executes → Python verifies the outcome.

---

## Intended Use Case

* Small internal automation (5–10 users)
* No dedicated RPA infrastructure
* No admin rights required
* Cheap “extra laptop” deployment
* Pilot / proof-of-concept automation

---



## Architecture

<img width="1140" height="1709" alt="workflow" src="https://github.com/user-attachments/assets/0e0950c7-cc59-40ca-9fb0-3e989c862a62" />

The diagram shows:

* How the Runtime and the RPA tool run independently
* How your RPA tool should be implemented

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

The recommended (production-like) setup is to run `main.py` from the RPA tool according to architecture diagram.  
The RPA tool starts and stops the runtime, which makes the robot behave as a single unit.

---

### Testing / development

Use the included dev tools to simulate real inputs and runtime behavior:

* `fake_jobs_generator.py` – to generate test jobs (emails / data)
* `rpa_tool_simulator.py` – to simulate the RPA tool and start main.py in the intended way

Manual start is also possible
```bash
python main.py
```

---

## Deployment requirements

- a dedicated machine or “extra laptop”
- a mailbox such as rpa@yourcompany.com
- an external RPA tool
- environment-specific setup for mail backend, ERP/query backend, job handlers, recording path, operating hours, and network health check

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

This project leverages the simplicity and rich ecosystem of Python for logic,
while relying on RPA tools for reliable UI automation.

---

#### Why not use an enterprise orchestrator?

Enterprise orchestrators (e.g. UiPath Orchestrator, Control Room, [orchestrator_rpa](https://github.com/daferferso/orchestrator_rpa), [openorchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator))

* Require infrastructure, setup, and licensing
* Are designed for large-scale, multi-bot environments

This project intentionally avoids that scope and runs on a single machine with simple file- and DB-based state.
If you need distributed execution, queues, or centralized control — this project is the wrong tool.

---

## License

MIT

---

## Status

Early-stage / experimental, but functional.

---

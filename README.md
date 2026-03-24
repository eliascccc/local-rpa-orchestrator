# LocalRPA Orchestrator

A lightweight local orchestrator for email- and data-driven automation, delegating execution to external screen-based RPA tools.

---

## Overview

**LocalRPA Orchestrator** is a minimal Python runtime designed to orchestrate automations tasks in small teams.

It runs locally on a standard operator machine and acts as the **control layer** between:

* incoming job triggers (email, shared inbox, data queries)
* and external RPA tools (UiPath, Power Automate, Blue Prism, etc.)

This project is intentionally designed for small, local deployments without infrastructure or licensing overhead.

---

## Key Idea

This project separates **orchestration** from **UI automation**:

* The **backend (this project)** handles:

  * job intake (email / data)
  * validation
  * decision logic
  * audit logging
  * system control

* The **front-end RPA tool** handles:

  * clicks
  * keyboard input
  * ERP/UI interaction

They communicate through a simple file-based IPC mechanism (`handover.txt`).

---

## Architecture

The diagram shows how:

* Python (backend) and RPA (front-end) run independently
* Both operate in their own loops
* State is synchronized via handover.txt
* Failures transition the system into safestop

## Features

* Email-driven job processing (own inbox)
* Shared inbox support (extensible)
* Data-driven jobs (ERP/query simulation)
* File-based IPC (`handover.txt`)
* SQLite audit logging (`audit.db`)
* Crash-safe mode (`safestop`)
* Manual reboot mechanism (`reboot.flag`)
* Network-aware execution
* Built-in screen recording (ffmpeg)
* Works without admin rights
* Easy to share the full runtime with an AI assistant
* Runs on both Windows and Linux

---

## Design Principles

* **Simplicity over scalability**
* **Local-first execution**
* **Fail fast and visibly**
* **No hidden state (file + DB only)**
* **Deterministic job lifecycle**
* **Cheap to deploy and operate**

---

## Job Sources

The system supports multiple job producers:

* Personal inbox (`own_inbox`)
* Shared mailbox (planned/partial)
* Scheduled jobs (ERP/data queries)

All sources produce standardized **candidates**, processed through a unified flow.

---

## Job Lifecycle

Jobs are tracked in SQLite (`audit.db`) with clear states:

* `REJECTED` – invalid request / no access
* `QUEUED` – waiting for RPA
* `RUNNING` – RPA executing
* `VERIFYING` – post-check (Python)
* `DONE` – success
* `FAIL` – error or verification failure

---

## IPC Protocol (handover.txt)

The orchestrator and RPA communicate via a shared JSON file.

### Example:

```json
{
  "ipc_state": "job_queued",
  "job_id": 20260324123001,
  "job_type": "job1"
}
```

### States:

* `idle`
* `job_queued`
* `job_running`
* `job_verifying`
* `safestop`

---

## Safety Model

* No resume: crashed jobs do not continue
* Watchdog timeout for RPA stalls
* All critical errors → `safestop`
* Manual recovery via:

  * `reboot.flag`
* System can kill external processes to reset environment

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

## Example Jobs

### Ping

Send:

```
Subject: ping
```

Reply:

```
PONG (robot online)
```

---

### Job1 (example)

Email body:

```
SKU: 123
Old material: ABC
New material: XYZ
```

Validated and passed to RPA for execution.

---

## Running the Project

### Requirements

* Python 3.14
* `openpyxl`
* `ffmpeg` (optional, for recording)

---

### Start

```bash
python main.py
```

---

### Test setup

Use included dev tools:

* `fake_jobs_generator.py`
* `frontend_rpa_simulator.py`

---

## Project Structure (simplified)

```
main.py
own_inbox/
shared_inbox/
handover.txt
audit.db
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

## Limitations

* Not designed for large-scale orchestration
* No distributed execution
* File-based IPC only
* Minimal error recovery (by design)

---

## Why not just use Robot Framework?

| Robot Framework                    | LocalRPA Orchestrator          |
| ---------------------------------- | ------------------------------ |
| Scalable test/automation framework | Lightweight local orchestrator |
| Requires setup/integration         | Runs on a single machine       |
| Code-driven automation             | Combines email + RPA tools     |
| General-purpose                    | Business-triggered workflows   |

---

## Philosophy

> This is not a full RPA platform.
> It is the simplest possible glue between:
>
> * business triggers (email/data)
> * and UI automation (RPA tools)

---

## License

MIT (recommended)

---

## Status

Early-stage / experimental, but functional.

---

readme written by AI

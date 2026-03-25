# LocalRPA Orchestrator

A lightweight local orchestrator that bridges business triggers and screen-based RPA.

---

## Overview

This project is a mini RPA orchestrator written in Python.

It is designed as a small-scale alternative where enterprise orchestrators would be unnecessary overhead,
focusing on clarity, ease of modification, and running on a single machine.

It does NOT replace RPA tools.

Instead, it orchestrates them:
you still need a real RPA tool (Power Automate, UiPath Studio, Blue Prism, etc.)
to perform screen-based automation.

---


## Typical examples

In an email-driven flow, a user sends a request to the robot. The orchestrator reads the request, validates the input, prepares the required payload, and writes the handover state. The front-end RPA tool then picks up the job and performs the UI actions.

Another common trigger is data-driven. In that case, the orchestrator discovers work by polling a query or another data source. When it finds a valid case, it prepares the required values and signals the front-end RPA to execute the task.

Example dashboard during runtime:
<img width="1209" height="635" alt="example_dash" src="https://github.com/user-attachments/assets/dc12a84b-c329-4b91-b402-387128197f9a" />

---

## Key Idea

This project separates **orchestration** from **UI automation**:

* The **back-end (this project)** handles:

  * job intake (email / data)
  * validation
  * decision logic
  * audit logging
  * system control

* The **front-end RPA tool** handles:

  * clicks
  * keyboard input
  * ERP/UI interaction

They communicate through a file-based IPC mechanism (`handover.json`).

---

## Architecture

<img width="1140" height="1709" alt="workflow" src="https://github.com/user-attachments/assets/8c94348e-be80-4a39-bc64-226e38f67167" />


The diagram defines the interaction:

* Python (back-end) and RPA (front-end) run independently
* Both operate in their own loops
* State is synchronized via handover.json
* Failures transition the system into safestop
* Your front-end RPA tool should follow this model


## Features

* Email-driven job processing (personal inbox)
* Shared inbox support (extensible)
* Data-driven jobs (ERP/query simulation)
* File-based IPC (`handover.json`)
* SQLite audit logging (`job_audit.db`)
* Crash-safe mode (`safestop`)
* Remote reboot mechanism (`reboot.flag`)
* Built-in screen recording (ffmpeg)
* Runs without administrator rights
* Cross-platform (Windows and Linux)
* Single-file runtime (`main.py`) for easy sharing and inspection

---

## Design Principles

* **Simplicity over scalability**
* **Local-first execution**
* **Fail fast and visibly**
* **Deterministic job lifecycle**
* **Minimal dependencies and infrastructure**
* **Transparent state (file + database)**
* **Cheap to deploy and operate**

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

## Handover concept

The system uses a file-based "handover" mechanism to transfer control between the Python orchestrator and the external RPA.

The handover file represents both:
- the current system state
- and the payload required for the next execution step

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
* `frontend_rpa_simulator.py`

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

## Limitations

* Not designed for large-scale orchestration
* No distributed execution
* File-based IPC only
* Minimal error recovery (by design)

---
## How it compares
| Tool / category                          | Primary role                                                                | How this project differs                                                                                       |
| ---------------------------------------- | --------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| **Robot Framework**                      | Executes automation/test logic                                              | This project focuses on intake, decision logic, audit, job state, and handover to an external RPA tool         |
| **Airflow-style workflow orchestrators** | Orchestrate scheduled/data workflows across tasks and systems               | This project is much smaller, local-first, and designed around business-triggered jobs plus screen-based RPA   |
| **Enterprise RPA orchestrators**         | Centralized control of bots, queues, schedules, credentials, and monitoring | This project intentionally avoids that scope and runs on a single machine with simple file- and DB-based state |

---

## Philosophy

This is not a full RPA platform.
It is the simplest possible layer between:

* business triggers (email/data)
* and UI automation (RPA tools)

---

## License

MIT (recommended)

---

## Status

Early-stage / experimental, but functional.

---

> I got help writing this readme

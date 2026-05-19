# maritime-telemetry-automation
Python-based ETL and data governance pipelines interfacing with cloud databases to automate compliance reporting.

Markdown
# Maritime Telemetry & Reporting Automation Pipeline

A production-ready Python automation suite designed to interface with cloud databases (PostgreSQL/AWS RDS) to capture real-time vessel telemetry, enforce strict data validation gates, mitigate cross-organizational data latency, and automate compliance document generation.

## Technical Architecture Highlights

* **Defensive Data Ingestion & Sanitization:** Developed robust ETL scripts leveraging `pandas` and Regular Expressions (`re`) to automatically parse, clean, and standardize unvetted client data and captain logs prior to database ingestion.
* **Infrastructure Latency Monitoring & Emergency Alerts:** Engineered a transaction monitoring script tracking backend process-duration deltas. Configured the utility to generate local system logs and execute automated high-priority SMTP alerts (`win32com`) if pipeline latency triggers anomalies.
* **Cross-Functional Data Synchronization:** Implemented automated data listeners that execute daily database queries to evaluate, aggregate, and report contract volumes across localized systems and regional international subsidiaries simultaneously.
* **Automated Spatial GIS Reporting Engine:** Managed the programmatic lifecycle execution of external mapping engines (**QGIS**) via binary flags, using relational outputs to dynamically inject spatial vector overlays straight into custom executive text reports.

## Core Technology Stack
* **Language:** Python 3
* **Database Interfacing:** `psycopg2`, `psycopg2.extras` (RealDictCursor)
* **Data Manipulation:** `pandas`, `numpy`
* **System & Workflow Automation:** `os`, `shutil`, `sys`, `subprocess`, `win32com`

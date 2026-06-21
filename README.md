# English 

# Lawn Care Field Operations Console
Turn scattered messages and mental notes into a lightweight dispatch and execution system for small lawn care crews.

![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Browser%20%2B%20Excel-green.svg)
![Tool](https://img.shields.io/badge/Tool-Operations%20Decision%20Support-orange.svg)

**A free, no-install dispatch and execution tracker that helps micro lawn care teams see today's workload, completion status, exceptions, and expected revenue in one place—available in both Browser and Excel formats.**

> **No signup. No installation. Free.**
>
> 🌐 Open in Browser → *HTML live demo link*  
> 📥 Download Excel → *Gumroad or GitHub Release link*

---

## Screenshots

### Browser Version

<!-- screenshot: browser version -->

*Operational dashboard showing today's workload, job status, exceptions, and expected revenue for quick owner review.*

### Excel Version

<!-- screenshot: excel version -->

*Mobile-friendly dispatch sheet used by field staff to update job status directly from the job site.*

---

## What It Helps You Track

- Planned jobs versus completed jobs for the day.
- Field exceptions that require owner follow-up before they are forgotten.
- Expected daily revenue based on actual scheduled work.
- Which crew member is responsible for each customer visit.
- Customer-specific instructions that prevent avoidable mistakes.
- The live execution status of today's route without phone calls or memory checks.

---

## Why I Built This

Most small lawn care businesses do not fail because they lack software. They struggle because operational information lives in too many places.

Customer addresses sit in someone's head.

Special instructions live inside old text messages.

Job assignments happen verbally in the truck.

Completion status depends on asking, "Did that one get done?"

The cognitive failure is simple: decisions are made using incomplete information.

Owners believe they have visibility because they remember most things most of the time. But once a second crew member joins, memory becomes an unreliable operating system.

Before:

A customer calls asking why service was skipped.

The owner spends ten minutes reconstructing events:

- Who handled the job?
- Was the gate locked?
- Was there a dog in the backyard?
- Was the crew even assigned?

After:

One dashboard immediately shows:

- Crew member assigned,
- Job status,
- Completion time,
- Exception notes,
- Expected payment impact.

This workbook is not another spreadsheet.

It productizes operational reasoning.

Instead of rebuilding the same dispatch process repeatedly through messages and memory, it provides a reusable framework for capturing the information required to make the next decision confidently.

---

## Common Lawn Care Problems This Solves

| Problem | Without This Tool | With This Tool |
|---|---|---|
| Customer instructions get forgotten | Crews miss gates, dogs, lock codes, or special requests | Persistent customer notes automatically appear during dispatch |
| Owners rely on verbal updates | Completion status requires repeated phone calls | Live job status is visible immediately |
| Revenue visibility is delayed | Income estimates happen after paperwork | Expected revenue updates throughout the day |
| Exception handling is inconsistent | Skipped jobs disappear until complaints arrive | Exceptions are highlighted for follow-up |
| Customer data is duplicated | Addresses and prices are copied repeatedly | Single-source customer records feed all dispatch activities |
| Growth beyond solo operation creates chaos | Memory breaks as team size increases | Structured workflows scale basic operations |

---

## Who This Is For

This tool is designed for:

- Solo lawn care operators adding their first employees.
- Small field-service teams managing daily dispatch manually.
- Owners who currently depend on messages, notes, and memory.
- Businesses wanting a low-cost operational layer before investing in larger systems.

This tool is **not** designed for:

- Large multi-branch service organizations.
- Businesses requiring enterprise CRM capabilities.
- Teams needing advanced route optimization or accounting systems.

No spreadsheet expertise is needed. Open the browser version and start tracking immediately, or use the Excel version offline.

---

## About

I build lightweight trackers and decision-support tools for situations with too many moving parts to hold in someone's head.

The question behind each tool is simple:

> What information needs to exist in one place so the next operational decision can be made confidently?

[Lawn Care Field Operations Console] is one example of this approach: productized reasoning packaged into familiar tools instead of requiring another software platform.

---

## Technical Details

<details>
<summary>For technical reviewers, Excel practitioners, and collaborators</summary>

### Workbook Architecture

| Layer | Sheet | Role |
|---|---|---|
| Input | `01_Customer_Master` | Customer master records |
| Transaction | `02_Job_Execution` | Daily dispatch and execution |
| Output | `03_Today_Dashboard` | Monitoring and decision support |

#### Data Flow

```text
Customer Master
      ↓
Dispatch Creation
      ↓
Field Execution Updates
      ↓
Dashboard Aggregation
      ↓
Owner Decisions
```

Validation flow:

```text
Customer Status Validation
        ↓
Customer ID Selection
        ↓
Automatic Lookup
        ↓
Field Status Updates
        ↓
Dashboard Calculations
```

Output dependencies:

- Dashboard KPIs depend on Job Execution.
- Job Execution depends on Customer Master.
- Customer Master remains protected as the single source of truth.

---

### Three Traps That Catch Even Experienced Owners

#### Trap 1: Assuming "No Complaints" Means Work Was Completed

Decision:
> The owner assumes the route was completed.

Faulty assumption:
> Silence equals execution.

Result:
Skipped jobs remain invisible.

Correct approach:
Track completion explicitly.

Outcome:
Exceptions become actionable instead of accidental.

<details>
<summary>Formula Reference</summary>

```excel
=COUNTIFS(
Job_Execution!B:B,TODAY(),
Job_Execution!I:I,"已完成"
)
```

```excel
=COUNTIFS(
Job_Execution!B:B,TODAY(),
Job_Execution!I:I,"跳过/异常"
)
```

</details>

---

#### Trap 2: Re-entering Customer Information

Decision:
> Staff manually copy addresses and notes.

Faulty model:
> Duplicate entry is harmless.

Result:
Different versions of the truth emerge.

Correct approach:
Use customer IDs as foreign keys.

Outcome:
One update propagates everywhere.

<details>
<summary>Formula Reference</summary>

```excel
=XLOOKUP(
CustomerID,
Customer_Master!A:A,
Customer_Master!C:C,
"未匹配"
)
```

</details>

---

#### Trap 3: Evaluating Revenue Only After the Day Ends

Decision:
> Financial visibility waits for bookkeeping.

Faulty assumption:
> Revenue tracking is accounting's job.

Result:
Owners discover shortfalls too late.

Correct approach:
Aggregate expected receipts continuously.

Outcome:
Operational decisions can be adjusted during the day.

<details>
<summary>Formula Reference</summary>

```excel
=SUMIF(
Job_Execution!B:B,
TODAY(),
Job_Execution!L:L
)
```

</details>

---

### Example Scenario

A three-person crew schedules six jobs for Monday.

#### Raw Inputs

| Customer | Price | Assigned To |
|---|---:|---|
| Smith | $50 | Employee A |
| Johnson | $60 | Employee B |
| Davis | $55 | Employee A |
| Brown | $45 | Owner |
| Wilson | $65 | Employee B |
| Garcia | $50 | Employee A |

Expected revenue:

```text
$325
```

#### Intermediate Results

Execution status:

- Completed: 4
- Exception: 1
- Pending: 1

Exception note:

```text
Johnson:
Backyard gate locked.
Front yard completed only.
Reschedule required.
```

Completion rate:

```text
4 ÷ 6 = 66.7%
```

#### Interpretation

Without structured tracking:

- The skipped job may disappear.
- Revenue shortfall remains hidden.
- Customer communication becomes reactive.

With the framework:

- Follow-up is immediate.
- Revenue expectations adjust.
- Re-dispatch decisions occur before complaints.

#### Decision Implications

Recommended actions:

1. Contact Johnson.
2. Schedule revisit.
3. Adjust revenue forecast.
4. Reallocate tomorrow's workload.

---

### Formula Reference

<details>
<summary>Customer Lookup Formulas</summary>

Purpose:
Retrieve customer information automatically.

```excel
=XLOOKUP(CustomerID,Master!A:A,Master!B:B)
=XLOOKUP(CustomerID,Master!A:A,Master!C:C)
=XLOOKUP(CustomerID,Master!A:A,Master!G:G)
=XLOOKUP(CustomerID,Master!A:A,Master!H:H)
```

</details>

<details>
<summary>Dashboard KPI Formulas</summary>

Purpose:
Monitor today's operations.

```excel
=COUNTIF(Job_Execution!B:B,TODAY())

=COUNTIFS(
Job_Execution!B:B,TODAY(),
Job_Execution!I:I,"已完成"
)

=COUNTIFS(
Job_Execution!B:B,TODAY(),
Job_Execution!I:I,"跳过/异常"
)

=IF(
N_total=0,
0,
N_completed/N_total
)

=SUMIF(
Job_Execution!B:B,
TODAY(),
Job_Execution!L:L
)
```

</details>

<details>
<summary>Dashboard Queue Formula</summary>

Purpose:
Display today's live task queue.

```excel
=FILTER(
Job_Execution!D:K,
Job_Execution!B:B=TODAY(),
"今日无排班任务"
)
```

</details>

---

### Validation Rules

| Field | Rule | Error Behavior |
|---|---|---|
| Customer ID | Unique primary key | Duplicate customer blocked |
| Customer Name | Required | Missing record flagged |
| Service Address | Required | Dispatch incomplete |
| Service Frequency | Dropdown only | Invalid value rejected |
| Customer Status | Active/Inactive only | Hidden from dispatch |
| Dispatch Date | Required date | Dashboard exclusion |
| Customer ID (Dispatch) | Must exist in master | Lookup returns unmatched |
| Responsible Staff | Dropdown only | Assignment blocked |
| Task Status | Controlled list | Invalid status rejected |
| Completion Time | Time format | Calculation inconsistency |
| Actual Amount | Numeric currency | Revenue aggregation affected |

</details>

---

## Other Tools in This Series

- Financial Dual-Track Budget Planner — align personal and business cash decisions.
- Inventory Forecasting Console — convert historical demand into replenishment actions.
- CRM Decision Frameworks — structure customer information without enterprise complexity.

GitHub Profile: *GitHub profile link*  
Gumroad Store: *Gumroad store link*

---

## License

This project is licensed under the **Apache License 2.0**.

You are free to use, modify, and distribute this work in accordance with the terms of the Apache License 2.0.

# Excel Systems Design: Complex Logic & Data Architecture

## Project Overview
This repository contains a collection of advanced spreadsheet architectures originally developed to automate complex Tabletop Role-Playing Game (TTRPG) systems (such as Pathfinder, D&D 3.5, and d20 Modern). 

While built for gaming, these sheets serve as robust case studies in **Data Architecture, System Design, and Logic Engineering**. Building a functional TTRPG character sheet requires handling highly relational data, enforcing strict data integrity, and translating hundreds of pages of conditional "business rules" into scalable spreadsheet logic.

**###🔗 [Download / View the Showcase Spreadsheet Here](./demo/TTRPG_Logic_Showcase.xlsx)###**
---

## Table of Contents
1. [Data Integrity & State Management](#1-data-integrity--state-management-multiclassing-engine)
2. [Dynamic Rules Engines](#2-dynamic-rules-engines-context-aware-attribution)
3. [Bidirectional Financial Modeling & Data Binning](#3-bidirectional-financial-modeling--data-binning-wealth-dc)
4. [End-to-End Data Pipeline & State-Dependent UI](#4-end-to-end-data-pipeline--state-dependent-ui-gestalt-engine)
5. [Boolean Eligibility Logic](#5-boolean-eligibility-logic-entitlement-mapping)

---

## 1. Data Integrity & State Management (Multiclassing Engine)

**The Challenge:** In systems allowing multiclassing, players may enter multiple rows for the same class as they level up. A standard `SUMIF` would double-count these entries, breaking the character's stats. 

**The Solution:** I developed a de-duplication and aggregation engine that enforces a "Single Version of Truth." It identifies unique entities, finds their maximum state (highest level), and only sums the valid records.

```excel
=CONCAT("+ ",@LET(u,UNIQUE($B$3:$B$14),m,MAXIFS($C$3:$C$14,$B$3:$B$14,u),SUM(MAXIFS(D$3:D$14,$B$3:$B$14,u,$C$3:$C$14,m))))
```

**Professional Translation:** This demonstrates **Data Normalization**. In a business context, this is equivalent to querying a transactional database to find the most recent interaction for a list of unique customers before aggregating their lifetime value. Using `LET` optimizes processing time by preventing recalculation of the `UNIQUE` array, while `CONCAT` dynamically formats the numeric output for the end-user UI.

---

## 2. Dynamic Rules Engines (Context-Aware Attribution)

**The Challenge:** Calculating an attack bonus requires the system to "know" the character's capabilities (feats) and the properties of the tool (weapon) simultaneously, applying different modifiers based on intersecting conditions.

**The Solution:** This formula uses nested logic and wildcard searching to build a conditional rules engine. It defaults to a standard modifier but dynamically switches if specific prerequisites are met.

```excel
=IF(OR($B15="Ranged",AND(COUNTIF($F$3:$G$5,"*"&"Weapon Finesse"&"*"),$B15="Light")),$C$11,$C$10)
```

**Professional Translation:** This mirrors **Automated Compliance Auditing or Incentive Logic**. The system evaluates a complex Boolean matrix (e.g., If Tool = X OR (Tool = Y AND User has Certification Z)) to dynamically route the data and retrieve the correct baseline metric based on intersecting conditions.

---

## 3. Bidirectional Financial Modeling & Data Binning (Wealth DC)

**The Challenge:** The d20 Modern system uses abstract "Purchase DCs" rather than actual dollar amounts. The sheet needed to seamlessly translate between abstract tiers and usable, real-world currency values based on dynamic user input.

**The Solution:** I built a two-way numerical transformation engine. The first formula uses an exact match lookup wrapped in a rounding function to convert a specific tier into a clean dollar amount. The second formula reverse-calculates raw financial input back into the correct system tier by evaluating where the input falls between defined minimum and maximum boundaries.

*Tier to Currency (Exact Match & Formatting):*
```excel
=MROUND(XLOOKUP($B$3,$K$15:$K$94,$N$15:$N$94,""),5)
```

*Currency to Tier (Range Match/Data Binning):*
```excel
=IFNA(IF($F$3<>"",LOOKUP($F$3,$L$15:$M$94,$K$15:$K$94),""),"")
```

**Professional Translation:** This engine performs two-way numerical transformation utilizing a 4-column reference architecture (Tier, Min, Max, Average). It translates abstract bands into standardized currency via exact match lookups, and reverse-calculates raw financial data back into system tiers using **Range-Based Lookups (Data Binning)**. This mirrors corporate financial models that map continuous variables (raw prices) into discrete analytical categories (pricing tiers).

---

## 4. End-to-End Data Pipeline & State-Dependent UI (Gestalt Engine)

**The Challenge:** In "Gestalt" systems, characters gain abilities from two separate classes simultaneously, including edge cases that require querying entirely different databases (e.g., Investigators). The system needed to enforce valid data entry at the source, extract this disparate data without altering the raw database structure, and merge it into a single, clean UI element.

**The Solution:** I built a three-stage ETL pipeline that begins with dynamic user validation. The front-end restricts inputs based on state changes, the extraction layer builds virtual composite keys to route and query the databases, and the presentation layer cleans and concatenates the final arrays.

*Stage 1: State-Dependent Data Validation (Cascading Drop-Downs)*
```excel
=INDIRECT(SUBSTITUTE($B7,"",""))
```
(This logic sits inside Data Validation, dynamically changing the Archetype drop-down list to strictly match the Class selected in $B7, preventing invalid database queries before they happen.)

*Stage 2: Conditional Extraction (Database Routing & Virtual Arrays)*
```excel
=IF(AND($B7="Investigator",$D7<>""),XLOOKUP($D7&"/"&$C7,$X$7:$X$36&"/"&$Y$7:$Y$36,$Z$7:$Z$36,""),XLOOKUP($B7&"/"&$C7,$O$7:$O$31&"/"&$P$7:$P$31,$U$7:$U$31,""))
```

*Stage 3: Transformation & UI Presentation*
```excel
=IFNA(IF(AND($B7<>"",$E7<>""),CONCAT($AD7," || ",$AE7),$AD7),"")
```

**Professional Translation:** This demonstrates a complete, closed-loop **ETL Pipeline with Front-End Governance**. In a corporate environment, Stage 1 enforces data integrity at the source via dependent validation (cascading menus). Stage 2 mirrors extracting data from multiple tables based on system flags, utilizing in-memory array concatenation (`$O$7:$O$31&"/"&$P$7:$P$31`) to create dynamic composite keys without needing hardcoded helper columns. Finally, Stage 3 transforms the disparate text fields into a single, user-friendly string while gracefully handling nulls.

---

## 5. Boolean Eligibility Logic (Entitlement Mapping)

**The Challenge:** Determining if a character unlocks a specific "Class Skill" requires cross-referencing their active classes against a strict, skill-specific list of prerequisites. Because every skill has a unique entitlement matrix (e.g., Appraise belongs strictly to Rogues, while Bluff is shared by both Rogues and Wizards), the system must evaluate shifting parameters on a row-by-row basis.

**The Solution:** This logic utilizes array evaluation to check the character's currently selected classes against a predefined array of approved classes for that specific skill, returning a clean binary output. 

*Example: "Bluff" Skill Eligibility Check*
```excel
=IFNA(IF(SUMPRODUCT(--('1_Data_Integrity'!$B$3:$B$14={"Rogue","Wizard"}))>0,"Yes","No"),"No")
```

**Professional Translation:** This demonstrates a hardcoded **Rules Engine** utilizing Array Constants and Double Unaries (--) for high-efficiency eligibility checking. In a corporate environment, this mirrors **Role-Based Access Control (RBAC)**. By comparing a user's active "roles" (e.g., Fighter, Wizard) against an array of approved roles for a specific "system" or asset (e.g., the Bluff skill), it securely converts Boolean `TRUE/FALSE` evaluations into a binary access gate, bypassing the need for massive, nested `OR` statements.

---

## Development Methodology
The logic within this repository demonstrates my ability to:
* Architect complex systems with decoupled data and presentation layers.
* Audit, refactor, and implement advanced formula solutions using modern Excel functions (`LET`, `UNIQUE`, dynamic arrays).
* Design resilient systems with comprehensive error handling (`IFNA`) and dynamic scalability.

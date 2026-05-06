# Excel Systems Design: Complex Logic & Data Architecture

## Project Overview
This repository contains a collection of advanced spreadsheet architectures originally developed to automate complex Tabletop Role-Playing Game (TTRPG) systems (such as Pathfinder, D&D 3.5, and d20 Modern). 

While built for gaming, these sheets serve as robust case studies in **Data Architecture, System Design, and Logic Engineering**. Building a functional TTRPG character sheet requires handling highly relational data, enforcing strict data integrity, and translating hundreds of pages of conditional "business rules" into scalable spreadsheet logic.

---

## Table of Contents
1. [Data Integrity & State Management](#1-data-integrity--state-management-multiclassing-engine)
2. [Dynamic Rules Engines](#2-dynamic-rules-engines-context-aware-attribution)
3. [Financial Modeling & Tiered Scaling](#3-financial-modeling--tiered-scaling-wealth-dc)
4. [Relational Data Architecture](#4-relational-data-architecture-composite-keys)
5. [Boolean Eligibility Logic](#5-boolean-eligibility-logic-array-processing)

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
=IFNA(MROUND(VLOOKUP($B$3,$K$15:$N$94,4,FALSE),5),"")
```

*Currency to Tier (Range Match/Data Binning):*
```excel
=IFNA(IF($F$3<>"",LOOKUP($F$3,$L$15:$M$94,$K$15:$K$94),""),"")
```

**Professional Translation:** This engine performs two-way numerical transformation utilizing a 4-column reference architecture (Tier, Min, Max, Average). It translates abstract bands into standardized currency via exact match lookups, and reverse-calculates raw financial data back into system tiers using **Range-Based Lookups (Data Binning)**. This mirrors corporate financial models that map continuous variables (raw prices) into discrete analytical categories (pricing tiers).

---

## 4. Conditional Data Pipeline & ETL Transformation (Gestalt Engine)

**The Challenge:** In "Gestalt" systems, characters gain abilities from two separate classes simultaneously, including edge cases that require querying entirely different databases. The system needed to extract this disparate data and merge it into a single, clean UI element without trailing delimiters if a character only possessed one class.

**The Solution:** I built a two-stage ETL (Extract, Transform, Load) pipeline. The extraction layer uses conditional database routing and composite keys to pull raw text, while the presentation layer cleanly concatenates the arrays using Boolean logic to handle nulls.

*Stage 1: Conditional Extraction (Database Routing & Composite Keys)*
```excel
=IFNA(IF(AND($B3="Investigator",$D3<>""), VLOOKUP($D3&"/"&$C3, Misc!$J$3:$M$122, 4, FALSE), VLOOKUP($B3&"/"&$C3, Classes!$A$2:$J$180, 9, FALSE)), "")
```

*Stage 2: Transformation & UI Presentation*
```excel
=IFNA(IF(AND($B3<>"",$E3<>""), CONCAT($AE3, " || ", $AF3), $AE3), "")
```

**Professional Translation:** This demonstrates a complete ETL Data Pipeline. In a corporate environment, this mirrors extracting data from multiple database tables based on specific system flags (conditional routing), using composite keys for exact record matching, and finally transforming disparate text fields into a single, user-friendly string for a front-end dashboard while gracefully handling null values.

---

## 5. Boolean Eligibility Logic (Array Processing)

**The Challenge:** Evaluating whether a character qualifies for an advanced path requires checking their entire history against a specific list of accepted prerequisites, while allowing for manual overrides.

**The Solution:** This logic uses array evaluation to cross-reference the character's history against a predefined list of qualifying classes, returning a clean binary output.

```excel
=IFNA(
  IF(
    SUMPRODUCT(--(Character!$B$3:$B$22={"Smart ","Dedicated ","Charismatic ","Infiltrator ","Personality ","Explorer "})) > 0, 
    "Yes",
    IF(AG31=B31, "Yes", "No")
  ), 
  "No"
)
```

**Professional Translation:** This utilizes **Array Constants and Double Unaries (`--`)** for high-efficiency eligibility checking. Instead of chaining massive `OR` statements, the logic compares the dataset against a defined array, converting Boolean `TRUE/FALSE` responses into integers for mathematical evaluation. This is highly scalable logic often used in HR systems or authorization gates.

---

## Development Methodology
The logic within this repository demonstrates my ability to:
* Architect complex systems with decoupled data and presentation layers.
* Audit, refactor, and implement advanced formula solutions using modern Excel functions (`LET`, `UNIQUE`, dynamic arrays).
* Design resilient systems with comprehensive error handling (`IFNA`) and dynamic scalability.

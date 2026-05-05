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
=@LET(
  u, UNIQUE($B$3:$B$22),
  m, MAXIFS($C$3:$C$22, $B$3:$B$22, u),
  SUM(MAXIFS(K$3:K$22, $B$3:$B$22, u, $C$3:$C$22, m))
)
```

**Professional Translation:** This demonstrates **Data Normalization**. In a business context, this is equivalent to querying a transactional database to find the most recent interaction for a list of unique customers before aggregating their lifetime value. Using `LET` optimizes processing time by preventing recalculation of the `UNIQUE` array.

---

## 2. Dynamic Rules Engines (Context-Aware Attribution)

**The Challenge:** Calculating an attack bonus requires the system to "know" the character's capabilities (feats) and the properties of the tool (weapon) simultaneously, applying different modifiers based on intersecting conditions.

**The Solution:** This formula uses nested logic and wildcard searching to build a conditional rules engine. It defaults to a standard modifier but dynamically switches if specific prerequisites are met.

```excel
=IF($D35>"", 
  IF(
    OR(
      $A35="Ranged", 
      AND(
        COUNTIF($F$3:$G$22, "*" & "Weapon Finesse" & "*"), 
        $A35="Light"
      )
    ), 
    CONCAT("+", SUM(NUMBERVALUE($C$24), NUMBERVALUE($B35))), 
    CONCAT("+", SUM(NUMBERVALUE($C$23), NUMBERVALUE($B35)))
  ), 
"")
```

**Professional Translation:** This mirrors **Automated Compliance Auditing or Incentive Logic**. The system evaluates a complex Boolean matrix (If Tool = X OR (Tool = Y AND User has Certification Z)) to determine the correct output. `NUMBERVALUE` enforces data typing, while `CONCAT` ensures a clean, user-friendly UI presentation.

---

## 3. Financial Modeling & Tiered Scaling (Wealth DC)

**The Challenge:** The d20 Modern system uses abstract "Purchase DCs" rather than actual dollar amounts. The sheet needed to translate an abstract tier into a usable, average real-world currency value.

**The Solution:**

```excel
=IFNA(MROUND(VLOOKUP($A$2, $K$5:$N$84, 4, FALSE), 5), "")
```

**Professional Translation:** This handles **Tiered Numerical Transformation**. Similar to translating credit score bands into interest rates, this logic extracts baseline data from a reference table and applies `MROUND` to clean and standardize the output for end-user readability (e.g., rounding to the nearest $5). 

---

## 4. Relational Data Architecture (Composite Keys)

**The Challenge:** Handling "Gestalt" characters requires retrieving data for a combination of two distinct classes. Since neither class alone is a unique identifier for the combined stats, standard lookups fail.

**The Solution:**

```excel
=IFNA(VLOOKUP($B3 & "/" & $C3, Classes!$A$2:$J$180, 4, FALSE), "")
```

**Professional Translation:** This is the spreadsheet equivalent of a **SQL JOIN on multiple conditions**. By concatenating two variables (`$B3 & "/" & $C3`), I created a **Composite Key** to ensure exact, unique record retrieval from a decoupled back-end database sheet (`Classes`).

---

## 5. Boolean Eligibility Logic (Array Processing)

**The Challenge:** Evaluating whether a character qualifies for an advanced path requires checking their entire history against a specific list of accepted prerequisites, while allowing for manual overrides.

**The Solution:**

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

# Curriculum progression pilot

This pilot examines what students are asked to do with data across **FA515 (semester 3) → FA503 (semester 4) → FA504 (semester 5)** at SDU. It follows two capabilities: preparing experimental data, and fitting and interpreting models. The purpose is to identify where expectations develop, where methods are practised again, and where the intended connection between courses needs clarification.

**Status: initial mapping from teaching materials, awaiting teacher review.** The pilot describes curriculum expectations, not student attainment. FA516 is outside the pilot because it is elective. The existing category mapping and published overview are unchanged.

## Working files

- [../assets/data/Pharma_DS_Progression.xlsx](../assets/data/Pharma_DS_Progression.xlsx): proposed six-semester expectations across all eight categories, with a separate teacher review record and candidate handovers. The proposal is broader than the sampled pilot; its timing and expectations need agreement. It is rendered alongside the exercise evidence in [../clean.html#pathway](../clean.html#pathway).
- [evidence.csv](../assets/data/evidence.csv): 11 provisional capability records from six activities. Open in a spreadsheet and enable text wrapping. Each record links to the exact source version and lines.
- [review.md](review.md): course-specific questions, a review record, and a joint decision agenda.
- [sources.json](../assets/data/sources.json): source paths, archive revision identifiers, file checksums, and retrieval date.

The sample covers two activities per course. It does not establish the full scope of any course. The public repositories contain R/Quarto materials; instructions, teaching support, and assessment requirements in itslearning need teacher confirmation.

## What to map

**Preparing experimental data** includes importing and inspecting tables, arranging observations, applying units and dilution corrections, and constructing quantities needed for analysis. It maps principally to the existing Data handling category, with Computing and programming where students implement the calculations.

**Fitting and interpreting models** includes specifying transformations and fits, selecting observations or models, interpreting parameters, and checking whether conclusions follow from the data. It maps principally to Analysis and modelling and Mathematics and statistics.

Record the expectation for a particular task, together with the support available. A complicated method with complete code can demand less implementation independence than a simple calculation students must construct. Repetition can consolidate a capability without increasing its level.

## Shared descriptors

Use these proposed descriptors separately; do not add them into a course score. They describe task demands and are not a validated scale.

| Dimension | Descriptor | Observable demand |
|---|---|---|
| Independence | `supplied` | Run or adapt parameters in a substantially complete implementation. |
| Independence | `completion` | Write missing calculations or code within a supplied sequence. |
| Independence | `planned` | Decide the analysis sequence and implement it without a task-specific recipe. |
| Judgement | `prescribed` | Carry out specified operations; no substantive method or validity decision is explicit. |
| Judgement | `bounded` | Make a defined choice or evaluate a stated condition, such as selecting a fitting interval or checking linearity. |
| Judgement | `open` | Justify an analysis approach and evaluate alternatives and limitations without a supplied decision procedure. |
| Transfer | `matched` | Apply the immediate example or procedure to closely matching data. |
| Transfer | `prompted` | Reuse a previously taught method in another context with an explicit cue connecting them. |
| Transfer | `unprompted` | Identify and adapt a relevant earlier method without that cue. |

Use `not_established` when the material does not resolve a dimension. In particular, similar methods in two courses do not by themselves establish that students are asked to transfer learning. Descriptors apply only to the work described in that row; split a row if important parts demand different kinds of performance. Code blanks alone establish neither conceptual difficulty nor independence in class.

## Initial comparison

| Capability | FA515 · semester 3 | FA503 · semester 4 | FA504 · semester 5 |
|---|---|---|---|
| Preparing experimental data | Complete imports and transformations; the vapour-pressure guide also addresses spreadsheet layout and instrument correction. | Dissolution supplies the cumulative correction; dosage uniformity requires students to complete assay and dilution calculations. | Caco-2 asks students to write cumulative transport and flux calculations; ASA stability combines experimental data from four temperatures. |
| Fitting and interpreting models | Kinetics asks students to compare reaction-order fits and inspect plots; vapour pressure explicitly reuses linearisation. | The sampled dissolution guide includes an **optional**, fully coded Weibull fit. The uniformity task applies acceptance rules rather than fitting a model. | Caco-2 asks students to select a steady-state interval; ASA stability specifies first-order and Arrhenius fits and asks students to check linearity. |

These observations support three questions for the pilot:

1. **Does implementation become less supplied?** The clearest candidate is the move from complete sample-removal correction in FA503 to writing the corresponding calculation in FA504. The Caco-2 guide explicitly identifies the earlier dissolution work. Compare records `503-D-P` and `504-C-P`.
2. **Where does modelling recur?** FA515 kinetics and FA504 ASA stability both use rate constants, Arrhenius fits, and shelf-life calculations. Their relationship is a candidate connection to confirm with teachers; the FA504 guide does not explicitly name FA515. Compare `515-K-M` and `504-A-M`.
3. **Which decisions remain with students?** FA515 already includes model comparison. FA504 adds a fitting-interval decision, while its ASA task supplies the reaction order. This is a change in the kind of judgement required, not evidence of a uniform increase.

The sampled FA503 modelling activity is optional. Do not use it as a required intermediate step, or infer that the course contains no other required modelling.

## Complete the pilot

1. Ask each course teacher to review their records using [review.md](review.md). Confirm the teaching year, actual use, required or optional participation, support, and expected output. Record corrections and additional evidence.
2. Discuss the two capabilities across all three courses. Agree what a later course may assume from an earlier one. Identify deliberate repetition, increased demands, and connections that remain undocumented.
3. Record each agreed connection with its two evidence IDs, dimension, rationale, and required status. Keep proposed teaching changes separate from descriptions of current practice.
4. Only then prepare a progression view for the overview. Show capabilities by semester, with descriptions and evidence links. Distinguish unreviewed evidence, optional opportunities, and teacher-confirmed required activities.

The pilot is ready for integration when all six sampled activities have a review outcome and each claimed progression has a documented comparison. Rejected or unused activities remain in the record, excluded from conclusions. A finding of consolidation or insufficient evidence is a valid result.

## Evidence and review conventions

`review_status` starts as `unreviewed`; use `confirmed`, `revised`, or `excluded` after review and record who reviewed it and when in `review.md`. `participation` starts as `unconfirmed`, except where the source explicitly says `optional`; change it to `required` only after confirmation. Optional work never establishes an expectation for all students.

Sources were retrieved on 5 November 2026. Links use the revision recorded in each downloaded archive, not the changing `main` branch. A snapshot of current materials describes an intended course sequence; it does not prove that one cohort encountered these exact versions. Keep the teaching year in the review record and add a new source entry if a revised exercise is reviewed. Workbook DS Version values describe course upgrades and are not progression levels.

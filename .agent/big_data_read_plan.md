FluentNPOI.Streaming — Architecture Specification
1. Purpose / Goal

This module provides a streaming (low-memory) Excel reading architecture for the FluentNPOI ecosystem.

Core principles:

Streaming module is READ-ONLY

Writing / formatting Excel is out of scope

Streaming must NOT load the entire workbook into memory

Streaming must be decoupled from NPOI internals

Fluent API is provided on top of a streaming pipeline

This module is intended for large Excel files (hundreds of thousands to millions of rows).

2. High-Level Design
FluentNPOI
│
├─ FluentNPOI.Core
│   └─ (existing workbook / sheet / writer APIs)
│
├─ FluentNPOI.Streaming
│   ├─ Abstractions
│   ├─ Rows
│   ├─ Mapping
│   ├─ Pipeline
│   └─ Extensions
│
└─ FluentNPOI.Streaming.ExcelDataReader
    └─ (concrete streaming reader implementation)


Key rule:

Streaming reads data → produces rows/DTOs → FluentNPOI.Core may consume results for writing

Dependency direction is strictly one-way.

3. Layer Responsibilities
3.1 Abstractions (Lowest layer – stable)

Purpose:

Define streaming contracts

No knowledge of Excel formats, DTOs, or Fluent APIs

Contains:

Interfaces for streaming readers

Interfaces for row/value access

Rules:

No reference to NPOI

No reference to ExcelDataReader

No business logic

3.2 Rows (Row access model)

Purpose:

Represent a single logical row in a streaming Excel file

Characteristics:

Read-only

Index-based column access only (0,1,2…)

No A1-style addressing

No style, formula, or merge handling

This is the primary data unit exposed to higher layers.

3.3 Mapping (Row → DTO)

Purpose:

Convert streaming rows into strongly-typed DTOs

Responsibilities:

Column-to-property mapping

Type conversion

Null/default handling

Validation (optional)

Non-responsibilities:

IO

Batching

Writing Excel

Design notes:

Mapping logic must be deterministic and testable

Mapping must not depend on streaming implementation

3.4 Pipeline (Fluent chaining)

Purpose:

Provide fluent, composable operations on a row stream

Conceptual flow:

Read → Select Sheet → Skip → Filter → Map → Batch → Consume


Characteristics:

Pipeline is lazy (iterator-based)

Each step wraps the previous enumerable

No buffering unless explicitly requested (e.g. Batch)

Pipeline does NOT:

Know how Excel is read

Know how Excel is written

3.5 Extensions (Syntax sugar)

Purpose:

Improve developer ergonomics

Examples:

SkipHeader

WhereRow

SelectRow

Batch helpers

Rules:

No business logic

No state

Pure pipeline transformations

4. Streaming Reader Implementations

Concrete Excel readers are implemented in separate packages.

Example:

FluentNPOI.Streaming.ExcelDataReader


Rules:

Depends on FluentNPOI.Streaming abstractions

Implements streaming reader interfaces

No Fluent API logic inside implementation

Replaceable without changing core architecture

This enables future support for:

OpenXML streaming

CSV streaming

Custom data sources

5. Explicit Non-Goals (Important)

This module intentionally does NOT support:

Writing Excel files

Editing existing Excel files

Cell styles or formatting

Formula evaluation

Merged cell reconstruction

Random access (rows are forward-only)

These responsibilities remain in FluentNPOI.Core.

6. Dependency Rules (Strict)

Streaming core MUST NOT depend on NPOI

Streaming core MUST NOT depend on ExcelDataReader

Writer/Core MUST NOT depend on Streaming

Concrete readers depend on Streaming abstractions only

Violation of these rules is considered an architecture error.

7. Conceptual Summary

FluentNPOI.Streaming is a streaming data pipeline,
not an Excel object model.

It treats Excel as:

A forward-only row stream

And exposes:

Fluent, composable operations over that stream
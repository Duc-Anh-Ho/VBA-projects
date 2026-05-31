# Gemini Conversation Context: Danh-Tools Project

## Next Action / Current Status

**Context:** We have finalized all project naming conventions, which are now documented in `CODING_CONVENTIONS.md`. Our agreed-upon standard is `PascalCase` for all names, with `Interface_` prefixes for interfaces and `Utils_` prefixes for utility classes.

**Next Action:** The immediate priority is to begin refactoring `FilesController.cls` by separating its generic utility functions.

**To-Do:**
1.  Create a new class file named `Utils_FileSystem.cls`.
2.  Move the `isValidVBAFileExtension` and `copyFileName` functions from `FilesController.cls` into the new `Utils_FileSystem.cls`.
3.  Refactor `FilesController.cls` to use the new utility class.

---

## Project Overview

This project, named "Danh-Tools", is a comprehensive VBA add-in for Microsoft Excel. It is well-structured, using a controller pattern to separate different areas of functionality. The project is designed to be modular and maintainable, with a clear separation of concerns.

## Key Accomplishments

1.  **`README.md` Creation:** I created a `README.md` file that documents the features of the project.
2.  **`REFACTORING_PLAN.md` Creation:** I created a prioritized refactoring plan to help improve the project's architecture and code quality, based on SOLID principles, OOP, and clean code best practices.
3.  **`AGILE_WORKFLOW.md` Creation:** We established a formal agile workflow for managing the project.
4.  **`CODING_CONVENTIONS.md` Updates:** We have had several discussions and updates to the coding conventions, finalizing our approach to naming and best practices.

---

## Test Class Refactoring (2025-10-09)

We began refactoring the testing framework (`TestUtils.cls`, `TestCode.cls`). This task is currently on hold while we address higher-priority architectural refactoring.

**Staged Files (Pending Review):**

*   `new_TestAddress.cls`
*   `new_TestUtils.cls`
*   `new_TestCode.cls`
*   `new_TestRunner.bas`

## Gemini & User Interaction Rules

*   When asked to apply changes to existing files, I will create new files with the prefix `new_` for comparison, instead of modifying the original files, unless specified otherwise.
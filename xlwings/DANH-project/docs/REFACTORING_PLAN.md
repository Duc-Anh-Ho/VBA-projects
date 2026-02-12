# Prioritized Refactoring Plan for Danh-Tools

## Introduction

This document provides a prioritized list of refactoring tasks to improve the quality, maintainability, and architecture of the Danh-Tools project. The tasks are ranked by importance and effort, allowing you to tackle the most impactful and easiest changes first.

## Prioritization Matrix

| Importance | Low Effort | Medium Effort | High Effort |
| :--- | :--- | :--- | :--- |
| **High** | 1. Use String-Specific Functions<br>2. Centralize Configuration<br>3. Improve Naming<br>4. Decompose `SystemUpdate` (Step 1) | 5. Implement Dependency Injection | 6. Refactor `SystemUpdate` God Class |
| **Medium** | 7. Add Explanatory Comments | 8. Adhere to the Open/Closed Principle (OCP) | 9. Write Unit Tests |
| **Low** | | | |

---

## Refactoring Tasks

### 1. Use String-Specific Functions (e.g., `Left$`)

-   **Priority:** High
-   **Effort:** Low

**Description:**

Replace generic `Variant` functions with their `String`-specific `$` counterparts (e.g., `Left` -> `Left$`) for better performance and type safety, as per the updated `CODING_CONVENTIONS.md`. The following violations were found via manual review:

*   **`Format(`**
    *   `EmailCDO.cls`
    *   `Shortcuts.bas`
*   **`Left(`**
    *   `PicturesController.cls`
    *   `Shortcuts.bas`
*   **`LTrim(`**
    *   `KeyboardShortcutForm.frm`
*   **`Str(`**
    *   `SystemUpdate.cls`
*   **`Trim(`**
    *   `MultipleReplaceForm.frm`
    *   `Shortcuts.bas`


### 2. Refactor `SystemUpdate` God Class (SRP Violation)

-   **Priority:** High
-   **Effort:** High

**Description:**

The `SystemUpdate` class violates the **Single Responsibility Principle (SRP)**. It currently acts as a "God Class," handling many unrelated tasks, which makes it difficult to maintain and test.

**Current Responsibilities:**
*   **Application State:** `speedOn`, `speedOff`
*   **Object Creation:** `createFileSystem`, `createWShell`, `createClipboard`
*   **System Information:** `getLastRow`, `getLastColumn`
*   **Error Handling:** `tackleErrors`
*   **UI Management:** `setStatusBar`

**Recommendation:**
Break `SystemUpdate` into smaller, focused service classes as outlined in the architecture plan. For example:
*   `ExcelApplicationService` (for `speedOn`, `speedOff`, `setStatusBar`)
*   `ScriptObjectFactory` (for all `CreateObject` calls)
*   `SheetUtils` (for `getLastRow`, `getLastColumn`)


### 3. Centralize Configuration & Remove Hardcoded Values

-   **Priority:** High
-   **Effort:** Low

**Description:**

Many modules contain hardcoded "magic strings" and paths. This makes the code rigid and hard to configure.

**Examples:**
*   **`ShareXController.cls`**: The path to the ShareX executable is hardcoded.
    ```vba
    Private Const SHAREX_PATH As String = "D:\Program Files\ShareX\"
    ```
*   **`FilesController.cls`**: The folder names for exporting VBA files are hardcoded.
    ```vba
    Private Const VBA_FOLDER As String = "\VBA-files-"
    Private Const MODULE_FOLDER As String = "\Modules\"
    ```
*   **`InternetConnector.cls`**: The default ping link is hardcoded.
    ```vba
    Private Const DEFAULT_LINK As String = "google.com"
    ```

**Recommendation:**
Create a `Config` worksheet and a `ConfigurationService` to read these values at runtime. This will make your add-in configurable without changing the code.


### 4. Implement Dependency Injection to Reduce Coupling

-   **Priority:** High
-   **Effort:** Medium

**Description:**

Your classes are tightly coupled because they create their own dependencies. This makes them difficult to test in isolation and reuse.

**Example (found in almost every controller):**
In `ChartsController.cls`, `FilesController.cls`, `PicturesController.cls`, and others:
```vba
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    ' ...
End Function
```
Here, the controller is directly responsible for creating its `InfoConstants` and `SystemUpdate` objects.

**Recommendation:**
Use the Dependency Injection (DI) pattern we established with `Utils_Address` and `Interface_Address`. Dependencies should be passed into the class during its initialization, not created within it.


### 5. Adhere to the Open/Closed Principle (OCP)

-   **Priority:** Medium
-   **Effort:** Medium

**Description:**

The Open/Closed Principle states that software entities should be open for extension but closed for modification. Your use of `Select Case` on object types violates this, as you must modify the procedure every time you want to support a new object type.

**Example:**
In `PicturesController.cls`:
```vba
Public Sub lockRatio(ByRef targetObject As Object)
    Select Case TypeName(targetObject)
        Case "Shape"
            Let targetObject.LockAspectRatio = isLockRatio
        Case Else
            Let targetObject.ShapeRange.LockAspectRatio = isLockRatio
    End Select
End Sub
```
**Recommendation:**
Use polymorphism. Create a common interface (e.g., `ILockable`) and have different "wrapper" classes for each object type (`ShapeWrapper`, `ChartWrapper`) that implement this interface. The controller can then work with the `ILockable` interface without needing to know the specific object type.


### 6. Write Unit Tests

-   **Priority:** Medium
-   **Effort:** High

**Description:**

Automated tests are essential for ensuring the quality and stability of your project, especially during a major refactoring. The changes you make in the steps above (especially Dependency Injection and refactoring `SystemUpdate`) will make your code much easier to test.

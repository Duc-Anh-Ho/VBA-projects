# Danh-Tools Coding Conventions

## 1. Introduction

This document outlines the official coding standards and best practices for the Danh-Tools VBA project. Adhering to these conventions is crucial for maintaining code quality, readability, and long-term maintainability.

## 2. Guiding Principles

- **Clarity:** Code should be as easy to read and understand as possible. Explicit is better than implicit.
- **Safety:** Code should be robust and free of subtle bugs or potential runtime errors.
- **Consistency:** The style and patterns used should be consistent across the entire project.

## 3. Language Best Practices

### 3.1. Use Fully Qualified Function Calls

To prevent ambiguity and make the code more self-documenting, always use the fully qualified path for built-in VBA functions.

- **Rule:** Specify the library and module (e.g., `VBA.FileSystem.Dir`).
- **Reason:** This practice guarantees that you are calling the correct function and protects against naming conflicts with custom functions or other libraries. It also improves clarity by grouping functions by their domain (`FileSystem`, `Strings`, `Interaction`).

**Example:**

```vba
' RECOMMENDED (Best)
Dim pathExists As Boolean
pathExists = VBA.FileSystem.Dir(PathName:=Path)

' AVOID (Good, but not the best)
pathExists = VBA.Dir(Path)

' AVOID (Can be ambiguous)
pathExists = Dir(Path)
```

### 3.2. Prefer String-Specific Functions (e.g., `Left$`)

To improve performance and type safety, always prefer the string-specific version of built-in functions, which are identifiable by the `$` suffix.

- **Rule:** Use `Left$`, `Mid$`, `Chr$`, etc., instead of `Left`, `Mid`, `Chr`.
- **Reason:** The `$` versions operate on and return `String` data types directly. This is faster and more memory-efficient than the `Variant` versions, which have additional processing overhead. It also makes the code more explicit and type-safe.
- **Exception:** The only time to use the `Variant` version (e.g., `Left`) is when you are intentionally handling a variable that may contain a `Null` value, as the `$` versions will raise an error in that case.

**Example:**

```vba
' RECOMMENDED (Faster, Type-Safe)
Dim myString As String
myString = Left$("Hello World", 5) ' Returns a String: "Hello"

' AVOID (Unless handling Nulls)
Dim myVariant As Variant
myVariant = Left("Hello World", 5) ' Returns a Variant/String: "Hello"
```

### 3.3. Conditional Logic: `If...Else` vs. `IIf`

The `IIf` function can introduce bugs because it always evaluates both its true and false arguments. Because of this risk, the following distinction applies:

*   **Core / Utility Modules (`...Controller`, `...Utils`, etc.):** In the core framework files, the use of `IIf` is **strictly forbidden**. All conditional logic **must** use the standard `If...Then...Else...End If` block to ensure maximum safety and predictability.

*   **Implementation Modules (e.g., `TestCode`, project-specific classes):** In modules written for specific, custom implementations, the use of `IIf` and abbreviated names is **permitted** for simple assignments, assuming the developer is aware of the risks. However, the `If...Then...Else...End If` block remains the **recommended practice** for clarity and safety.

### 3.4. Use the Correct Access Modifier (Public, Friend, Private)

To ensure proper encapsulation and a clean project API, it's critical to use the most restrictive access modifier possible.

- **Rule:** Default to `Private`. Only use `Friend` or `Public` when necessary.

#### `Private`
- **When to use:** For any method or property that is only used *inside* the class where it is defined. This should be your default choice.
- **Purpose:** Hides internal implementation details.

#### `Friend`
- **When to use:** For methods or properties that need to be called by other modules *within this project*, but should NOT be exposed to external projects (e.g., other workbooks).
- **Purpose:** Defines the internal API of your add-in. `TesterUtil` is a perfect example of a class whose methods should be `Friend`.

#### `Public`
- **When to use:** Only for methods and properties that are meant to be the official, external-facing API of your add-in.
- **Purpose:** Exposes functionality to be consumed by other VBA projects or applications.

By following this "Principle of Least Privilege," you create a much cleaner and more robust separation between your project's internal workings and its public interface.

## 4. Naming Conventions

### 4.1. Class Prefixes for Grouping

To ensure clarity, safety, and organization in the VBA project explorer, we will use explicit, verbose prefixes for special class types.

-   **Utility Classes:** `Utils_` prefix (e.g., `Utils_FileSystem`, `Utils_Address`).
-   **Interfaces:** `Interface_` prefix (e.g., `Interface_Address`).

**Reasoning:** While non-standard compared to other languages, this approach provides maximum clarity and safety within the project. The explicit prefixes make it easy to identify a class's role and prevent accidental misuse or deletion during maintenance. This consistency in prefix *style* (`Word_`) is prioritized for this project.

### 4.2. Casing Convention

- **Rule:** All file names and the code names within them (classes, interfaces, modules) must use **`PascalCase`**. Variable names should use **`camelCase`**.

**Examples:**
- **Class:** `Utils_FileSystem`
- **Interface:** `Interface_Address`
- **Variable:** `Dim fileSystem As Utils_FileSystem`

## 5. Architectural Principles

### 5.1. The Single Responsibility Principle (SRP)

A class should have one, and only one, reason to change. This means every class should have a single, well-defined job or responsibility.

- **Guideline:** Instead of counting lines of code, focus on making each class do just one thing well.
- **The "And" Test:** If you describe what your class does and you have to use the word "and" (e.g., "It manages the file system **and** the clipboard **and** the application state"), it's a sign that the class has too many responsibilities and should be split into smaller, more focused classes.
- **Analogy:** Avoid creating "Swiss Army Knife" classes that do many unrelated things. Instead, create a "Toolbox" of specialized classes that each do one job perfectly.

**Example from our project:**

-   **Bad (violates SRP):** The original `SystemUpdate.cls` handles too many different jobs.
-   **Good (follows SRP):** The new `ScriptObjectUtils.cls` has only one job: to create scripting objects.

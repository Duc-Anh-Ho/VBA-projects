# Architectural Refactoring Plan for Danh-Tools

## 1. Introduction

This document provides a comprehensive plan for refactoring the Danh-Tools VBA project. The goal is to evolve the existing codebase into a more robust, scalable, and maintainable application by applying professional software engineering principles, including Object-Oriented Programming (OOP), modern design patterns, and clean code practices.

While the project is well-structured and modular, these recommendations will help elevate it to a professional standard, making it easier to extend, test, and maintain in the long run.

## 2. High-Level Architectural Vision: A Service-Oriented Approach

The current architecture is controller-based, which is a great start. We can enhance this by moving towards a more **Service-Oriented Architecture (SOA)** within the VBA environment.

This means we will:

1.  **Decompose "God" Objects:** Break down monolithic classes like `SystemUpdate` into smaller, single-responsibility services (e.g., `FileSystemService`, `ClipboardService`, `ExcelApplicationService`).
2.  **Implement Dependency Injection (DI):** Instead of controllers creating their own dependencies (e.g., `New SystemUpdate`), these services will be passed into them. This is the cornerstone of a loosely coupled and testable system.
3.  **Establish Clear Layers:** We will formalize the layers of the application:
    *   **Presentation/UI Layer:** Forms and the Ribbon XML. This layer should only contain UI logic.
    *   **Application Layer (Controllers):** The existing `...Controller` classes. They will act as orchestrators, taking user input from the UI layer and using services from the layer below to perform actions. They will not contain business logic themselves.
    *   **Domain/Service Layer:** The new, smaller service classes that encapsulate the core business logic and interactions with Excel or the system.
    *   **Data/Infrastructure Layer:** This layer would handle data persistence if needed (e.g., reading/writing to the `Config` sheet).

This layered approach, combined with DI, will significantly improve the project's structure and testability.

## 3. Prioritized Refactoring Roadmap

Here is a step-by-step roadmap, starting with the highest-impact, lowest-effort changes.

### Step 1: Centralize Configuration (High Priority, Low Effort)

**Problem:** Hardcoded values (e.g., paths, constants) are scattered throughout the codebase, making it difficult to configure and maintain.

**Solution:**
1.  **Create a `Config` worksheet:** A hidden sheet to store all configuration key-value pairs.
2.  **Create a `ConfigurationService` class:** This class will be responsible for reading the configuration from the `Config` sheet on startup and providing a simple `GetValue(key)` method.
3.  **Refactor controllers to use `ConfigurationService`:** Use Dependency Injection to provide the `ConfigurationService` to any class that needs it.

### Step 2: Implement a Dependency Injection (DI) Container (High Priority, Medium Effort)

**Problem:** Classes are tightly coupled because they create their own dependencies (e.g., `Set info = New InfoConstants`). This makes them difficult to test and reuse.

**Solution:** We will create a simple DI "container" to manage the lifecycle of our services.

1.  **Create a `Services` module:** This standard module will act as our DI container.
2.  **Create public factory functions:** For each service, we will create a public function in the `Services` module that returns a single, cached instance of that service (singleton pattern).

    ```vba
    ' In module "Services"
    Private pConfigService As ConfigurationService

    Public Function GetConfigurationService() As ConfigurationService
        If pConfigService Is Nothing Then
            Set pConfigService = New ConfigurationService
        End If
        Set GetConfigurationService = pConfigService
    End Function
    ```
3.  **Refactor controllers to use the service locator:**

    ```vba
    ' In a controller class
    Private Sub Class_Initialize()
        Set config = GetConfigurationService()
        ' ...
    End Sub
    ```

***Note (2025-10-10):** A successful DI pattern was implemented for the `Utils_Address` class. It now depends on a `Interface_Address` interface, with `PJ1_Address` being the concrete implementation that is "injected" during initialization. This serves as an excellent practical template for future DI refactoring across the project.*

### Step 3: Decompose the `SystemUpdate` God Class (High Priority, High Effort)

**Problem:** `SystemUpdate` violates the Single Responsibility Principle by handling everything from file system operations to Excel state management.

**Solution:** Break it down into a set of cohesive, single-responsibility services.

1.  **Identify Responsibilities:** Group the methods in `SystemUpdate` into logical domains.
2.  **Create New Service Classes:** Create the following new classes, moving the relevant methods from `SystemUpdate` into them:
    *   `ExcelApplicationService`: For methods that interact with the `Application` object (e.g., `speedOn`, `speedOff`, `hasWorkPlace`).
    *   `FileSystemService`: For file and folder operations (e.g., `createFileSystem`, `getFolder`).
    *   `ClipboardService`: For clipboard interactions (e.g., `getClipboard`, `setClipboard`).
    *   `ArrayUtilsService`: For array manipulation functions.
3.  **Refactor `SystemUpdate`:** The `SystemUpdate` class can either be eliminated or become a facade that delegates calls to the new services (though elimination is preferred).
4.  **Update Controllers:** Update all controllers to depend on the new, smaller services instead of `SystemUpdate`.

### Step 4: Introduce a Testing Framework (Medium Priority, High Effort)

**Problem:** The lack of automated tests makes refactoring risky and new feature development error-prone.

**Solution:** We will integrate a simple, yet effective, testing strategy.

1.  **Create a `TesterUtil` class:** This class will contain assertion helpers and test running logic.
2.  **Adopt a Naming Convention:** Test procedures will be named `Test_ClassName_MethodName_Scenario`.
3.  **Write the First Tests:** Start by writing tests for the new service classes, as they are decoupled and easy to test.
4.  **Build a Test Suite:** Gradually create a comprehensive suite of tests that can be run automatically to validate the application's behavior.

This plan provides a clear path to a more professional and maintainable architecture. I am ready to start implementing these changes, beginning with Step 1.

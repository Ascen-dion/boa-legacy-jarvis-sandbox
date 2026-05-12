# Bank of America POC: Legacy UFT "Jarvis" Framework Sandbox

## 📌 Project Overview
This repository contains a 1-to-1 localized replica of the legacy Bank of America "Jarvis" automated testing framework. Built using UFT One and VBScript, this sandbox mirrors the exact enterprise architecture currently utilized in production, including keyword-driven XML flows, Excel-based execution controllers, and a custom file-routing execution engine.

---

## 🏗️ Core Architecture & Components
The "Jarvis" framework is a highly structured, Keyword-Driven and Data-Driven hybrid. It relies on decoupling the test data, the execution flow, and the physical UI interactions into separate, modular files.

### 1. The Execution Engine (State Tracking)
Instead of relying purely on a test runner, the framework uses physical directories to manage the state of test execution, simulating an enterprise server queue:
* `NewRequest/`: Where tests queue up before execution.
* `Working/`: Where the active test tracker file is moved during execution.
* `Completed/`: Where the tracker file is stored upon successful run completion.

### 2. Master Control Data (`Master_TestData.xlsx`)
An Excel workbook acts as the brain of the execution. It contains:
* Execution flags (`Yes/No`) dictating which tests run.
* Test Data (Usernames, Passwords, Cities) fed directly into the test functions.
* Mapping to the specific XML Flow Template required for the test case.

### 3. Input Flow Templates (`.xml`)
XML files define the exact sequence of keywords (business steps) for a given scenario. This allows non-technical users to build new test flows without touching VBScript.

### 4. Function Libraries (`.qfl`)
Modular VBScript files containing the actual execution logic. These functions map to the XML keywords and interact with the application strictly via Logical Names defined in the Object Repository.

### 5. Object Repository (`.tsr`)
Strict reliance on binary shared object repositories. Zero physical properties (like `devname` or `text`) are hardcoded in the scripts. The script calls the logical name, and UFT looks up the physical properties in the `.tsr` file.

### 6. Configuration Management (`ConfigTemplates\`)
Hierarchical XML templates separate framework-level rules (timeouts, logging) from application-level targets (environments, database strings), allowing for seamless environment swapping without code modifications.

---

## 🛠️ Framework Execution Instructions

### Prerequisites
* **Micro Focus / OpenText UFT One** installed (Run as Administrator).
* **WPF & .NET Add-ins** enabled in the UFT Add-in Manager.
* **Target Application** (e.g., OpenText Flight GUI Application) installed.

### How to Run a Test Batch
1. Open **UFT One**.
2. Navigate to **File > Open > GUI Test** and open the test package located inside the `DriverScript` folder.
3. Open the `Action1` tab to view the driver code.
4. **CRITICAL:** Ensure Microsoft Excel is completely closed. If the `Master_TestData.xlsx` file is open, it creates a hidden `~$` lock file that will crash the script.
5. Ensure the target application is closed before starting the run.
6. Click the **Run** button (F5) in UFT.
7. Upon completion, navigate to `Reports\Consolidated\` to view the dynamic Extent-style execution dashboard and individual test step traces.
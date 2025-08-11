# MergeDocuments (.NET)

**Merge multiple Word (.docx) files, update content controls and add dynamic footers** using only standard .NET libraries.

---

## Overview
This C# console application:
- Merges multiple `.docx` files into a single consolidated document.
- Replaces content controls (SDTs) with supplied values.
- Adds dynamic footers for each appended document showing the filename and running page number.

> **Note:** DOCX → PDF conversion is **not included** in this solution because converting DOCX to PDF requires either Microsoft Word (Interop) or a third-party library/service. I can add a PDF conversion step if allowed to use Interop or an external library.

---

## Prerequisites
- .NET SDK 8.0+ (or the version used in the project)
- Windows recommended for Interop-based PDF (if later used)

---

## Configuration
Modify replacements for content controls in the `Program.cs` file under the `replacements` dictionary.

---

## How to run
1. Place source `.docx` files in the `Docs` folder (project root).
2. Build and run:
```bash
dotnet build
dotnet run --project ./MergeDocuments
```

---

## Output
After running the application, the merged Word document will be created at: `\bin\Debug\net8.0\Docs\merged.docx`

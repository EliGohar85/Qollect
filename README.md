# Qollect

Collect your Qlik app metadata in one click.

---

## Why Qollect?

Qollect removes the hassle of hunting through your Qlik app to find object definitions or chart metadata. It's perfect for:

- **Documentation & Governance** - Create quick snapshots to support audits, doc updates, or team handoffs.  
- **Development & Impact Analysis** - Check where things are used before making changes; find unused or duplicated objects.  
- **Onboarding & Collaboration** - Share a clear, structured overview of your app with new team members or stakeholders.

---

## Features

- Collects key Qlik app metadata components - dimensions, measures, fields, charts, variables and sheets - into a structured format.  
- **One-click** metadata collection inside your Qlik Sense app.  
- Clean, minimal output - ideal for sharing, reviewing, or maintaining documentation.

---

## Installation

1. Download the latest `Qollect.zip` from the **[Releases page](https://github.com/EliGohar85/Qollect/releases)**.  
2. Open the Qlik Sense Management Console (QMC) and navigate to **Extensions**.  
3. Upload `Qollect.zip`.  
4. In your app, drag the **Qollect** visualization onto a sheet, click it, and you’ll have your metadata in just one click.

---

## Contributions & Feedback

Do you have a feature idea or have you found a bug?  
- Open an issue or submit a PR on **GitHub:** https://github.com/EliGohar85/Qollect

---

## Project Links

- **GitHub (source & releases):** https://github.com/EliGohar85/Qollect
- **Regarden Marketplace:** https://www.regarden.io/extensions/3170c06d-0cbb-462f-a3e6-5303f9763249  
- **Ko-fi (support):** https://ko-fi.com/eligohar

---

## Support

If Qollect saves you time, consider supporting development:  
**Ko-fi:** https://ko-fi.com/eligohar

---

## Changelog

### 1.3.2
- Fixed Fields tagged as $key are now always reported as USED and marked as Key
- Added Full support for Container objects, including nested charts and their dimensions and measures
- Improved Field usage detection for charts inside containers

### 1.3.1
- Fixed detection of field names with spaces (like [Total Sales])
- Fixed detection of unbracketed dotted fields (like Orders.Total)
- Improved overall field parsing logic

### 1.3.0
- Added **"Items"** column to the Charts sheet - showing all associated dimensions, measures, fields, expressions, and alternate items.
- Added new **"Script" sheet** - Summarizes key script activity per tab (including LOAD/STORE/JOIN operations, RESIDENT usage, QVD references, and variables).

### 1.2.0
- Added **Export Script as QVS file** - allows exporting your app’s script directly to a .qvs file for easier reuse and documentation.
- Fixed **alternative dimensions/measures detection** - Qollect now correctly detects and displays whether an alternative dimension or measure was used.

### 1.1.0
- Added **App Overview** sheet (first in workbook) summarizing: # of Dimensions, # of Measures, # of Fields, # of Sheets, # of Charts, # of Variables.  
- Added **field usage mapping** ("Used In": Chart, Set analysis, Dimension, Measure, Variable) and **unused field highlighting**.  
- **Unused master items (Dimensions/Measures):** highlighting and usage counts.

---

## License

Distributed under the MIT License. See `LICENSE` for details.

---

Made with ❤️ by **Eli Gohar**

# Navisworks Plugin Development — Interactive Tutorial

A comprehensive, interactive web-based tutorial for building **Autodesk Navisworks plugins** with **C#** and **.NET Framework**. Covers the full plugin lifecycle — from project setup and multi-version builds to advanced API usage, clash detection, and production deployment.

![Navisworks Versions](https://img.shields.io/badge/Navisworks-2022--2026-blue?style=flat-square)
![.NET Framework](https://img.shields.io/badge/.NET_Framework-4.8-purple?style=flat-square)
![Language](https://img.shields.io/badge/Language-C%23-green?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)

---

## 🚀 Quick Start

No build step, no dependencies, no server required — just open the file:

```
index.html
```

Double-click to open in any modern browser (Chrome, Edge, Firefox).

---

## ✨ Features

| Feature | Description |
|---|---|
| **14 Chapters** | Structured learning path from beginner to advanced |
| **100+ Code Snippets** | Copy-ready C# and XML examples |
| **Interactive Class Diagram** | Clickable API hierarchy with detail modals |
| **Plugin Type Tabs** | Side-by-side comparison of all 4 plugin types |
| **Full-Text Search** | Instant search across all sections (Ctrl+K) |
| **Sidebar Navigation** | Persistent table of contents with scroll tracking |
| **Reading Progress** | Visual progress ring in the top navigation |
| **Interactive Checklist** | Step-by-step "Add New Command" workflow |
| **Dark Theme** | Premium glassmorphism UI with micro-animations |
| **Single-File** | Fully self-contained HTML — CSS & JS inlined |

---

## 📖 Chapters

| # | Chapter | Topics |
|---|---------|--------|
| 01 | **Prerequisites & Setup** | Visual Studio, .NET Framework, Navisworks versions |
| 02 | **Project Structure** | Interactive file tree with tooltips |
| 03 | **Creating the VS Project** | Class Library, multi-version `.csproj` configuration |
| 04 | **API References** | DLL references, conditional ItemGroups, `<Private>False</Private>` |
| 05 | **Plugin Types & Classes** | CommandHandler, AddIn, EventWatcher, DockPane + class hierarchy diagram |
| 06 | **Ribbon UI — XAML** | XAML layout, NWRibbonButton, three-way sync rule |
| 07 | **Deployment** | Post-build events, AppData plugin folder, F5 debugging |
| 08 | **Working with the API** | Document, Selection, Search, Properties, COM write-back, Model Traversal |
| 09 | **Clash Detection** | ClashTest, SelectionA/B, TestsClear() bug workaround |
| 10 | **Advanced Features** | Selection Sets, Viewpoints, Color Override, Bounding Boxes |
| 11 | **Data Models** | POCO classes — ElementInfo, ElementPosition |
| 12 | **Debugging & Troubleshooting** | Attach to process, common gotchas table |
| 13 | **New Command Checklist** | Interactive checklist with progress tracking |
| 14 | **Quick Reference Card** | At-a-glance summary tables for plugin types, attributes, API entry points |

---

## 🛠 Tech Stack

The tutorial itself is built with:

- **HTML5** — Semantic structure
- **CSS3** — Custom properties, glassmorphism, grid, animations
- **Vanilla JavaScript** — IntersectionObserver, clipboard API, keyboard shortcuts
- **[Google Fonts](https://fonts.google.com/)** — Inter, Outfit, JetBrains Mono (CDN)
- **[Highlight.js](https://highlightjs.org/)** — C# and XML syntax highlighting (CDN)

> **No frameworks, no build tools, no npm.** Just open `index.html`.

---

## 📁 Project Structure

```
Naviswork API toturial/
├── index.html                      # Complete tutorial (single-file SPA)
├── NAVISWORKS_PLUGIN_TUTORIAL.md   # Source markdown reference
└── README.md                       # This file
```

---

## 🎯 Who Is This For?

- **.NET developers** new to the Navisworks API
- **BIM specialists** looking to automate workflows with custom plugins
- **AEC tech teams** building internal tooling on Navisworks Manage
- **Students** learning Autodesk plugin architecture

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + K` | Focus search bar |
| `Escape` | Close search / modals |

---

## 📋 Key API Patterns Covered

```csharp
// Get active document
Document doc = Application.ActiveDocument;

// Search by property
Search search = new Search();
search.SearchConditions.Add(
    SearchCondition.HasPropertyByDisplayName("Element", "Material")
        .EqualValue(new VariantData("Concrete")));
search.Selection.SelectAll();
ModelItemCollection results = search.FindAll(doc, true);

// Write properties via COM bridge (managed API is read-only)
ComApiBridge.State.SetUserDefined(item, "MyCategory", "MyProp", "MyValue");
```

---

## 👤 Author

**Bassel Taha**
Full Stack Developer & BIM Automation Specialist | Civil Engineering Background

3+ years building production-grade AEC technology — from cloud microservices on Azure AKS to desktop add-ins and web apps with .NET, React, C#, and Python. Expert in Autodesk Platform Services (APS), BIM workflow automation, and integrating Generative AI into engineering pipelines.

- 📧 [BasselTaha98@gmail.com](mailto:BasselTaha98@gmail.com)
- 📱 +20 122 648 0472
- 🔗 [GitHub — Bassel-Taha](https://github.com/Bassel-Taha)
- 💼 [LinkedIn — Bassel-Taha-Keshk](https://linkedin.com/in/Bassel-Taha-Keshk)
- 📍 UAE — Abu Dhabi

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

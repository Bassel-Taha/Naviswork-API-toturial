# Navisworks Plugin Development — Comprehensive Tutorial

> A complete, generic guide to building Autodesk Navisworks plugins with C# and .NET.  
> Covers everything from scratch: project setup, plugin types, ribbon UI, commands, model access, clash detection, Excel integration, logging, deployment, and debugging.

---

## Table of Contents

1. [Prerequisites & Environment Setup](#1-prerequisites--environment-setup)
2. [Project Structure Overview](#2-project-structure-overview)
3. [Creating the Visual Studio Project](#3-creating-the-visual-studio-project)
4. [Navisworks API References (Multi-Version)](#4-navisworks-api-references-multi-version)
5. [NuGet Packages](#5-nuget-packages)
6. [Plugin Types Explained](#6-plugin-types-explained)
7. [CommandHandlerPlugin — The Primary Plugin Class](#7-commandhandlerplugin--the-primary-plugin-class)
8. [AddInPlugin — The Automation Plugin](#8-addinplugin--the-automation-plugin)
9. [EventWatcherPlugin — Application Menu & Lifecycle Hooks](#9-eventwatcherplugin--application-menu--lifecycle-hooks)
10. [Registering Commands with Attributes](#10-registering-commands-with-attributes)
11. [Ribbon UI — The XAML File](#11-ribbon-ui--the-xaml-file)
12. [Deploying to Navisworks (Post-Build Event)](#12-deploying-to-navisworks-post-build-event)
13. [app.config — Assembly Binding Redirects](#13-appconfig--assembly-binding-redirects)
14. [Working with the Navisworks API](#14-working-with-the-navisworks-api)
    - [14.1 Accessing the Active Document](#141-accessing-the-active-document)
    - [14.2 Getting Selected Items](#142-getting-selected-items)
    - [14.3 Searching the Model with Search API](#143-searching-the-model-with-search-api)
    - [14.4 Reading Properties from ModelItems](#144-reading-properties-from-modelitems)
    - [14.5 Writing / Setting Properties via COM API](#145-writing--setting-properties-via-com-api)
    - [14.6 Traversing the Model Hierarchy](#146-traversing-the-model-hierarchy)
15. [Clash Detection API](#15-clash-detection-api)
16. [Excel Integration (EPPlus & ExcelDataReader)](#16-excel-integration-epplus--exceldatareader)
17. [Progress Forms (Windows Forms UI)](#17-progress-forms-windows-forms-ui)
18. [Logging with Serilog](#18-logging-with-serilog)
19. [WPF Command Handler for Application Menu](#19-wpf-command-handler-for-application-menu)
20. [Data Models](#20-data-models)
21. [Adding New Commands — Step-by-Step Checklist](#21-adding-new-commands--step-by-step-checklist)
22. [Debugging Tips](#22-debugging-tips)
23. [Common Gotchas & Troubleshooting](#23-common-gotchas--troubleshooting)
24. [Advanced API Features & Plugins](#24-advanced-api-features--plugins)
    - [24.1 DockPanePlugin — Custom Dockable Windows](#241-dockpaneplugin--custom-dockable-windows)
    - [24.2 Selection Sets](#242-selection-sets)
    - [24.3 Saved Viewpoints](#243-saved-viewpoints)
    - [24.4 Overriding Item Colors & Transparency](#244-overriding-item-colors--transparency)
    - [24.5 Bounding Boxes & Geometry Constraints](#245-bounding-boxes--geometry-constraints)

---

## 1. Prerequisites & Environment Setup

### Required Software

| Tool | Notes |
|------|-------|
| Visual Studio 2019/2022 | Community edition is fine |
| Autodesk Navisworks Manage | Version 2022–2026 supported |
| .NET Framework 4.8 (or 4.7.2 for NW2022) | Must match Navisworks version |
| NuGet Package Manager | Included in Visual Studio |

### Supported Navisworks Versions

This tutorial targets **NW 2022 through 2026** using conditional build configurations. Each version has its own DLLs in a separate install path.

### DLL Locations

```
C:\Program Files\Autodesk\Navisworks Manage 2022\
C:\Program Files\Autodesk\Navisworks Manage 2023\
C:\Program Files\Autodesk\Navisworks Manage 2024\
C:\Program Files\Autodesk\Navisworks Manage 2025\
C:\Program Files\Autodesk\Navisworks Manage 2026\
```

---

## 2. Project Structure Overview

```
MyPlugin/
├── MyCommands.cs          ← CommandHandlerPlugin (ribbon button handler)
├── MyAddin.cs             ← AddInPlugin (automation/scripting entry point)
├── MyEventPlugin.cs       ← EventWatcherPlugin (app menu injection)
├── MyCommandHandler.cs    ← ICommand for WPF menu items
├── Logger.cs              ← Serilog logging wrapper
├── app.config             ← Assembly binding redirects
├── packages.config        ← NuGet package list
├── MyPlugin.csproj        ← Multi-version project file
│
├── Commands/              ← One file per functional domain
│   ├── ExportCommands.cs  ← Export-related commands
│   ├── InspectCommands.cs ← Model inspection commands
│   ├── ReportCommands.cs  ← Reporting commands
│   └── UtilityCommands.cs ← Helper/utility commands
│
├── Helpers/               ← Reusable API abstractions
│   ├── NavisHelper.cs     ← Model search, property read/write
│   ├── ClashHelper.cs     ← Clash detection runner
│   └── ExcelHelper.cs     ← Excel read/write utilities
│
├── Models/                ← Data transfer objects
│   ├── ElementInfo.cs     ← Generic element data model
│   ├── ElementPosition.cs ← Spatial position data
│   └── ReportRow.cs       ← Report output row model
│
├── Views/                 ← Windows Forms dialogs
│   ├── ProgressForm.cs
│   ├── ProgressForm.Designer.cs
│   ├── ProgressForm.resx
│   └── BatchProgressForm.cs
│
├── Utils/
│   └── ExcelReader.cs     ← EPPlus-based Excel parser
│
├── Resources/
│   └── PluginIcon.ico     ← Icon embedded as resource
│
└── en-US/                 ← Navisworks XAML ribbon definition + icons
    ├── MyPlugin.xaml      ← Ribbon layout (tabs, panels, buttons)
    └── PluginIcon.ico
```

---

## 3. Creating the Visual Studio Project

### Step 1: Create a Class Library project

1. Open Visual Studio → **New Project → Class Library (.NET Framework)**
2. **Name**: `MyPlugin`
3. **Framework**: `.NET Framework 4.8`
4. Click **Create**

> **Important:** Navisworks runs in-process. Your plugin is a `.dll`, not an `.exe`. Always use `OutputType = Library`.

### Step 2: Edit the `.csproj` for Multi-Version Support

Replace the default property groups with conditional ones. This lets you select a target Navisworks version from the build configuration dropdown in Visual Studio.

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"
          Condition="Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')" />

  <!-- Default settings -->
  <PropertyGroup>
    <Configuration Condition=" '$(Configuration)' == '' ">Debug2024</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">x64</Platform>
    <ProjectGuid>{YOUR-GUID-HERE}</ProjectGuid>
    <OutputType>Library</OutputType>
    <RootNamespace>MyPlugin</RootNamespace>
    <AssemblyName>MyPlugin</AssemblyName>
    <UseWPF>true</UseWPF>
    <UseWindowsForms>true</UseWindowsForms>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
  </PropertyGroup>

  <!-- Configuration for Navisworks 2022 -->
  <PropertyGroup Condition="$(Configuration.Contains('2022'))">
    <NavisworksVersion>2022</NavisworksVersion>
    <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
    <OutputPath>bin\$(Configuration)\x64\</OutputPath>
    <PlatformTarget>x64</PlatformTarget>
    <LangVersion>7.3</LangVersion>
  </PropertyGroup>

  <!-- Configuration for Navisworks 2024 -->
  <PropertyGroup Condition="$(Configuration.Contains('2024'))">
    <NavisworksVersion>2024</NavisworksVersion>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
    <OutputPath>bin\$(Configuration)\x64\</OutputPath>
    <PlatformTarget>x64</PlatformTarget>
    <LangVersion>7.3</LangVersion>
  </PropertyGroup>

  <!-- Configuration for Navisworks 2025 -->
  <PropertyGroup Condition="$(Configuration.Contains('2025'))">
    <NavisworksVersion>2025</NavisworksVersion>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
    <OutputPath>bin\$(Configuration)\x64\</OutputPath>
    <PlatformTarget>x64</PlatformTarget>
    <LangVersion>7.3</LangVersion>
  </PropertyGroup>

  <!-- Add more versions as needed -->

  <!-- Import CSharp targets -->
  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
</Project>
```

> **Tip:** Add a Build Configuration in Visual Studio named `Debug2024`, `Debug2025`, etc. matching the condition strings.

---

## 4. Navisworks API References (Multi-Version)

Add the references conditionally per Navisworks version. Each version installs its DLLs into its own install folder.

```xml
<!-- For Navisworks 2024 -->
<ItemGroup Condition="$(Configuration.Contains('2024'))">

  <!-- Core UI (AdWindows = ribbon, application menu) -->
  <Reference Include="AdWindows">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\AdWindows.dll</HintPath>
    <Private>False</Private>
  </Reference>

  <!-- Primary managed API -->
  <Reference Include="Autodesk.Navisworks.Api">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\Autodesk.Navisworks.Api.dll</HintPath>
    <Private>False</Private>
  </Reference>

  <!-- Roamer (ribbon controls for NW) -->
  <Reference Include="navisworks.gui.roamer">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\navisworks.gui.roamer.dll</HintPath>
    <Private>False</Private>
  </Reference>

  <!-- Automation API (for scripting / batch) -->
  <Reference Include="Autodesk.Navisworks.Automation">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\Autodesk.Navisworks.Automation.dll</HintPath>
    <Private>False</Private>
  </Reference>

  <!-- Clash detection -->
  <Reference Include="Autodesk.Navisworks.Clash">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\Autodesk.Navisworks.Clash.dll</HintPath>
    <Private>False</Private>
  </Reference>

  <!-- COM API (for property write-back) -->
  <Reference Include="Autodesk.Navisworks.ComApi">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\Autodesk.Navisworks.ComApi.dll</HintPath>
    <Private>False</Private>
  </Reference>

  <!-- COM Interop (embed types) -->
  <Reference Include="Autodesk.Navisworks.Interop.ComApi">
    <HintPath>c:\Program Files\Autodesk\Navisworks Manage 2024\Autodesk.Navisworks.Interop.ComApi.dll</HintPath>
    <EmbedInteropTypes>True</EmbedInteropTypes>
  </Reference>

</ItemGroup>

<!-- Always-included WPF assemblies -->
<ItemGroup>
  <Reference Include="WindowsBase" />
  <Reference Include="PresentationCore" />
  <Reference Include="PresentationFramework" />
  <Reference Include="System.Windows.Forms" />
  <Reference Include="System.Drawing" />
  <Reference Include="System.Xaml" />
  <Reference Include="System" />
  <Reference Include="System.Data" />
  <Reference Include="System.Xml" />
  <Reference Include="Microsoft.CSharp" />
</ItemGroup>
```

> **Key Rule:** Set `<Private>False</Private>` on all Navisworks DLLs so they are **not** copied to the output. They already exist in the Navisworks install directory.

---

## 5. NuGet Packages

These are recommended packages for common Navisworks plugin scenarios. Add what you need:

```xml
<!-- packages.config -->
<?xml version="1.0" encoding="utf-8"?>
<packages>
  <!-- Excel read/write -->
  <package id="EPPlus" version="4.5.3.3" targetFramework="net48" />
  <package id="ExcelDataReader" version="3.7.0" targetFramework="net48" />
  <package id="ExcelDataReader.DataSet" version="3.7.0" targetFramework="net48" />

  <!-- CSV parsing -->
  <package id="CsvHelper" version="33.0.1" targetFramework="net48" />

  <!-- Structured logging -->
  <package id="Serilog" version="4.0.2" targetFramework="net48" />
  <package id="Serilog.Sinks.Console" version="6.0.0" targetFramework="net48" />
  <package id="Serilog.Sinks.File" version="6.0.0" targetFramework="net48" />
  <package id="Serilog.Sinks.Seq" version="8.0.0" targetFramework="net48" />

  <!-- IFC export (optional) -->
  <package id="Xbim.Essentials" version="6.0.445" targetFramework="net48" />
  <package id="Xbim.Ifc" version="6.0.445" targetFramework="net48" />
  <package id="Xbim.Ifc4" version="6.0.445" targetFramework="net48" />

  <!-- Lightweight ORM -->
  <package id="Dapper" version="2.1.35" targetFramework="net48" />

  <!-- Runtime helpers -->
  <package id="System.Memory" version="4.5.5" targetFramework="net48" />
  <package id="System.Buffers" version="4.5.1" targetFramework="net48" />
</packages>
```

Install via NuGet Package Manager Console:

```powershell
Install-Package EPPlus -Version 4.5.3.3
Install-Package Serilog -Version 4.0.2
Install-Package Serilog.Sinks.File -Version 6.0.0
```

---

## 6. Plugin Types Explained

Navisworks supports three base plugin types. You declare which type you are using the `[Plugin]` attribute and the base class you inherit.

| Type | Base Class | Purpose |
|------|-----------|---------|
| **CommandHandler** | `CommandHandlerPlugin` | Registers ribbon buttons and handles clicks |
| **AddIn** | `AddInPlugin` | Entry point for automation / scripting mode |
| **EventWatcher** | `EventWatcherPlugin` | Runs code on load/unload, modifies app menu |
| **DockPane** | `DockPanePlugin` | Creates a custom dockable panel / tool window inside Navisworks |

You can have **multiple plugin classes in a single DLL**, each serving a different role.

---

## 7. CommandHandlerPlugin — The Primary Plugin Class

This is the plugin type that wires ribbon buttons to C# methods. It's the most important class in any Navisworks addin.

### Key Attributes

| Attribute | Role |
|-----------|------|
| `[PluginAttribute]` | Declares the plugin name, developer ID, tooltip, and display name |
| `[RibbonLayoutAttribute]` | Points to the XAML file that defines the ribbon layout |
| `[RibbonTab]` | Declares which tabs this plugin creates in the Navisworks ribbon |
| `[Command]` | Registers each individual button command ID |

### Full Example: `MyCommands.cs`

```csharp
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api;
using System.Windows.Forms;

namespace MyPlugin
{
    // Plugin identity
    [PluginAttribute("MyPluginCommands",          // Unique plugin name (no spaces)
                     "MyCompany",                  // Developer ID (must be <= 4 chars or GUID)
                     ToolTip = "My Navisworks Plugin",
                     DisplayName = "My Plugin")]

    // Point to the XAML file that defines ribbon layout
    [RibbonLayoutAttribute("MyPlugin.xaml")]

    // Declare ribbon tabs (IDs must match the XAML Tab Ids)
    [RibbonTab("ID_MyMainTab", DisplayName = "My Plugin")]
    [RibbonTab("ID_MySecondTab", DisplayName = "My Second Tab")]

    // Register every button command (IDs must match XAML button Ids and switch cases)
    [Command("ID_CMD_HelloWorld", CanToggle = true, DisplayName = "Hello World")]
    [Command("ID_CMD_ExportData", CanToggle = true, DisplayName = "Export Data")]

    public class MyCommands : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            try
            {
                switch (name)
                {
                    case "ID_CMD_HelloWorld":
                        MessageBox.Show("Hello from Navisworks Plugin!");
                        break;

                    case "ID_CMD_ExportData":
                        MyExporter.RunExport();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Plugin Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0; // Must always return 0
        }
    }
}
```

> **Critical Rules:**
> - The `CommandHandlerPlugin.ExecuteCommand(string name, ...)` receives the command ID as the `name` argument.
> - The method must return `0`.
> - Every `ID_CMD_*` string in a `[Command]` attribute, in the `switch`, and in the XAML button's `Id` must be **identical**.

---

## 8. AddInPlugin — The Automation Plugin

This is used when Navisworks is started in **automation mode** (e.g., from a script or command line). It also enables calling the plugin from an external process.

```csharp
using Autodesk.Navisworks.Api.Plugins;

namespace MyPlugin
{
    [PluginAttribute("MyAddin",
                     "MyCompany",
                     ToolTip = "My Addin",
                     DisplayName = "My Addin")]
    [AddInPlugin(AddInLocation.None)]  // 'None' = don't show a button for this addin itself
    internal class MyAddin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            // Called when running in automation mode
            if (Autodesk.Navisworks.Api.Application.IsAutomated && parameters.Length == 0)
            {
                throw new InvalidOperationException("No parameters provided in automation mode");
            }

            // Dispatch based on the first parameter (the command ID)
            switch (parameters[0])
            {
                case "ID_CMD_HelloWorld":
                    MyCommands_Impl.HelloWorld(false);
                    break;

                case "ID_CMD_ExportData":
                    MyCommands_Impl.ExportData(false);
                    break;
            }

            return 0;
        }
    }
}
```

### AddInLocation Options

```csharp
AddInLocation.None          // Plugin not accessible from the UI directly
AddInLocation.AddIn         // Appears in the Add-Ins menu
```

---

## 9. EventWatcherPlugin — Application Menu & Lifecycle Hooks

Used to inject items into the **File/Application menu** and to run code when the plugin loads or unloads.

```csharp
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Windows;
using System;
using System.Windows;
using System.Windows.Media.Imaging;

namespace MyPlugin
{
    [Plugin("MyEventPlugin",
            "MyCompany",
            DisplayName = "My Event Plugin")]
    internal class MyEventPlugin : EventWatcherPlugin
    {
        public override void OnLoaded()
        {
            try
            {
                // Access the application (File) menu
                var appMenu = Autodesk.Windows.ComponentManager.ApplicationMenu;

                if (appMenu != null)
                {
                    // Create a new application menu item
                    var menuItem = new ApplicationMenuItem
                    {
                        Text = "My Custom Action",
                        Id = "ID_CMD_HelloWorld",
                        Image = new BitmapImage(new Uri(
                            "pack://application:,,,/MyPlugin;component/Resources/MyIcon.ico")),
                        LargeImage = new BitmapImage(new Uri(
                            "pack://application:,,,/MyPlugin;component/Resources/MyIcon.ico")),
                        CommandHandler = new MyCommandHandler(),
                        IsActive = true,
                        IsVisible = true,
                        UID = "ID_CMD_HelloWorld"
                    };

                    appMenu.MenuContent.Items.Add(menuItem);
                    appMenu.SaveState();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding menu item: " + ex.Message);
            }
        }

        public override void OnUnloading()
        {
            // Cleanup code here (optional)
        }
    }
}
```

---

## 10. Registering Commands with Attributes

Every button in the ribbon needs three things in sync:

1. A `[Command]` attribute on the `CommandHandlerPlugin` class
2. An entry in `ExecuteCommand`'s switch block
3. A matching button definition in the XAML file

```csharp
// 1. Attribute (on the class, before the class declaration)
[Command("ID_CMD_MyButton", CanToggle = true, DisplayName = "My Button")]

// 2. Switch case (inside ExecuteCommand)
case "ID_CMD_MyButton":
    MyFeature.Run();
    break;
```

```xml
<!-- 3. XAML ribbon button (in en-US/MyPlugin.xaml) -->
<local:NWRibbonButton x:Uid="BtnMyButton" Id="ID_CMD_MyButton"
    Size="Large"
    Image="MyIcon.ico"
    ShowText="True"
    Text="My Button"
    Orientation="Vertical"/>
```

---

## 11. Ribbon UI — The XAML File

The XAML file defines the entire visual layout of your ribbon. It is referenced by `[RibbonLayoutAttribute("MyPlugin.xaml")]` and must be placed in an `en-US` sub-folder at the same level as the plugin DLL when deployed.

### File: `en-US/MyPlugin.xaml`

```xml
<?xml version="1.0" encoding="utf-8" ?>
<RibbonControl
    x:Uid="ADNRibbonMyPlugin"
    xmlns="clr-namespace:Autodesk.Windows;assembly=AdWindows"
    xmlns:wpf="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:adwi="clr-namespace:Autodesk.Internal.Windows;assembly=AdWindows"
    xmlns:system="clr-namespace:System;assembly=mscorlib"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:local="clr-namespace:Autodesk.Navisworks.Gui.Roamer.AIRLook;assembly=navisworks.gui.roamer">

    <!-- Tab 1 -->
    <RibbonTab Id="ID_MyMainTab" KeyTip="M" Title="My Plugin">
        <!-- Panel 1 inside Tab 1 -->
        <RibbonPanel x:Uid="RibbonPanel_Main">
            <RibbonPanelSource x:Uid="RibbonPanelSource_Main" KeyTip="A" Title="Actions">

                <!-- Large button with icon and label below -->
                <local:NWRibbonButton x:Uid="BtnHelloWorld"
                    Id="ID_CMD_HelloWorld"
                    Size="Large"
                    KeyTip="H"
                    Image="MyIcon.ico"
                    ShowText="True"
                    Text="Hello World"
                    Orientation="Vertical"/>

                <!-- Another button in the same panel -->
                <local:NWRibbonButton x:Uid="BtnExportData"
                    Id="ID_CMD_ExportData"
                    Size="Large"
                    KeyTip="E"
                    Image="MyIcon.ico"
                    ShowText="True"
                    Text="Export Data"
                    Orientation="Vertical"/>

            </RibbonPanelSource>
        </RibbonPanel>
    </RibbonTab>

    <!-- Tab 2 with multiple panels -->
    <RibbonTab Id="ID_MySecondTab" KeyTip="S" Title="My Second Tab">
        <RibbonPanel x:Uid="RibbonPanel_Reports">
            <RibbonPanelSource x:Uid="RibbonPanelSource_Reports" KeyTip="R" Title="Reports">
                <local:NWRibbonButton x:Uid="BtnReport"
                    Id="ID_CMD_GenerateReport"
                    Size="Large"
                    KeyTip="R"
                    Image="MyIcon.ico"
                    ShowText="True"
                    Text="Generate Report"
                    Orientation="Vertical"/>
            </RibbonPanelSource>
        </RibbonPanel>
    </RibbonTab>

</RibbonControl>
```

### XAML Build Action

In the `.csproj`, mark the XAML and icon as **Content** and set them to always copy to output:

```xml
<ItemGroup>
  <Content Include="en-US\MyPlugin.xaml">
    <Generator>MSBuild:Compile</Generator>
    <CopyToOutputDirectory>Always</CopyToOutputDirectory>
  </Content>
  <Content Include="en-US\MyIcon.ico">
    <CopyToOutputDirectory>Always</CopyToOutputDirectory>
  </Content>
</ItemGroup>
```

### Icon as Embedded Resource (for app menu / WPF)

```xml
<ItemGroup>
  <Resource Include="Resources\MyIcon.ico">
    <CopyToOutputDirectory>Always</CopyToOutputDirectory>
  </Resource>
</ItemGroup>
```

---

## 12. Deploying to Navisworks (Post-Build Event)

Navisworks loads plugins from this AppData path on startup:

```
%AppData%\Autodesk\Navisworks Manage {Year}\Plugins\{PluginName}\
```

The `en-US` sub-folder must be inside the plugin folder:

```
%AppData%\Autodesk\Navisworks Manage 2024\Plugins\MyPlugin\
    ├── MyPlugin.dll
    ├── EPPlus.dll
    ├── Serilog.dll
    └── en-US\
        ├── MyPlugin.xaml
        └── MyIcon.ico
```

### Automated Post-Build Event (multi-version)

Add this to your `.csproj`. It uses the `$(NavisworksVersion)` property set by the build configuration:

```xml
<PropertyGroup>
  <PostBuildEvent>
    REM Create plugin folder
    mkdir "$(APPDATA)\Autodesk\Navisworks Manage $(NavisworksVersion)\Plugins\$(TargetName)"

    REM Copy all built DLLs and files
    copy "$(TargetDir)\*" "$(APPDATA)\Autodesk\Navisworks Manage $(NavisworksVersion)\Plugins\$(TargetName)\"

    REM Create en-US subfolder
    mkdir "$(APPDATA)\Autodesk\Navisworks Manage $(NavisworksVersion)\Plugins\$(TargetName)\en-US"

    REM Copy XAML and icons from project source
    copy "$(ProjectDir)en-US\*" "$(APPDATA)\Autodesk\Navisworks Manage $(NavisworksVersion)\Plugins\$(TargetName)\en-US\"
  </PostBuildEvent>
</PropertyGroup>
```

### Debug Startup (Launch Navisworks directly from Visual Studio)

```xml
<PropertyGroup Condition="$(Configuration.Contains('2024'))">
  <StartAction>Program</StartAction>
  <StartProgram>C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe</StartProgram>
  <DebugType>full</DebugType>
  <Optimize>false</Optimize>
  <DebugWorkingDirectory>C:\Program Files\Autodesk\Navisworks Manage 2024\</DebugWorkingDirectory>
</PropertyGroup>
```

When you press **F5** in Visual Studio, it will build, copy the DLL, and launch Navisworks automatically.

---

## 13. app.config — Assembly Binding Redirects

Some NuGet packages have version conflicts. Use binding redirects in `app.config`:

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <runtime>
    <gcAllowVeryLargeObjects enabled="true" />
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">

      <dependentAssembly>
        <assemblyIdentity name="System.Memory" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-4.0.1.2" newVersion="4.0.1.2" />
      </dependentAssembly>

      <dependentAssembly>
        <assemblyIdentity name="System.Runtime.CompilerServices.Unsafe"
                          publicKeyToken="b03f5f7f11d50a3a" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-6.0.0.0" newVersion="6.0.0.0" />
      </dependentAssembly>

      <dependentAssembly>
        <assemblyIdentity name="System.Buffers" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-4.0.3.0" newVersion="4.0.3.0" />
      </dependentAssembly>

    </assemblyBinding>
  </runtime>
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.8" />
  </startup>
</configuration>
```

---

## 14. Working with the Navisworks API

### 14.1 Accessing the Active Document

```csharp
using Autodesk.Navisworks.Api;

// Get the active Navisworks document
Document doc = Application.ActiveDocument;

// Check if any file is open
if (doc == null) return;
```

### 14.2 Getting Selected Items

```csharp
// Get currently selected items in the scene
ModelItemCollection selectedItems = doc.CurrentSelection.SelectedItems;

if (selectedItems.Count == 0)
{
    MessageBox.Show("Please select at least one item.", "No Selection",
        MessageBoxButtons.OK, MessageBoxIcon.Warning,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly, false);
    return;
}
```

### 14.3 Searching the Model with Search API

The `Search` API lets you find items by property name, value, or condition.

#### Search for Items Having a Specific Property

```csharp
Search search = new Search();
// Replace "Item Data" and "Element ID" with the actual category/property names
// visible in the Navisworks Properties panel for your model source
SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Item Data", "Element ID");
search.SearchConditions.Add(condition);
search.Locations = SearchLocations.DescendantsAndSelf;
search.Selection.SelectAll();

ModelItemCollection results = search.FindAll(Application.ActiveDocument, true);
```

#### Search for Items with a Specific Property Value

```csharp
Search search = new Search();
// Example: find all items whose "Material" property equals "Concrete"
SearchCondition condition = SearchCondition
    .HasPropertyByDisplayName("Element", "Material")
    .EqualValue(new VariantData("Concrete"))
    .IgnoreStringValueCase();
search.SearchConditions.Add(condition);
search.Selection.SelectAll();
ModelItemCollection concreteItems = search.FindAll(Application.ActiveDocument, true);
```

#### Search with Numeric Comparison

```csharp
// Example: find items whose "Elevation" property is >= 100.0
SearchCondition aboveZ = SearchCondition
    .HasPropertyByDisplayName("Geometry", "Elevation")
    .CompareWith(SearchConditionComparison.NumericGreaterThanOrEqual, new VariantData(100.0));
search.SearchConditions.Add(aboveZ);
```

#### Search with OR Logic (Multiple condition groups)

```csharp
Search search = new Search();

// Each group = one "OR" alternative
// Example: find items of type "Pipe" OR "Duct" from an "Element" category
List<SearchCondition> group1 = new List<SearchCondition>
{
    SearchCondition.HasPropertyByDisplayName("Element", "Category")
        .EqualValue(new VariantData("Pipe")).IgnoreStringValueCase()
};
List<SearchCondition> group2 = new List<SearchCondition>
{
    SearchCondition.HasPropertyByDisplayName("Element", "Category")
        .EqualValue(new VariantData("Duct")).IgnoreStringValueCase()
};

search.SearchConditions.AddGroup(group1);  // group1 OR group2
search.SearchConditions.AddGroup(group2);

search.Selection.SelectAll();
var result = search.FindAll(Application.ActiveDocument, true);
```

#### Search on a Sub-selection

```csharp
// First pass: get a broad set (e.g. all structural items)
var structuralItems = GetStructuralItems();

// Second pass: narrow down within that set using another condition
Search subSearch = new Search();
subSearch.SearchConditions.Add(someOtherCondition);
subSearch.Selection.CopyFrom(structuralItems); // search WITHIN structuralItems only
var filteredItems = subSearch.FindAll(Application.ActiveDocument, true);
```

### 14.4 Reading Properties from ModelItems

```csharp
// Read a single property from a ModelItem
ModelItem item = ...;

// Method 1: Using FindPropertyByDisplayName (category + property)
// Use the exact category and property names shown in Navisworks Properties panel
var prop = item.PropertyCategories.FindPropertyByDisplayName("Item", "Element ID");
if (prop != null)
{
    string value = prop.Value.ToDisplayString();
    double numValue = prop.Value.ToDouble();
}

// Method 2: Iterate all categories and properties
foreach (PropertyCategory category in item.PropertyCategories)
{
    string categoryName = category.DisplayName;
    foreach (DataProperty property in category.Properties)
    {
        string propName = property.DisplayName;
        string propValue = property.Value.ToDisplayString();
    }
}
```

### 14.5 Writing / Setting Properties via COM API

The managed API is read-only. To write custom (user-defined) properties, you must use the COM API bridge.

```csharp
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

public static bool SetParameterValue(ModelItem item, string category, string attribute, dynamic value)
{
    try
    {
        // Get the COM state
        InwOpState10 oState = Autodesk.Navisworks.Api.ComApi.ComApiBridge.State;

        // Convert .NET ModelItem → COM path
        InwOaPath oPath = Autodesk.Navisworks.Api.ComApi.ComApiBridge.ToInwOaPath(item);

        // Get the GUI property node (true = create if not found)
        InwGUIPropertyNode2 propNode = (InwGUIPropertyNode2)oState.GetGUIPropertyNode(oPath, true);

        int index = 1;
        bool found = false;

        // Iterate existing user-defined categories
        foreach (InwGUIAttribute2 cat in propNode.GUIAttributes())
        {
            if (cat.ClassUserName.ToUpper() == category.ToUpper() && cat.UserDefined)
            {
                // Build updated property list
                InwOaPropertyVec propVec = (InwOaPropertyVec)oState.ObjectFactory(
                    nwEObjectType.eObjectType_nwOaPropertyVec);

                foreach (InwOaProperty existingProp in cat.Properties())
                {
                    InwOaProperty newProp = (InwOaProperty)oState.ObjectFactory(
                        nwEObjectType.eObjectType_nwOaProperty);
                    newProp.name = existingProp.name;
                    newProp.value = existingProp.value;

                    if (newProp.name == attribute)
                    {
                        newProp.value = value;  // Update the value
                        found = true;
                    }
                    propVec.Properties().Add(newProp);
                }

                // Add new attribute if not found
                if (!found)
                {
                    InwOaProperty addProp = (InwOaProperty)oState.ObjectFactory(
                        nwEObjectType.eObjectType_nwOaProperty);
                    addProp.name = attribute;
                    addProp.value = value;
                    propVec.Properties().Add(addProp);
                }

                // Commit: index > 0 = update existing, 0 = create new
                propNode.SetUserDefined(index, cat.ClassUserName, cat.ClassName, propVec);
                return true;
            }
            index++;
        }

        // Category not found — create new category with one property
        InwOaPropertyVec newPropVec = (InwOaPropertyVec)oState.ObjectFactory(
            nwEObjectType.eObjectType_nwOaPropertyVec);
        InwOaProperty newProperty = (InwOaProperty)oState.ObjectFactory(
            nwEObjectType.eObjectType_nwOaProperty);
        newProperty.name = attribute;
        newProperty.UserName = attribute;
        newProperty.value = value;
        newPropVec.Properties().Add(newProperty);

        propNode.SetUserDefined(0, category, category, newPropVec);
        return true;
    }
    catch (Exception ex)
    {
        MessageBox.Show(ex.Message);
        return false;
    }
}
```

### 14.6 Traversing the Model Hierarchy

```csharp
// Access all models in the document
foreach (Model model in Application.ActiveDocument.Models)
{
    ModelItem root = model.RootItem;

    // Get all geometric leaf items (has geometry = true)
    var geometricItems = root.Descendants.Where(x => x.HasGeometry);
}

// Recursive descent helper
public static List<ModelItem> GetAllLeafItems(ModelItem item)
{
    var result = new List<ModelItem>();
    foreach (ModelItem child in item.Children)
    {
        if (!child.Children.Any())
            result.Add(child);           // leaf node
        else
            result.AddRange(GetAllLeafItems(child));  // recurse
    }
    return result;
}

// Find the root model for any ModelItem
public static Model GetModelForItem(ModelItem item)
{
    if (item.Parent == null) return item.Model;
    return GetModelForItem(item.Parent);
}
```

---

## 15. Clash Detection API

Clash detection allows you to programmatically perform interference checks between two sets of elements.

```csharp
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api;

public static IList<ClashResult> RunClashTest(
    string testName,
    ModelItemCollection selectionA,
    ModelItemCollection selectionB,
    ClashTestType testType = ClashTestType.Hard,
    double tolerance = 0.01)
{
    Document document = Application.ActiveDocument;
    DocumentClash documentClash = document.GetClash();
    DocumentClashTests clashTests = documentClash.TestsData;

    // Create a new clash test
    ClashTest newTest = new ClashTest();
    newTest.DisplayName = testName;
    newTest.SelectionA.Selection.CopyFrom(selectionA);
    newTest.SelectionB.Selection.CopyFrom(selectionB);
    newTest.TestType = testType;
    newTest.Tolerance = tolerance;

    // Clear existing tests and run (Navisworks API bug: must clear first)
    clashTests.TestsClear();
    clashTests.TestsAddCopy(newTest);
    clashTests.TestsRunAllTests();

    // Re-read from document (required to get actual results)
    newTest = clashTests.Tests[0] as ClashTest;

    // Collect results
    var results = new List<ClashResult>();
    foreach (SavedItem item in newTest.Children)
    {
        // Handle grouped results
        if (item is ClashResultGroup group)
        {
            foreach (SavedItem groupItem in group.Children)
            {
                if (groupItem is ClashResult clashResult)
                    results.Add(clashResult);
            }
        }

        // Handle individual results
        if (item is ClashResult result)
            results.Add(result);
    }

    return results;
}

// Usage example: check interference between two categories of elements
var groupA = GetSetA();  // e.g. structural beams
var groupB = GetSetB();  // e.g. HVAC ducts
var clashes = RunClashTest("Beams vs Ducts", groupA, groupB);

foreach (var clash in clashes)
{
    ModelItem item1 = clash.Item1;
    ModelItem item2 = clash.Item2;
    // Process the clash pair...
}
```

### ClashTestType Options

```csharp
ClashTestType.Hard       // Objects physically overlap
ClashTestType.Clearance  // Objects are within a specified distance
ClashTestType.Duplicates // Identical overlapping geometry
```

---

## 16. Excel Integration (EPPlus & ExcelDataReader)

### Reading Excel with EPPlus

```csharp
using OfficeOpenXml;
using System.Data;
using System.IO;

public DataTable ReadExcelToDataTable(string filePath, bool hasHeader = true)
{
    using (var package = new ExcelPackage(new FileInfo(filePath)))
    {
        var worksheet = package.Workbook.Worksheets[1]; // 1-indexed

        DataTable table = new DataTable();

        // Add columns from header row
        foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
        {
            table.Columns.Add(hasHeader ? cell.Text : $"Column{cell.Start.Column}");
        }

        int startRow = hasHeader ? 2 : 1;
        for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
        {
            var wsRow = worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column];
            DataRow dataRow = table.Rows.Add();
            foreach (var cell in wsRow)
            {
                dataRow[cell.Start.Column - 1] = cell.Text;
            }
        }

        return table;
    }
}
```

### Writing Excel with EPPlus

```csharp
public void WriteExcelReport(string outputPath, List<MyDataModel> data)
{
    using (var package = new ExcelPackage())
    {
        var ws = package.Workbook.Worksheets.Add("Report");

        // Write headers
        ws.Cells[1, 1].Value = "Element ID";
        ws.Cells[1, 2].Value = "Name";
        ws.Cells[1, 3].Value = "Value";

        // Style the header row
        using (var range = ws.Cells[1, 1, 1, 3])
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
        }

        // Write data
        for (int i = 0; i < data.Count; i++)
        {
            int row = i + 2;
            ws.Cells[row, 1].Value = data[i].ElementId;
            ws.Cells[row, 2].Value = data[i].Name;
            ws.Cells[row, 3].Value = data[i].Value;
        }

        // Auto-fit column widths
        ws.Cells[ws.Dimension.Address].AutoFitColumns();

        package.SaveAs(new FileInfo(outputPath));
    }

    MessageBox.Show($"Report saved to:\n{outputPath}", "Export Complete",
        MessageBoxButtons.OK, MessageBoxIcon.Information,
        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
}
```

### Showing a File Save Dialog

```csharp
using System.Windows.Forms;

public static string ShowSaveExcelDialog(string title = "Save Excel Report")
{
    using (var dialog = new SaveFileDialog())
    {
        dialog.Title = title;
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
        dialog.DefaultExt = "xlsx";
        dialog.FileName = $"Report_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";

        if (dialog.ShowDialog() == DialogResult.OK)
            return dialog.FileName;

        return null;
    }
}
```

---

## 17. Progress Forms (Windows Forms UI)

Always show a progress dialog when processing hundreds or thousands of elements.

### Simple Progress Bar Form

```csharp
// Views/ProgressForm.cs
using System.Windows.Forms;

namespace MyPlugin
{
    public partial class ProgressForm : Form
    {
        private string _format;

        public ProgressForm(string caption, string format, int max)
        {
            _format = format;
            InitializeComponent();
            Text = caption;
            label1.Text = format == null ? caption : string.Format(format, 0);
            progressBar1.Minimum = 0;
            progressBar1.Maximum = max;
            progressBar1.Value = 0;
            Show();
            Application.DoEvents();
        }

        public void Increment()
        {
            ++progressBar1.Value;
            if (_format != null)
                label1.Text = string.Format(_format, progressBar1.Value);
            Application.DoEvents(); // Allow UI to repaint
        }
    }
}
```

### Usage Pattern

```csharp
using (var progressForm = new ProgressForm("Processing Items", "Processed {0} items...", items.Count))
{
    foreach (var item in items)
    {
        ProcessItem(item);
        progressForm.Increment();
    }
}
```

### Batch Progress Form (Two-Level)

When processing items in batches (e.g., grouped by model or category):

```csharp
public class BatchProgressForm : Form
{
    private ProgressBar batchBar, overallBar;
    private Label batchLabel, overallLabel;
    private int _totalBatches, _currentBatch;
    private int _totalOverall, _processedOverall;

    public BatchProgressForm(int totalBatches, int totalOverall)
    {
        _totalBatches = totalBatches;
        _totalOverall = totalOverall;
        InitControls();
    }

    private void InitControls()
    {
        this.Text = "Processing Progress";
        this.Width = 420; this.Height = 220;

        batchLabel = new Label { AutoSize = true, Location = new Point(20, 20) };
        this.Controls.Add(batchLabel);

        batchBar = new ProgressBar { Location = new Point(20, 50), Width = 360, Minimum = 0 };
        this.Controls.Add(batchBar);

        overallLabel = new Label { AutoSize = true, Location = new Point(20, 90) };
        this.Controls.Add(overallLabel);

        overallBar = new ProgressBar
        { Location = new Point(20, 120), Width = 360, Minimum = 0, Maximum = _totalOverall };
        this.Controls.Add(overallBar);
    }

    public void StartBatch(int batchNumber, int itemsInBatch)
    {
        _currentBatch = batchNumber;
        batchBar.Maximum = itemsInBatch;
        batchBar.Value = 0;
        batchLabel.Text = $"Batch {batchNumber} of {_totalBatches}";
        Application.DoEvents();
    }

    public void IncrementBatch()
    {
        batchBar.Value++;
        Application.DoEvents();
    }

    public void IncrementOverall()
    {
        _processedOverall++;
        overallBar.Value = _processedOverall;
        overallLabel.Text = $"Overall: {_processedOverall} / {_totalOverall}";
        Application.DoEvents();
    }
}
```

---

## 18. Logging with Serilog

A clean logging setup using Serilog, writing to a rolling file and optionally to a Seq server.

### Logger Wrapper Class

```csharp
// logger.cs
using Serilog;
using System;
using System.IO;

namespace MyPlugin
{
    public static class Logger
    {
        static Logger()
        {
            // Log to %TEMP%\MyPlugin\Logs\
            string logDir = Path.Combine(Path.GetTempPath(), "MyPlugin", "Logs");
            Directory.CreateDirectory(logDir);
            string logFile = Path.Combine(logDir, "app_log.txt");

            Log.Logger = new LoggerConfiguration()
                .Enrich.FromLogContext()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    logFile,
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}"
                )
                // Optionally write to Seq (structured log viewer)
                // .WriteTo.Seq("http://localhost:5341")
                .CreateLogger();
        }

        public static void Information(string message) => Log.Information(message);
        public static void Error(string message, Exception ex = null)
        {
            if (ex != null) Log.Error(ex, message);
            else Log.Error(message);
        }
        public static void Debug(string message) => Log.Debug(message);
        public static void Warning(string message) => Log.Warning(message);
        public static void CloseAndFlush() => Log.CloseAndFlush();
    }
}
```

### Usage

```csharp
Logger.Information("Starting export for 500 items");
Logger.Debug($"Processing item: {item.DisplayName}");

try { /* ... */ }
catch (Exception ex)
{
    Logger.Error("Export failed", ex);
    MessageBox.Show($"Error: {ex.Message}");
}

Logger.Information("Export complete");
```

---

## 19. WPF Command Handler for Application Menu

When using `EventWatcherPlugin` to add items to the application menu, commands are executed via `ICommand` handlers:

```csharp
// ExportCommandHandler.cs
using System;
using System.Windows;
using System.Windows.Input;

public class ExportCommandHandler : ICommand
{
    // Fired when CanExecute changes (not needed in most cases)
    public event EventHandler CanExecuteChanged;

    // Controls whether the button is enabled
    public bool CanExecute(object parameter) => true;

    // Called when the button is clicked
    public void Execute(object parameter)
    {
        try
        {
            MyExporter.RunExport();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Export failed: " + ex.Message);
        }
    }
}
```

---

## 20. Data Models

Define plain C# classes (POCOs) to hold model data you extract from Navisworks:

```csharp
// Models/ElementInfo.cs
using Autodesk.Navisworks.Api;

namespace MyPlugin.Models
{
    public class ElementInfo
    {
        // Reference back to the Navisworks item
        public ModelItem ModelItem { get; set; }

        // User-defined properties (map these to whatever your model exposes)
        public string ElementId { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }

        // Bounding box centre (computed from item.BoundingBox())
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public double CenterZ { get; set; }

        // Calculated fields
        public string Orientation { get; set; }   // "X", "Y", or "Z" axis alignment
        public string ElementType { get; set; }    // e.g. "BEAM", "COLUMN", "PIPE"
    }
}
```

```csharp
// Models/ElementPosition.cs
// Generic model for storing spatial extents of any element
namespace MyPlugin.Models
{
    public class ElementPosition
    {
        public ModelItem ModelItem { get; set; }
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public double CenterZ { get; set; }
        public double TopElevation { get; set; }
        public double BottomElevation { get; set; }
        public string Orientation { get; set; }  // "X", "Y", or "Z"
    }
}
```

---

## 21. Adding New Commands — Step-by-Step Checklist

When adding a new button/feature to the plugin:

- [ ] **1. Plan the command ID** — Use `ID_CMD_` prefix and descriptive name (e.g., `ID_CMD_ExportReport`)

- [ ] **2. Add `[Command]` attribute** — On the `CommandHandlerPlugin` class:
  ```csharp
  [Command("ID_CMD_ExportReport", CanToggle = true, DisplayName = "Export Report")]
  ```

- [ ] **3. Add `switch case`** — In `ExecuteCommand` inside `MyCommands.cs`:
  ```csharp
  case "ID_CMD_ExportReport": MyReporter.ExportReport(); break;
  ```

- [ ] **4. Add `switch case`** — Also in `MyAddin.cs` (for automation support):
  ```csharp
  case "ID_CMD_ExportReport": MyReporter.ExportReport(false); break;
  ```

- [ ] **5. Add button to XAML** — In `en-US/MyPlugin.xaml`, inside the appropriate `RibbonPanelSource`:
  ```xml
  <local:NWRibbonButton x:Uid="BtnExportReport" Id="ID_CMD_ExportReport"
      Size="Large" Image="MyIcon.ico" ShowText="True"
      Text="Export Report" Orientation="Vertical"/>
  ```

- [ ] **6. Implement the feature** — Create or update the relevant class in `Commands/`:
  ```csharp
  public static class MyReporter
  {
      public static void ExportReport(bool UI = true)
      {
          Logger.Information("ExportReport started");
          var doc = Application.ActiveDocument;
          // ... implementation
      }
  }
  ```

- [ ] **7. Build and test** — Press F5 to launch Navisworks and verify the button appears

---

## 22. Debugging Tips

### Attach to Running Navisworks

If Navisworks is already open:
1. In Visual Studio: **Debug → Attach to Process**
2. Select `Roamer.exe`
3. Set breakpoints and click the ribbon button

### Launching from Visual Studio

Configure the project to launch `Roamer.exe` on F5 (see section 12). Your DLL will be automatically copied by the post-build event before launch.

### MessageBox for Quick Debugging

```csharp
MessageBox.Show(
    $"Items Found: {items.Count}",
    "Debug",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button1,
    MessageBoxOptions.DefaultDesktopOnly,  // Required for Navisworks context
    false
);
```

> **Always use `MessageBoxOptions.DefaultDesktopOnly`** when calling `MessageBox.Show` from a plugin. Otherwise, the dialog may not appear on top.

### Check the Log File

```
%TEMP%\MyPlugin\Logs\app_log.txt
```

### Print to Debug Output

```csharp
System.Diagnostics.Debug.WriteLine($"Processing: {item.DisplayName}");
```

Visible in Visual Studio's **Output** window while debugging.

---

## 23. Common Gotchas & Troubleshooting

| Issue | Cause | Solution |
|-------|-------|----------|
| Plugin doesn't appear in ribbon | XAML not deployed to `en-US` folder or command IDs mismatch | Verify post-build event ran; check XAML `Id` matches `[Command]` attribute |
| `MessageBox` doesn't appear | Missing `MessageBoxOptions.DefaultDesktopOnly` | Always pass this flag |
| Property write fails silently | Using managed API (read-only) | Use COM API bridge (`SetUserDefined`) instead |
| Clash test returns 0 results | Known NW API bug: must clear tests before adding | Always call `TestsClear()` before `TestsAddCopy()` |
| DLL conflicts / `FileLoadException` | NuGet package version conflicts | Add binding redirects to `app.config` |
| Plugin loads old DLL after rebuild | Post-build copy failed (permissions) | Run Visual Studio as Administrator |
| `Private>False</Private>` missing | Navisworks DLLs copied to output dir | Set `<Private>False</Private>` on all NW references |
| Build config doesn't target correct NW | Wrong configuration selected | Check Navisworks version string in configuration condition |
| `NullReferenceException` on `ActiveDocument` | No file is open in Navisworks | Always null-check `Application.ActiveDocument` |
| LINQ not available | Missing `System.Linq` using | Add `using System.Linq;` |
| COM interop error on property write | `EmbedInteropTypes` must be `True` | Set `<EmbedInteropTypes>True</EmbedInteropTypes>` in csproj |

---

## Quick Reference: Plugin Registration Summary

```
Plugin Type           → Base Class              → Purpose
──────────────────────────────────────────────────────────
CommandHandlerPlugin  → CommandHandlerPlugin    → Ribbon buttons
AddInPlugin           → AddInPlugin             → Automation entry point
EventWatcherPlugin    → EventWatcherPlugin      → App menu + lifecycle
DockPanePlugin        → DockPanePlugin          → Custom dockable tool windows

Attribute             → Target         → Effect
──────────────────────────────────────────────────────────
[PluginAttribute]     → Plugin class   → Plugin name + developer ID
[RibbonLayoutAttribute] → CommandHandler → XAML file reference
[RibbonTab]           → CommandHandler → Declares a ribbon tab
[Command]             → CommandHandler → Registers a button command
[AddInPlugin]         → AddIn class    → Sets addin location in UI
[DockPanePlugin]      → DockPane class → Declares dockable pane location
```

---

## 24. Advanced API Features & Plugins

---

### 24.1 DockPanePlugin — Custom Dockable Windows

A `DockPanePlugin` lets you embed a fully custom WPF or WinForms user control directly into the Navisworks workspace as a **native dockable panel** — exactly like the built-in Properties or Selection Tree panels.

#### Step 1 — Create the User Control

Create a standard WPF `UserControl` (or WinForms `UserControl`) that contains your panel's UI.

```xml
<!-- Views/MyPanelControl.xaml -->
<UserControl x:Class="MyPlugin.Views.MyPanelControl"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Grid>
        <StackPanel Margin="8">
            <TextBlock Text="My Dockable Panel" FontWeight="Bold" FontSize="14" Margin="0,0,0,8"/>
            <Button Content="Refresh Model Info" Name="BtnRefresh" Click="BtnRefresh_Click"
                    Padding="8,4" Margin="0,0,0,4"/>
            <ListBox Name="LstItems" Height="200"/>
            <TextBlock Name="TxtStatus" Text="Ready" Margin="0,8,0,0" Foreground="Gray"/>
        </StackPanel>
    </Grid>
</UserControl>
```

```csharp
// Views/MyPanelControl.xaml.cs
using Autodesk.Navisworks.Api;
using System.Windows.Controls;

namespace MyPlugin.Views
{
    public partial class MyPanelControl : UserControl
    {
        public MyPanelControl()
        {
            InitializeComponent();
        }

        private void BtnRefresh_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            LstItems.Items.Clear();
            var doc = Application.ActiveDocument;
            if (doc == null) { TxtStatus.Text = "No document open."; return; }

            // Populate list with selected item display names
            foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
                LstItems.Items.Add(item.DisplayName);

            TxtStatus.Text = $"Found {LstItems.Items.Count} selected items.";
        }
    }
}
```

#### Step 2 — Create the DockPanePlugin Class

```csharp
// MyDockPane.cs
using Autodesk.Navisworks.Api.Plugins;
using MyPlugin.Views;
using System.Windows.Controls;

namespace MyPlugin
{
    [Plugin("MyDockPane",
            "MyCompany",
            DisplayName = "My Panel")]
    [DockPanePlugin(
        500,              // Default width in pixels
        300,              // Default height in pixels
        FixedSize = false // Allow the user to resize it
    )]
    public class MyDockPane : DockPanePlugin
    {
        private MyPanelControl _control;

        /// <summary>
        /// Called once when Navisworks first creates the pane.
        /// Return the WPF/WinForms control to host inside it.
        /// </summary>
        public override System.Windows.FrameworkElement CreatePanelControl()
        {
            _control = new MyPanelControl();
            return _control;
        }
    }
}
```

#### Step 3 — Show the Pane from a Ribbon Button

Use `Application.Gui.ShowPane()` to open the dock pane from any command:

```csharp
case "ID_CMD_ShowMyPanel":
    // The string argument is: "AssemblyName.Namespace.ClassName"
    Application.Gui.ShowPane("MyPlugin.MyDockPane");
    break;
```

> **Tips:**
> - The argument to `ShowPane` is the **fully-qualified plugin class name** (namespace + class).
> - To hide the pane programmatically: `Application.Gui.HidePane("MyPlugin.MyDockPane");`
> - You can host **WinForms** controls too by wrapping them in a `System.Windows.Forms.Integration.WindowsFormsHost`.

---

### 24.2 Selection Sets

Navisworks lets users save named **Selection Sets** and **Search Sets** in the Sets panel. You can read and write them programmatically.

#### Reading Existing Selection / Search Sets

```csharp
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;

public static void ListAllSelectionSets()
{
    Document doc = Application.ActiveDocument;
    DocumentSavedItemsRoot setsRoot = doc.SavedItems;

    // Recursively visit all saved items (sets can be nested in folders)
    PrintSets(setsRoot.Value);
}

private static void PrintSets(SavedItemCollection items)
{
    foreach (SavedItem item in items)
    {
        if (item is GroupItem folder)
        {
            System.Diagnostics.Debug.WriteLine($"[Folder] {folder.DisplayName}");
            PrintSets(folder.Children);  // recurse into folder
        }
        else if (item is SelectionSet selSet)
        {
            System.Diagnostics.Debug.WriteLine($"[SelectionSet] {selSet.DisplayName}");

            // Apply it to the current selection
            // doc.CurrentSelection.SelectedItems = selSet.GetSelectedItems(doc);
        }
        else if (item is SearchSet searchSet)
        {
            System.Diagnostics.Debug.WriteLine($"[SearchSet] {searchSet.DisplayName}");
        }
    }
}
```

#### Applying a Named Selection Set

```csharp
public static void ApplySelectionSetByName(string setName)
{
    Document doc = Application.ActiveDocument;

    // Find the set by name (flat search across all sets)
    foreach (SavedItem item in doc.SavedItems.Value)
    {
        if (item is SelectionSet selSet &&
            selSet.DisplayName.Equals(setName, StringComparison.OrdinalIgnoreCase))
        {
            // Activate the selection set
            doc.CurrentSelection.SelectedItems.CopyFrom(selSet.GetSelectedItems(doc));
            return;
        }
    }

    MessageBox.Show($"Selection set '{setName}' not found.",
        "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning,
        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
}
```

#### Saving the Current Selection as a New Set

```csharp
public static void SaveCurrentSelectionAsSet(string newSetName)
{
    Document doc = Application.ActiveDocument;
    ModelItemCollection selected = doc.CurrentSelection.SelectedItems;

    if (selected.Count == 0)
    {
        MessageBox.Show("Nothing selected.", "Warning",
            MessageBoxButtons.OK, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
        return;
    }

    // Create a new selection set
    SelectionSet newSet = new SelectionSet(selected);
    newSet.DisplayName = newSetName;

    // Add it to the root of the Sets panel
    doc.SavedItems.Value.Add(newSet);

    MessageBox.Show($"Saved \"{newSetName}\" with {selected.Count} items.",
        "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information,
        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
}
```

---

### 24.3 Saved Viewpoints

Saved viewpoints capture the camera position, orientation, and render mode so users can return to key views instantly.

#### Iterating All Saved Viewpoints

```csharp
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;

public static void ListSavedViewpoints()
{
    Document doc = Application.ActiveDocument;
    DocumentSavedViewsRoot viewpointsRoot = doc.SavedViewpoints;

    WalkViewpoints(viewpointsRoot.Value);
}

private static void WalkViewpoints(SavedItemCollection items)
{
    foreach (SavedItem item in items)
    {
        if (item is GroupItem folder)
        {
            System.Diagnostics.Debug.WriteLine($"[Folder] {folder.DisplayName}");
            WalkViewpoints(folder.Children);
        }
        else if (item is SavedViewpoint vp)
        {
            System.Diagnostics.Debug.WriteLine($"[Viewpoint] {vp.DisplayName}");
        }
    }
}
```

#### Applying a Saved Viewpoint by Name

```csharp
public static void GoToViewpoint(string viewpointName)
{
    Document doc = Application.ActiveDocument;

    foreach (SavedItem item in doc.SavedViewpoints.Value)
    {
        if (item is SavedViewpoint vp &&
            vp.DisplayName.Equals(viewpointName, StringComparison.OrdinalIgnoreCase))
        {
            // Apply the viewpoint to the active view
            doc.SavedViewpoints.CurrentSavedViewpoint = vp;
            return;
        }
    }

    MessageBox.Show($"Viewpoint '{viewpointName}' not found.",
        "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning,
        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, false);
}
```

#### Saving the Current Camera as a New Viewpoint

```csharp
public static void SaveCurrentView(string name)
{
    Document doc = Application.ActiveDocument;

    // Snapshot the current camera state
    SavedViewpoint newVp = new SavedViewpoint(doc.SavedViewpoints.CurrentViewpoint);
    newVp.DisplayName = name;

    doc.SavedViewpoints.Value.Add(newVp);
}
```

---

### 24.4 Overriding Item Colors & Transparency

You can programmatically change the **display colour and transparency** of model items without altering the source model data. This is useful for highlighting search results, clash groups, or status-based colouring.

#### Override the Color of Selected Items

```csharp
using Autodesk.Navisworks.Api;

public static void OverrideItemColor(
    ModelItemCollection items,
    byte r, byte g, byte b,
    float transparency = 0f)   // 0 = opaque, 1 = fully transparent
{
    Document doc = Application.ActiveDocument;

    // Build the override: custom colour + transparency
    Color navisColor = new Color(r / 255.0, g / 255.0, b / 255.0);
    var colorOverride = new ColorTransparency(navisColor, transparency);

    doc.Models.OverridePermanentColor(items, colorOverride);
}

// Usage: highlight selected items in red
var selected = Application.ActiveDocument.CurrentSelection.SelectedItems;
OverrideItemColor(selected, 255, 0, 0, 0f);
```

#### Reset Color/Transparency Overrides

```csharp
public static void ResetItemColors(ModelItemCollection items)
{
    Document doc = Application.ActiveDocument;
    doc.Models.ResetPermanentColor(items);
}

// Clear ALL overrides in the entire document
public static void ResetAllColors()
{
    Document doc = Application.ActiveDocument;
    // Collect all model items
    Search search = new Search();
    search.Selection.SelectAll();
    var all = search.FindAll(doc, false);
    doc.Models.ResetPermanentColor(all);
}
```

#### Make Items Semi-Transparent

```csharp
public static void MakeItemsTransparent(ModelItemCollection items, float transparency = 0.7f)
{
    // Keep original colour but apply transparency (0 = opaque, 1 = invisible)
    Application.ActiveDocument.Models.OverridePermanentTransparency(items, transparency);
}
```

> **Note:** These overrides are **session-only** — they do not get saved with the `.nwd` file unless you explicitly use `SaveAs`. If you need persistent overrides, write them as a Clash/Redline or use custom property categories.

---

### 24.5 Bounding Boxes & Geometry Constraints

Every `ModelItem` with geometry has a **bounding box** that describes its axis-aligned extents in world coordinates. This is extremely useful for spatial queries, proximity detection, and centre-of-gravity calculations.

#### Extracting the Bounding Box of a ModelItem

```csharp
using Autodesk.Navisworks.Api;

public static void PrintBoundingBox(ModelItem item)
{
    if (!item.HasGeometry)
    {
        System.Diagnostics.Debug.WriteLine($"{item.DisplayName}: no geometry");
        return;
    }

    BoundingBox3D bbox = item.FindFirstGeometryLayer<Autodesk.Navisworks.Api.Geometry.PrimitiveLibrary>()
        ?.BoundingBox;

    // Simpler approach: use the composite BoundingBox from the Model API
    BoundingBox3D bb = item.BoundingBox();

    Point3D min = bb.Min;   // Lower-left-front corner
    Point3D max = bb.Max;   // Upper-right-back corner

    // Compute the centre
    double cx = (min.X + max.X) / 2.0;
    double cy = (min.Y + max.Y) / 2.0;
    double cz = (min.Z + max.Z) / 2.0;

    // Compute extents (dimensions)
    double width  = max.X - min.X;
    double depth  = max.Y - min.Y;
    double height = max.Z - min.Z;

    System.Diagnostics.Debug.WriteLine(
        $"{item.DisplayName}: " +
        $"Min=({min.X:F3}, {min.Y:F3}, {min.Z:F3}) " +
        $"Max=({max.X:F3}, {max.Y:F3}, {max.Z:F3}) " +
        $"Centre=({cx:F3}, {cy:F3}, {cz:F3}) " +
        $"Size=({width:F3} x {depth:F3} x {height:F3})");
}
```

#### Aggregate Bounding Box for a Group of Items

```csharp
public static BoundingBox3D GetGroupBoundingBox(ModelItemCollection items)
{
    double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
    double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

    foreach (ModelItem item in items)
    {
        if (!item.HasGeometry) continue;

        BoundingBox3D bb = item.BoundingBox();
        if (bb.Min.X < minX) minX = bb.Min.X;
        if (bb.Min.Y < minY) minY = bb.Min.Y;
        if (bb.Min.Z < minZ) minZ = bb.Min.Z;
        if (bb.Max.X > maxX) maxX = bb.Max.X;
        if (bb.Max.Y > maxY) maxY = bb.Max.Y;
        if (bb.Max.Z > maxZ) maxZ = bb.Max.Z;
    }

    return new BoundingBox3D(
        new Point3D(minX, minY, minZ),
        new Point3D(maxX, maxY, maxZ));
}
```

#### Using Bounding Boxes for Proximity Search

This pattern (used extensively in the SAIPEM repo) finds items whose centre falls within a spatial offset of a target element:

```csharp
public static ModelItemCollection GetItemsWithinOffset(
    ModelItemCollection candidateItems,
    ModelItem targetItem,
    double offset)
{
    BoundingBox3D targetBB = targetItem.BoundingBox();
    double cx = (targetBB.Min.X + targetBB.Max.X) / 2.0;
    double cy = (targetBB.Min.Y + targetBB.Max.Y) / 2.0;
    double cz = (targetBB.Min.Z + targetBB.Max.Z) / 2.0;

    // Use the Search API with numeric comparison on stored Centre properties
    Search search = new Search();
    var X1 = SearchCondition.HasPropertyByDisplayName("SAIPEM", "Center.X")
        .CompareWith(SearchConditionComparison.NumericLessThanOrEqual,    new VariantData(cx + offset));
    var X2 = SearchCondition.HasPropertyByDisplayName("SAIPEM", "Center.X")
        .CompareWith(SearchConditionComparison.NumericGreaterThanOrEqual,  new VariantData(cx - offset));
    var Y1 = SearchCondition.HasPropertyByDisplayName("SAIPEM", "Center.Y")
        .CompareWith(SearchConditionComparison.NumericLessThanOrEqual,    new VariantData(cy + offset));
    var Y2 = SearchCondition.HasPropertyByDisplayName("SAIPEM", "Center.Y")
        .CompareWith(SearchConditionComparison.NumericGreaterThanOrEqual,  new VariantData(cy - offset));

    search.SearchConditions.Add(X1);
    search.SearchConditions.Add(X2);
    search.SearchConditions.Add(Y1);
    search.SearchConditions.Add(Y2);
    search.Selection.CopyFrom(candidateItems);

    return search.FindAll(Application.ActiveDocument, true);
}
```

> **Tip:** For best performance on large models, pre-compute and store bounding box centres as custom properties (using `SetParameterValue` from Section 14.5), then use the `Search` API to filter spatially — this avoids iterating every item in C# code and is significantly faster.

---

*This is a standalone, generic tutorial for building Autodesk Navisworks plugins with C# and .NET Framework.*  
*Applies to Navisworks Manage 2022–2026.*

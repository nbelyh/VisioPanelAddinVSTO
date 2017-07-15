# Extended Visio Addin Project

This extension contains a project template to create a extneded 
Microsoft Visio Addin based on Visual Tools for Office or vanilla COM Shared AddIn in C# and VB.NET 

Published on:
https://marketplace.visualstudio.com/items?itemName=NikolayBelyh.ExtendedVisioAddinProject

In addition to standard functionality, the project generated with this extension features the following:

- User interface to start with
- A TaskPane (docking panel), and a button to control it.
- Support for custom images for the buttons.
- Optional support for state (enabled/disabled) for the buttons.
- Optional support for legacy Visio version (command bar with buttons)
- Optional support for ribbon designer

The add-in generated should work for Visio 2003 SP3 (COM addin only), 2007 SP3, 2010, 2013, 2016.

- Both x32 and x64 Visio versions are supported (the code compiles to Any CPU).
- Support per user/per machine install in one MSI
- Visio files publishing functionality ("featured" Templates/Stencil on the start page and in menu)
- Visio Addon (VSL) publishing support
- Visio Help files (CHM) publishign support
- Localization support (non-english UI for the installer)

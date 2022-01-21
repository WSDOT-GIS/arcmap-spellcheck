Spell Check for ArcMap
======================

**[Esri is deprecating ArcMap](https://support.esri.com/en/arcmap-esri-plan), so this project is no longer active**

Requirements
------------

* ArcMap 10.5.1 or higher
* User must have Microsoft Word installed. (Tested with Word 2016. Other versions may or may not work.)

### Development Requirements ###

* Visual Studio 2015
* Visual Studio 2013 Isolated Shell
* ArcObjects .NET SDK

The idea for this add-in was based on a Visual Basic for Applications (VBA) code example by Esri: [How to spell check text elements in ArcMap].

This version differs from the original sample on which it was based in the following ways:

* Written in C# instead of VBA.
* Is implemented as an ArcMap add-in instead of a VBA macro.
* This version of the tool checks the spelling of all text elements as well as the names of all maps and layers, in the current ArcMap document.
* The original version would only check a single selected text element, and would crash if no text elements were selected.

[How to spell check text elements in ArcMap]:http://help.arcgis.com/en/sdk/10.0/vba_desktop/conceptualhelp/index.html#//0001000000qt000000

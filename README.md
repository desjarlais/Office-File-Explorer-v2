# Office-File-Explorer-v2

The purpose of this tool is to provide potential file specific troubleshooting of Office Open Xml formatted documents for Word, Excel and PowerPoint (.docx, .dotx, .docm, .xlsx, .xlst, .xlsm, .pptx, .pptm).

The tool can also perform a series of checks for known document corruptions in the underlying xml tags to possibly fix and make the file open in Office.  

This tool is built for .NET Core 6 and you should be prompted to download from the Microsoft site if you don't have .NET 6 installed.

## Note
Keep in mind if you use this on a production document and choose to use something that changes or removes data, you should be working on a copy of the file, not the original.  

## Help
If you need assistance (find a bug, have a question or any suggestions or feedback), please report them using the [Issues tab](https://github.com/desjarlais/Office-File-Explorer-v2/issues)

## Main Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofe2.jpg?raw=true)

## List of features

Use the Wiki for more information about the features listed here -> [Wiki](https://github.com/desjarlais/Office-File-Explorer-v2/wiki)

### Word
* Display the following document content: (content controls, styles, hyperlinks, List Templates, fonts, footnotes, endnotes, document properties, revisions/tracked changes, comments, field codes, bookmarks)
* Delete content (headers / footers, orphan list templates, page breaks, comments, hidden text, footnotes, endnotes, unused styles)
* Convert Macro enabled file (.docm) to non-macro enabled (.docx)
* Set Print Orientation
* Change Default Template
* Accept All Revisions
* Fix document corruptions (bookmarks, revisions, endnotes, list templates, table properties, comments, hyperlinks, content controls, math formulas)
* Remove PII

### Excel
* List function to display (links, comments, worksheets, hidden rows & columns, shared strings, connections, defined names, hyperlinks)
* Delete content (comments, links)
* Convert Macro enabled file (.xlsm) to non-macro enabled (.xlsx) 
* Convert Strict .xlsx format to non-Strict .xlsx format
* Added a Sheet Viewer form to look at cell values, formulas and font information for cells

### PowerPoint
* List function to display (hyperlinks, slide titles, slide text, comments, slide transitions)
* Convert Macro enabled file (.pptm) to non-macro enabled (.pptx)
* Reset note page size to default value (if the button doesn't fix the issue, go to File | Settings and enable the notes master checkbox)
* Reset note page size to custom value based on another presentation file

### Shared
* List function to display (Ole Objects, shapes, package parts)
* View Custom Document Properties
* View embedded images in file
* View Ribbon/Backstage customizations (customUI)
* Search and Replace
* Add custom properties for a file
* Change theme for a file
* Validate underlying xml file details
* View Custom Xml - view custom xml files in Office documents

## Batch File Processing Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofe2batch.jpg?raw=true)

### Authentication Note
There is an ability to login with the Batch Processing Window, this is a work in progress, see the following wiki for more information
[Authentication](https://github.com/desjarlais/Office-File-Explorer-v2/wiki/Authentication-in-Batch-Processing)

### Batch File Processing Features
* Change Theme (Word, Excel, PowerPoint)
* Add Custom Properties (Word, Excel, PowerPoint)
* Delete Custom Properties (Word, Excel, PowerPoint)
* Reset note page size to default value (PowerPoint)
* Fix corrupt bookmarks (Word)
* Fix corrupt revisions (Word)
* Fix corrupt table properties (Word)
* Fix corrupt hyperlinks (Word)
* Update Quick Part Namespaces (Word)
* Remove Personally Identifiable Information (Word)
* Remove PII On Save (PowerPoint)
* Set OpenByDefault = false (Word)
* Convert Strict xlsx format to non-Strict xlsx format (Excel)
* Delete RequestStatus xml node - (Word, Excel, PowerPoint)

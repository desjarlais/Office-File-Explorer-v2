# Office-File-Explorer-v2

The purpose of this tool is to provide potential file specific troubleshooting of Office Open Xml formatted documents for Word, Excel, PowerPoint and Outlook (.docx, .dotx, .docm, .doc, .xlsx, .xlst, .xlsm, .xls, .pptx, .pptm, .ppt, .msg).  It also includes the ability to view the underlying xml files and modify them directly.  

The tool can also perform a series of checks for known document corruptions in the underlying xml tags to possibly fix and make the file open in Office.  

This tool is built for .NET Core 6 and you should be prompted to download from the Microsoft site if you don't have .NET 6 installed.

## Note
Keep in mind if you use this on a production document and choose to use something that changes or removes data, you should be working on a copy of the file, not the original.  

## Main Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/a570b6d425ae62a896722d391efaa957683f9ecb/img/ofcscreenshot.png)

## List of features

Use the Wiki for more information about the features listed here -> [Wiki](https://github.com/desjarlais/Office-File-Explorer-v2/wiki)

### Word
* Display the following document content: (content controls, styles, hyperlinks, List Templates, fonts, footnotes, endnotes, document properties, revisions/tracked changes, comments, field codes, bookmarks)
* Delete content (headers / footers, orphan list templates, page breaks, comments, hidden text, footnotes, endnotes, unused styles)
* Convert Macro enabled file (.docm) to non-macro enabled (.docx)
* Set Print Orientation
* Change Default Template
* Accept All Revisions
* Fix content control namespaces
* Fix document corruptions (see [Fix Document Feature](https://github.com/desjarlais/Office-File-Explorer-v2/wiki/Fix-Document-Feature) for more information)
* Remove PII

### Excel
* List function to display (links, comments, worksheets, hidden rows & columns, shared strings, connections, defined names, hyperlinks)
* Delete content (comments, links)
* Convert Macro enabled file (.xlsm) to non-macro enabled (.xlsx) 
* Convert Strict .xlsx format to non-Strict .xlsx format
* Added a Sheet Viewer form to look at cell values, formulas and font information for cells

### PowerPoint
* List function to display (hyperlinks, slide titles, slide text, comments, slide transitions, fonts)
* Convert Macro enabled file (.pptm) to non-macro enabled (.pptx)
* Reset note page size to default value (if the button doesn't fix the issue, go to File | Settings and enable the notes master checkbox)
* Reset note page size to custom value based on another presentation file

### Shared
* View Sensitivity Labeled files (credit to [OpenMcdf](https://github.com/ironfede/openmcdf) for the heavy lifting parsing the compound file format)
* List function to display (Ole Objects, shapes, package parts)
* View Custom Document Properties
* View embedded images in file
* View Ribbon/Backstage customizations (customUI)
* Search and Replace
* Add custom properties for a file
* Change theme for a file
* Validate underlying xml file details
* View Custom Xml - view custom xml files in Office documents
* View Custom UI - borrowed most of the code from [office-custom-ui-editor](https://github.com/OfficeDev/office-custom-ui-editor)

## Batch File Processing Window 
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofe2batch.jpg?raw=true)

More info in the wiki here -> [Batch Processing Wiki](https://github.com/desjarlais/Office-File-Explorer-v2/wiki/Batch-Processing)

### Batch File Processing Features
* Change Attached Template (Word)
*	Add Custom Properties (Word, Excel, PowerPoint)
*	Convert Strict To Non-Strict (Excel)
*	Fix Corrupt Bookmarks (Word)
*	Remove PII On Save (PowerPoint)
*	Remove PII (Word)
*	Set OpenByDefault = False (Word)
*	Delete Custom Props (Word, Excel, PowerPoint)
*	Remove Custom Title Prop (Word)
*	Fix Corrupt Revisions (Word)
*	Delete RequestStatus (Word, Excel, PowerPoint)
*	Change Theme (Word, Excel, PowerPoint)
*	Update Quick Part Namespaces (Word)
*	Fix Corrupt Hyperlinks (Word)
*	Fix Notes Page Size (PowerPoint)
*	Fix Table Grid Props (Word)
*	Fix Corrupt Comments (Word)
*	Reset Bullet Tab Margins (PowerPoint)
*	Check For Digital Signatures (Word, Excel, PowerPoint)
*	Fix Footer Spacing (Word)
*	Remove Custom File Props  (Word)
*	Fix Corrupt Table Cells (Word)
*	Remove Custom Xml (Word, Excel, PowerPoint)
*	Fix Duplicate Custom Xml (Word)

## Help
If you need assistance (find a bug, have a question or any suggestions or feedback), please report them using the [Issues tab](https://github.com/desjarlais/Office-File-Explorer-v2/issues)

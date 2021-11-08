# Office-File-Explorer-v2

The purpose of this tool is to provide potential file specific troubleshooting of Office Open Xml formatted documents for Word, Excel and PowerPoint (.docx, .dotx, .docm, .xlsx, .xlst, .xlsm, .pptx, .pptm).

I've also added a couple of known document fixes where the xml tags in the underlying file format need to be tweaked to make the file openable again.  

## Note
Keep in mind if you use this on a production document and choose to use something that changes or removes data, you should be working on a copy of the file, not the original.  

## Help
If you need assistance (find a bug, have a question or any suggestions or feedback), please report them using the [Issues tab](https://github.com/desjarlais/Office-File-Explorer/issues)

## Main Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofe2.jpg?raw=true)

## List of features

### Word
* List function to display (content controls, styles, hyperlinks, List Templates, fonts, footnotes, endnotes, document properties, authors, revisions/tracked changes, comments, field codes, bookmarks, paragraphs, paragraph styles)
* Delete content (headers / footers, orphan list templates, page breaks, comments, hidden text, footnotes, endnotes)
* Convert Macro enabled file (.docm) to non-macro enabled (.docx)
* Fix corrupt documents
* Fix corrupt bookmarks
* Fix corrupt revisions
* Fix corrupt endnotes
* Fix orphan / unused ListTemplates (Numbering)
* Fix corrupt table properties
* Fix corrupt commments
* Remove PII
* Delete unused styles

### Excel
* List function to display (links, comments, worksheets, hidden rows & columns, shared strings, cell values, connections, defined names)
* Delete content (comments, links)
* Convert Macro enabled file (.xlsm) to non-macro enabled (.xlsx) 
* Convert Strict .xlsx format to non-Strict .xlsx format

### PowerPoint
* List function to display (hyperlinks, slide titles, slide text, comments)
* Convert Macro enabled file (.pptm) to non-macro enabled (.pptx)
* Reset note page size to default value (if the button doesn't fix the issue, go to File | Settings and enable the notes master checkbox)
* Remove PII
* Fix Presentation - fix corrupt notes slides / pages

### Shared
* List function to display (Ole Objects, shapes, custom properties, package parts)
* Add custom properties for a file
* Change theme for a file
* Validate underlying xml file details
* View Custom Xml - view custom xml files in Office documents

## Batch File Processing Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofe2batch.jpg?raw=true)

### Batch File Processing (following features can be used to change many documents at one time)
* Change Theme (Word, Excel, PowerPoint)
* Add Custom Properties (Word, Excel, PowerPoint)
* Reset note page size to default value (PowerPoint)
* Fix corrupt bookmarks (Word)
* Fix corrupt revisions (Word)
* Fix corrupt table properties (Word)
* Remove Personally Identifiable Information (Word)
* Remove PII (PowerPoint)
* Convert Strict xlsx format to non-Strict xlsx format (Excel)
* Delete RequestStatus xml node - (Word, Excel, PowerPoint)

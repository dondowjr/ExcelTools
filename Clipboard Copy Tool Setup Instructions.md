# Clipboard Copy Tool Setup Instructions

The Clipboard Copy Tool is a handy excel workbook that allows you to copy text to the clipboard by simply clicking into a cell. 

The tool is usedful when you are filling out forms where there is a significant amount of repeated data to fill out. For example, you could user this tool for mailing address, pre-written email clips, phone numbers, or other standarized text.

## Setting up a new Workbook

The layout Workbook is faily simple. There are three main components, the first is the on/off cell that allows you to turn on/off the copy functionality for the workbook.  The next two components are simply a two column table that represent a title/description of the text that is to be copied, and then the text itself. It will make a lot of sense once you create the workbook.

1. Open excel and start a new workbook.
2. Turn off gridlines - Click on View, then uncheck the Gridlines option in the Show group box. 
3. Add the following text to the spreadsheet:
	a.	B2: “Copying is Off”, also bold the text.
	b.	B4: “Title/Description”, also bold the text
	c.	C5: “Text to Copy”, also bold the text.
4.	Add all boarders to:
	a.	B2 (one cell)
	b.	B5:C100
5.	Resize columns:
	a.	Make column A about the width of the default row height
	b.	Expand column B to at least fit the Title/Description text.
	c.	Expand column C to be 4x the size of column A
6.	Save the your work as a macro enabled workbook:
a.	Click on File > Save As >  Browse
b.	Give your spreadsheet a name
c.	Change the Save as Type to: Excel Macro Enabled Workbook
d.	Save
7.	Enable the Developer “tab”
a.	File > Options
b.	Customize Ribbon
c.	Select “Main Tabs” in the “Customize the Ribbon” drop list (right side of the window.)
d.	Find the “Developer” option in the list and check it.
e.	Click Ok
f.	Open the code window for sheet 1
i.	Developer > Visual Basic
ii.	Under the VBAProject node in the tree (left side), find the “Sheet1” node and double click it. This will open the code window for Sheet1

8.	Paste in the following VBA code.



## Adding the VBA Code




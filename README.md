# excelsaver
This program selects every Excel Process (from list of Processes) and checks if there are opened workbooks. If yes, it saves every Excel document as a new file (path: C:\ExcelSaver). If cell is in edit mode, it has to move app window foreground first then it simulates 'enter' keypress. Then it closes Excel app, cleans memory and reopens all saved documents. 

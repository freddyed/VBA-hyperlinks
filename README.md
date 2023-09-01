# Extracting Hyperlinks with VBA
Showcasing hyperlinks sourced from Excel cells

**First step:**

Open Excel and navigate to the "Developer" tab.


**Second step:**

Click on "Visual Basic" to access the VBA editor.


**Third step:**

Insert a new module within the editor.


**Fourth step:**

Copy and paste the following VBA code into the module. Once done, close the windows.

```vba
Function ExtraerHipervinculo(celda As Range) As String
  On Error Resume Next
  ExtraerHipervinculo = celda.Hyperlinks(1).Address
  On Error GoTo 0 
End Function
```
**Fifth step:**

Utilize the function "=ExtraerHipervinculo" to commence the hyperlink extraction process. Rejoice in the results!

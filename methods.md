
### methods
ref. : https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/object-model-excel-vba-reference
##### add sheet
> ActiveWorkbook.Sheets.Add after:=Worksheets(Worksheets.Count)

##### autofilter ( http://www.contextures.com/xlautofilter03.html )
> ThisWorkbook.Sheets(1).AutoFilterMode = False  
> With Range(Cells(11, "L"), Cells(lst_row, "AL"))  
> &#160;&#160;&#160;&#160;.AutoFilter  
> &#160;&#160;&#160;&#160;.AutoFilter Field:=1, Criteria1:="=*bank*", Operator:=xlOr, Criteria2:="=*city*"  
> &#160;&#160;&#160;&#160;.AutoFilter Field:=25, Criteria1:="<>#N/A"  
> End With
 
##### 'check for filter, turn on if none exists  
> If Not ActiveSheet.AutoFilterMode Then  
> &#160;&#160;&#160;&#160;ActiveSheet.Range("A1").AutoFilter  
> End If
 
##### 'removes AutoFilter if one exists  
> ActiveSheet.AutoFilterMode = False

##### calculation manu/auto
> Application.Calculation = xlManual
> Application.Calculation = xlAutomatic

#####  close workbook
> ActiveWorkbook.Close SaveChanges:=True
> ActiveWorkbook.Close SaveChanges:=False

##### copy
> Range(Cells(12, "AH"), Cells(12, "AL")).Copy Range(Cells(13, "AH"), Cells(lst_row, "AL"))
> Application.CutCopyMode = False

##### copy sheet
> Workbooks(fl_priv).Sheets("SheetName").Copy Before:=Workbooks(fl_comb).Sheets(5)

##### multi-line messagebox
> MsgBox " ******first line******" & vbCrLf & _  
> &#160;&#160;&#160;&#160;" ******second line******" _  
> &#160;&#160;&#160;&#160;, vbYesNo

##### open file by Excel
> Workbooks.Open Filename:=ThisWorkbook.Path & "\settle.csv"

##### open webpage by
> Shell "D:\Programs\FirefoxPortable\FirefoxPortable.exe http://www.web.net/index.html", vbNormalFocus

##### save book
> Workbooks.Add
> ActiveWorkbook.SaveAs Filename:="E:\" & filename & ".xlsx"

#### select folder to open files
> Set Fold = CreateObject("shell.application").BrowseForFolder(0, "Please select folder:", 0, 0) 'from desktop
> If Fold Is Nothing Then Exit Sub 'exit while no selection
> filepath = Fold.Items.Item.Path & "\" 'set data path

#### select file to open
> FileToOpen = Application.GetOpenFilename (Title:="Please choose a file to open", FileFilter:="Excel Files *.xls* (*.xls*),")
> If FileToOpen = False Then
>     MsgBox "No file selected.", vbExclamation, "Sorry!"
>     Exit Sub
> Else
>     Workbooks.Open Filename:=FileToOpen
> End If

##### select special cells only
> Selection.SpecialCells(xlCellTypeVisible).Select
> Selection.SpecialCells(xlCellTypeBlanks).Select

##### Send key
 ref. : https://docs.microsoft.com/en-us/office/vba/api/excel.application.sendkeys

> Application.SendKeys ("{F2}")  
> Application.SendKeys ("0")  
> Application.SendKeys ("%~")  
> Application.SendKeys ("{RETURN} ")

##### text to columns (column 2 as text)
> Range("A2:A10").Select  
> Selection.TextToColumns DataType:=xlDelimited, Tab:=True, FieldInfo:=Array(Array(1, 1), Array(2, 2))

##### whole range
> Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell)).Select

##### mailmerge start auto merge
> With ActiveDocument.MailMerge
> &#160;&#160;&#160;&#160;.Destination = wdSendToNewDocument
> &#160;&#160;&#160;&#160;.SuppressBlankLines = True
> &#160;&#160;&#160;&#160;With .DataSource
> &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;.FirstRecord = 1
> &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;.LastRecord = 1
> &#160;&#160;&#160;&#160;End With
> &#160;&#160;&#160;&#160;.Execute Pause:=False
> End With

##### icident
put ThisWorkbook.Close under Private Sub Workbook_Open() in ThisWorkbook - Workbook by mistake

solution:
in a temp file run VBA code
Application.EnableEvents = False

will stop VBA run

then use below code to restore
Application.EnableEvents = True

ref. http://cn.voidcc.com/question/p-bmkwvtpf-mw.html

### methods
ref. : https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/object-model-excel-vba-reference
##### add sheet
> ActiveWorkbook.Sheets.Add after:=Worksheets(Worksheets.Count)

##### autofilter ( http://www.contextures.com/xlautofilter03.html )
> ThisWorkbook.Sheets(1).AutoFilterMode = False  
> With Range(Cells(11, "L"), Cells(lst_row, "AL"))  
>     .AutoFilter  
>     .AutoFilter Field:=1, Criteria1:="=*bank*", Operator:=xlOr, Criteria2:="=*city*"  
>     .AutoFilter Field:=25, Criteria1:="<>#N/A"  
> End With
 
##### 'check for filter, turn on if none exists  
> If Not ActiveSheet.AutoFilterMode Then  
> ActiveSheet.Range("A1").AutoFilter  
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
> " ******second line******" _  
> , vbYesNo

##### open file by Excel
> Workbooks.Open Filename:=ThisWorkbook.Path & "\settle.csv"

##### open webpage by
> Shell "D:\Programs\FirefoxPortable\FirefoxPortable.exe http://report.gserver.net/WebReport", vbNormalFocus

##### save book
> Workbooks.Add
> ActiveWorkbook.SaveAs Filename:="E:\" & sv_book & ".xlsx"

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
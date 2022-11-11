Sub copiaDatos()
Dim libroDatos As Workbook

Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
    Range("U1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.ListObject.QueryTable.Delete
    Selection.ClearContents
    Range("U1").Select
'copiado de informacion de otro libro al actual
 
    
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\procesar.xlsx")
    
    libroDatos.Sheets(3).Range("A1:R1001").Copy
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
        
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\procesar.xlsx")
    libroDatos.Sheets(1).Range("N1:O331").Copy
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
    Range("S1").Select
    ActiveSheet.Paste
    Range("S331").Select
    
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\procesar.xlsx")
    libroDatos.Sheets(2).Range("N2:O671").Copy
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
    Range("S332").Select
    ActiveSheet.Paste
    Range("A1").Select
    
'limpiado de datos y acomodo de columnas
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
     Columns("N:N").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.TextToColumns Destination:=Range("Total__2[[#Headers],[INMUEBLE]]") _
        , DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=True, OtherChar:="-", FieldInfo:=Array(Array(1, 1 _
        ), Array(2, 1), Array(3, 9)), TrailingMinusNumbers:=True
    Columns("N:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("S:T").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


'LA SIGUIENTE LINEA HOMOLOGA LOS ESPACIOS DE "SAT " POR "SAT"
    Columns("M:M").Select
    Selection.Replace What:="SAT ", Replacement:="SAT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Selection.Replace What:="SA", Replacement:="SAT", LookAt:=xlPart, _
        'SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        'ReplaceFormat:=False
 
'LA SIGUIENTE LINEA HOMOLOGA LOS ESPACIOS DE "PRODECON  " POR "PRODECON"
    Columns("M:M").Select
    Selection.Replace What:="PRODECON ", Replacement:="PRODECON", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 
 
        
'LA SIGUIENTE LINEA BORRA DE LA COLUMNA "M" LOS VALORES DE "SAT"
   Sheets("Total").Select
   col = "M"
    texto = "SAT"    '
    valor = texto
    If IsNumeric(texto) Then valor = Val(texto)
    If IsDate(texto) Then valor = CDate(texto)    '
    Application.ScreenUpdating = False
    For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
    If LCase(Cells(i, "M")) = LCase(valor) Then
    Rows(i).Delete
    End If
    Next
    Application.ScreenUpdating = True
    
    
    Sheets("Total").Select
   col = "H"
    texto = "MESA DE SERVICIOS PRODECON"    '
    valor = texto
    If IsNumeric(texto) Then valor = Val(texto)
    If IsDate(texto) Then valor = CDate(texto)    '
    Application.ScreenUpdating = False
    For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
    If LCase(Cells(i, "H")) = LCase(valor) Then
    Rows(i).Delete
    End If
    Next
    Application.ScreenUpdating = True
    
    
     Sheets("Total").Select
   col = "M"
    texto = "PRODECON"    '
    valor = texto
    If IsNumeric(texto) Then valor = Val(texto)
    If IsDate(texto) Then valor = CDate(texto)    '
    Application.ScreenUpdating = False
    For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
    If LCase(Cells(i, "M")) = LCase(valor) Then
    Rows(i).Delete
    End If
    Next
    Application.ScreenUpdating = True
    MsgBox "Filas eliminadas", vbInformation, "DAM"
    
    
End Sub
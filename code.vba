Sub copiaDatos()
    Dim libroDatos As Workbook
    
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
    
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
    col = "A"
    


    
    
    
    
    
    
End Sub



'SEGUNDA MACRO##################################################################################################
Sheets("Hoja1").Select      'nombre de la hoja con la información
col = "A"                   'columna para aplicar la condición
'texto de la condición
'Para una fecha: "10/07/2017" el formato debe ser dd/mm/aaaa
'Para un número: "123"
texto = "ave"    '
valor = texto
If IsNumeric(texto) Then valor = Val(texto)
If IsDate(texto) Then valor = CDate(texto)    '
Application.ScreenUpdating = False
For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
If LCase(Cells(i, "A")) = LCase(valor) Then
Rows(i).Delete
End If
Next
Application.ScreenUpdating = True
MsgBox "Filas eliminadas", vbInformation, "DAM"
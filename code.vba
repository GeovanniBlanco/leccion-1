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
    Range("N331").Select
    
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\procesar.xlsx")
    libroDatos.Sheets(2).Range("N2:O671").Copy
    Set libroDatos = Workbooks.Open("C:\Users\DRS0034\Documents\procesar\total.xlsm")
    Range("N332").Select
    ActiveSheet.Paste
    Range("A1").Select
    
    
    
    
    
End Sub
Attribute VB_Name = "mainClass"
Dim startDate As Date
Dim endDate As Date

Sub main()

    Dim wsControl As Worksheet

    ' Imposta i riferimenti ai fogli
    Set wsControl = ThisWorkbook.Worksheets("CONTROL CENTER")

    ' Legge le date di inizio e fine fornite dall'utente
    startDate = wsControl.Range("C7").Value
    endDate = wsControl.Range("E7").Value

    ' Creo il layout del report per i profumi
    Call GeneraReportProfumi
    Call FormuleReport(ActiveWorkbook.Worksheets("PERFUME"))
    
'    ' Creo il layout del report per i deodoranti
    Call GeneraReportGiftsets
    Call FormuleReport(ActiveWorkbook.Worksheets("GIFTSET"))
'
'    ' Ripeto per i Gift Sets e i Body Mist
    Call GeneraReportTester
    Call FormuleReport(ActiveWorkbook.Worksheets("TESTER"))

    Call stampaReport

End Sub

Private Sub FormuleReport(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim skuStartRow As Long
    Dim formulaStringJ As String
    Dim formulaStringK As String
    Dim firstCellJ As Boolean
    Dim firstCellK As Boolean
    
    ' Trovare l'ultima riga popolata nella colonna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Scorrere tutte le righe della colonna A
    For i = 1 To lastRow
        ' Controllare se la cella in colonna A contiene 5 trattini (riga di separazione)
        If ws.Cells(i, "A").Value = "-----" Then
            formulaStringJ = "=SUM("  ' Iniziare la formula per la colonna J
            formulaStringK = "=SUM("  ' Iniziare la formula per la colonna K
            firstCellJ = True  ' Variabile per gestire la prima cella da inserire nella formula J
            firstCellK = True  ' Variabile per gestire la prima cella da inserire nella formula K
            
            ' Iniziare a scorrere a ritroso fino a trovare un'altra riga con 5 trattini
            skuStartRow = i - 1
            Do While skuStartRow > 0 And ws.Cells(skuStartRow, "A").Value <> "-----" And ws.Cells(skuStartRow, "A").Value <> "SKU"
                ' Per la colonna J: Aggiungere le celle della colonna G se colonna I = 0 e colonna E <> "SUPPLY"
                If ws.Cells(skuStartRow, "I").Value = 0 And ws.Cells(skuStartRow, "E").Value <> "SUPPLY" Then
                    ' Aggiungere un separatore di virgola se non è la prima cella
                    If Not firstCellJ Then
                        formulaStringJ = formulaStringJ & ","
                    End If
                    ' Aggiungere il riferimento alla cella G corrispondente
                    formulaStringJ = formulaStringJ & "G" & skuStartRow
                    firstCellJ = False  ' Dopo la prima cella, usare il separatore
                End If
                
                ' Per la colonna K: Aggiungere le celle della colonna G se colonna I = 0 e colonna E = "SUPPLY"
                If ws.Cells(skuStartRow, "I").Value = 0 And ws.Cells(skuStartRow, "E").Value = "SUPPLY" Then
                    ' Aggiungere un separatore di virgola se non è la prima cella
                    If Not firstCellK Then
                        formulaStringK = formulaStringK & ","
                    End If
                    ' Aggiungere il riferimento alla cella G corrispondente
                    formulaStringK = formulaStringK & "G" & skuStartRow
                    firstCellK = False  ' Dopo la prima cella, usare il separatore
                End If
                
                skuStartRow = skuStartRow - 1
            Loop
            
            ' Chiudere la formula per la colonna J e inserire nella colonna J se ci sono celle da sommare
            If Not firstCellJ Then
                formulaStringJ = formulaStringJ & ")"
                ws.Cells(i, "J").Formula = formulaStringJ
            Else
                ' Se non ci sono celle da sommare per la colonna J, inserire "0"
                ws.Cells(i, "J").Formula = "=0"
            End If
            
            ' Chiudere la formula per la colonna K e inserire nella colonna K se ci sono celle da sommare
            If Not firstCellK Then
                formulaStringK = formulaStringK & ")"
                ws.Cells(i, "K").Formula = formulaStringK
            Else
                ' Se non ci sono celle da sommare per la colonna K, inserire "0"
                ws.Cells(i, "K").Formula = "=0"
            End If
        End If
    Next i
End Sub

Private Sub GeneraReportProfumi()
    ' Definizione delle variabili
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim wsClients As Worksheet
    Dim wsSuppliers As Worksheet
    Dim rngTable As ListObject, rngClients As ListObject, rngSuppliers As ListObject
    Dim cell As Range
    Dim skuList As Collection
    Dim sku As Variant, regDate As Variant, docDate As Variant, docNumber As Variant
    Dim transaction As Variant, customerSupplier As Variant, pieces As Variant, amount As Variant, price As Double
    Dim i As Long, resultRow As Long
    Dim lastSku As String, clientName As String

    ' Imposta i riferimenti ai fogli
    Set wsData = ThisWorkbook.Worksheets("datiBHPC")
    Set wsResult = ThisWorkbook.Worksheets("PERFUME")
    Set wsClients = ThisWorkbook.Worksheets("clienti")
    Set wsSuppliers = ThisWorkbook.Worksheets("fornitori")
    Set rngTable = wsData.ListObjects("movimentiProfumi")
    Set rngClients = wsClients.ListObjects("clientiBHPC")
    Set rngSuppliers = wsSuppliers.ListObjects("fornitoriBHPC")

    ' Inizializza la collezione per gli SKU
    Set skuList = New Collection

    ' Scansiona la tabella e registra i dati con date comprese tra startDate e endDate
    For Each cell In rngTable.ListColumns("DT#REG#").DataBodyRange
        If cell.Value >= startDate And cell.Value <= endDate Then
            sku = cell.Offset(0, rngTable.ListColumns("SKU CODE").Index - rngTable.ListColumns("DT#REG#").Index).Value
            If sku Like "P#####" Or sku Like "PBM###" Then
                regDate = cell.Value
                docDate = cell.Offset(0, rngTable.ListColumns("DT#DOC#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                docNumber = cell.Offset(0, rngTable.ListColumns("N#DOC#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                transaction = cell.Offset(0, rngTable.ListColumns("CAUSALE MOVIM#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                customerSupplier = cell.Offset(0, rngTable.ListColumns("CLI/FOR NUMBER").Index - rngTable.ListColumns("DT#REG#").Index).Value
                pieces = cell.Offset(0, rngTable.ListColumns("QUANTITA'").Index - rngTable.ListColumns("DT#REG#").Index).Value
                amount = cell.Offset(0, rngTable.ListColumns("IMPORTO NETTO").Index - rngTable.ListColumns("DT#REG#").Index).Value
                price = cell.Offset(0, rngTable.ListColumns("PRICE").Index - rngTable.ListColumns("DT#REG#").Index).Value
    
                ' Correggi i valori della colonna "CAUSALE MOVIM#"
                Select Case transaction
                    Case "VENDITA"
                        transaction = "SALE"
                    Case "CARICO DA FORNI"
                        transaction = "SUPPLY"
                    Case "CAMPIONATURA GR"
                        transaction = "SAMPLES"
                    Case "SCARICO COMPONE"
                        transaction = "USED FOR GIFT SETS"
                    Case "ESISTENZA INIZI"
                        ' Escludi il record se la causale è "ESISTENZA INIZI"
                        GoTo NextCell
                    Case "CARICO  INTERNO"
                        ' Escludi il record se la causale è "CARICO  INTERNO"
                        GoTo NextCell
                    Case "SCARICO INTERNO"
                        ' Escludi il record se la causale è "SCARICO INTERNO"
                        GoTo NextCell
                    Case "CARICO DA PRODU"
                        ' Escludi il record se la causale è "CARICO DA PRODU"
                        GoTo NextCell
                End Select
    
                ' Cerca il valore di customerSupplier nella tabella clienti o fornitori e sostituisci con la ragione sociale
                clientName = ""
                If transaction <> "SUPPLY" Then
                    For Each clientCell In rngClients.ListColumns("CODICE").DataBodyRange
                        If clientCell.Value = customerSupplier Then
                            clientName = clientCell.Offset(0, rngClients.ListColumns("RAGIONE SOCIALE").Index - rngClients.ListColumns("CODICE").Index).Value
                            Exit For
                        End If
                    Next clientCell
                Else
                    For Each supplierCell In rngSuppliers.ListColumns("CODICE").DataBodyRange
                        If supplierCell.Value = customerSupplier Then
                            clientName = supplierCell.Offset(0, rngSuppliers.ListColumns("RAGIONE SOCIALE").Index - rngSuppliers.ListColumns("CODICE").Index).Value
                            Exit For
                        End If
                    Next supplierCell
                End If
    
                skuList.Add Array(sku, regDate, docDate, docNumber, transaction, clientName, pieces, amount, price)
            End If
        End If
NextCell:
    Next cell

    ' Svuota il foglio RESULT
    wsResult.Cells.Clear

    ' Scrivi l'intestazione del report
    wsResult.Range("A1").Value = "SKU"
    wsResult.Range("B1").Value = "DATE"
    wsResult.Range("C1").Value = "DATE DOC"
    wsResult.Range("D1").Value = "N.DOC"
    wsResult.Range("E1").Value = "TRANSACTION"
    wsResult.Range("F1").Value = "CUSTOMER/SUPPLIER"
    wsResult.Range("G1").Value = "PIECES"
    wsResult.Range("H1").Value = "AMOUNT"
    wsResult.Range("I1").Value = "PRICE"
    wsResult.Range("J1").Value = "TOTAL FOC given"
    wsResult.Range("K1").Value = "TOTAL FOC received"
    wsResult.Range("A1:K1").Font.Bold = True
    wsResult.Range("A1:K1").WrapText = True
    wsResult.Rows(1).RowHeight = 29
    wsResult.Rows(1).VerticalAlignment = xlCenter
    wsResult.Columns("J:K").ColumnWidth = 11
    wsResult.Columns("B:C").ColumnWidth = 11.5
    wsResult.Columns("E").ColumnWidth = 20
    wsResult.Range("J1:K1").Interior.Color = RGB(255, 255, 0)

    ' Scrivi i dati nel foglio RESULT con riga vuota al cambio di SKU
    resultRow = 2
    If skuList.Count > 0 Then
        lastSku = ""
        For i = 1 To skuList.Count
            If skuList(i)(0) <> lastSku Then
                If lastSku <> "" Then
                    ' Inserisci una riga vuota con trattini al cambio di SKU
                    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).Value = "-----"
                    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).HorizontalAlignment = xlCenter
                    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Interior.Color = RGB(255, 255, 0)
                    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Font.Bold = True
                    resultRow = resultRow + 1
                End If
                lastSku = skuList(i)(0)
            End If
            wsResult.Cells(resultRow, 1).Value = skuList(i)(0)
            wsResult.Cells(resultRow, 2).Value = skuList(i)(1)
            wsResult.Cells(resultRow, 3).Value = skuList(i)(2)
            wsResult.Cells(resultRow, 4).Value = skuList(i)(3)
            wsResult.Cells(resultRow, 5).Value = skuList(i)(4)
            wsResult.Cells(resultRow, 6).Value = skuList(i)(5)
            wsResult.Cells(resultRow, 7).Value = skuList(i)(6)
            wsResult.Cells(resultRow, 8).Value = skuList(i)(7)
            wsResult.Cells(resultRow, 9).Value = skuList(i)(8)
            resultRow = resultRow + 1
        Next i
    End If
    
        ' Inserisci una riga vuota con trattini come ultima riga
        wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).Value = "-----"
        wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).HorizontalAlignment = xlCenter
        wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Interior.Color = RGB(255, 255, 0)
        wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Font.Bold = True

    ' Layout del foglio RESULT
    With wsResult
        .Activate
        .Rows(1).Font.Bold = True
        .Rows(2).Select
        .Application.ActiveWindow.FreezePanes = True
        .Range("A:K").HorizontalAlignment = xlCenter
        ' Formatta colonne
        .Columns("B:C").NumberFormat = "dd/mm/yyyy"
        .Columns("G").NumberFormat = "#,##0"
        .Columns("G").HorizontalAlignment = xlCenter
        .Columns("H:I").NumberFormat = "_-* #,##0.00 [$€-it-IT]_-;-* #,##0.00 [$€-it-IT]_-;_-* ""-""?? [$€-it-IT]_-;_-@_-"
        .Columns("J:K").NumberFormat = "0"
        .Columns("J:K").Font.Bold = True
        ' Formattazione condizionale per la colonna I
        With .Range("I2", .Range("I100000").End(xlUp)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
            .Interior.Color = RGB(242, 220, 219)
            .Font.Color = RGB(192, 0, 0)
            .Font.Bold = True
        End With
    End With
    
    ' Controlla se i filtri sono già attivi
    If wsResult.AutoFilterMode = False Then
        ' Applica il filtro alle colonne da A a K
        wsResult.Range("A1:K1").AutoFilter
    Else
        ' Se i filtri sono già attivi, li rimuove e li riapplica
        wsResult.AutoFilterMode = False
        wsResult.Range("A1:K1").AutoFilter
    End If

End Sub

Private Sub GeneraReportGiftsets()
    ' Definizione delle variabili
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim wsClients As Worksheet
    Dim wsSuppliers As Worksheet
    Dim rngTable As ListObject, rngClients As ListObject, rngSuppliers As ListObject
    Dim cell As Range
    Dim skuList As Collection
    Dim sku As Variant, regDate As Variant, docDate As Variant, docNumber As Variant
    Dim transaction As Variant, customerSupplier As Variant, pieces As Variant, amount As Variant, price As Double
    Dim i As Long, resultRow As Long
    Dim lastSku As String, clientName As String

    ' Imposta i riferimenti ai fogli
    Set wsData = ThisWorkbook.Worksheets("datiBHPC")
    Set wsResult = ThisWorkbook.Worksheets("GIFTSET")
    Set wsClients = ThisWorkbook.Worksheets("clienti")
    Set wsSuppliers = ThisWorkbook.Worksheets("fornitori")
    Set rngTable = wsData.ListObjects("movimentiProfumi")
    Set rngClients = wsClients.ListObjects("clientiBHPC")
    Set rngSuppliers = wsSuppliers.ListObjects("fornitoriBHPC")

    ' Inizializza la collezione per gli SKU
    Set skuList = New Collection

    ' Scansiona la tabella e registra i dati con date comprese tra startDate e endDate
    For Each cell In rngTable.ListColumns("DT#REG#").DataBodyRange
        If cell.Value >= startDate And cell.Value <= endDate Then
            sku = cell.Offset(0, rngTable.ListColumns("SKU CODE").Index - rngTable.ListColumns("DT#REG#").Index).Value
            If sku Like "GS????" Then
                regDate = cell.Value
                docDate = cell.Offset(0, rngTable.ListColumns("DT#DOC#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                docNumber = cell.Offset(0, rngTable.ListColumns("N#DOC#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                transaction = cell.Offset(0, rngTable.ListColumns("CAUSALE MOVIM#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                customerSupplier = cell.Offset(0, rngTable.ListColumns("CLI/FOR NUMBER").Index - rngTable.ListColumns("DT#REG#").Index).Value
                pieces = cell.Offset(0, rngTable.ListColumns("QUANTITA'").Index - rngTable.ListColumns("DT#REG#").Index).Value
                amount = cell.Offset(0, rngTable.ListColumns("IMPORTO NETTO").Index - rngTable.ListColumns("DT#REG#").Index).Value
                price = cell.Offset(0, rngTable.ListColumns("PRICE").Index - rngTable.ListColumns("DT#REG#").Index).Value
    
                ' Correggi i valori della colonna "CAUSALE MOVIM#"
                Select Case transaction
                    Case "VENDITA"
                        transaction = "SALE"
                    Case "CARICO DA FORNI"
                        transaction = "SUPPLY"
                    Case "CAMPIONATURA GR"
                        transaction = "SAMPLES"
                    Case "SCARICO COMPONE"
                        transaction = "USED FOR GIFT SETS"
                    Case "ESISTENZA INIZI"
                        ' Escludi il record se la causale è "ESISTENZA INIZI"
                        GoTo NextCell
                    Case "CARICO  INTERNO"
                        ' Escludi il record se la causale è "CARICO  INTERNO"
                        GoTo NextCell
                    Case "SCARICO INTERNO"
                        ' Escludi il record se la causale è "SCARICO INTERNO"
                        GoTo NextCell
                    Case "CARICO DA PRODU"
                        ' Escludi il record se la causale è "CARICO DA PRODU"
                        GoTo NextCell
                End Select
    
                ' Cerca il valore di customerSupplier nella tabella clienti o fornitori e sostituisci con la ragione sociale
                clientName = ""
                If transaction <> "SUPPLY" Then
                    For Each clientCell In rngClients.ListColumns("CODICE").DataBodyRange
                        If clientCell.Value = customerSupplier Then
                            clientName = clientCell.Offset(0, rngClients.ListColumns("RAGIONE SOCIALE").Index - rngClients.ListColumns("CODICE").Index).Value
                            Exit For
                        End If
                    Next clientCell
                Else
                    For Each supplierCell In rngSuppliers.ListColumns("CODICE").DataBodyRange
                        If supplierCell.Value = customerSupplier Then
                            clientName = supplierCell.Offset(0, rngSuppliers.ListColumns("RAGIONE SOCIALE").Index - rngSuppliers.ListColumns("CODICE").Index).Value
                            Exit For
                        End If
                    Next supplierCell
                End If
    
                skuList.Add Array(sku, regDate, docDate, docNumber, transaction, clientName, pieces, amount, price)
            End If
        End If
NextCell:
    Next cell

    ' Svuota il foglio RESULT
    wsResult.Cells.Clear

    ' Scrivi l'intestazione del report
    wsResult.Range("A1").Value = "SKU"
    wsResult.Range("B1").Value = "DATE"
    wsResult.Range("C1").Value = "DATE DOC"
    wsResult.Range("D1").Value = "N.DOC"
    wsResult.Range("E1").Value = "TRANSACTION"
    wsResult.Range("F1").Value = "CUSTOMER/SUPPLIER"
    wsResult.Range("G1").Value = "PIECES"
    wsResult.Range("H1").Value = "AMOUNT"
    wsResult.Range("I1").Value = "PRICE"
    wsResult.Range("J1").Value = "TOTAL FOC given"
    wsResult.Range("K1").Value = "TOTAL FOC received"
    wsResult.Range("A1:K1").Font.Bold = True
    wsResult.Range("A1:K1").WrapText = True
    wsResult.Rows(1).RowHeight = 29
    wsResult.Rows(1).VerticalAlignment = xlCenter
    wsResult.Columns("J:K").ColumnWidth = 11
    wsResult.Columns("B:C").ColumnWidth = 11.5
    wsResult.Columns("E").ColumnWidth = 20
    wsResult.Range("J1:K1").Interior.Color = RGB(255, 255, 0)

    ' Scrivi i dati nel foglio RESULT con riga vuota al cambio di SKU
    resultRow = 2
    If skuList.Count > 0 Then
        lastSku = ""
        For i = 1 To skuList.Count
            If skuList(i)(0) <> lastSku Then
                If lastSku <> "" Then
                    ' Inserisci una riga vuota con trattini al cambio di SKU
                    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).Value = "-----"
                    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).HorizontalAlignment = xlCenter
                    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Interior.Color = RGB(255, 255, 0)
                    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Font.Bold = True
                    resultRow = resultRow + 1
                End If
                lastSku = skuList(i)(0)
            End If
            wsResult.Cells(resultRow, 1).Value = skuList(i)(0)
            wsResult.Cells(resultRow, 2).Value = skuList(i)(1)
            wsResult.Cells(resultRow, 3).Value = skuList(i)(2)
            wsResult.Cells(resultRow, 4).Value = skuList(i)(3)
            wsResult.Cells(resultRow, 5).Value = skuList(i)(4)
            wsResult.Cells(resultRow, 6).Value = skuList(i)(5)
            wsResult.Cells(resultRow, 7).Value = skuList(i)(6)
            wsResult.Cells(resultRow, 8).Value = skuList(i)(7)
            wsResult.Cells(resultRow, 9).Value = skuList(i)(8)
            resultRow = resultRow + 1
        Next i
    End If
    
    ' Inserisci una riga vuota con trattini come ultima riga
    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).Value = "-----"
    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).HorizontalAlignment = xlCenter
    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Interior.Color = RGB(255, 255, 0)
    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Font.Bold = True

    ' Layout del foglio RESULT
    With wsResult
        .Activate
        .Rows(1).Font.Bold = True
        .Rows(2).Select
        .Application.ActiveWindow.FreezePanes = True
        .Range("A:K").HorizontalAlignment = xlCenter
        ' Formatta colonne
        .Columns("B:C").NumberFormat = "dd/mm/yyyy"
        .Columns("G").NumberFormat = "#,##0"
        .Columns("G").HorizontalAlignment = xlCenter
        .Columns("H:I").NumberFormat = "_-* #,##0.00 [$€-it-IT]_-;-* #,##0.00 [$€-it-IT]_-;_-* ""-""?? [$€-it-IT]_-;_-@_-"
        .Columns("J:K").NumberFormat = "0"
        .Columns("J:K").Font.Bold = True
        ' Formattazione condizionale per la colonna I
        With .Range("I2", .Range("I100000").End(xlUp)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
            .Interior.Color = RGB(242, 220, 219)
            .Font.Color = RGB(192, 0, 0)
            .Font.Bold = True
        End With
    End With
    
    ' Controlla se i filtri sono già attivi
    If wsResult.AutoFilterMode = False Then
        ' Applica il filtro alle colonne da A a K
        wsResult.Range("A1:K1").AutoFilter
    Else
        ' Se i filtri sono già attivi, li rimuove e li riapplica
        wsResult.AutoFilterMode = False
        wsResult.Range("A1:K1").AutoFilter
    End If

End Sub

Private Sub GeneraReportTester()
    ' Definizione delle variabili
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim wsClients As Worksheet
    Dim wsSuppliers As Worksheet
    Dim rngTable As ListObject, rngClients As ListObject, rngSuppliers As ListObject
    Dim cell As Range
    Dim skuList As Collection
    Dim sku As Variant, regDate As Variant, docDate As Variant, docNumber As Variant
    Dim transaction As Variant, customerSupplier As Variant, pieces As Variant, amount As Variant, price As Double
    Dim i As Long, resultRow As Long
    Dim lastSku As String, clientName As String

    ' Imposta i riferimenti ai fogli
    Set wsData = ThisWorkbook.Worksheets("datiBHPC")
    Set wsResult = ThisWorkbook.Worksheets("TESTER")
    Set wsClients = ThisWorkbook.Worksheets("clienti")
    Set wsSuppliers = ThisWorkbook.Worksheets("fornitori")
    Set rngTable = wsData.ListObjects("movimentiProfumi")
    Set rngClients = wsClients.ListObjects("clientiBHPC")
    Set rngSuppliers = wsSuppliers.ListObjects("fornitoriBHPC")

    ' Inizializza la collezione per gli SKU
    Set skuList = New Collection

    ' Scansiona la tabella e registra i dati con date comprese tra startDate e endDate
    For Each cell In rngTable.ListColumns("DT#REG#").DataBodyRange
        If cell.Value >= startDate And cell.Value <= endDate Then
            sku = cell.Offset(0, rngTable.ListColumns("SKU CODE").Index - rngTable.ListColumns("DT#REG#").Index).Value
            If sku Like "PT####" Then
                regDate = cell.Value
                docDate = cell.Offset(0, rngTable.ListColumns("DT#DOC#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                docNumber = cell.Offset(0, rngTable.ListColumns("N#DOC#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                transaction = cell.Offset(0, rngTable.ListColumns("CAUSALE MOVIM#").Index - rngTable.ListColumns("DT#REG#").Index).Value
                customerSupplier = cell.Offset(0, rngTable.ListColumns("CLI/FOR NUMBER").Index - rngTable.ListColumns("DT#REG#").Index).Value
                pieces = cell.Offset(0, rngTable.ListColumns("QUANTITA'").Index - rngTable.ListColumns("DT#REG#").Index).Value
                amount = cell.Offset(0, rngTable.ListColumns("IMPORTO NETTO").Index - rngTable.ListColumns("DT#REG#").Index).Value
                price = cell.Offset(0, rngTable.ListColumns("PRICE").Index - rngTable.ListColumns("DT#REG#").Index).Value
    
                ' Correggi i valori della colonna "CAUSALE MOVIM#"
                Select Case transaction
                    Case "VENDITA"
                        transaction = "SALE"
                    Case "CARICO DA FORNI"
                        transaction = "SUPPLY"
                    Case "CAMPIONATURA GR"
                        transaction = "SAMPLES"
                    Case "SCARICO COMPONE"
                        transaction = "USED FOR GIFT SETS"
                    Case "ESISTENZA INIZI"
                        ' Escludi il record se la causale è "ESISTENZA INIZI"
                        GoTo NextCell
                    Case "CARICO  INTERNO"
                        ' Escludi il record se la causale è "CARICO  INTERNO"
                        GoTo NextCell
                    Case "SCARICO INTERNO"
                        ' Escludi il record se la causale è "SCARICO INTERNO"
                        GoTo NextCell
                    Case "CARICO DA PRODU"
                        ' Escludi il record se la causale è "CARICO DA PRODU"
                        GoTo NextCell
                End Select
    
                ' Cerca il valore di customerSupplier nella tabella clienti o fornitori e sostituisci con la ragione sociale
                clientName = ""
                If transaction <> "SUPPLY" Then
                    For Each clientCell In rngClients.ListColumns("CODICE").DataBodyRange
                        If clientCell.Value = customerSupplier Then
                            clientName = clientCell.Offset(0, rngClients.ListColumns("RAGIONE SOCIALE").Index - rngClients.ListColumns("CODICE").Index).Value
                            Exit For
                        End If
                    Next clientCell
                Else
                    For Each supplierCell In rngSuppliers.ListColumns("CODICE").DataBodyRange
                        If supplierCell.Value = customerSupplier Then
                            clientName = supplierCell.Offset(0, rngSuppliers.ListColumns("RAGIONE SOCIALE").Index - rngSuppliers.ListColumns("CODICE").Index).Value
                            Exit For
                        End If
                    Next supplierCell
                End If
    
                skuList.Add Array(sku, regDate, docDate, docNumber, transaction, clientName, pieces, amount, price)
            End If
        End If
NextCell:
    Next cell

    ' Svuota il foglio RESULT
    wsResult.Cells.Clear

    ' Scrivi l'intestazione del report
    wsResult.Range("A1").Value = "SKU"
    wsResult.Range("B1").Value = "DATE"
    wsResult.Range("C1").Value = "DATE DOC"
    wsResult.Range("D1").Value = "N.DOC"
    wsResult.Range("E1").Value = "TRANSACTION"
    wsResult.Range("F1").Value = "CUSTOMER/SUPPLIER"
    wsResult.Range("G1").Value = "PIECES"
    wsResult.Range("H1").Value = "AMOUNT"
    wsResult.Range("I1").Value = "PRICE"
    wsResult.Range("J1").Value = "TOTAL FOC given"
    wsResult.Range("K1").Value = "TOTAL FOC received"
    wsResult.Range("A1:K1").Font.Bold = True
    wsResult.Range("A1:K1").WrapText = True
    wsResult.Rows(1).RowHeight = 29
    wsResult.Rows(1).VerticalAlignment = xlCenter
    wsResult.Columns("J:K").ColumnWidth = 11
    wsResult.Columns("B:C").ColumnWidth = 11.5
    wsResult.Columns("E").ColumnWidth = 20
    wsResult.Range("J1:K1").Interior.Color = RGB(255, 255, 0)

    ' Scrivi i dati nel foglio RESULT con riga vuota al cambio di SKU
    resultRow = 2
    If skuList.Count > 0 Then
        lastSku = ""
        For i = 1 To skuList.Count
            If skuList(i)(0) <> lastSku Then
                If lastSku <> "" Then
                    ' Inserisci una riga vuota con trattini al cambio di SKU
                    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).Value = "-----"
                    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).HorizontalAlignment = xlCenter
                    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Interior.Color = RGB(255, 255, 0)
                    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Font.Bold = True
                    resultRow = resultRow + 1
                End If
                lastSku = skuList(i)(0)
            End If
            wsResult.Cells(resultRow, 1).Value = skuList(i)(0)
            wsResult.Cells(resultRow, 2).Value = skuList(i)(1)
            wsResult.Cells(resultRow, 3).Value = skuList(i)(2)
            wsResult.Cells(resultRow, 4).Value = skuList(i)(3)
            wsResult.Cells(resultRow, 5).Value = skuList(i)(4)
            wsResult.Cells(resultRow, 6).Value = skuList(i)(5)
            wsResult.Cells(resultRow, 7).Value = skuList(i)(6)
            wsResult.Cells(resultRow, 8).Value = skuList(i)(7)
            wsResult.Cells(resultRow, 9).Value = skuList(i)(8)
            resultRow = resultRow + 1
        Next i
    End If
    
    ' Inserisci una riga vuota con trattini come ultima riga
    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).Value = "-----"
    wsResult.Range(wsResult.Cells(resultRow, 1), wsResult.Cells(resultRow, 9)).HorizontalAlignment = xlCenter
    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Interior.Color = RGB(255, 255, 0)
    wsResult.Range(wsResult.Cells(resultRow, 10), wsResult.Cells(resultRow, 11)).Font.Bold = True

    ' Layout del foglio RESULT
    With wsResult
        .Activate
        .Rows(1).Font.Bold = True
        .Rows(2).Select
        .Application.ActiveWindow.FreezePanes = True
        .Range("A:K").HorizontalAlignment = xlCenter
        ' Formatta colonne
        .Columns("B:C").NumberFormat = "dd/mm/yyyy"
        .Columns("G").NumberFormat = "#,##0"
        .Columns("G").HorizontalAlignment = xlCenter
        .Columns("H:I").NumberFormat = "_-* #,##0.00 [$€-it-IT]_-;-* #,##0.00 [$€-it-IT]_-;_-* ""-""?? [$€-it-IT]_-;_-@_-"
        .Columns("J:K").NumberFormat = "0"
        .Columns("J:K").Font.Bold = True
        ' Formattazione condizionale per la colonna I
        With .Range("I2", .Range("I100000").End(xlUp)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
            .Interior.Color = RGB(242, 220, 219)
            .Font.Color = RGB(192, 0, 0)
            .Font.Bold = True
        End With
    End With
    
    ' Controlla se i filtri sono già attivi
    If wsResult.AutoFilterMode = False Then
        ' Applica il filtro alle colonne da A a K
        wsResult.Range("A1:K1").AutoFilter
    Else
        ' Se i filtri sono già attivi, li rimuove e li riapplica
        wsResult.AutoFilterMode = False
        wsResult.Range("A1:K1").AutoFilter
    End If

End Sub

Sub QuickSort(arr() As Variant, low As Long, high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant

    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high

        Do While i <= j
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop

            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop

        Call QuickSort(arr, low, j)
        Call QuickSort(arr, i, high)
    End If
End Sub

Sub stampaReport()
    Dim wb As Workbook
    Dim wsPerfume As Worksheet
    Dim wsGiftset As Worksheet
    Dim wsBodymist As Worksheet
    Dim folderPath As String
    Dim savePath As String
    Dim fileDialog As fileDialog
    
    ' Imposta i riferimenti ai fogli
    Set wsPerfume = ActiveWorkbook.Sheets("PERFUME")
    Set wsGiftset = ActiveWorkbook.Sheets("GIFTSET")
    Set wsBodymist = ActiveWorkbook.Sheets("TESTER")
    
    ' Apri la finestra di dialogo per selezionare il percorso di salvataggio
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    ' Se l'utente seleziona una cartella
    If fileDialog.Show = -1 Then
        folderPath = fileDialog.SelectedItems(1)
    Else
        MsgBox "Nessuna cartella selezionata. Report non stampati. E' comunque possibili consultarli dal Workbook corrente."
        Exit Sub
    End If
    
    ' Crea una nuova cartella di lavoro per salvare i fogli
    Set wb = Workbooks.Add
    
    ' Copia i fogli nella nuova cartella di lavoro e assicurati che siano nell'ordine desiderato
    wsBodymist.Copy before:=wb.Sheets(1)
    wsGiftset.Copy before:=wb.Sheets(1)
    wsPerfume.Copy before:=wb.Sheets(1)
    
    ' Elimina il foglio vuoto creato di default
    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name <> "PERFUME" And ws.Name <> "GIFTSET" And ws.Name <> "TESTER" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Imposta il percorso e il nome del file
    savePath = folderPath & "\FOC Sent & Received " & Format(startDate, "DDMMYY") & " " & Format(endDate, "DDMMYY") & ".xlsx"
    
    ' Salva il file
    On Error GoTo errorHandler
    wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    MsgBox "Il report è stato salvato con successo in " & savePath
    Exit Sub
    
errorHandler:
    MsgBox "Errore inatteso. Assicurarsi di non stare tentando di sovrascrivere un file aperto"
    Exit Sub
    
End Sub



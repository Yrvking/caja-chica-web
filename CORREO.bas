Attribute VB_Name = "CORREO"
' Script modularizado para enviar documentos de pago MARKETING
' Modo: Abre Outlook con todo listo pero NO envía automáticamente

Option Explicit

Sub EnviarDocumentosPago()
    Dim ws As Worksheet, tbl As ListObject
    Dim rng As Range, cell As Range, row As ListRow
    Dim OutApp As Object, OutMail As Object
    Dim folderRaiz As String, resumenPath As String, zipPath As String
    Dim mailBody As String
    Dim listaFilas As Collection

    ' 1. Selección de rango válido en columna F. ENVÍO
    Set ws = ThisWorkbook.Sheets("BASE DE DATOS GASTOS")
    Set tbl = ws.ListObjects("Tabla3")
    Set rng = Intersect(Selection, tbl.ListColumns("F. ENVÍO").Range)
    If rng Is Nothing Then MsgBox "Selecciona celdas en la columna F. ENVÍO": Exit Sub

    ' 2. Recolectar filas seleccionadas
    Set listaFilas = New Collection
    For Each cell In rng.Cells
        If Not cell.EntireRow.Hidden Then
            Set row = tbl.ListRows(cell.row - tbl.Range.row)
            listaFilas.Add row
        End If
    Next cell

    If listaFilas.Count = 0 Then MsgBox "No hay filas seleccionadas.": Exit Sub
    If Not ValidarFilasPagoLote(listaFilas, tbl) Then Exit Sub

    ' 3. Seleccionar carpeta origen de documentos
    folderRaiz = SeleccionarCarpetaZIP()
    If folderRaiz = "" Then Exit Sub

    ' 4. Generar resumen Excel y ZIP
    resumenPath = GenerarResumenExcel(listaFilas, tbl)
    zipPath = GenerarZIPCarpetaCompleta(folderRaiz)

    ' 5. Generar cuerpo de correo HTML
    mailBody = GenerarCuerpoHTML(listaFilas, tbl, folderRaiz)

    ' 6. Crear correo Outlook sin enviar
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = "durbina@padovasac.com;pmendoza@constructorapadova.pe;mcosta@padovasac.com"
        .CC = "hlopez@padovasac.com;jguevara@padovasac.com;sdenegri@constructorapadova.pe;cegoavil@constructorapadova.pe;emescate@constructorapadova.pe"
        .Subject = "PAGO VÍA TRANSFERENCIA: Pagos varios OBRA - " & Format(Date, "dd/mm/yyyy")
        .HTMLBody = mailBody
        If Dir(resumenPath) <> "" Then .Attachments.Add resumenPath
        If Dir(zipPath) <> "" Then .Attachments.Add zipPath
        .Display ' Solo mostrar, NO enviar
    End With

    ' Eliminar archivos temporales
    On Error Resume Next
    Kill resumenPath
    Kill zipPath
    On Error GoTo 0
End Sub

Function ValidarFilasPagoLote(lista As Collection, tbl As ListObject) As Boolean
    Dim camposObligatorios As Variant, colName As Variant
    Dim row As ListRow, colIndex As Integer
    Dim errores As String, faltante As String
    Dim filaIdx As Long
    errores = ""

    camposObligatorios = Array("SERIE", "N°", "RUC", "F. EMISIÓN", "PROYECTO", "PC", "FACTURADO A", "MONEDA", "IMPORTE")

    For Each row In lista
        filaIdx = row.Index
        faltante = ""

        For Each colName In camposObligatorios
            On Error Resume Next
            colIndex = tbl.ListColumns(colName).Index
            On Error GoTo 0

            If colIndex = 0 Then
                errores = errores & "?? Columna '" & colName & "' no existe en la tabla." & vbCrLf
            ElseIf Trim(row.Range(1, colIndex).Value) = "" Then
                faltante = faltante & colName & ", "
            End If
        Next colName

        If Len(faltante) > 0 Then
            errores = errores & "Fila " & filaIdx + 1 & ": faltan [" & Left(faltante, Len(faltante) - 2) & "]" & vbCrLf
        End If
    Next row

    If errores <> "" Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("Se encontraron campos incompletos:" & vbCrLf & vbCrLf & errores & vbCrLf & _
                           "¿Deseas continuar de todas formas?", vbYesNo + vbExclamation, "Validación de datos")
        ValidarFilasPagoLote = (respuesta = vbYes)
    Else
        ValidarFilasPagoLote = True
    End If
End Function

Function SeleccionarCarpetaZIP() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = "Selecciona la carpeta donde están los documentos a pagar"
        If .Show <> -1 Then
            SeleccionarCarpetaZIP = ""
        Else
            SeleccionarCarpetaZIP = .SelectedItems(1)
            If Right(SeleccionarCarpetaZIP, 1) <> "\" Then SeleccionarCarpetaZIP = SeleccionarCarpetaZIP & "\"
        End If
    End With
End Function

Function GenerarResumenExcel(lista As Collection, tbl As ListObject) As String
    Dim wb As Workbook, ws As Worksheet, colName As Variant, row As ListRow
    Dim columnNames As Variant, i As Integer, r As Integer, c As Integer, lastCol As Integer, lastRow As Integer
    columnNames = Array("F. EMISIÓN", "TIPO", "SERIE", "N°", "RUC", "RAZON SOCIAL", "PROYECTO", "PC", _
                     "DESCRIPCION", "FACTURADO A", "MONEDA", "SUBTOTAL", "IGV", "IMPORTE", _
                     "TIPO DET", "PORCENTAJE", "DETRACCION", "BANCO", "CC", "CCI")

    Set wb = Workbooks.Add: Set ws = wb.Sheets(1)
    r = 1: c = 1
    For Each colName In columnNames
        ws.Cells(r, c).Value = colName: c = c + 1
    Next colName

    For Each row In lista
        r = r + 1: c = 1
        For Each colName In columnNames
            ws.Cells(r, c).Value = row.Range(1, tbl.ListColumns(colName).Index).Value
            c = c + 1
        Next colName
    Next row

    ' Crear formato de tabla
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = UBound(columnNames) + 1
    ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes).Name = "ResumenTabla"
    ws.ListObjects("ResumenTabla").TableStyle = "TableStyleMedium9"

    Dim ruta As String
    ruta = Environ("Temp") & "\ResumenPago_" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
    wb.SaveAs ruta, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    GenerarResumenExcel = ruta
End Function

' NUEVA FUNCIÓN: Comprime toda la carpeta seleccionada
Function GenerarZIPCarpetaCompleta(carpetaRaiz As String) As String
    Dim zipName As String
    zipName = Environ("Temp") & "\DocumentosPago_" & Format(Now, "yyyymmddhhmmss") & ".zip"
    Call CreateZipFileDirecto(carpetaRaiz, zipName)
    If Dir(zipName) <> "" Then
        GenerarZIPCarpetaCompleta = zipName
    Else
        GenerarZIPCarpetaCompleta = ""
    End If
End Function

Sub CreateZipFileDirecto(carpetaOrigen As String, rutaZIP As String)
    Dim ps As String, shell As Object, fso As Object, timeout As Double
    ' Quitar barra invertida final si existe
    If Right(carpetaOrigen, 1) = "\" Then carpetaOrigen = Left(carpetaOrigen, Len(carpetaOrigen) - 1)
    ps = "Compress-Archive -Path '" & carpetaOrigen & "\*' -DestinationPath '" & rutaZIP & "' -Force"
    Set shell = CreateObject("WScript.Shell")
    shell.Run "powershell -command """ & ps & """", 0, True

    ' Esperar a que el archivo ZIP exista (máx. 5 segundos)
    Set fso = CreateObject("Scripting.FileSystemObject")
    timeout = Timer + 5
    Do While Not fso.FileExists(rutaZIP)
        DoEvents
        If Timer > timeout Then Exit Do
    Loop
End Sub

Function GenerarCuerpoHTML(lista As Collection, tbl As ListObject, carpetaRaiz As String) As String
    Dim row As ListRow, colName As Variant, colSet As Variant, colCuentaSet As Variant
    Dim html As String, fila As String, i As Integer

    colSet = Array("F. EMISIÓN", "TIPO", "SERIE", "N°", "RUC", "RAZON SOCIAL", "PROYECTO", "PC", _
                  "DESCRIPCION", "FACTURADO A", "MONEDA", "SUBTOTAL", "IGV", "IMPORTE", "TIPO DET")
    colCuentaSet = Array("RAZON SOCIAL", "BANCO", "MONEDA", "CC", "CCI", "C. DETRACCION")

    html = "<html><body style='font-family:""Arial Nova Cond"", Arial, sans-serif; font-size:10pt;'>"
    html = html & "<p>Buen día David/Pedro,</p>"
    html = html & "<p>Adjunto los pagos pertenecientes al día " & Format(Date, "dd/mm/yyyy") & " del área de obra.</p>"
    'html = html & "<p>Detallamos enlace donde se encuentran los documentos para pago.</p>"
    'html = html & "<p><a href='https://inversionespadova-my.sharepoint.com/:f:/g/personal/sistemas_padovasac_com/Eq6KxHlrMgBDj5F21QHXEFMB6hdTKXrW1C0MHUUqaXilOA?e=IOCbaG'>Enlace a Documentos 2025</a>.</p>"
    html = html & "<style>" & _
        "body, table, td, th {font-family:'Arial Nova Cond', Arial, sans-serif; font-size:8pt;}" & _
        "table {border-collapse:collapse;width:100%;}" & _
        "th {background-color:yellow;}" & _
        "td {padding:4px; border:1px solid #ccc;}" & _
        "</style>"

    ' Primera tabla
    html = html & "<table><tr>"
    For Each colName In colSet: html = html & "<th>" & colName & "</th>": Next colName
    html = html & "</tr>"

    For Each row In lista
        fila = "<tr>"
        For Each colName In colSet
            fila = fila & "<td>" & row.Range(1, tbl.ListColumns(colName).Index).Value & "</td>"
        Next colName
        fila = fila & "</tr>"
        html = html & fila
    Next row
    html = html & "</table><br><br>"

    ' Segunda tabla - Cuentas
    html = html & "<table border='1'><tr>"
    For Each colName In colCuentaSet
        html = html & "<th style='background-color: #FFCCCC;'>" & colName & "</th>"
    Next colName
    html = html & "</tr>"

    For Each row In lista
        fila = "<tr>"
        For Each colName In colCuentaSet
            fila = fila & "<td>" & row.Range(1, tbl.ListColumns(colName).Index).Value & "</td>"
        Next colName
        fila = fila & "</tr>"
        html = html & fila
    Next row
    html = html & "</table><br><br>"
    ' Mensaje de cierre después de la segunda tabla
    html = html & "<p>Muchas gracias<br>Saludos.<br>Equipo de obra</p>"
    html = html & "</body></html>"
    GenerarCuerpoHTML = html
End Function



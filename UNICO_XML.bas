Attribute VB_Name = "UNICO_XML"
Sub ProcesarVariosXMLs()
    Dim fDialog As FileDialog
    Dim i As Long
    Dim archivo As Variant
    Dim filaEliminar As Range
    Dim filasAEliminar As Range
    Dim respuesta As VbMsgBoxResult
    Const clave As String = "PRUEBA2025YRV"
    Set ws = ActiveSheet
    ' Desproteger si está protegida
    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect Password:=clave
        If ws.ProtectContents Then
            MsgBox "No se pudo desproteger la hoja. Verifica la contraseña.", vbCritical
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' Seleccionar múltiples archivos XML
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Selecciona uno o más archivos XML"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Archivos XML", "*.xml"
        If .Show <> -1 Then Exit Sub
        If .SelectedItems.Count = 0 Then Exit Sub
    End With

    ' Procesar cada archivo seleccionado
    For Each archivo In fDialog.SelectedItems
        Call ProcesarArchivoXML(CStr(archivo))
    Next archivo

    MsgBox "Todos los archivos seleccionados han sido procesados.", vbInformation
        ' ?? Establecer permisos antes de proteger nuevamente
    With ws.UsedRange
        .Locked = True ' Bloquear todo primero
    End With

    ' ?? Desbloquear solo las columnas autorizadas
    ws.Range("A:L").Locked = False
    ws.Range("N:Q").Locked = False
    ws.Range("S:S").Locked = False
    ws.Range("W:X").Locked = False
    ws.Range("Z:AA").Locked = False

    ' ?? Proteger la hoja nuevamente
    ws.Protect Password:=clave, UserInterfaceOnly:=True
End Sub

Sub ProcesarArchivoXML(ByVal archivoXML As String)
    Dim xmlDoc As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim proyectos As Object
    Dim splitID() As String
    Dim invoiceID As String
    Dim lastRow As Long

    ' ---------- Referencia a hoja y tabla ----------
    Set ws = ThisWorkbook.Sheets("BASE DE DATOS GASTOS")
    Set tbl = ws.ListObjects("Tabla3")

    ' ---------- Diccionario de proyectos ----------
    Set proyectos = CreateObject("Scripting.Dictionary")
    proyectos.Add "INVERSIONES ANZIO SAC", "VILLAVICENCIO"
    proyectos.Add "INVERSIONES BORGO SAC", "MONTE UMBROSO 122"
    proyectos.Add "INVERSIONES FIDENZA SAC", "SANTA BEATRIZ"
    proyectos.Add "INVERSIONES FIDENZA SAC.", "SANTA BEATRIZ"
    proyectos.Add "INVERSIONES FIDENZA S.A.C.", "SANTA BEATRIZ"
    proyectos.Add "INVERSIONES FORLI SAC", "MARISCAL SUCRE 296"
    proyectos.Add "INVERSIONES GODIA SAC", "LUIS PASTEUR 1228"
    proyectos.Add "INVERSIONES MAJANO SAC", "LUIS PASTEUR 1250"
    proyectos.Add "INVERSIONES MELCEN SAC", "SUNNY LA MOLINA"
    proyectos.Add "INVERSIONES MELCEN SAC.", "SUNNY LA MOLINA"
    proyectos.Add "INVERSIONES MELCEN S.A.C.", "SUNNY LA MOLINA"
    proyectos.Add "INVERSIONES NOVARA SAC", "NOVARA CONDOMINIO"
    proyectos.Add "INVERSIONES PADOVA SAC", "LOMAS DE CARABAYLLO 5"
    proyectos.Add "INVERSIONES PADOVA SAC.", "LOMAS DE CARABAYLLO 5"
    proyectos.Add "INVERSIONES  PADOVA SAC", "LOMAS DE CARABAYLLO 5"
    proyectos.Add "INVERSIONES PADOVA S.A.C.", "LOMAS DE CARABAYLLO 5"
    proyectos.Add "INVERSIONES RAVENA SAC", "TUDELA & VARELA 445"
    proyectos.Add "INVERSIONES SALETTI SAC", "HELIO"
    proyectos.Add "INVERSIONES SAURIS SAC", "LITORAL 900"
    proyectos.Add "INVERSIONES SAURIS SAC.", "LITORAL 900"
    proyectos.Add "INVERSIONES SAURIS S.A.C.", "LITORAL 900"
    proyectos.Add "INVERSIONES VASTO SAC", "SAN MARTIN 230"
    proyectos.Add "CONSTRUCTORA PADOVA SAC", "CONSTRUCTORA"
    proyectos.Add "CONSTRUCTORA  PADOVA SAC", "CONSTRUCTORA"
    proyectos.Add "CONSTRUCTORA PADOVA S.A.C.", "CONSTRUCTORA"

    ' ---------- Cargar XML ----------
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.Load archivoXML
    xmlDoc.SetProperty "SelectionNamespaces", _
        "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' " & _
        "xmlns:cac='urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'"

    On Error GoTo Salir

    invoiceID = ""
    On Error Resume Next
    invoiceID = xmlDoc.SelectSingleNode("/*[local-name()='Invoice']/*[local-name()='ID']").Text
    On Error GoTo 0
    If invoiceID = "" Or InStr(invoiceID, "-") = 0 Then GoTo Salir
    splitID = Split(invoiceID, "-")

    Dim descripcionValue As String, facturadoValue As String
    Dim rucValue As String, subtotalValue As String
    Dim importeValue As String, fechaValue As String
    Dim monedaValue As String, bienValue As String
    Dim razonSocialValue As String, porcentajeDetraccion As String
    Dim proyectoValue As String
    Dim nodeMoneda As Object

    descripcionValue = ""
    On Error Resume Next
    descripcionValue = xmlDoc.SelectSingleNode("//cac:InvoiceLine/cac:Item/cbc:Description").Text
    If descripcionValue = "" Then descripcionValue = xmlDoc.SelectSingleNode("//cac:InvoiceLine/cbc:Note").Text
    On Error GoTo 0

    facturadoValue = ""
    On Error Resume Next
    facturadoValue = xmlDoc.SelectSingleNode("//cac:AccountingCustomerParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName").Text
    If facturadoValue = "" Then facturadoValue = xmlDoc.SelectSingleNode("//cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name").Text
    facturadoValue = Replace(facturadoValue, ".", "")
    On Error GoTo 0

    rucValue = ""
    On Error Resume Next
    rucValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID").Text
    If rucValue = "" Then rucValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cbc:CustomerAssignedAccountID").Text
    On Error GoTo 0

    subtotalValue = ""
    On Error Resume Next
    subtotalValue = xmlDoc.SelectSingleNode("//cac:LegalMonetaryTotal/cbc:LineExtensionAmount").Text
    On Error GoTo 0

    importeValue = ""
    On Error Resume Next
    importeValue = xmlDoc.SelectSingleNode("//cac:LegalMonetaryTotal/cbc:PayableAmount").Text
    On Error GoTo 0

    fechaValue = ""
    On Error Resume Next
    fechaValue = xmlDoc.SelectSingleNode("//cbc:IssueDate").Text
    On Error GoTo 0

    monedaValue = ""
    On Error Resume Next
    Set nodeMoneda = xmlDoc.SelectSingleNode("//cac:LegalMonetaryTotal/cbc:PayableAmount")
    If Not nodeMoneda Is Nothing And Not nodeMoneda.Attributes Is Nothing Then
        monedaValue = nodeMoneda.Attributes.getNamedItem("currencyID").Text
    End If
    If monedaValue = "" Then
        Set nodeMoneda = xmlDoc.SelectSingleNode("//*[local-name()='DocumentCurrencyCode']")
        If Not nodeMoneda Is Nothing Then monedaValue = nodeMoneda.Text
    End If
    If monedaValue = "PEN" Then
        monedaValue = "SOL"
    ElseIf monedaValue = "USD" Then
        monedaValue = "DOLARES"
    End If
    On Error GoTo 0

    bienValue = ""
    porcentajeDetraccion = ""
    On Error Resume Next
    bienValue = xmlDoc.SelectSingleNode("//cac:PaymentTerms/cbc:PaymentMeansID").Text
    Select Case bienValue
        Case "019", "020", "022", "025", "027", "030", "037"
        Case Else: bienValue = ""
    End Select
    If bienValue <> "" Then
        Dim nodePercentDetraccion As Object
        Set nodePercentDetraccion = xmlDoc.SelectSingleNode("//cac:PaymentTerms/cbc:PaymentPercent")
        If Not nodePercentDetraccion Is Nothing Then
            porcentajeDetraccion = Trim(nodePercentDetraccion.Text)
        Else
            Set nodePercentDetraccion = xmlDoc.SelectSingleNode("//cac:InvoiceLine//cac:TaxSubtotal/cbc:Percent")
            If Not nodePercentDetraccion Is Nothing Then porcentajeDetraccion = Trim(nodePercentDetraccion.Text)
        End If
    End If
    On Error GoTo 0

    razonSocialValue = ""
    On Error Resume Next
    razonSocialValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName").Text
    If razonSocialValue = "" Then razonSocialValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name").Text
    On Error GoTo 0

    If proyectos.Exists(facturadoValue) Then
        proyectoValue = proyectos(facturadoValue)
    Else
        proyectoValue = ""
    End If

    ' RxH
    Dim es_recibo_honorarios As Boolean
    es_recibo_honorarios = False
    Dim nodeTaxCategory As Object, nodePercent As Object
    On Error Resume Next
    Set nodeTaxCategory = xmlDoc.SelectSingleNode("//cac:InvoiceLine/cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory[cbc:ID='RET 4TA']")
    If Not nodeTaxCategory Is Nothing Then
        Set nodePercent = nodeTaxCategory.ParentNode.SelectSingleNode("cbc:Percent")
        If Not nodePercent Is Nothing Then
            If Trim(nodePercent.Text) = "8.00" Or Trim(nodePercent.Text) = "8" Then
                es_recibo_honorarios = True
            End If
        End If
    End If
    On Error GoTo 0

    ' Agregar a la tabla
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, tbl.ListColumns("SERIE").Index).Value = splitID(0)
        .Cells(1, tbl.ListColumns("N°").Index).Value = CStr(Val(splitID(1)))
        .Cells(1, tbl.ListColumns("DESCRIPCION").Index).Value = descripcionValue
        .Cells(1, tbl.ListColumns("FACTURADO A").Index).Value = facturadoValue
        .Cells(1, tbl.ListColumns("RUC").Index).Value = rucValue
        '.Cells(1, tbl.ListColumns("PC").Index).Formula = "=IFERROR(VLOOKUP([@RUC],Tabla9[#All],3,0),""SIN CUENTA"")"
        .Cells(1, tbl.ListColumns("BANCO").Index).Formula = "=IFERROR(INDEX(CUENTAS[BANCO],MATCH(CONCAT([@RUC],[@MONEDA]),CUENTAS[BUSQUEDA],0)),""--"")"
        .Cells(1, tbl.ListColumns("CC").Index).Formula = "=IFERROR(INDEX(CUENTAS[CC/CA],MATCH(CONCAT([@RUC],[@MONEDA]),CUENTAS[BUSQUEDA],0)),""--"")"
        .Cells(1, tbl.ListColumns("CCI").Index).Formula = "=IFERROR(INDEX(CUENTAS[CCI],MATCH(CONCAT([@RUC],[@MONEDA]),CUENTAS[BUSQUEDA],0)),""--"")"
        .Cells(1, tbl.ListColumns("SUBTOTAL").Index).Value = subtotalValue
        .Cells(1, tbl.ListColumns("IMPORTE").Index).Value = importeValue
        .Cells(1, tbl.ListColumns("F. EMISIÓN").Index).Value = fechaValue
        .Cells(1, tbl.ListColumns("MONEDA").Index).Value = monedaValue
        .Cells(1, tbl.ListColumns("TIPO DET").Index).Value = bienValue
        .Cells(1, tbl.ListColumns("PORCENTAJE").Index).Value = IIf(bienValue <> "" And porcentajeDetraccion <> "", porcentajeDetraccion & "%", "")
        .Cells(1, tbl.ListColumns("TIPO").Index).Value = IIf(es_recibo_honorarios, "RxH", "FACTURA")
        .Cells(1, tbl.ListColumns("F. PROVISIÓN").Index).Value = Date
        .Cells(1, tbl.ListColumns("RAZON SOCIAL").Index).Value = razonSocialValue
        .Cells(1, tbl.ListColumns("PROYECTO").Index).Value = proyectoValue
    End With

    ' Hipervínculo a la carpeta
    lastRow = tbl.ListRows.Count
    Dim targetCell As Range
    Set targetCell = tbl.ListColumns("F. PROVISIÓN").DataBodyRange.Cells(lastRow)
    Dim displayText As String
    displayText = Format(targetCell.Value, "dd/mm/yyyy")
    On Error Resume Next
    targetCell.Hyperlinks.Delete
    ws.Hyperlinks.Add Anchor:=targetCell, Address:=Left(archivoXML, InStrRev(archivoXML, "\")), TextToDisplay:=displayText
    On Error GoTo 0

Salir:
   ' MsgBox "Proceso finalizado.", vbInformation
End Sub





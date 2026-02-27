Attribute VB_Name = "XML"
Sub ExtractDataFromXMLs()
    Dim carpetaXML As String
    Dim archivoXML As String
    Dim xmlDoc As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim invoiceID As String
    Dim splitID() As String
    Dim lastRow As Long
    Dim fDialog As FileDialog
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
    
        ' Diccionario de proyectos por razón social
    Dim proyectos As Object
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

    ' ---------- 1. Seleccionar carpeta ----------
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Selecciona la carpeta que contiene los archivos XML"
        If .Show <> -1 Then Exit Sub
        carpetaXML = .SelectedItems(1) & "\"
    End With

    ' ---------- 2. Validar existencia de XMLs ----------
    archivoXML = Dir(carpetaXML & "*.xml")
    If archivoXML = "" Then
        MsgBox "No se encontraron archivos XML en la carpeta seleccionada.", vbExclamation
        Exit Sub
    End If

    ' ---------- 3. Referencias a hoja y tabla ----------
    Set ws = ThisWorkbook.Sheets("BASE DE DATOS GASTOS")
    Set tbl = ws.ListObjects("Tabla3")

    ' ---------- 4. Procesar cada archivo XML ----------
    Do While archivoXML <> ""
        Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
        xmlDoc.async = False
        xmlDoc.validateOnParse = False
        xmlDoc.Load carpetaXML & archivoXML
        xmlDoc.SetProperty "SelectionNamespaces", _
            "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' " & _
            "xmlns:cac='urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'"

        On Error GoTo SiguienteArchivo

        invoiceID = ""
        On Error Resume Next
        invoiceID = xmlDoc.SelectSingleNode("/*[local-name()='Invoice']/*[local-name()='ID']").Text
        On Error GoTo 0
        If invoiceID = "" Or InStr(invoiceID, "-") = 0 Then GoTo SiguienteArchivo
        splitID = Split(invoiceID, "-")

        Dim descripcionValue As String, facturadoValue As String
        Dim rucValue As String, subtotalValue As String
        Dim importeValue As String, fechaValue As String
        Dim monedaValue As String, bienValue As String
        Dim razonSocialValue As String, porcentajeDetraccion As String
        Dim proyectoValue As String

        ' --------- Descripción ---------
        descripcionValue = ""
        On Error Resume Next
        descripcionValue = xmlDoc.SelectSingleNode("//cac:InvoiceLine/cac:Item/cbc:Description").Text
        If descripcionValue = "" Then descripcionValue = xmlDoc.SelectSingleNode("//cac:InvoiceLine/cbc:Note").Text
        On Error GoTo 0

        ' --------- Facturado a ---------
        facturadoValue = ""
        On Error Resume Next
        facturadoValue = xmlDoc.SelectSingleNode("//cac:AccountingCustomerParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName").Text
        If facturadoValue = "" Then facturadoValue = xmlDoc.SelectSingleNode("//cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name").Text
        On Error GoTo 0
        facturadoValue = Replace(facturadoValue, ".", "")

        ' --------- RUC ---------
        rucValue = ""
        On Error Resume Next
        rucValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID").Text
        If rucValue = "" Then rucValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cbc:CustomerAssignedAccountID").Text
        On Error GoTo 0

        ' --------- Subtotal ---------
        subtotalValue = ""
        On Error Resume Next
        subtotalValue = xmlDoc.SelectSingleNode("//cac:LegalMonetaryTotal/cbc:LineExtensionAmount").Text
        On Error GoTo 0

        ' --------- Importe ---------
        importeValue = ""
        On Error Resume Next
        importeValue = xmlDoc.SelectSingleNode("//cac:LegalMonetaryTotal/cbc:PayableAmount").Text
        On Error GoTo 0

        ' --------- Fecha ---------
        fechaValue = ""
        On Error Resume Next
        fechaValue = xmlDoc.SelectSingleNode("//cbc:IssueDate").Text
        On Error GoTo 0

        ' --------- Moneda ---------
        monedaValue = ""
        On Error Resume Next
        Dim nodeMoneda As Object
        Set nodeMoneda = xmlDoc.SelectSingleNode("//cac:LegalMonetaryTotal/cbc:PayableAmount")
        If Not nodeMoneda Is Nothing Then
            If Not nodeMoneda.Attributes Is Nothing Then
                If nodeMoneda.Attributes.Length > 0 Then
                    monedaValue = nodeMoneda.Attributes.getNamedItem("currencyID").Text
                End If
            End If
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

        ' --------- Código y Porcentaje de Detracción ---------
        bienValue = ""
        porcentajeDetraccion = ""
        On Error Resume Next
        bienValue = xmlDoc.SelectSingleNode("//cac:PaymentTerms/cbc:PaymentMeansID").Text
        Select Case bienValue
            Case "019", "020", "022", "025", "027", "030", "037"
            Case Else
                bienValue = ""
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

        ' --------- RAZÓN SOCIAL ---------
        razonSocialValue = ""
        On Error Resume Next
        razonSocialValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName").Text
        If razonSocialValue = "" Then
            razonSocialValue = xmlDoc.SelectSingleNode("//cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name").Text
        End If
        On Error GoTo 0

        ' --------- Obtener nombre del proyecto ----------
        proyectoValue = ""
        If proyectos.Exists(facturadoValue) Then
            proyectoValue = proyectos(facturadoValue)
        End If

        ' --------- ¿Es Recibo por Honorarios? ---------
        Dim es_recibo_honorarios As Boolean
        es_recibo_honorarios = False
        Dim nodeTaxCategory As Object
        Dim nodePercent As Object

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

        ' --------- Agregar fila a la tabla ----------
        Set newRow = tbl.ListRows.Add
        newRow.Range(, tbl.ListColumns("SERIE").Index).Value = splitID(0)
        newRow.Range(, tbl.ListColumns("N°").Index).Value = CStr(Val(splitID(1)))
        newRow.Range(, tbl.ListColumns("DESCRIPCION").Index).Value = descripcionValue
        newRow.Range(, tbl.ListColumns("FACTURADO A").Index).Value = facturadoValue
        newRow.Range(, tbl.ListColumns("RUC").Index).Value = rucValue
        'newRow.Range(, tbl.ListColumns("PC").Index).Formula = "=IFERROR(VLOOKUP([@RUC],Tabla9[#All],3,0),""SIN CUENTA"")"
        'newRow.Range(1, tbl.ListColumns("BANCO").Index).Formula = "=IFERROR(XLOOKUP(CONCAT([@RUC],[@MONEDA]), 'CUENTAS'!H:H, 'CUENTAS'!C:C), ""--"")"
        'newRow.Range(1, tbl.ListColumns("CC").Index).Formula = "=IFERROR(XLOOKUP(CONCAT([@RUC],[@MONEDA]), 'CUENTAS'!H:H, 'CUENTAS'!E:E), ""--"")"
        'newRow.Range(1, tbl.ListColumns("CCI").Index).Formula = "=IFERROR(XLOOKUP(CONCAT([@RUC],[@MONEDA]), 'CUENTAS'!H:H, 'CUENTAS'!F:F), ""--"")"
        newRow.Range(1, tbl.ListColumns("BANCO").Index).Formula = "=IFERROR(INDEX(CUENTAS[BANCO],MATCH(CONCAT([@RUC],[@MONEDA]),CUENTAS[BUSQUEDA],0)),""--"")"
        newRow.Range(1, tbl.ListColumns("CC").Index).Formula = "=IFERROR(INDEX(CUENTAS[CC/CA],MATCH(CONCAT([@RUC],[@MONEDA]),CUENTAS[BUSQUEDA],0)),""--"")"
        newRow.Range(1, tbl.ListColumns("CCI").Index).Formula = "=IFERROR(INDEX(CUENTAS[CCI], MATCH(CONCAT([@RUC], [@MONEDA]), CUENTAS[BUSQUEDA], 0)),""--"")"
        newRow.Range(, tbl.ListColumns("SUBTOTAL").Index).Value = subtotalValue
        newRow.Range(, tbl.ListColumns("IMPORTE").Index).Value = importeValue
        newRow.Range(, tbl.ListColumns("F. EMISIÓN").Index).Value = fechaValue
        newRow.Range(, tbl.ListColumns("MONEDA").Index).Value = monedaValue
        newRow.Range(, tbl.ListColumns("TIPO DET").Index).Value = bienValue
        If bienValue <> "" And porcentajeDetraccion <> "" Then
            newRow.Range(, tbl.ListColumns("PORCENTAJE").Index).Value = porcentajeDetraccion & "%"
        End If
        If es_recibo_honorarios Then
            newRow.Range(, tbl.ListColumns("TIPO").Index).Value = "RxH"
        Else
            newRow.Range(, tbl.ListColumns("TIPO").Index).Value = "FACTURA"
        End If
        newRow.Range(, tbl.ListColumns("F. PROVISIÓN").Index).Value = Date
        newRow.Range(, tbl.ListColumns("RAZON SOCIAL").Index).Value = razonSocialValue
        newRow.Range(, tbl.ListColumns("PROYECTO").Index).Value = proyectoValue

        ' --------- Hipervínculo a carpeta ----------
        lastRow = tbl.ListRows.Count
        Dim targetCell As Range
        Set targetCell = tbl.ListColumns("F. PROVISIÓN").DataBodyRange.Cells(lastRow)
        Dim displayText As String
        displayText = Format(targetCell.Value, "dd/mm/yyyy")
        On Error Resume Next
        targetCell.Hyperlinks.Delete
        On Error GoTo 0
        ws.Hyperlinks.Add Anchor:=targetCell, Address:=carpetaXML, TextToDisplay:=displayText

SiguienteArchivo:
        archivoXML = Dir
        On Error GoTo 0
    Loop

    MsgBox "Proceso finalizado.", vbInformation
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



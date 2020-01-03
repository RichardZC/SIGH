VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptERecaudAltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Recaudación de Altas
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim rsReporte As New ADODB.Recordset
Dim rsRsTmp1 As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasReportes As New ReglasReportes
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New ERecaudAltas
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim mrs_Tmp1 As New Recordset
Dim mrs_Tmp As New Recordset
Dim iFila As Long: Dim lnNroDias As Integer
Dim lnTotal As Double: Dim lnPrecio As Double
Dim lcPaciente As String: Dim lcDias As String
Dim lnIdPaciente As Long: Dim lnIdServicio As Long: Dim lnNroHistorias As Long
Dim lnSubTotal As Double: Dim lnPago As Double: Dim lnSaldo As Double: Dim lnTsaldo As Double
Dim lnImpSIS As Double: Dim lnImpExo  As Double: Dim lnImpSOAT As Double: Dim lnImpTotal As Double
Dim lcTipoPaciente As String
Dim lbSeguir As Boolean: Dim lcDocum As String: Dim lcBuscar As String: Dim lnDctos As Double
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    On Error GoTo ManejadorErrorExcel
    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    GenerarRecordsetTemporal
    GenerarRecordsetDctos
    'Filtra los Datos
    Set mrs_Tmp1 = mo_ReglasReportes.AtencionesSeleccionarPorFechasDeEgresoAdministrativo(mda_FechaInicio, mda_FechaFin)
    lnNroHistorias = mrs_Tmp1.RecordCount
    If lnNroHistorias = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        lnTotal = 0: lnTsaldo = 0
        mrs_Tmp1.MoveFirst
        Do While Not mrs_Tmp1.EOF
           lnImpExo = 0: lnImpSIS = 0: lnImpSOAT = 0: lnPago = 0: lnSubTotal = 0
           'Farmacia
           Set mrs_Tmp = FacturacionBienesInsumosSeleccionarPorIdAtencion(mrs_Tmp1.Fields!idAtencion)
           If mrs_Tmp.RecordCount > 0 Then
              mrs_Tmp.MoveFirst
              Do While Not mrs_Tmp.EOF
                    Select Case mrs_Tmp!IdEstadoFacturacion
                    Case 4   'Solo los Pagados
                       lnImpExo = lnImpExo + IIf(IsNull(mrs_Tmp!importeEXO), 0, mrs_Tmp!importeEXO)             'Exonerado
                       lnImpSIS = lnImpSIS + IIf(IsNull(mrs_Tmp!ImporteSIS), 0, mrs_Tmp!ImporteSIS)             'SIS
                       lnImpSOAT = lnImpSOAT + IIf(IsNull(mrs_Tmp!ImporteSOAT), 0, mrs_Tmp!ImporteSOAT)             'SOAT
                       lnPago = lnPago + mrs_Tmp!totalPorPagar
                       lnSubTotal = lnSubTotal + mrs_Tmp!TotalPagar
                    Case 1, 3  'Solo Registrados, Pendientes de pago
                       lnSubTotal = lnSubTotal + mrs_Tmp!TotalPagar
                    End Select
                    mrs_Tmp.MoveNext
              Loop
              lcTipoPaciente = "Paciente Normal"
              If lnImpSIS > 0 Then
                 lcTipoPaciente = "Paciente SIS"
              ElseIf lnImpSOAT > 0 Then
                 lcTipoPaciente = "Paciente SOAT"
              End If
           End If
           mrs_Tmp.Close
           
           'Servicios
           Set mrs_Tmp = FacturacionServiciosSeleccionarPorIdAtencion(mrs_Tmp1.Fields!idAtencion)
           lnDctos = 0
           If mrs_Tmp.RecordCount > 0 Then
              mrs_Tmp.MoveFirst
              Do While Not mrs_Tmp.EOF
                    Select Case mrs_Tmp!IdEstadoFacturacion
                    Case 4   'Solo los Pagados
                       lnImpExo = lnImpExo + IIf(IsNull(mrs_Tmp!importeEXO), 0, mrs_Tmp!importeEXO)             'Exonerado
                       lnImpSIS = lnImpSIS + IIf(IsNull(mrs_Tmp!ImporteSIS), 0, mrs_Tmp!ImporteSIS)             'SIS
                       lnImpSOAT = lnImpSOAT + IIf(IsNull(mrs_Tmp!ImporteSOAT), 0, mrs_Tmp!ImporteSOAT)             'SOAT
                       lnPago = lnPago + mrs_Tmp!totalPorPagar
                       lnSubTotal = lnSubTotal + mrs_Tmp!TotalPagar
                    Case 1, 3  'Solo Registrados, Pendientes de pago
                       lnSubTotal = lnSubTotal + mrs_Tmp!TotalPagar
                    End Select
                    
                    'Suma Descuentos de la CABECERA - COMPROBANTE PAGO

                    
                    mrs_Tmp.MoveNext
              Loop
              lcTipoPaciente = "Paciente Normal"
              If lnImpSIS > 0 Then
                 lcTipoPaciente = "Paciente SIS"
              ElseIf lnImpSOAT > 0 Then
                 lcTipoPaciente = "Paciente SOAT"
              End If
           End If
           mrs_Tmp.Close
           
           lnSubTotal = lnSubTotal - lnDctos
           lnImpTotal = lnSubTotal - lnImpSIS - lnImpExo - lnImpSOAT
           lnSaldo = lnImpTotal - lnPago
           
           lnTotal = lnTotal + lnPago
           lnTsaldo = lnTsaldo + lnSaldo
           
           rsReporte.AddNew
           rsReporte.Fields!FechaEgresoAdministrativo = mrs_Tmp1.Fields!FechaEgresoAdministrativo
           rsReporte.Fields!FechaIngreso = mrs_Tmp1.Fields!FechaIngreso
           rsReporte.Fields!Paciente = Trim(mrs_Tmp1.Fields("ApellidoPaterno").Value) & " " & mrs_Tmp1.Fields("ApellidoMaterno").Value & " " & mrs_Tmp1.Fields("PrimerNombre").Value
           rsReporte.Fields!nrohistoriaclinica = mrs_Tmp1.Fields!nrohistoriaclinica
           rsReporte.Fields!DServicio = mrs_Tmp1.Fields("CodServ").Value & " - " & mrs_Tmp1.Fields("dServicio").Value
           rsReporte.Fields!tipoPaciente = lcTipoPaciente
           rsReporte.Fields!Subtotal = lnSubTotal
           rsReporte.Fields!SIS = lnImpSIS
           rsReporte.Fields!EXO = lnImpExo
           rsReporte.Fields!Total = lnImpTotal
           rsReporte.Fields!Pago = lnPago
           rsReporte.Fields!Saldo = lnSaldo
           rsReporte.Update
           mrs_Tmp1.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ERecaudAlta.ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/OpenOffice.ods"
            '
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
            '

            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ERecaudAlta.xls")
            oWorkBookPlantilla.Worksheets("ERecaudAlta").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
        Else
            oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
        End If
        iFila = 6: lnTotal = 0: lnSaldo = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("FechaEgresoAdministrativo").Value)
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("Paciente").Value)
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte.Fields("nroHistoriaClinica").Value)
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields("TipoPaciente").Value)
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula(rsReporte.Fields("dservicio").Value)
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(rsReporte.Fields("FechaIngreso").Value)
            Call Feuille.getcellbyposition(12, iFila - 1).setFormula(rsReporte.Fields("FechaEgresoAdministrativo").Value)
            Call Feuille.getcellbyposition(13, iFila - 1).setFormula(rsReporte.Fields("subTotal").Value)
            Call Feuille.getcellbyposition(14, iFila - 1).setFormula(rsReporte.Fields("SIS").Value)
            Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields("EXO").Value)
            Call Feuille.getcellbyposition(16, iFila - 1).setFormula(rsReporte.Fields("Total").Value)
            Call Feuille.getcellbyposition(17, iFila - 1).setFormula(rsReporte.Fields("Pago").Value)
            Call Feuille.getcellbyposition(18, iFila - 1).setFormula(rsReporte.Fields("saldo").Value)
        Else
            oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("FechaEgresoAdministrativo").Value
            oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("Paciente").Value
            oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields("nroHistoriaClinica").Value
            oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields("TipoPaciente").Value
            oWorkSheet.Cells(iFila, 10).Value = rsReporte.Fields("dservicio").Value
            oWorkSheet.Cells(iFila, 12).Value = rsReporte.Fields("FechaIngreso").Value
            oWorkSheet.Cells(iFila, 13).Value = rsReporte.Fields("FechaEgresoAdministrativo").Value
            oWorkSheet.Cells(iFila, 14).Value = rsReporte.Fields("subTotal").Value
            oWorkSheet.Cells(iFila, 15).Value = rsReporte.Fields("SIS").Value
            oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields("EXO").Value
            oWorkSheet.Cells(iFila, 17).Value = rsReporte.Fields("Total").Value
            oWorkSheet.Cells(iFila, 18).Value = rsReporte.Fields("Pago").Value
            oWorkSheet.Cells(iFila, 19).Value = rsReporte.Fields("saldo").Value
        End If
            lnSaldo = lnSaldo + rsReporte.Fields("saldo").Value
            lnTotal = lnTotal + rsReporte.Fields("Pago").Value
            iFila = iFila + 1
            rsReporte.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":S" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Historias Clínicas: " + Trim(Str(lnNroHistorias)))
            Call Feuille.getcellbyposition(17, iFila - 1).setFormula(lnTotal)
            Call Feuille.getcellbyposition(18, iFila - 1).setFormula(lnSaldo)
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 19
            oWorkSheet.Cells(iFila, 2).Value = "Nº Historias Clínicas: " + Trim(Str(lnNroHistorias))
            oWorkSheet.Cells(iFila, 18).Value = lnTotal
            oWorkSheet.Cells(iFila, 19).Value = lnSaldo
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 19
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
    End If
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub

Sub GenerarRecordsetTemporal()
    With rsReporte
          .Fields.Append "FechaEgresoAdministrativo", adDate
          .Fields.Append "FechaIngreso", adDate
          .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
          .Fields.Append "NroHistoriaClinica", adInteger, 4, adFldIsNullable
          .Fields.Append "dServicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "TipoPaciente", adVarChar, 50, adFldIsNullable
          .Fields.Append "SubTotal", adDouble, , adFldIsNullable
          .Fields.Append "SIS", adDouble, , adFldIsNullable
          .Fields.Append "EXO", adDouble, , adFldIsNullable
          .Fields.Append "Total", adDouble, , adFldIsNullable
          .Fields.Append "Pago", adDouble, , adFldIsNullable
          .Fields.Append "Saldo", adDouble, , adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Sub GenerarRecordsetDctos()
    With rsRsTmp1
          .Fields.Append "Docume", adBSTR, 150
          .Fields.Append "Dctos", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub







Function FacturacionServiciosSeleccionarPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServiciosSeleccionarPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServiciosSeleccionarPorIdAtencion"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServiciosSeleccionarPorIdAtencion = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesInsumosSeleccionarPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesInsumosSeleccionarPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesInsumosSeleccionarPorIdAtencion"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesInsumosSeleccionarPorIdAtencion = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



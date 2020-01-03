VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEExoneraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Exoneraciones
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
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasCaja As New ReglasCaja

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
    Dim oFormulario As New EExoneraciones
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
Dim lnTpacientes As Long: Dim lnPacientes As Long
Dim lnTImpExo As Double: Dim lnTPago As Double
Dim lcDServicio As String
Dim lcDocumento As String: Dim lnDctos As Double: Dim lnIdCuentaAtencion As Long
Dim lbContinua As Boolean: Dim lnIdServicioEgreso As Long
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
       'Farmacia
       Set mrs_Tmp = ml_ReglasFacturacion.FarmaciaConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo(mda_FechaInicio, mda_FechaFin + 1)
       If mrs_Tmp.RecordCount > 0 Then
          mrs_Tmp.MoveFirst
          Do While Not mrs_Tmp.EOF
                lnImpExo = 0: lnImpSIS = 0: lnImpSOAT = 0: lnPago = 0: lnSubTotal = 0
                If IsNull(mrs_Tmp.Fields!idCuentaAtencion) Then
                   mrs_Tmp.MoveNext
                Else
                lnIdCuentaAtencion = mrs_Tmp.Fields!idCuentaAtencion
                lnIdServicioEgreso = mrs_Tmp.Fields!IdServicioEgreso
                lcDServicio = mrs_Tmp.Fields!CodServicio & " - " & mrs_Tmp.Fields!DServicio
                'exoneraciones
                Set mrs_Tmp1 = ml_ReglasFacturacion.FacturacionBienesFinanciamientosPorIdCuenta(mrs_Tmp.Fields!idCuentaAtencion)
                If mrs_Tmp1.RecordCount > 0 Then
                   mrs_Tmp1.MoveFirst
                   Do While Not mrs_Tmp1.EOF
                      If mrs_Tmp1.Fields!idTipoFinanciamiento = 9 Then
                         lnImpExo = lnImpExo + mrs_Tmp1.Fields!TotalFinanciado
                      End If
                      mrs_Tmp1.MoveNext
                   Loop
                End If
                lnPago = 0
                Do While Not mrs_Tmp.EOF And lnIdCuentaAtencion = mrs_Tmp.Fields!idCuentaAtencion
                    Select Case mrs_Tmp!IdEstadoFacturacion
                    Case 4   'Solo los Pagados
                       lnPago = lnPago + mrs_Tmp!TotalPagar
                    End Select
                    mrs_Tmp.MoveNext
                    If mrs_Tmp.EOF Then
                       Exit Do
                    End If
                Loop
                rsReporte.AddNew
                rsReporte.Fields!idServicio = lnIdServicioEgreso
                rsReporte.Fields!IdCuenta = lnIdCuentaAtencion
                rsReporte.Fields!DServicio = lcDServicio
                rsReporte.Fields!ImpExonerado = lnImpExo
                rsReporte.Fields!impPagado = lnPago
                rsReporte.Update
                End If
          Loop
       End If
       mrs_Tmp.Close
       
       'Servicios
       Set mrs_Tmp = ml_ReglasFacturacion.ServiciosConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo(mda_FechaInicio, mda_FechaFin + 1)
       If mrs_Tmp.RecordCount > 0 Then
          mrs_Tmp.MoveFirst
          Do While Not mrs_Tmp.EOF
                lnIdCuentaAtencion = mrs_Tmp.Fields!idCuentaAtencion
                lnIdServicioEgreso = mrs_Tmp.Fields!IdServicioEgreso
                lcDServicio = mrs_Tmp.Fields!CodServicio & " - " & mrs_Tmp.Fields!DServicio
                'exoneraciones
                Set mrs_Tmp1 = ml_ReglasFacturacion.FacturacionServicioFinanciamientosPorIdCuenta(mrs_Tmp.Fields!idCuentaAtencion)
                If mrs_Tmp1.RecordCount > 0 Then
                   mrs_Tmp1.MoveFirst
                   Do While Not mrs_Tmp1.EOF
                      If mrs_Tmp1.Fields!idTipoFinanciamiento = 9 Then
                         lnImpExo = lnImpExo + mrs_Tmp1.Fields!TotalFinanciado
                      End If
                      mrs_Tmp1.MoveNext
                   Loop
                End If
                lnPago = 0
                Do While Not mrs_Tmp.EOF And lnIdCuentaAtencion = mrs_Tmp.Fields!idCuentaAtencion
                    Select Case mrs_Tmp!IdEstadoFacturacion
                    Case 4   'Solo los Pagados
                       lnPago = lnPago + mrs_Tmp!Total
                    End Select
                    mrs_Tmp.MoveNext
                    If mrs_Tmp.EOF Then
                       Exit Do
                    End If
                Loop
                rsReporte.AddNew
                rsReporte.Fields!idServicio = lnIdServicioEgreso
                rsReporte.Fields!IdCuenta = lnIdCuentaAtencion
                rsReporte.Fields!DServicio = lcDServicio
                rsReporte.Fields!ImpExonerado = lnImpExo
                rsReporte.Fields!impPagado = lnPago
          Loop
       End If
       mrs_Tmp.Close
       If rsReporte.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
       Else
            rsReporte.Sort = "dservicio"
                If lbEsOpenOffice = True Then
                    'Abre el archivo ExcelOpenOffice
                    lcArchivoExcel = App.Path + "\Plantillas\EExoneraciones.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
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
                    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EExoneraciones.xls")
                    oWorkBookPlantilla.Worksheets("EExoneraciones").Copy Before:=oWorkBook.Sheets(1)
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
            iFila = 6: lnTotal = 0: lnTpacientes = 0: lnTImpExo = 0: lnTPago = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               lnTotal = lnTotal + 1
               lcDServicio = rsReporte.Fields!DServicio
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lnTotal)
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcDServicio)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = lnTotal
                        oWorkSheet.Cells(iFila, 3).Value = lcDServicio
                    End If
               lnImpExo = 0: lnPago = 0: lnPacientes = 0
               Do While Not rsReporte.EOF And lcDServicio = rsReporte.Fields!DServicio
                  If rsReporte!ImpExonerado > 0 Then
                        lnPacientes = lnPacientes + 1
                        lnImpExo = lnImpExo + rsReporte!ImpExonerado
                        lnPago = lnPago + rsReporte!impPagado
                  End If
                  rsReporte.MoveNext
                  If rsReporte.EOF Then
                     Exit Do
                  End If
               Loop
               If lnImpExo > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnPacientes)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnPago + lnImpExo)
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnImpExo)
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnPago)
                    Else
                        oWorkSheet.Cells(iFila, 5).Value = lnPacientes
                        oWorkSheet.Cells(iFila, 6).Value = lnPago + lnImpExo
                        oWorkSheet.Cells(iFila, 7).Value = lnImpExo
                        oWorkSheet.Cells(iFila, 8).Value = lnPago
                    End If
                    
                    lnTpacientes = lnTpacientes + lnPacientes
                    lnTImpExo = lnTImpExo + lnImpExo
                    lnTPago = lnTPago + lnPago
                    
                    iFila = iFila + 1
               End If
            Loop
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnTpacientes)
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnTPago + lnTImpExo)
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnTImpExo)
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnTPago)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 8
                oWorkSheet.Cells(iFila, 5).Value = lnTpacientes
                oWorkSheet.Cells(iFila, 6).Value = lnTPago + lnTImpExo
                oWorkSheet.Cells(iFila, 7).Value = lnTImpExo
                oWorkSheet.Cells(iFila, 8).Value = lnTPago
             End If
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 8
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
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
          .Fields.Append "idServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "idCuenta", adInteger, 4, adFldIsNullable
          .Fields.Append "dServicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "ImpExonerado", adDouble, , adFldIsNullable
          .Fields.Append "ImpPagado", adDouble, , adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub


Sub CrearReportePorEmpleadoDeServicioSocial(lnIdEmpleado As Long)
Dim rsReporte As New Recordset
Dim rsTmp1 As New Recordset
Dim lcSql As String, lnTotal As Double
        On Error GoTo ManejadorError
        With rsReporte
              .Fields.Append "Fecha", adDate
              .Fields.Append "Empleado", adVarChar, 100, adFldIsNullable
              .Fields.Append "Boleta", adVarChar, 50, adFldIsNullable
              .Fields.Append "Cajero", adVarChar, 100, adFldIsNullable
              .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
              .Fields.Append "Cuenta", adInteger, 4, adFldIsNullable
              .Fields.Append "Tipo", adVarChar, 10, adFldIsNullable
              .Fields.Append "Importe", adDouble, , adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With
        'Filtra los Datos
        lnTotal = 0
        Set rsTmp1 = ml_ReglasFacturacion.FactOrdenServicioPagosXImporteExoneradoMayorAcero(lnIdEmpleado, mda_FechaInicio, mda_FechaFin)
        If rsTmp1.RecordCount > 0 Then
           rsTmp1.MoveFirst
           Do While Not rsTmp1.EOF
              rsReporte.AddNew
              rsReporte.Fields!Fecha = rsTmp1.Fields!fechaCreacion
              rsReporte.Fields!Empleado = rsTmp1.Fields!empExonera
              rsReporte.Fields!Importe = rsTmp1.Fields!ImporteExonerado
              If Not IsNull(IdComprobantePago) Then
                    rsReporte.Fields!Boleta = Trim(rsTmp1.Fields!nroSerie) & "-" & rsTmp1.Fields!NroDocumento
                    rsReporte.Fields!Cajero = rsTmp1.Fields!EmpCajero
                    rsReporte.Fields!Paciente = rsTmp1.Fields!razonSocial
                    rsReporte.Fields!cuenta = rsTmp1.Fields!idCuentaAtencion
                    rsReporte.Fields!Tipo = "Servicio"
              End If
              rsReporte.Update
              lnTotal = lnTotal + rsTmp1.Fields!ImporteExonerado
              rsTmp1.MoveNext
           Loop
        End If
        rsTmp1.Close
        '
        Set rsTmp1 = ml_ReglasFacturacion.FactOrdenesBienesXimporteExoneradoMayorAcero(lnIdEmpleado, mda_FechaInicio, mda_FechaFin)
        If rsTmp1.RecordCount > 0 Then
           rsTmp1.MoveFirst
           Do While Not rsTmp1.EOF
              rsReporte.AddNew
              rsReporte.Fields!Fecha = rsTmp1.Fields!fechaCreacion
              rsReporte.Fields!Empleado = rsTmp1.Fields!empExonera
              rsReporte.Fields!Importe = rsTmp1.Fields!ImporteExonerado
              If Not IsNull(IdComprobantePago) Then
                    rsReporte.Fields!Boleta = Trim(rsTmp1.Fields!nroSerie) & "-" & rsTmp1.Fields!NroDocumento
                    rsReporte.Fields!Cajero = rsTmp1.Fields!EmpCajero
                    rsReporte.Fields!Paciente = rsTmp1.Fields!razonSocial
                    rsReporte.Fields!cuenta = rsTmp1.Fields!idCuentaAtencion
                    rsReporte.Fields!Tipo = "Servicio"
              End If
              rsReporte.Update
              lnTotal = lnTotal + rsTmp1.Fields!ImporteExonerado
              rsTmp1.MoveNext
           Loop
        End If
        rsTmp1.Close
        rsReporte.Sort = "Fecha asc"
        '
        Set RpEExoneraciones.DataSource = rsReporte
        RpEExoneraciones.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        RpEExoneraciones.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        RpEExoneraciones.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        RpEExoneraciones.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
        RpEExoneraciones.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
        RpEExoneraciones.Sections("cabecera").Controls("lblTitulo").Caption = "Exoneraciones por Empleado de Servicio Social"
        RpEExoneraciones.Sections("cabecera").Controls("lblSubTitulo").Caption = ml_TextoDelFiltro
        
        Set RpEExoneraciones.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
        RpEExoneraciones.Sections("piePag").Controls("lblTotal").Caption = "N° " & Trim(Str(rsReporte.RecordCount))
        RpEExoneraciones.Sections("piePag").Controls("lblImporteT").Caption = Trim(Str(lnTotal))
        RpEExoneraciones.Orientation = rptOrientLandscape
        RpEExoneraciones.Show 1
        'debb-27/05/2015
        Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
        mo_reglasComunes.grabaTablaAuditoria ("RpEExoneraciones: " & _
                                       Mid(ml_TextoDelFiltro, IIf(InStr(ml_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))
        Set mo_reglasComunes = Nothing
        '
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Resume
    Exit Sub
End Sub

'debb-18/05/2016
Sub CrearReportePorFechasBoletas()
Dim rsReporte As New Recordset
Dim rsTmp1 As New Recordset
Dim lcSql As String, lnTotal As Double, ldFecha As Date, lnRegistros As Long
        On Error GoTo ManejadorError
        With rsReporte
              .Fields.Append "movNumero", adVarChar, 9, adFldIsNullable
              .Fields.Append "movTipo", adVarChar, 1, adFldIsNullable
              .Fields.Append "idproducto", adInteger
              .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
              .Fields.Append "FechaVencimiento", adDate
              .Fields.Append "item", adInteger
              .Fields.Append "cantidad", adInteger
              .Fields.Append "precio", adDouble
              .Fields.Append "total", adDouble
              .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
              .Fields.Append "codigo", adVarChar, 10, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
              .Fields.Append "Presentacion", adVarChar, 100, adFldIsNullable
              .Fields.Append "idAlmcen", adInteger
              .Fields.Append "Totalactual", adDouble, , adFldIsNullable
              .Fields.Append "CantidadSaldo", adDouble
              .Fields.Append "CantidadFaltante", adDouble
              .Fields.Append "CantidadSobrante", adDouble
              .Fields.Append "totalgen", adDouble
              .Fields.Append "totalf", adDouble, , adFldIsNullable
              .Fields.Append "totals", adDouble, , adFldIsNullable
              .Fields.Append "FormaFarmaceutica", adVarChar, 100, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With
        lnRegistros = 0
        'Servicios
        Set rsTmp1 = mo_ReglasCaja.CajaComprobantesPagoXfechasExoneradosServicios(mda_FechaInicio, mda_FechaFin)
        If rsTmp1.RecordCount > 0 Then
           rsTmp1.MoveFirst
           Do While Not rsTmp1.EOF
If Val(rsTmp1!NroDocumento) = 1375727 Then
lcSql = ""
End If
              lnRegistros = lnRegistros + 1
              rsReporte.AddNew
              rsReporte.Fields!FechaVencimiento = Format(rsTmp1!fechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY)
              rsReporte.Fields!Lote = rsTmp1!nroSerie + "-" + rsTmp1!NroDocumento
              rsReporte.Fields!Nombre = Left(rsTmp1!razonSocial, 150)
              rsReporte.Fields!Presentacion = Left(rsTmp1!empExonera, 100)
              rsReporte.Fields!totals = rsTmp1!Exoneraciones + rsTmp1!Total
              rsReporte.Fields!totalf = rsTmp1!Exoneraciones
              rsReporte.Fields!Totalactual = rsTmp1!Total
              rsReporte.Update
              ldFecha = rsTmp1.Fields!fechaCobranza
              lcSql = rsTmp1!nroSerie + " " + rsTmp1!NroDocumento
              Do While Not rsTmp1.EOF And ldFecha = rsTmp1.Fields!fechaCobranza And lcSql = (rsTmp1!nroSerie + " " + rsTmp1!NroDocumento)
                 rsTmp1.MoveNext
                 If rsTmp1.EOF Then
                    Exit Do
                 End If
              Loop
           Loop
        End If
        rsTmp1.Close
        'Farmacia
        Set rsTmp1 = mo_ReglasCaja.CajaComprobantesPagoXfechasExoneradosFarmacia(mda_FechaInicio, mda_FechaFin)
        If rsTmp1.RecordCount > 0 Then
           rsTmp1.MoveFirst
           Do While Not rsTmp1.EOF
If Val(rsTmp1!NroDocumento) = 159979 Then
lcSql = ""
End If
              lnRegistros = lnRegistros + 1
              rsReporte.AddNew
              rsReporte.Fields!FechaVencimiento = Format(rsTmp1!fechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY)
              rsReporte.Fields!Lote = rsTmp1!nroSerie + " " + rsTmp1!NroDocumento
              rsReporte.Fields!Nombre = Left(rsTmp1!razonSocial, 150)
              rsReporte.Fields!Presentacion = Left(rsTmp1!empExonera, 100)
              rsReporte.Fields!totals = rsTmp1!Exoneraciones + rsTmp1!Total
              rsReporte.Fields!totalf = rsTmp1!Exoneraciones
              rsReporte.Fields!Totalactual = rsTmp1!Total
              rsReporte.Update
              ldFecha = rsTmp1.Fields!fechaCobranza
              lcSql = rsTmp1!nroSerie + " " + rsTmp1!NroDocumento
              Do While Not rsTmp1.EOF And ldFecha = rsTmp1.Fields!fechaCobranza And lcSql = (rsTmp1!nroSerie + " " + rsTmp1!NroDocumento)
                 rsTmp1.MoveNext
                 If rsTmp1.EOF Then
                    Exit Do
                 End If
              Loop
           Loop
        End If
        rsTmp1.Close
        If rsReporte.RecordCount = 0 Then
            MsgBox "No existe información con esos datos", vbInformation, ""
        Else
            rsReporte.Sort = "FechaVencimiento,lote"
            Dim oRpt As New rCrystal
            oRpt.TextoDelFiltro = ml_TextoDelFiltro
            oRpt.TipoReporte = "EExoneracionXboleta"
            Set oRpt.RecordSet_mrs_Tmp = rsReporte
            oRpt.Show 1
            
            '
            Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
            mo_reglasComunes.grabaTablaAuditoria ("Exoneraciones x Boleta: " & _
                                           Mid(ml_TextoDelFiltro, IIf(InStr(ml_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))
            Set mo_reglasComunes = Nothing
        
        End If
        
        
        
        '
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Resume
    Exit Sub
End Sub








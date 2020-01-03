VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptESisSoatExoConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Liquidación
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oRsPacientes As New ADODB.Recordset
Dim ml_IdServicioCE As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim ml_TextoDelFiltro1 As String
Dim mda_FechaInicio As String
Dim mda_FechaFin As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_IdPlan As Long
Dim ml_IdEstadoCuenta As Long
Public Event ProgressActualizaValor(ByRef lnValorActual As Long, ByRef lnValorTotal As Long)
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let IdEstadoCuenta(lValue As Long)
    ml_IdEstadoCuenta = lValue
End Property
Property Let IdPlan(lValue As Long)
    ml_IdPlan = lValue
End Property
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Let FechaInicio(daValue As String)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As String)
    mda_FechaFin = daValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let TextoDelFiltro1(lValue As String)
    ml_TextoDelFiltro1 = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdServicioCE(lValue As Long)
    ml_IdServicioCE = lValue
End Property
Property Let IdResponsable(lValue As Long)
    ml_IdResponsable = lValue
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New ESisSoatExoConv
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lbAgregarCuentasSinAlta As Boolean, lnHwnd As Long)
Dim iFila As Long
Dim lnNumHistorias As Long
Dim oRsFarmacia As Recordset
Dim oRsServicio As Recordset
Dim oRsTmp1 As New Recordset
Dim lnNroHistoriaClinica As Long: Dim lcPaciente As String: Dim lcEstadoCuenta As String
Dim lnCuenta As Long: Dim lcServicioH As String: Dim ldFechaIngreso As Date
Dim lnTotal As Double: Dim ldFechaLiq As Date: Dim lnTotServicios As Double
Dim lnTotFarmacia As Double, lnTotTotal As Double, lnTotPagosAdelantados As Double
Dim lnAdelantos As Double, lnPagosAdelantados As Double, lnFor As Integer
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oConexion As New Connection
Dim lRecordCount As Long, f As Long
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
GenerarRecordsetTemporal

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
    

    'Filtra datos de Servicios (Hosp/Emerg)
    For lnFor = 1 To 2
        If lnFor = 2 Then
           If lbAgregarCuentasSinAlta = False Then
              Exit For
           End If
           Set oRsServicio = ml_ReglasFacturacion.AtencionesSeleccionarPorFuenteFinanciamientoFingresoSinAlta1(CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdPlan, ml_IdEstadoCuenta)
        Else
           Set oRsServicio = ml_ReglasFacturacion.atencionesSeleccionarPorFuenteFinanciamientoFAltaMedica(CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdPlan, ml_IdEstadoCuenta)
        End If
        lRecordCount = oRsServicio.RecordCount
        If lRecordCount > 0 Then
           f = 0
           oRsServicio.MoveFirst
           Do While Not oRsServicio.EOF
              f = f + 1
              RaiseEvent ProgressActualizaValor(f, lRecordCount)
              '
              lnNroHistoriaClinica = oRsServicio.Fields!nroHistoriaClinica
              lcPaciente = Trim(oRsServicio.Fields!ApellidoPaterno) & " " & Trim(oRsServicio.Fields!apellidoMaterno) & " " & Trim(oRsServicio.Fields!PrimerNombre) & IIf(lnFor = 2, " (Cta sin AM)", "")
              lcEstadoCuenta = oRsServicio.Fields!DestadoCta  'IIf(oRsServicio.Fields!IdEstado = 5, "Cerrado", "Abierto")
              lnCuenta = oRsServicio.Fields!idCuentaAtencion
              lcServicioH = IIf(oRsServicio.Fields!idTipoServicio = 1, "Consulta Externa", IIf(oRsServicio.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia"))
              ldFechaIngreso = IIf(lnFor = 2, oRsServicio.Fields!FechaIngreso, oRsServicio.Fields!fechaEgreso)
              lnTotServicios = ml_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnCuenta)
              lnTotFarmacia = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnCuenta)
              lnPagosAdelantados = ml_ReglasFacturacion.RetornaImporteDePagosAdelantadosPorNroCuenta(lnCuenta, oConexion)
              oRsPacientes.AddNew
              oRsPacientes.Fields!Key = Trim(Str(lnNroHistoriaClinica)) & Trim(Str(lnCuenta))
              oRsPacientes.Fields!nroHistoriaClinica = lnNroHistoriaClinica
              oRsPacientes.Fields!Paciente = lcPaciente
              oRsPacientes.Fields!cuentaNro = lnCuenta
              oRsPacientes.Fields!cuentaEstado = lcEstadoCuenta
              oRsPacientes.Fields!ServicioH = lcServicioH
              oRsPacientes.Fields!servicioF = ldFechaIngreso
              oRsPacientes.Fields!liqFecha = ldFechaLiq
              oRsPacientes.Fields!liqFarmacia = lnTotFarmacia
              oRsPacientes.Fields!liqServicios = lnTotServicios
              oRsPacientes.Fields!liqTotal = lnTotServicios + lnTotFarmacia - lnPagosAdelantados
              oRsPacientes.Fields!PagosAdelantados = lnPagosAdelantados
              oRsPacientes.Update
              oRsServicio.MoveNext
           Loop
        End If
    Next
    'Filtra datos de Servicios (CE)
    Set oRsServicio = ml_ReglasFacturacion.atencionesSeleccionarPorFuenteFinanciamientoFingreso(CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdPlan, ml_IdEstadoCuenta)
    oRsServicio.Filter = "idTipoServicio=1"
    lRecordCount = oRsServicio.RecordCount
    If lRecordCount > 0 Then
       f = 0
       oRsServicio.MoveFirst
       Do While Not oRsServicio.EOF
          f = f + 1
          RaiseEvent ProgressActualizaValor(f, lRecordCount)
          '
          lnNroHistoriaClinica = oRsServicio.Fields!nroHistoriaClinica
          lcPaciente = Trim(oRsServicio.Fields!ApellidoPaterno) & " " & Trim(oRsServicio.Fields!apellidoMaterno) & " " & Trim(oRsServicio.Fields!PrimerNombre)
          lcEstadoCuenta = oRsServicio.Fields!DestadoCta  'IIf(oRsServicio.Fields!IdEstado = 5, "Cerrado", "Abierto")
          lnCuenta = oRsServicio.Fields!idCuentaAtencion
          lcServicioH = IIf(oRsServicio.Fields!idTipoServicio = 1, "Consulta Externa", IIf(oRsServicio.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia"))
          ldFechaIngreso = oRsServicio.Fields!FechaIngreso
          lnTotServicios = ml_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnCuenta)
          lnTotFarmacia = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnCuenta)
          lnPagosAdelantados = ml_ReglasFacturacion.RetornaImporteDePagosAdelantadosPorNroCuenta(lnCuenta, oConexion)
          oRsPacientes.AddNew
          oRsPacientes.Fields!Key = Trim(Str(lnNroHistoriaClinica)) & Trim(Str(lnCuenta))
          oRsPacientes.Fields!nroHistoriaClinica = lnNroHistoriaClinica
          oRsPacientes.Fields!Paciente = lcPaciente
          oRsPacientes.Fields!cuentaNro = lnCuenta
          oRsPacientes.Fields!cuentaEstado = lcEstadoCuenta
          oRsPacientes.Fields!ServicioH = lcServicioH
          oRsPacientes.Fields!servicioF = ldFechaIngreso
          oRsPacientes.Fields!liqFecha = ldFechaLiq
          oRsPacientes.Fields!liqFarmacia = lnTotFarmacia
          oRsPacientes.Fields!liqServicios = lnTotServicios
          oRsPacientes.Fields!liqTotal = lnTotServicios + lnTotFarmacia - lnPagosAdelantados
          oRsPacientes.Fields!PagosAdelantados = lnPagosAdelantados
          oRsPacientes.Update
          oRsServicio.MoveNext
       Loop
    End If
    oRsPacientes.Filter = "liqTotal>0"
    lnNumHistorias = oRsPacientes.RecordCount
    '
    If oRsPacientes.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\EliqSisSoatExoConv.ods"
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
            oRsPacientes.Sort = "paciente,cuentaNro"
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EliqSisSoatExoConv.xls")
            oWorkBookPlantilla.Worksheets("liqSisSoatExoConv").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 1).setFormula(ml_TextoDelFiltro1)
            Call Feuille.getcellbyposition(1, 2).setFormula(ml_TextoDelFiltro)
        Else
            oWorkSheet.Cells(2, 2).Value = ml_TextoDelFiltro1
            oWorkSheet.Cells(3, 2).Value = ml_TextoDelFiltro
        End If
            iFila = 7
            lnNumHistorias = 0
            lnTotServicios = 0
            lnTotFarmacia = 0
            lnTotTotal = 0
            lnTotPagosAdelantados = 0
            oRsPacientes.MoveFirst
            Do While Not oRsPacientes.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(oRsPacientes.Fields("nroHistoriaClinica").Value)
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(oRsPacientes.Fields!Paciente)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(oRsPacientes.Fields!cuentaNro)
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(oRsPacientes.Fields!ServicioH)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(oRsPacientes.Fields!servicioF)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsPacientes!nroHistoriaClinica)), True)
                    oWorkSheet.Cells(iFila, 3).Value = oRsPacientes.Fields!Paciente
                    oWorkSheet.Cells(iFila, 4).Value = oRsPacientes.Fields!cuentaNro
                    oWorkSheet.Cells(iFila, 5).Value = oRsPacientes.Fields!ServicioH
                    oWorkSheet.Cells(iFila, 6).Value = oRsPacientes.Fields!servicioF
                End If
               lnCuenta = oRsPacientes.Fields!cuentaNro
               lnNumHistorias = lnNumHistorias + 1
               Do While Not oRsPacientes.EOF And lnCuenta = oRsPacientes.Fields!cuentaNro
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(oRsPacientes.Fields!liqFarmacia)
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula(oRsPacientes.Fields!liqServicios)
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(oRsPacientes.Fields!PagosAdelantados)
                    Call Feuille.getcellbyposition(9, iFila - 1).setFormula(oRsPacientes.Fields!liqTotal)
                Else
                    oWorkSheet.Cells(iFila, 7).Value = oRsPacientes.Fields!liqFarmacia
                    oWorkSheet.Cells(iFila, 8).Value = oRsPacientes.Fields!liqServicios
                    oWorkSheet.Cells(iFila, 9).Value = oRsPacientes.Fields!PagosAdelantados
                    oWorkSheet.Cells(iFila, 10).Value = oRsPacientes.Fields!liqTotal
                End If
                    lnTotFarmacia = lnTotFarmacia + oRsPacientes.Fields!liqFarmacia
                    lnTotServicios = lnTotServicios + oRsPacientes.Fields!liqServicios
                    lnTotTotal = lnTotTotal + oRsPacientes.Fields!liqTotal
                    lnTotPagosAdelantados = lnTotPagosAdelantados + oRsPacientes.Fields!PagosAdelantados
                    iFila = iFila + 1
                    'Gastos en Servicios
                    'Gastos en Farmacia
                    oRsPacientes.MoveNext
                    If oRsPacientes.EOF Then
                       Exit Do
                    End If
               Loop
            Loop
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":L" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Historias: " & Trim(Str(lnNumHistorias)))
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnTotFarmacia)
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnTotServicios)
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(lnTotPagosAdelantados)
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(lnTotTotal)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 12
                oWorkSheet.Cells(iFila, 2).Value = "Nro Historias: " & Trim(Str(lnNumHistorias))
                oWorkSheet.Cells(iFila, 7).Value = lnTotFarmacia
                oWorkSheet.Cells(iFila, 8).Value = lnTotServicios
                oWorkSheet.Cells(iFila, 9).Value = lnTotPagosAdelantados
                oWorkSheet.Cells(iFila, 10).Value = lnTotTotal
            End If
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 10
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
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
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
    Resume
End Sub




Sub GenerarRecordsetTemporal()
    On Error Resume Next
    If oRsPacientes.State = 1 Then Set oRsPacientes = Nothing
    With oRsPacientes
          .Fields.Append "key", adVarChar, 150
          .Fields.Append "NroHistoriaClinica", adUnsignedBigInt
          .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
          .Fields.Append "CuentaNro", adUnsignedBigInt
          .Fields.Append "CuentaEstado", adVarChar, 50, adFldIsNullable
          .Fields.Append "ServicioH", adVarChar, 50, adFldIsNullable
          .Fields.Append "ServicioF", adDate
          .Fields.Append "LiqFecha", adDate
          .Fields.Append "LiqFarmacia", adDouble
          .Fields.Append "LiqServicios", adDouble
          .Fields.Append "LiqTotal", adDouble
          .Fields.Append "PagosAdelantados", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub








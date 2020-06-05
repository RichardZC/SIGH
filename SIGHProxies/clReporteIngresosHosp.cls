VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clReporteIngrHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Ingresos Hospitalarios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
        
        
        
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdmiServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReporteUtil As New ReporteUtil
'Dim mo_ProgressRpt As XP_ProgressBar
Dim ml_idServicio As Long
Dim ml_IdDepartamento As Long
Dim ml_IdEspecialidad As Long
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_IdTipoNroHistoria As Long      '100=Todos tipos, 200=no debe considear Temporal alojamiento conjunto
Dim ml_IdTipoEspecialidad As Integer
Dim ml_TextoDelFiltro  As String
Dim ml_IdPlan As Long
Dim ml_IdTipoReporte  As Long
Dim ml_idTipoServicio
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_MostrarReporte As Boolean
Property Let mostrarReporte(lIdValue As Boolean)
    mo_MostrarReporte = lIdValue
End Property
Property Let idTipoServicio(lIdValue As Long)
    ml_idTipoServicio = lIdValue
End Property
Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property


Property Let IdPlan(lValue As Long)
    ml_IdPlan = lValue
End Property

Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let IdTipoEspecialidad(lValue As Long)
    ml_IdTipoEspecialidad = lValue
End Property

Property Let IdTipoNroHistoria(lValue As Long)
    ml_IdTipoNroHistoria = lValue
End Property
Property Let IdDepartamento(lValue As Long)
    ml_IdDepartamento = lValue
End Property
Property Let IdEspecialidad(lValue As Long)
    ml_IdEspecialidad = lValue
End Property
Property Let idServicio(lValue As Long)
    ml_idServicio = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property
'Property Set progressRpt(oValue As XP_ProgressBar)
'    Set mo_ProgressRpt = oValue
'End Property


Sub CrearReportePacientesSISconMas180diasEstancia(lnNroDiasEstancia As Long, lnHwnd As Long)
    On Error GoTo errCRJ
    Dim rsReporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oConexionExterna As New Connection
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lbContinuar As Boolean
    If ml_IdPlan = sghFuenteFinanciamiento.sghFFSIS Then
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.CursorLocation = adUseClient
       oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       With rsReporte
          .Fields.Append "nroCuenta", adVarChar, 10, adFldIsNullable
          .Fields.Append "HistoriaClinica", adVarChar, 10, adFldIsNullable
          .Fields.Append "Paciente", adVarChar, 50, adFldIsNullable
          .Fields.Append "FechaIngreso", adDate
          .Fields.Append "horaIngreso", adVarChar, 5, adFldIsNullable
          .Fields.Append "servicioIngreso", adVarChar, 30, adFldIsNullable
          .Fields.Append "Estancia", adInteger
          .LockType = adLockOptimistic
          .Open
        End With
        Set oRsTmp = mo_AdminReportes.ReporteIngresosHospitalariosXfuente(lnNroDiasEstancia, ml_IdPlan, _
                                                                             mda_FechaInicio, mda_FechaFin)
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              lbContinuar = True
              Set oRsTmp1 = mo_AdminReportes.SisFuaAtencionFuasPorCuenta(Val(oRsTmp!nroCuenta), oConexionExterna)
              If oRsTmp1.RecordCount > 0 Then
                 oRsTmp1.MoveFirst
                 If DateDiff("d", oRsTmp1!fechafinal, mda_FechaFin) < lnNroDiasEstancia Then
                    lbContinuar = False
                 End If
              End If
              oRsTmp1.Close
              If lbContinuar = True Then
                    rsReporte.AddNew
                    rsReporte.Fields!nroCuenta = oRsTmp!nroCuenta
                    rsReporte.Fields!HistoriaClinica = HCigualDNI_DevuelveHistoriaConCerosIzquierda(oRsTmp!HistoriaClinica, False)
                    rsReporte.Fields!Paciente = oRsTmp!Paciente
                    rsReporte.Fields!FechaIngreso = oRsTmp!FechaIngreso
                    rsReporte.Fields!HoraIngreso = oRsTmp!HoraIngreso
                    rsReporte.Fields!servicioIngreso = oRsTmp!servicioIngreso
                    rsReporte.Fields!Estancia = oRsTmp!Estancia
                    rsReporte.Update
              End If
              oRsTmp.MoveNext
           Loop
        End If
    Else
        Set rsReporte = mo_AdminReportes.ReporteIngresosHospitalariosXfuente(lnNroDiasEstancia, ml_IdPlan, _
                                                                             mda_FechaInicio, mda_FechaFin)
    End If
    If rsReporte.RecordCount > 0 Then
       mo_AdminReportes.ExportarRecordSetAexcel rsReporte, "PACIENTES SIN ALTA MEDICA QUE PASARON MAS DE " & _
                                                Trim(str(lnNroDiasEstancia)) & " DIAS DE ESTANCIA HOSPITALARIA", _
                                                ml_TextoDelFiltro, "", lnHwnd, True, True
    End If
    Set rsReporte = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oConexionExterna = Nothing
    Set lcBuscaParametro = Nothing
errCRJ:
End Sub

'***************daniel barrantes**************
'***************Ingresos de Pacientes, se incluye tambien EMERGENCIA
'***************
'Sub CrearReporteIngresosHospitalarios(lnHwnd As Long)
'Dim rsReporte As New Recordset
'Dim rsServicio As New Recordset
'Dim iFila As Long
'Dim lnNumHistorias As Long
'Dim lnIdPaciente As Long
'Dim lcPaciente As String
'Dim lnNumTotal As Long
'Dim lcCodServEgreso  As String: Dim lcServEgreso As String: Dim lbContinuar As Boolean
'Dim lcBuscaParametro As New SIGHDatos.Parametros
'Dim lbEsOpenOffice As Boolean
'Dim lcNombre As String, lcSql As String
'
'lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
'On Error GoTo ManejadorErrorExcel
'
'    If lbEsOpenOffice = True Then
'        Dim ServiceManager As Object
'        Dim Desktop As Object
'        Dim Document As Object
'        Dim Feuille As Object
'        Dim Plage As Object
'        Dim args()
'        Dim Chemin As String
'        Dim Fichier As String
'        Dim lcArchivoExcel As String
'        Dim PrintArea(0)
'        Dim Style As Object
'        Dim Border As Object
'        'encabezado
'        Dim PageStyles As Object
'        Dim Sheet As Object
'        Dim StyleFamilies As Object
'        Dim DefPage As Object
'        Dim Htext As Object
'        Dim Hcontent As Object
'        Dim ret As Long
'    Else
'        Dim oExcel As Excel.Application
'        Dim oWorkBookPlantilla As Workbook
'        Dim oWorkBook As Workbook
'        Dim oWorkSheet As Worksheet
'    End If
'    'Filtra los Datos
'    Set rsReporte = mo_AdminReportes.ReporteIngresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
'    If ml_IdPlan > 0 Then
'       rsReporte.Filter = "idFuenteFinanciamiento=" & ml_IdPlan
'    End If
'    lnNumTotal = rsReporte.RecordCount
'    If lnNumTotal = 0 Then
'        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
'    Else
'
'        If lbEsOpenOffice = True Then
'            'Abre el archivo ExcelOpenOffice
'            lcArchivoExcel = App.Path + "\Plantillas\HIngreso_Hosp.ods"
''            FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
''            Chemin = "file:///" & App.Path & "\Plantillas\"
''            Chemin = Replace(Chemin, "\", "/")
''            Fichier = Chemin & "/OpenOffice.ods"
'            '
'            Fichier = Format(Time, "hhmmss") & ".ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
'            lcArchivoExcel = Fichier
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/" & lcArchivoExcel
'            '
'
'            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
'            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
'            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
'            Set Feuille = Document.getSheets().getByIndex(0)
'            'Encabezado de Pagina
'            mo_CabeceraReportes.CabeceraReportes Document, True
'            ' Pone la ventana en primer plano, pasándole el Hwnd
'            ret = SetForegroundWindow(lnHwnd)
'        Else
'
'            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'            'Crea nueva hoja
'            Set oWorkBook = oExcel.Workbooks.Add
'            'Abre, copia y cierra la plantilla
'            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HIngreso_Hosp.xls")
'            oWorkBookPlantilla.Worksheets("RPT_INGRESO_HOSP").Copy Before:=oWorkBook.Sheets(1)
'            oWorkBookPlantilla.Close
'
'            'Activa la primera hoja
'            Set oWorkSheet = oWorkBook.Sheets(1)
'            oWorkSheet.PageSetup.CenterHeader = "Reporte de Ingreso de Epicrisis" & Chr(13) & "Establecimiento: " & lcBuscaParametro.SeleccionaFilaParametro(205)
'            Set oWorkSheet = oWorkBook.Sheets(1)
'            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
'        End If
'            If lbEsOpenOffice = True Then
'                Call Feuille.getcellbyposition(0, 0).setFormula("REPORTE DE INGRESOS     " & ml_TextoDelFiltro)
'            Else
'                oWorkSheet.Cells(1, 1).Value = "REPORTE DE INGRESOS     " & ml_TextoDelFiltro
'            End If
'            iFila = 4
'            lnNumTotal = 0
'            Do While Not rsReporte.EOF
'                lbContinuar = True
'                Select Case ml_IdTipoNroHistoria
'                Case 100    'Todos los Tipos
'                Case 200    'No debe considerar el "Temporal Alojamiento Conjunto"
'                    If rsReporte.Fields("IdTipoNumeracion").Value = 6 Then
'                       lbContinuar = False
'                    End If
'                Case Else   'Solo un Tipo
'                    If rsReporte.Fields("IdTipoNumeracion").Value <> ml_IdTipoNroHistoria Then
'                       lbContinuar = False
'                    End If
'                End Select
'                If lbContinuar Then
'                    lnNumTotal = lnNumTotal + 1
'                    lcCodServEgreso = ""
'                    lcServEgreso = ""
'                    If Not IsNull(rsReporte.Fields("idServicioEgreso").Value) Then
'                        Set rsServicio = ServiciosSelecionarPorId(rsReporte.Fields("idServicioEgreso").Value)
'                        lcCodServEgreso = rsServicio.Fields("codigo").Value
'                        lcServEgreso = rsServicio.Fields("nombre").Value
'                        rsServicio.Close
'                    End If
'                    If lbEsOpenOffice = True Then
'                        Call Feuille.getcellbyposition(0, iFila - 1).setFormula(rsReporte.Fields("nroHistoriaClinica").Value)
'                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("TipoNumeracion").Value)
'                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("ApellidoPaterno").Value)
'                        Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("ApellidoMaterno").Value)
'                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("PrimerNombre").Value)
'                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("SegundoNombre").Value))
'                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte.Fields("Sexo").Value)
'                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(rsReporte.Fields("edad").Value)
'                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields("tedad").Value)
'                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("dpto").Value))
'                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("prov").Value))
'                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("dist").Value))
'                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("CentP").Value))
'                        Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields("FechaIngreso").Value)
'                        Call Feuille.getcellbyposition(16, iFila - 1).setFormula(rsReporte.Fields("HoraIngreso").Value)
'                        Call Feuille.getcellbyposition(17, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("FechaEgreso").Value))
'                        Call Feuille.getcellbyposition(18, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("HoraEgreso").Value))
'                        Call Feuille.getcellbyposition(19, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("CondicionAlta").Value))
'                        Call Feuille.getcellbyposition(20, iFila - 1).setFormula(rsReporte.Fields("CodServIng").Value)
'                        Call Feuille.getcellbyposition(21, iFila - 1).setFormula(rsReporte.Fields("ServicioIngreso").Value)
'                        Call Feuille.getcellbyposition(22, iFila - 1).setFormula(lcCodServEgreso)
'                        Call Feuille.getcellbyposition(23, iFila - 1).setFormula(lcServEgreso)
'                        Call Feuille.getcellbyposition(24, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields("TipoAlta").Value), "", rsReporte.Fields("TipoAlta").Value))
'                        Call Feuille.getcellbyposition(25, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields("DxPrincipal").Value), "", rsReporte.Fields("DxPrincipal").Value))
'                        Call Feuille.getcellbyposition(26, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields("Ddx").Value), "", rsReporte.Fields("Ddx").Value))
'                    Else
'                        oWorkSheet.Cells(iFila, 1).Value = rsReporte.Fields("nroHistoriaClinica").Value
'                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("TipoNumeracion").Value
'                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("ApellidoPaterno").Value
'                        oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("ApellidoMaterno").Value
'                        oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("PrimerNombre").Value
'                        oWorkSheet.Cells(iFila, 6).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("SegundoNombre").Value)
'                        oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields("Sexo").Value
'                        oWorkSheet.Cells(iFila, 8).Value = rsReporte.Fields("edad").Value
'                        oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields("tedad").Value
'                        oWorkSheet.Cells(iFila, 10).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("dpto").Value)
'                        oWorkSheet.Cells(iFila, 11).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("prov").Value)
'                        oWorkSheet.Cells(iFila, 12).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("dist").Value)
'                        oWorkSheet.Cells(iFila, 13).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("CentP").Value)
'                        oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields("FechaIngreso").Value
'                        oWorkSheet.Cells(iFila, 17).Value = Left(rsReporte.Fields("HoraIngreso").Value, 5)
'                        oWorkSheet.Cells(iFila, 18).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("FechaEgreso").Value)
'                        oWorkSheet.Cells(iFila, 19).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("HoraEgreso").Value)
'                        oWorkSheet.Cells(iFila, 20).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("CondicionAlta").Value)
'                        oWorkSheet.Cells(iFila, 21).Value = rsReporte.Fields("CodServIng").Value
'                        oWorkSheet.Cells(iFila, 22).Value = rsReporte.Fields("ServicioIngreso").Value
'                        oWorkSheet.Cells(iFila, 23).Value = lcCodServEgreso
'                        oWorkSheet.Cells(iFila, 24).Value = lcServEgreso
'                        oWorkSheet.Cells(iFila, 25).Value = rsReporte.Fields("TipoAlta").Value
'                        oWorkSheet.Cells(iFila, 26).Value = rsReporte.Fields("DxPrincipal").Value
'                        oWorkSheet.Cells(iFila, 27).Value = rsReporte.Fields("Ddx").Value
'        '                oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("").Value
'                    End If
'                    iFila = iFila + 1
'                    'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1
'                End If
'                rsReporte.MoveNext
'            Loop
'            iFila = iFila + 1
'            If lbEsOpenOffice = True Then
'                Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":V" & CStr(iFila))
'                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
'                Call Feuille.getcellbyposition(0, iFila - 1).setFormula("Nº Historias Clínicas: " + Trim(Str(lnNumTotal)))
'            Else
'                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 27
'                oWorkSheet.Cells(iFila, 1).Value = "Nº Historias Clínicas: " + Trim(Str(lnNumTotal))
'            End If
'            'Falta que salga el nombre del responsable
'            If lbEsOpenOffice = True Then
'                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
'                PrintArea(0).Sheet = 0
'                PrintArea(0).startcolumn = 1
'                PrintArea(0).StartRow = 0
'                PrintArea(0).EndColumn = 27
'                PrintArea(0).EndRow = iFila
'                Call Feuille.SetPrintAreas(PrintArea())
'                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
'                Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
'                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
'            Else
'                If oWorkSheet.PageSetup.PrintArea <> "" Then
'                   oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
'                End If
'                oExcel.Visible = True
'                oWorkSheet.PrintPreview
'                'oWorkSheet.PrintOut
'            End If
'    End If
'    If lbEsOpenOffice = True Then
'        'Liberar Memoria
'        Set Plage = Nothing
'        Set Feuille = Nothing
'        Set Document = Nothing
'        Set Desktop = Nothing
'        Set ServiceManager = Nothing
'        Set Style = Nothing
'        Set Border = Nothing
'        'encabezado de pagina
'        Set PageStyles = Nothing
'        Set Sheet = Nothing
'        Set StyleFamilies = Nothing
'        Set DefPage = Nothing
'        Set Htext = Nothing
'        Set Hcontent = Nothing
'    Else
'        'Liberar memoria
'        Set oExcel = Nothing
'        Set oWorkBookPlantilla = Nothing
'        Set oWorkBook = Nothing
'        Set oWorkSheet = Nothing
'    End If
'Exit Sub
'ManejadorErrorExcel:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'    End Select
'    Exit Sub
'End Sub

'debb-15/04/2016
Sub CrearReporteIngresosHospitalarios(lnHwnd As Long, lnConsumoActualFarmaciaServicios As Double, _
                                      lcHoraInicio As String, lcHoraFin As String)
Dim rsReporte As New Recordset
Dim rsServicio As New Recordset
Dim rsGravedad As New Recordset
Dim rsTmpCuentas As New Recordset
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim oConexion As New Connection
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lcCodServEgreso  As String: Dim lcServEgreso As String: Dim lbContinuar As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnConsumoFarmServicioCta As Double, lcGravedad As String, lcNroEmergencia As String
Dim lcUsuario As String, ldUsuarioFecha As String, lbPrimeraCuenta As Boolean, lcFuenteFinanciamiento As String

oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighentidades.CadenaConexion

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
    'Filtra los Datos
    Set rsReporte = mo_AdminReportes.ReporteIngresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, _
                    mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
                    
                    
                    
    If ml_IdPlan > 0 Then
       rsReporte.Filter = "idFuenteFinanciamiento=" & ml_IdPlan
    End If
    lnNumTotal = rsReporte.RecordCount
    If lnNumTotal = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
    
       With rsTmpCuentas
          .Fields.Append "nroCuenta", adInteger
          .LockType = adLockOptimistic
          .Open
        End With
    
        Set rsGravedad = mo_AdminServComunes.TipoGravedadAtencionSeleccionarTodos()
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HIngreso_Hosp.ods"
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
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HIngreso_Hosp.xls")
            oWorkBookPlantilla.Worksheets("RPT_INGRESO_HOSP").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close

            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            oWorkSheet.PageSetup.CenterHeader = "Reporte de Ingreso de Epicrisis" & Chr(13) & "Establecimiento: " & lcBuscaParametro.SeleccionaFilaParametro(205)
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, 0).setFormula("REPORTE DE INGRESOS     " & ml_TextoDelFiltro)
                Call Feuille.getcellbyposition(26, 2).Width(50)
                Call Feuille.getcellbyposition(27, 2).setFormula("Nro Cuenta")
                Call Feuille.getcellbyposition(28, 2).setFormula("Consumo S/.")
                Call Feuille.getcellbyposition(29, 2).setFormula("Usuario Digitación")
                Call Feuille.getcellbyposition(30, 2).setFormula("Fecha Digitación")
                If ml_IdTipoEspecialidad = 2 Then
                   Call Feuille.getcellbyposition(31, 2).setFormula("Gravedad")
                   Call Feuille.getcellbyposition(32, 2).setFormula("Emergencia N°")
                End If
            Else
                oWorkSheet.Cells(1, 1).Value = "REPORTE DE INGRESOS     " & ml_TextoDelFiltro
                oWorkSheet.Cells(3, 28).Value = "Nro Cuenta"
                oWorkSheet.Cells(3, 29).Value = "Consumo S/."
                oWorkSheet.Cells(3, 30).Value = "Usuario Digitación"
                oWorkSheet.Cells(3, 31).Value = "Fecha Digitación"
                If ml_IdTipoEspecialidad = 2 Then
                   oWorkSheet.Cells(3, 32).Value = "Gravedad"
                   oWorkSheet.Cells(3, 33) = "Emergencia N°"
                End If
                
            End If
            If lnConsumoActualFarmaciaServicios > 0 Then
               If lbEsOpenOffice = True Then

                  Call Feuille.getcellbyposition(28, 2).setFormula("Consumo S/.")
               Else
                  oWorkSheet.Cells(3, 27).ColumnWidth = 50

               End If
            End If
            iFila = 4
            lnNumTotal = 0
            '
            lbPrimeraCuenta = True
            Do While Not rsReporte.EOF
                lbContinuar = True
                Select Case ml_IdTipoNroHistoria
                Case 100    'Todos los Tipos
                Case 200    'No debe considerar el "Temporal Alojamiento Conjunto"
                    If rsReporte.Fields("IdTipoNumeracion").Value = 6 Then
                       lbContinuar = False
                    End If
                Case Else   'Solo un Tipo
                    If rsReporte.Fields("IdTipoNumeracion").Value <> ml_IdTipoNroHistoria Then
                       lbContinuar = False
                    End If
                End Select
                '
                lnConsumoFarmServicioCta = 0
                If lbContinuar = True And lnConsumoActualFarmaciaServicios > 0 Then
                   If IsNull(rsReporte!FechaEgreso) Then
                        lnConsumoFarmServicioCta = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(rsReporte!idCuentaAtencion, oConexion, True)
                        If lnConsumoFarmServicioCta < lnConsumoActualFarmaciaServicios Then
                           lbContinuar = False
                        End If
                   Else
                        lbContinuar = False
                   End If
                End If
                '
                If lbContinuar = True Then
                    If lbPrimeraCuenta = True Then
                        lbPrimeraCuenta = False
                        rsTmpCuentas.AddNew
                        rsTmpCuentas.Fields!nroCuenta = rsReporte!idCuentaAtencion
                        rsTmpCuentas.Update
                    Else
                        rsTmpCuentas.MoveFirst
                        rsTmpCuentas.Find "nroCuenta=" & rsReporte!idCuentaAtencion
                        If rsTmpCuentas.EOF Then
                          rsTmpCuentas.AddNew
                          rsTmpCuentas.Fields!nroCuenta = rsReporte!idCuentaAtencion
                          rsTmpCuentas.Update
                        Else
                          lbContinuar = False
                        End If
                    End If
                End If
                '
                If lbContinuar = True And _
                   Not (CDate(Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte!HoraIngreso) >= CDate(Format(mda_FechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & lcHoraInicio) And _
                   CDate(Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte!HoraIngreso) <= CDate(Format(mda_FechaFin, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & lcHoraFin)) Then
                   lbContinuar = False
                End If
                '
                If lbContinuar Then
                    
                    lnNumTotal = lnNumTotal + 1
                    lcCodServEgreso = ""
                    lcServEgreso = ""
                    If Not IsNull(rsReporte.Fields("idServicioEgreso").Value) Then
                        Set rsServicio = ServiciosSelecionarPorId(rsReporte.Fields("idServicioEgreso").Value)
                        lcCodServEgreso = rsServicio.Fields("codigo").Value
                        lcServEgreso = rsServicio.Fields("nombre").Value
                        rsServicio.Close
                    End If
                    '
                    If ml_IdTipoEspecialidad = 2 Then
                       lcGravedad = ""
                       If Not IsNull(rsReporte!idTipoGravedad) Then
                          rsGravedad.MoveFirst
                          rsGravedad.Find "idTipoGravedad=" & rsReporte.Fields!idTipoGravedad
                          If Not rsGravedad.EOF Then
                             lcGravedad = rsGravedad!Descripcion
                          End If
                       End If
                       lcNroEmergencia = ""   'debb-13/07/2016
                       Set rsServicio = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarXidAtencion(rsReporte!idAtencion)
                       If rsServicio.RecordCount > 0 Then
                          lcNroEmergencia = IIf(IsNull(rsServicio!emergenciaCorrelativo), "", rsServicio!emergenciaCorrelativo)
                       End If
                       rsServicio.Close
                    End If
                    '
                    lcUsuario = "": ldUsuarioFecha = ""
                    Set rsServicio = mo_ReglasDeSeguridad.AuditoriaFiltrarCitasPorIdAtencion(rsReporte!idAtencion, _
                                           IIf(ml_IdTipoEspecialidad = 2, sghAdmisionEmergencia, sghAdmisionHospitalizacion), _
                                           oConexion)
                    If rsServicio.RecordCount > 0 Then
                       lcUsuario = IIf(IsNull(rsServicio!dUsuario), "", rsServicio!dUsuario)
                       ldUsuarioFecha = IIf(IsNull(rsServicio!fechaHora), "", Format(rsServicio!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
                    End If
                    rsServicio.Close
                    '
                    lcFuenteFinanciamiento = ""
                    Set rsServicio = mo_AdminServComunes.FuentesFinanciamientoSegunFiltro("idFuenteFinanciamiento=" & rsReporte!idFuenteFinanciamiento)
                    If rsServicio.RecordCount > 0 Then
                       lcFuenteFinanciamiento = Trim(rsServicio!Descripcion)
                    End If
                    rsServicio.Close
                    
                    '
                    
                    
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(0, iFila - 1).setFormula(rsReporte.Fields("nroHistoriaClinica").Value)
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("TipoNumeracion").Value)
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("ApellidoPaterno").Value)
                        Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("ApellidoMaterno").Value)
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("PrimerNombre").Value)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("SegundoNombre").Value))
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte.Fields("Sexo").Value)
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(rsReporte.Fields("edad").Value)
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields("tedad").Value)
                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("dpto").Value))
                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("prov").Value))
                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("dist").Value))
                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("CentP").Value))
                        Call Feuille.getcellbyposition(15, iFila - 1).setFormula(rsReporte.Fields("FechaIngreso").Value)
                        Call Feuille.getcellbyposition(16, iFila - 1).setFormula(rsReporte.Fields("HoraIngreso").Value)
                        Call Feuille.getcellbyposition(17, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("FechaEgreso").Value))
                        Call Feuille.getcellbyposition(18, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("HoraEgreso").Value))
                        Call Feuille.getcellbyposition(19, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte.Fields("CondicionAlta").Value))
                        Call Feuille.getcellbyposition(20, iFila - 1).setFormula(rsReporte.Fields("CodServIng").Value)
                        Call Feuille.getcellbyposition(21, iFila - 1).setFormula(rsReporte.Fields("ServicioIngreso").Value)
                        Call Feuille.getcellbyposition(22, iFila - 1).setFormula(lcCodServEgreso)
                        Call Feuille.getcellbyposition(23, iFila - 1).setFormula(lcServEgreso)
                        Call Feuille.getcellbyposition(24, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields("TipoAlta").Value), "", rsReporte.Fields("TipoAlta").Value))
                        Call Feuille.getcellbyposition(25, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields("DxPrincipal").Value), "", rsReporte.Fields("DxPrincipal").Value))
                        Call Feuille.getcellbyposition(26, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields("Ddx").Value), "", rsReporte.Fields("Ddx").Value))
                        Call Feuille.getcellbyposition(27, iFila - 1).setFormula(rsReporte!idCuentaAtencion)
                        If lnConsumoFarmServicioCta > 0 Then
                            Call Feuille.getcellbyposition(28, iFila - 1).setFormula(lnConsumoFarmServicioCta)
                        End If
                        Call Feuille.getcellbyposition(29, iFila - 1).setFormula(lcUsuario)
                        Call Feuille.getcellbyposition(30, iFila - 1).setFormula(ldUsuarioFecha)
                        If ml_IdTipoEspecialidad = 2 Then
                           Call Feuille.getcellbyposition(31, iFila - 1).setFormula(lcGravedad)
                           Call Feuille.getcellbyposition(32, iFila - 1).setFormula(lcNroEmergencia)
                        End If
                    Else
                        oWorkSheet.Cells(iFila, 1).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(rsReporte.Fields("nroHistoriaClinica").Value, True)
                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("TipoNumeracion").Value
                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("ApellidoPaterno").Value
                        oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("ApellidoMaterno").Value
                        oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("PrimerNombre").Value
                        oWorkSheet.Cells(iFila, 6).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("SegundoNombre").Value)
                        oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields("Sexo").Value
                        oWorkSheet.Cells(iFila, 8).Value = rsReporte.Fields("edad").Value
                        oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields("tedad").Value
                        oWorkSheet.Cells(iFila, 10).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("dpto").Value)
                        oWorkSheet.Cells(iFila, 11).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("prov").Value)
                        oWorkSheet.Cells(iFila, 12).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("dist").Value)
                        oWorkSheet.Cells(iFila, 13).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("CentP").Value)
                        oWorkSheet.Cells(iFila, 15).Value = lcFuenteFinanciamiento
                        'SCCQ 29/05/2020 Cambio19 Inicio
                        oWorkSheet.Cells(iFila, 16).Value = Format(rsReporte.Fields("FechaIngreso").Value, sighentidades.DevuelveFechaSoloFormato_DMY)
                        'SCCQ 29/05/2020 Cambio19 Fin
                        oWorkSheet.Cells(iFila, 17).Value = Left(rsReporte.Fields("HoraIngreso").Value, 5)
                        'SCCQ 29/05/2020 Cambio19 Inicio
                        oWorkSheet.Cells(iFila, 18).Value = mo_ReporteUtil.NullToVacio(Format(rsReporte.Fields("FechaEgreso").Value, sighentidades.DevuelveFechaSoloFormato_DMY))
                        'SCCQ 29/05/2020 Cambio19 Fin
                        oWorkSheet.Cells(iFila, 19).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("HoraEgreso").Value)
                        oWorkSheet.Cells(iFila, 20).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("CondicionAlta").Value)
                        oWorkSheet.Cells(iFila, 21).Value = rsReporte.Fields("CodServIng").Value
                        oWorkSheet.Cells(iFila, 22).Value = rsReporte.Fields("ServicioIngreso").Value
                        oWorkSheet.Cells(iFila, 23).Value = lcCodServEgreso
                        oWorkSheet.Cells(iFila, 24).Value = lcServEgreso
                        oWorkSheet.Cells(iFila, 25).Value = rsReporte.Fields("TipoAlta").Value
                        oWorkSheet.Cells(iFila, 26).Value = rsReporte.Fields("DxPrincipal").Value
                        oWorkSheet.Cells(iFila, 27).Value = rsReporte.Fields("Ddx").Value
                        oWorkSheet.Cells(iFila, 28).Value = rsReporte!idCuentaAtencion
                        If lnConsumoFarmServicioCta > 0 Then
                            oWorkSheet.Cells(iFila, 29).Value = lnConsumoFarmServicioCta
                        End If
                        oWorkSheet.Cells(iFila, 30).Value = lcUsuario
                        oWorkSheet.Cells(iFila, 31).Value = ldUsuarioFecha
                        If ml_IdTipoEspecialidad = 2 Then
                           oWorkSheet.Cells(iFila, 32).Value = lcGravedad
                           oWorkSheet.Cells(iFila, 33).Value = lcNroEmergencia
                        End If
                    End If
                    iFila = iFila + 1
                    'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1
                End If
                rsReporte.MoveNext
            Loop
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":V" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula("Nº Historias Clínicas: " + Trim(str(lnNumTotal)))
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 27
                oWorkSheet.Cells(iFila, 1).Value = "Nº Historias Clínicas: " + Trim(str(lnNumTotal))
            End If
            'Falta que salga el nombre del responsable
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 27
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
                'oWorkSheet.PrintOut
            End If
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
    
oConexion.Close
Set oConexion = Nothing
Set rsReporte = Nothing
Set rsServicio = Nothing
Set mo_ReglasFacturacion = Nothing
Set lcBuscaParametro = Nothing
Set rsGravedad = Nothing
Set mo_ReglasAdmision = Nothing
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






Function ServiciosSelecionarPorId(lnServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ServiciosSelecionarPorId = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosSelecionarPorId = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub EjecutaFormulario()
    Dim oFormulario As New ReporteIngresosHosp
    oFormulario.IdTipoReporte = ml_IdTipoReporte
    oFormulario.cmbConsiderar.ListIndex = ml_idTipoServicio
    oFormulario.mostrarReporte = mo_MostrarReporte
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

Sub CrearReporteIngresosEmergencia(lnHwnd As Long, lnConsumoActualFarmaciaServicios As Double, _
                                      lcHoraInicio As String, lcHoraFin As String)
Dim rsReporte As New Recordset
Dim rsServicio As New Recordset
Dim rsGravedad As New Recordset
Dim rsTmpCuentas As New Recordset
Dim rsProcedimientos As New Recordset
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim oConexion As New Connection
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lcCodServEgreso  As String: Dim lcServEgreso As String: Dim lbContinuar As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnConsumoFarmServicioCta As Double, lcGravedad As String, lcNroEmergencia As String
Dim lcUsuario As String, ldUsuarioFecha As String, lbPrimeraCuenta As Boolean
Dim lcDireccionDomicilio As String, lcAcompaniante As String
Dim lcEstadoLlego As String, lcFormaLlego As String, lcFuenteFinanciamiento As String

oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighentidades.CadenaConexion

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

    If lbEsOpenOffice = True Then
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    'Filtra los Datos
    Set rsReporte = mo_AdminReportes.ReporteIngresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, _
                    mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    If ml_IdPlan > 0 Then
       rsReporte.Filter = "idFuenteFinanciamiento=" & ml_IdPlan
    End If
    lnNumTotal = rsReporte.RecordCount
    If lnNumTotal = 0 Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
    
           With rsTmpCuentas
              .Fields.Append "nroCuenta", adInteger
              .LockType = adLockOptimistic
              .Open
            End With
        
            Set rsGravedad = mo_AdminServComunes.TipoGravedadAtencionSeleccionarTodos()
            If lbEsOpenOffice = True Then
            Else
        
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HIngreso_Emerg.xls")
                oWorkBookPlantilla.Worksheets("RPT_INGRESO_EMERG").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
    
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                oWorkSheet.PageSetup.CenterHeader = "Establecimiento: " & lcBuscaParametro.SeleccionaFilaParametro(205)
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            oWorkSheet.Cells(1, 1).Value = ml_TextoDelFiltro
            iFila = 5
            lnNumTotal = 0
            '
            lbPrimeraCuenta = True
            Do While Not rsReporte.EOF
                lbContinuar = True
                Select Case ml_IdTipoNroHistoria
                Case 100    'Todos los Tipos
                Case 200    'No debe considerar el "Temporal Alojamiento Conjunto"
                    If rsReporte.Fields("IdTipoNumeracion").Value = 6 Then
                       lbContinuar = False
                    End If
                Case Else   'Solo un Tipo
                    If rsReporte.Fields("IdTipoNumeracion").Value <> ml_IdTipoNroHistoria Then
                       lbContinuar = False
                    End If
                End Select
                '
                lnConsumoFarmServicioCta = 0
                If lbContinuar = True And lnConsumoActualFarmaciaServicios > 0 Then
                   If IsNull(rsReporte!FechaEgreso) Then
                        lnConsumoFarmServicioCta = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(rsReporte!idCuentaAtencion, oConexion, True)
                        If lnConsumoFarmServicioCta < lnConsumoActualFarmaciaServicios Then
                           lbContinuar = False
                        End If
                   Else
                        lbContinuar = False
                   End If
                End If
                '
                If lbContinuar = True Then
                    If lbPrimeraCuenta = True Then
                        lbPrimeraCuenta = False
                        rsTmpCuentas.AddNew
                        rsTmpCuentas.Fields!nroCuenta = rsReporte!idCuentaAtencion
                        rsTmpCuentas.Update
                    Else
                        rsTmpCuentas.MoveFirst
                        rsTmpCuentas.Find "nroCuenta=" & rsReporte!idCuentaAtencion
                        If rsTmpCuentas.EOF Then
                          rsTmpCuentas.AddNew
                          rsTmpCuentas.Fields!nroCuenta = rsReporte!idCuentaAtencion
                          rsTmpCuentas.Update
                        Else
                          lbContinuar = False
                        End If
                    End If
                End If
                '
                If lbContinuar = True And _
                   Not (CDate(Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte!HoraIngreso) >= CDate(Format(mda_FechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & lcHoraInicio) And _
                   CDate(Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & rsReporte!HoraIngreso) <= CDate(Format(mda_FechaFin, sighentidades.DevuelveFechaSoloFormato_DMY) & " " & lcHoraFin)) Then
                   lbContinuar = False
                End If
                '
                If lbContinuar Then
                    
                    lnNumTotal = lnNumTotal + 1
                    lcCodServEgreso = ""
                    lcServEgreso = ""
                    If Not IsNull(rsReporte.Fields("idServicioEgreso").Value) Then
                        Set rsServicio = ServiciosSelecionarPorId(rsReporte.Fields("idServicioEgreso").Value)
                        lcCodServEgreso = rsServicio.Fields("codigo").Value
                        lcServEgreso = rsServicio.Fields("nombre").Value
                        rsServicio.Close
                    End If
                    '
                    lcAcompaniante = ""
                    lcDireccionDomicilio = ""
                    If ml_IdTipoEspecialidad = 2 Then
                       lcGravedad = ""
                       If Not IsNull(rsReporte!idTipoGravedad) Then
                          rsGravedad.MoveFirst
                          rsGravedad.Find "idTipoGravedad=" & rsReporte.Fields!idTipoGravedad
                          If Not rsGravedad.EOF Then
                             lcGravedad = rsGravedad!Descripcion
                          End If
                       End If
                       lcNroEmergencia = ""   'debb-13/07/2016
                       Set rsServicio = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarXidAtencion(rsReporte!idAtencion)
                       If rsServicio.RecordCount > 0 Then
                          lcNroEmergencia = IIf(IsNull(rsServicio!emergenciaCorrelativo), "", rsServicio!emergenciaCorrelativo)
                          lcDireccionDomicilio = IIf(IsNull(rsServicio!DireccionDomicilio), "", rsServicio!DireccionDomicilio)
                          lcAcompaniante = IIf(IsNull(rsServicio!NombreAcompaniante), "", rsServicio!NombreAcompaniante)
                       End If
                       rsServicio.Close
                    End If
                    '
                    lcUsuario = "": ldUsuarioFecha = ""
                    Set rsServicio = mo_ReglasDeSeguridad.AuditoriaFiltrarCitasPorIdAtencion(rsReporte!idAtencion, _
                                           IIf(ml_IdTipoEspecialidad = 2, sghAdmisionEmergencia, sghAdmisionHospitalizacion), _
                                           oConexion)
                    If rsServicio.RecordCount > 0 Then
                       lcUsuario = IIf(IsNull(rsServicio!dUsuario), "", rsServicio!dUsuario)
                       ldUsuarioFecha = IIf(IsNull(rsServicio!fechaHora), "", Format(rsServicio!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
                    End If
                    rsServicio.Close
                    '
                    lcFuenteFinanciamiento = ""
                    Set rsServicio = mo_AdminServComunes.FuentesFinanciamientoSegunFiltro("idFuenteFinanciamiento=" & rsReporte!idFuenteFinanciamiento)
                    If rsServicio.RecordCount > 0 Then
                       lcFuenteFinanciamiento = Trim(rsServicio!Descripcion)
                    End If
                    rsServicio.Close
                    
                    '
                    lcEstadoLlego = "": lcFormaLlego = ""
                    Set rsProcedimientos = mo_ReglasAdmision.AtencionesEmergenciaXidAtencion(rsReporte!idAtencion, oConexion)
                    If rsProcedimientos.RecordCount > 0 Then
                       lcFormaLlego = IIf(rsProcedimientos!comoLlego = 1, "CAMINANDO", _
                                     IIf(rsProcedimientos!comoLlego = 2, "SILLA DE RUEDAS", _
                                     IIf(rsProcedimientos!comoLlego = 3, "CAMILLA", "OTROS")))
                       lcEstadoLlego = IIf(rsProcedimientos!idEstadoLlegada = 1, "DESPIERTO", _
                                       IIf(rsProcedimientos!idEstadoLlegada = 2, "ETILICO", "INCONSCIENTE"))
                    End If
                    rsProcedimientos.Close
                    '
                    If lbEsOpenOffice = True Then
                    Else
                        oWorkSheet.Cells(iFila, 1).Value = lcNroEmergencia
                        oWorkSheet.Cells(iFila, 2).Value = Day(rsReporte.Fields("FechaIngreso").Value)
                        oWorkSheet.Cells(iFila, 3).Value = Month(rsReporte.Fields("FechaIngreso").Value)
                        oWorkSheet.Cells(iFila, 4).Value = Year(rsReporte.Fields("FechaIngreso").Value)
                        oWorkSheet.Cells(iFila, 5).Value = rsReporte!HoraIngreso
                        oWorkSheet.Cells(iFila, 6).Value = Trim(rsReporte!ApellidoPaterno) & " " & Trim(rsReporte!ApellidoMaterno) & " " & Trim(rsReporte!PrimerNombre) & " " & mo_ReporteUtil.NullToVacio(rsReporte.Fields("SegundoNombre").Value)
                        oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields("edad").Value & " " & rsReporte.Fields("tedad").Value
                        oWorkSheet.Cells(iFila, 8).Value = IIf(Left(rsReporte.Fields("Sexo").Value, 1) = "M", "X", " ")
                        oWorkSheet.Cells(iFila, 9).Value = IIf(Left(rsReporte.Fields("Sexo").Value, 1) <> "M", "X", " ")
                        oWorkSheet.Cells(iFila, 10).Value = lcDireccionDomicilio
                        oWorkSheet.Cells(iFila, 11).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("dist").Value)
                        oWorkSheet.Cells(iFila, 12).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("prov").Value)
                        oWorkSheet.Cells(iFila, 13).Value = lcEstadoLlego
                        oWorkSheet.Cells(iFila, 14).Value = lcFormaLlego
                        'oWorkSheet.Cells(iFila, 15).Value = lcAcompaniante
                        oWorkSheet.Cells(iFila, 15).Value = lcFuenteFinanciamiento
                        oWorkSheet.Cells(iFila, 20).Value = mo_ReporteUtil.NullToVacio(rsReporte!DxPrincipal) & " " & mo_ReporteUtil.NullToVacio(rsReporte.Fields("Ddx").Value)
                        oWorkSheet.Cells(iFila, 24).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("FechaEgreso").Value)
                        oWorkSheet.Cells(iFila, 25).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("HoraEgreso").Value)
                        Set rsProcedimientos = mo_AdminServComunes.AtencionesSeleccionarMedicoPorCuenta(rsReporte!idCuentaAtencion)
                        If rsProcedimientos.RecordCount > 0 Then
                            If lbEsOpenOffice = True Then
                            Else
                                oWorkSheet.Cells(iFila, 26).Value = IIf(IsNull(rsProcedimientos!dmedico), "", rsProcedimientos!dmedico)
                            End If
                        End If
                        rsProcedimientos.Close
                        oWorkSheet.Cells(iFila, 29).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte!NroHistoriaClinica)), True)
                        oWorkSheet.Cells(iFila, 30).Value = "'" & rsReporte!idCuentaAtencion
                    End If
                    iFila = iFila + 1
                    'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1
                End If
                rsReporte.MoveNext
            Loop
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 27
                oWorkSheet.Cells(iFila, 1).Value = "Nº Historias Clínicas: " + Trim(str(lnNumTotal))
            End If
            'Falta que salga el nombre del responsable
            If lbEsOpenOffice = True Then
            Else
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
                'oWorkSheet.PrintOut
            End If
    End If
    If lbEsOpenOffice = True Then
    Else
        'Liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    
oConexion.Close
Set oConexion = Nothing
Set rsProcedimientos = Nothing
Set rsReporte = Nothing
Set rsServicio = Nothing
Set mo_ReglasFacturacion = Nothing
Set lcBuscaParametro = Nothing
Set rsGravedad = Nothing
Set mo_ReglasAdmision = Nothing
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





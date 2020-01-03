VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptIngresosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdmiServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_ProgressRpt As XP_ProgressBar
Dim ml_idServicio As Long
Dim ml_IdDepartamento As Long
Dim ml_IdEspecialidad As Long
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_IdTipoNroHistoria As Long      '100=Todos tipos, 200=no debe considear Temporal alojamiento conjunto
Dim ml_IdTipoEspecialidad As Integer
Dim ml_TextoDelFiltro  As String
Dim ml_IdPlan As Long

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
Property Set progressRpt(oValue As XP_ProgressBar)
    Set mo_ProgressRpt = oValue
End Property




'***************daniel barrantes**************
'***************Ingresos de Pacientes, se incluye tambien EMERGENCIA
'***************
Sub CrearReporteIngresosHospitalarios()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim rsServicio As New Recordset
Dim iFila As Integer
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lcCodServEgreso  As String: Dim lcServEgreso As String: Dim lbContinuar As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
On Error GoTo ManejadorErrorExcel
    'Filtra los Datos
    Set rsReporte = mo_AdminReportes.ReporteIngresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    If ml_IdPlan > 0 Then
       rsReporte.Filter = "idFuenteFinanciamiento=" & ml_IdPlan
    End If
    lnNumTotal = rsReporte.RecordCount
    If lnNumTotal = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HIngreso_Hosp.xls")
            oWorkBookPlantilla.Worksheets("RPT_INGRESO_HOSP").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            
            mo_ProgressRpt.Min = 0
            mo_ProgressRpt.Max = lnNumTotal
            mo_ProgressRpt.Value = 0
            mo_ProgressRpt.ShowText = True
            mo_ProgressRpt.Color = vbGreen
            
            
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            oWorkSheet.PageSetup.CenterHeader = "Reporte de Ingreso de Epicrisis" & Chr(13) & "Establecimiento: " & lcBuscaParametro.SeleccionaFilaParametro(205)
            oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
            oWorkSheet.Cells(1, 1).Value = ml_TextoDelFiltro
            iFila = 4
            lnNumTotal = 0
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
                    oWorkSheet.Cells(iFila, 1).Value = rsReporte.Fields("nroHistoriaClinica").Value
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
                    oWorkSheet.Cells(iFila, 16).Value = rsReporte.Fields("FechaIngreso").Value
                    oWorkSheet.Cells(iFila, 17).Value = rsReporte.Fields("HoraIngreso").Value
                    oWorkSheet.Cells(iFila, 18).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("FechaEgreso").Value)
                    oWorkSheet.Cells(iFila, 19).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("HoraEgreso").Value)
                    oWorkSheet.Cells(iFila, 20).Value = mo_ReporteUtil.NullToVacio(rsReporte.Fields("CondicionAlta").Value)
                    oWorkSheet.Cells(iFila, 21).Value = rsReporte.Fields("CodServIng").Value
                    oWorkSheet.Cells(iFila, 22).Value = rsReporte.Fields("ServicioIngreso").Value
                    oWorkSheet.Cells(iFila, 23).Value = lcCodServEgreso
                    oWorkSheet.Cells(iFila, 24).Value = lcServEgreso
                    oWorkSheet.Cells(iFila, 25).Value = rsReporte.Fields("TipoAlta").Value
                    oWorkSheet.Cells(iFila, 26).Value = rsReporte.Fields("DxPrincipal").Value
                    oWorkSheet.Cells(iFila, 27).Value = rsReporte.Fields("Ddx").Value
    '                oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("").Value
                    iFila = iFila + 1
                    mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1
                End If
                rsReporte.MoveNext
            Loop
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 22
            oWorkSheet.Cells(iFila, 1).Value = "Nº Historias Clínicas: " + Trim(Str(lnNumTotal))
            'Falta que salga el nombre del responsable
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If

            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
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
    oConexion.Open SIGHEntidades.CadenaConexion
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


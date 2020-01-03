VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptMovimientoHistorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ml_IdGrupoMovimiento As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdmiServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReporteUtil As New ReporteUtil
'Dim mo_ProgressRpt As XP_ProgressBar

Property Let IdGrupoMovimiento(lValue As Long)
    ml_IdGrupoMovimiento = lValue
End Property
'Property Set progressRpt(oValue As XP_ProgressBar)
'    Set mo_ProgressRpt = oValue
'End Property

Sub CrearReporteMovimientoHistoria()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim oRange As Range
Dim iFila As Integer
Dim lIdServicio As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lIdServicioDestino As Long
On Error GoTo ManejadorError


    'Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    Set oExcel = New Excel.Application
    
    'Crea nueva hoja
    Set oWorkBook = oExcel.Workbooks.Add
    
    'Abre, copia y cierra la plantilla
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCMovimientoHistorias.xls")
    oWorkBookPlantilla.Worksheets("MovimientoHistorias").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    
    'Activa la primera hoja
    Set oWorkSheet = oWorkBook.Sheets(1)
    'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
       
    'Crea la cabecera
    Set rsReporte = mo_AdminReportes.ObtenerCabeceraDeMovimiento(ml_IdGrupoMovimiento)
    
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(rsReporte!IdEmpleadoArchivo)
    If Not oDOEmpleado Is Nothing Then
        oWorkSheet.Cells(5, 5).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(rsReporte!IdEmpleadoTransporte)
    If Not oDOEmpleado Is Nothing Then
        oWorkSheet.Cells(6, 5).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(rsReporte!IdEmpleadoRecepcion)
    If Not oDOEmpleado Is Nothing Then
        oWorkSheet.Cells(7, 5).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    
    'Fecha y hora del movimiento
    oWorkSheet.Cells(4, 15).Value = "'" & Format(rsReporte!FechaMovimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oWorkSheet.Cells(5, 15).Value = "'" & Format(rsReporte!FechaMovimiento, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    
    'Fecha hora de reporte
    oWorkSheet.Cells(1, 14).Value = "'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oWorkSheet.Cells(1, 17).Value = "'" & Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    
    
    Set rsReporte = mo_AdminReportes.ObtenerHistoriasClinicasDeEsteMovimiento(ml_IdGrupoMovimiento)
    oWorkSheet.Cells(6, 15).Value = "'" & rsReporte.Fields!motivo
    
    iFila = 10
    lIdServicioDestino = 0
'    mo_ProgressRpt.Min = 0
'    mo_ProgressRpt.Max = rsReporte.RecordCount
'    mo_ProgressRpt.Value = 0
    
    Do While Not rsReporte.EOF
        'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        
        oWorkSheet.Cells(iFila, 1).Value = "'" & IIf(lIdServicioDestino <> rsReporte!idServicioDestino, rsReporte!ServicioDestino, "")
        oWorkSheet.Cells(iFila, 6).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroHistoriaClinica)
        oWorkSheet.Cells(iFila, 8).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
        oWorkSheet.Cells(iFila, 14).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ServicioOrigen)
        oWorkSheet.Cells(iFila, 19).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroFolios)
        oWorkSheet.Cells(iFila, 21).Value = "'" & Format(mo_ReporteUtil.NullToVacio(rsReporte!FechaRequerida), SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        
        Set oRange = oExcel.Range("A" & iFila & ":" & "E" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
        Set oRange = oExcel.Range("F" & iFila & ":" & "G" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
        Set oRange = oExcel.Range("H" & iFila & ":" & "M" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
        Set oRange = oExcel.Range("N" & iFila & ":" & "R" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
        Set oRange = oExcel.Range("S" & iFila & ":" & "T" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
        Set oRange = oExcel.Range("U" & iFila & ":" & "V" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
        
        lIdServicioDestino = rsReporte!idServicioDestino
        
        iFila = iFila + 1
        rsReporte.MoveNext
    Loop
    
    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 10, 1, iFila - 1, 21
    
    'Falta que salga el nombre del responsable
    If oWorkSheet.PageSetup.PrintArea <> "" Then
       oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
    End If
    oExcel.Visible = True
    oWorkSheet.PrintPreview
    oWorkSheet.PrintOut
    'Set oExcel = Nothing
    
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub


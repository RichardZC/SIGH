VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCuposAsignadosPorServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ml_IdEmpleado As Long
Dim mda_FechaRequeridaDesde As Date
Dim mda_FechaRequeridaHasta As Date
Dim mda_FechaSolicitudDesde As Date
Dim mda_FechaSolicitudHasta As Date
Dim ml_IdTipoServicio As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_ProgressRpt As XP_ProgressBar
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim rsReporte1 As New Recordset
Property Let IdEmpleado(lValue As Long)
    ml_IdEmpleado = lValue
End Property
Property Let FechaRequeridaDesde(daValue As Date)
    mda_FechaRequeridaDesde = daValue
End Property
Property Let FechaRequeridaHasta(daValue As Date)
    mda_FechaRequeridaHasta = daValue
End Property
Property Let FechaSolicitudDesde(daValue As Date)
    mda_FechaSolicitudDesde = daValue
End Property
Property Let FechaSolicitudHasta(daValue As Date)
    mda_FechaSolicitudHasta = daValue
End Property
Property Set progressRpt(oValue As XP_ProgressBar)
    Set mo_ProgressRpt = oValue
End Property
Property Let idTipoServicio(daValue As Date)
    ml_IdTipoServicio = daValue
End Property


Sub EjecutaFormulario()
    Dim oFormulario As New CuposAsignadosReporte
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub



Sub CrearReporteCuposAsignados()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim oRange As Range
Dim iFila As Integer
Dim lIdServicio As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
On Error GoTo ManejadorError

    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    
    'Crea nueva hoja
    Set oWorkBook = oExcel.Workbooks.Add
    
    'Abre, copia y cierra la plantilla
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEcuposAsignados.xls")
    oWorkBookPlantilla.Worksheets("CuposAsignados").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    
    'Activa la primera hoja
    Set oWorkSheet = oWorkBook.Sheets(1)
    'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
    oWorkSheet.Cells(1, 7).Value = "'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oWorkSheet.Cells(1, 8).Value = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    oWorkSheet.Cells(5, 3).Value = mda_FechaRequeridaDesde
    oWorkSheet.Cells(5, 6).Value = mda_FechaRequeridaHasta
    Set rsReporte = mo_AdminReportes.CantidadHistoriaSolicitadasPorServicio(mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_IdTipoServicio)
    rsReporte.Filter = "idTipoServicio=1"
    iFila = 8
    mo_ProgressRpt.Min = 0
    mo_ProgressRpt.Max = rsReporte.RecordCount
    mo_ProgressRpt.Value = 0
    Dim rsSolicitudes  As Recordset
    
    Do While Not rsReporte.EOF
        mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        oWorkSheet.Cells(iFila, 2).Value = rsReporte!Codigo
        oWorkSheet.Cells(iFila, 3).Value = "'" & rsReporte!Servicio
        oWorkSheet.Cells(iFila, 7).Value = rsReporte!Cantidad
        Set oRange = oExcel.Range("C" & iFila & ":" & "F" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
        iFila = iFila + 1
        rsReporte.MoveNext
    Loop
    
    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 8, 2, iFila - 1, 7
    
    'Falta que salga el nombre del responsable
    If oWorkSheet.PageSetup.PrintArea <> "" Then
        oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
    End If
    oExcel.Visible = True
    oWorkSheet.PrintPreview
    'oWorkSheet.PrintOut
    
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

Sub CrearReporteCuposAsignadosVaciosPorMedico()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim rsReporte2 As New ADODB.Recordset
Dim oRange As Range
Dim iFila As Integer
Dim lIdServicio As Integer: Dim lnCuposLlenos As Integer
Dim lnIdMedico As Long: Dim lnCuposTotal As Integer: Dim lnCuposAsignados As Integer
Dim lnTotal As Integer: Dim lnLlenos As Integer: Dim lnVacios As Integer
Dim lnTTotal As Integer: Dim lnTLlenos As Integer: Dim lnTVacios As Integer
Dim lnIdServicio As Long
Dim lcServicio As String: Dim lcMedico As String
Dim lHoraInicio As Long
Dim lHoraFin  As Long
Dim lTiempoPromedio As Long
Dim lHoraSiguiente As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
On Error GoTo ManejadorError
    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    'Crea nueva hoja
    Set oWorkBook = oExcel.Workbooks.Add
    'Abre, copia y cierra la plantilla
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEcuposPorMedico.xls")
    oWorkBookPlantilla.Worksheets("CuposPorMedico").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    'Activa la primera hoja
    Set oWorkSheet = oWorkBook.Sheets(1)
    'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
    oWorkSheet.Cells(1, 7).Value = "'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oWorkSheet.Cells(1, 8).Value = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    oWorkSheet.Cells(5, 3).Value = mda_FechaSolicitudDesde
    oWorkSheet.Cells(5, 6).Value = mda_FechaSolicitudHasta
    Set rsReporte = CuposProgramadosPorMedicosPorFechas(mda_FechaSolicitudDesde, mda_FechaSolicitudHasta)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        GenerarRecordsetTemporal
'        mo_ProgressRpt.Min = 0
'        mo_ProgressRpt.Max = rsReporte.RecordCount
'        mo_ProgressRpt.Value = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lnIdMedico = rsReporte.Fields!idMedico
            lnIdServicio = rsReporte.Fields!IdEspecialidad
            lcServicio = rsReporte.Fields!Nombre
            lcMedico = Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!Nombres)
            lnCuposTotal = 0
            Do While Not rsReporte.EOF And lnIdMedico = rsReporte.Fields!idMedico
              lHoraInicio = ConvertirAMinutos(rsReporte.Fields!HoraInicio)
              lHoraFin = ConvertirAMinutos(rsReporte.Fields!HoraFin)
              lTiempoPromedio = rsReporte.Fields!TiempoPromedioAtencion
              lHoraSiguiente = lHoraInicio
              
              Do While lHoraSiguiente < lHoraFin
                    lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
                    lnCuposTotal = lnCuposTotal + 1
                    lHoraInicio = lHoraSiguiente
              Loop
              rsReporte.MoveNext
              If rsReporte.EOF Then
                 Exit Do
              End If
            Loop
            Set rsReporte2 = DevuelveCantidadCitasYaAsignadas(mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, lnIdMedico)
            rsReporte2.Filter = "idEstadoAtencion<>0"
            lnCuposLlenos = rsReporte2.RecordCount
            rsReporte1.AddNew
            rsReporte1.Fields!DServicio = lcServicio
            rsReporte1.Fields!idServicio = lnIdServicio
            rsReporte1.Fields!idMedico = lnIdMedico
            rsReporte1.Fields!dMedico = lcMedico
            rsReporte1.Fields!cuposTotal = lnCuposTotal
            rsReporte1.Fields!cuposLlenos = lnCuposLlenos
            rsReporte1.Fields!cuposVacios = lnCuposTotal - lnCuposLlenos
            rsReporte1.Update
            If rsReporte.EOF Then
               Exit Do
            End If
        Loop
        iFila = 8
        rsReporte1.Sort = "dServicio,dMedico"
        rsReporte1.MoveFirst
        lnTTotal = 0: lnTLlenos = 0: lnTVacios = 0
        Do While Not rsReporte1.EOF
           lcServicio = rsReporte1.Fields!DServicio
           lnIdServicio = rsReporte1.Fields!idServicio
           oWorkSheet.Cells(iFila, 2).Value = lcServicio
           lnTotal = 0: lnLlenos = 0: lnVacios = 0
           Do While Not rsReporte1.EOF And lcServicio = rsReporte1.Fields!DServicio And lnIdServicio = rsReporte1.Fields!idServicio
              oWorkSheet.Cells(iFila, 4).Value = rsReporte1.Fields!dMedico
              oWorkSheet.Cells(iFila, 7).Value = rsReporte1.Fields!cuposTotal
              oWorkSheet.Cells(iFila, 8).Value = rsReporte1.Fields!cuposLlenos
              oWorkSheet.Cells(iFila, 9).Value = rsReporte1.Fields!cuposVacios
              lnTotal = lnTotal + rsReporte1.Fields!cuposTotal
              lnLlenos = lnLlenos + rsReporte1.Fields!cuposLlenos
              lnVacios = lnVacios + rsReporte1.Fields!cuposVacios
              lnTTotal = lnTTotal + rsReporte1.Fields!cuposTotal
              lnTLlenos = lnTLlenos + rsReporte1.Fields!cuposLlenos
              lnTVacios = lnTVacios + rsReporte1.Fields!cuposVacios
              iFila = iFila + 1
              rsReporte1.MoveNext
              If rsReporte1.EOF Then
                 Exit Do
              End If
           Loop
           oWorkSheet.Cells(iFila, 7).Value = lnTotal
           oWorkSheet.Cells(iFila, 8).Value = lnLlenos
           oWorkSheet.Cells(iFila, 9).Value = lnVacios
           mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 9
           iFila = iFila + 2
           If rsReporte1.EOF Then
               Exit Do
           End If
        Loop
        oWorkSheet.Cells(iFila, 7).Value = lnTTotal
        oWorkSheet.Cells(iFila, 8).Value = lnTLlenos
        oWorkSheet.Cells(iFila, 9).Value = lnTVacios
        If oWorkSheet.PageSetup.PrintArea <> "" Then
            oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
    End If
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



Function CuposProgramadosPorMedicosPorFechas(ldFechaIni As Date, ldFechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand  As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
Dim sSQL As String
    Set CuposProgramadosPorMedicosPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    sSQL = "SELECT     dbo.Medicos.IdMedico, dbo.ProgramacionMedica.HoraInicio, dbo.Especialidades.IdEspecialidad, dbo.ProgramacionMedica.HoraFin, " & _
                     " dbo.ProgramacionMedica.IdEspecialidad AS Expr1, dbo.Especialidades.Nombre, dbo.EspecialidadCE.TiempoPromedioAtencion," & _
                     " dbo.ProgramacionMedica.Fecha, dbo.ProgramacionMedica.IdProgramacion, dbo.Empleados.ApellidoPaterno, dbo.Empleados.ApellidoMaterno," & _
                     " dbo.Empleados.Nombres" & _
           " FROM         dbo.ProgramacionMedica INNER JOIN" & _
                     " dbo.Medicos ON dbo.ProgramacionMedica.IdMedico = dbo.Medicos.IdMedico INNER JOIN" & _
                     " dbo.Especialidades ON dbo.ProgramacionMedica.IdEspecialidad = dbo.Especialidades.IdEspecialidad INNER JOIN" & _
                     " dbo.EspecialidadCE ON dbo.EspecialidadCE.IdEspecialidad = dbo.Especialidades.IdEspecialidad LEFT OUTER JOIN" & _
                     " dbo.Empleados ON dbo.Medicos.IdEmpleado = dbo.Empleados.IdEmpleado" & _
            " Where ProgramacionMedica.IdTipoServicio = 1" & _
            " and ProgramacionMedica.Fecha Between ('" & ldFechaIni & "') and ('" & ldFechaFin & "')" & _
            " order by Medicos.IdMedico"
    With oCommand
        .CommandType = adCmdText
        .CommandTimeout = 150
        Set .ActiveConnection = oConexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CuposProgramadosPorMedicosPorFechas = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub GenerarRecordsetTemporal()
    With rsReporte1
          .Fields.Append "dServicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "IdMedico", adInteger, 4, adFldIsNullable
          .Fields.Append "dMedico", adVarChar, 150, adFldIsNullable
          .Fields.Append "CuposTotal", adInteger
          .Fields.Append "cuposLlenos", adInteger
          .Fields.Append "cuposVacios", adInteger
          .LockType = adLockOptimistic
          .Open
          .Sort = "dServicio,dMedico"
    End With
End Sub

Function DevuelveCantidadCitasYaAsignadas(ldFechaIni As Date, ldFechaFin As Date, idMedico As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand  As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
Dim sSQL As String
    Set DevuelveCantidadCitasYaAsignadas = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    sSQL = "SELECT  dbo.Atenciones.*" & _
            " From dbo.Atenciones" & _
            " Where dbo.Atenciones.idtipoServicio=1 and FechaIngreso Between ('" & ldFechaIni & "') and ('" & ldFechaFin & "')" & _
            "       and idMedicoIngreso=" & idMedico
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveCantidadCitasYaAsignadas = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function ConvertirAMinutos(sHora As String) As Long
Dim sHoras() As String
        
        sHoras = Split(sHora, ":")
        ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))
        
End Function
Function ConvertirAHora(lMinutos As Long) As String
Dim sHoras() As String
Dim lMin As Long
Dim lHora As Long
        
        lMin = lMinutos Mod 60
        lHora = (lMinutos - lMin) / 60
        
        ConvertirAHora = Right("00" & lHora, 2) + ":" + Right("00" & lMin, 2)
        
End Function


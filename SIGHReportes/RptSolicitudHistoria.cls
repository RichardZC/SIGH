VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptSolicitudHistoria"
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
Dim ml_HoraReqIni As String
Dim ml_HoraReqFin As String
Dim ml_HoraSoliIni As String
Dim ml_HoraSoliFin As String
Dim ml_Historicos As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IncluyeHistoriasQueSalieron As Boolean
Dim ml_SoloCitasPagadas As Boolean
Property Let SoloCitasPagadas(lValue As Boolean)
    ml_SoloCitasPagadas = lValue
End Property
Property Let IncluyeHistoriasQueSalieron(lValue As Boolean)
    ml_IncluyeHistoriasQueSalieron = lValue
End Property
Property Let Historicos(lValue As Boolean)
    ml_Historicos = lValue
End Property
Property Let HoraSoliFin(lValue As String)
    ml_HoraSoliFin = lValue
End Property
Property Let HoraSoliIni(lValue As String)
    ml_HoraSoliIni = lValue
End Property
Property Let HoraReqFin(lValue As String)
    ml_HoraReqFin = lValue
End Property
Property Let HoraReqIni(lValue As String)
    ml_HoraReqIni = lValue
End Property

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

'***************daniel barrantes**************
'***************Se muestra o NO "datos históricos de Pacientes"
'***************Filtro por HORA de Cita
Sub CrearReporteHistoriaSolicitadas()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim oRsCitaPagada As New Recordset
Dim oRange As Range
Dim iFila As Integer
Dim lIdServicio As Integer
Dim lbContinuar As Boolean
Dim lcHora As String
Dim lIdServicioDestino As Long
Dim lcSql As String
Dim lcPago As String
On Error GoTo ManejadorError

    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    
    'Crea nueva hoja
    Set oWorkBook = oExcel.Workbooks.Add
    
    'Abre, copia y cierra la plantilla
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCHistoriasSolicitadas.xls")
    oWorkBookPlantilla.Worksheets("HistoriasSolicitadas").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    
    'Activa la primera hoja
    Set oWorkSheet = oWorkBook.Sheets(1)
    'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(ml_IdEmpleado)
    If Not oDOEmpleado Is Nothing Then
        'oWorkSheet.Cells(4, 4).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.nombres
    End If
    If mda_FechaRequeridaDesde <> 0 Then
         oWorkSheet.Cells(4, 10).Value = "   F.Requerida: " & mda_FechaRequeridaDesde & "  -  " & mda_FechaRequeridaHasta
    End If
    oWorkSheet.Cells(4, 24).Value = "Hora Registro Cita: " & ml_HoraReqIni & " - " & ml_HoraReqFin
    If mda_FechaSolicitudDesde <> 0 Then
         oWorkSheet.Cells(5, 12).Value = "F.Solicitada: " & mda_FechaSolicitudDesde & "  -  " & mda_FechaSolicitudHasta
    End If
    oWorkSheet.Cells(1, 18).Value = "'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oWorkSheet.Cells(1, 21).Value = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    If ml_SoloCitasPagadas = True Then
       oWorkSheet.Cells(5, 2).Value = "(sólo Pagadas para IAFA: Particular)"
    End If
    
    Set rsReporte = mo_AdminReportes.FiltrarHistoriaSolicitadas(ml_IdEmpleado, mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_IdTipoServicio, ml_IncluyeHistoriasQueSalieron)
    If rsReporte.RecordCount = 0 Then
       Exit Sub
    End If
    
    iFila = 8
    lIdServicio = 0
    mo_ProgressRpt.Min = 0
    mo_ProgressRpt.Max = rsReporte.RecordCount
    mo_ProgressRpt.Value = 0
    Dim rsSolicitudes  As Recordset
    rsReporte.MoveFirst
    Do While Not rsReporte.EOF
        mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        lcPago = "   "
        lbContinuar = True
        If Not IsNull(rsReporte!FechaHoraReg) Then
            lcHora = Format(rsReporte!HoraRequerida, SIGHEntidades.DevuelveHoraSoloFormato_HM)
            If Not (lcHora >= ml_HoraReqIni And lcHora <= ml_HoraReqFin) Then
               lbContinuar = False
            End If
        End If
        If lbContinuar And rsReporte!idAtencion > 0 Then
            lcSql = " SELECT     dbo.Atenciones.FechaIngreso, dbo.Atenciones.HoraIngreso, dbo.Atenciones.IdTipoServicio, dbo.Atenciones.IdPaciente," & _
                   " dbo.Atenciones.IdServicioIngreso, dbo.TiposFinanciamiento.Descripcion, dbo.FactOrdenServicio.IdEstadoFacturacion," & _
                   " dbo.FactOrdenServicio.IdOrden, dbo.FactOrdenServicio.IdEstadoFacturacion,dbo.Atenciones.IdFormaPago" & _
                   " FROM         dbo.Atenciones RIGHT OUTER JOIN" & _
                   " dbo.FactOrdenServicio ON dbo.Atenciones.IdCuentaAtencion = dbo.FactOrdenServicio.IdCuentaAtencion LEFT OUTER JOIN" & _
                   " dbo.TiposFinanciamiento ON dbo.Atenciones.IdFormaPago = dbo.TiposFinanciamiento.IdTipoFinanciamiento" & _
                   " Where  dbo.Atenciones.idFormaPago=1 and  dbo.Atenciones.idAtencion=" & rsReporte!idAtencion
             Select Case ml_IdTipoServicio
             Case "1"   'ce
                  lcSql = lcSql & " and (dbo.Atenciones.IdTipoServicio = 1)"
             Case "3"   'hospitalizacion
                  lcSql = lcSql & " and (dbo.Atenciones.IdTipoServicio = 3)"
             Case Else   'emergencia
                  lcSql = lcSql & " and (dbo.Atenciones.IdTipoServicio in (2,4))"
             End Select
             oRsCitaPagada.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
             lcSql = " "
             If oRsCitaPagada.RecordCount > 0 Then
                If oRsCitaPagada.Fields!IdEstadoFacturacion = 4 Then
                   lcPago = "-P-"
                Else
                   lcPago = "-N-"
                    If ml_SoloCitasPagadas = True Then
                       lbContinuar = False
                    End If
                End If
             End If
             oRsCitaPagada.Close
        End If
        If lbContinuar Then
            oWorkSheet.Cells(iFila, 1).Value = IIf(lIdServicio <> rsReporte!idServicio, rsReporte!Servicio, "")
            'If Trim(lcPago) = "" Then
               oWorkSheet.Cells(iFila, 8).Value = "'" & Trim(lcPago) & Trim(Str(rsReporte!NroHistoriaClinica))
            'Else
            '   oWorkSheet.Cells(iFila, 8).Value = Trim(lcPago) & Trim(Str(rsReporte!NroHistoriaClinica))
            'End If
            oWorkSheet.Cells(iFila, 10).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
            oWorkSheet.Cells(iFila, 18).Value = "'" & rsReporte!TipoHistoria
            oWorkSheet.Cells(iFila, 20).Value = "'" & Format(rsReporte!FechaRequerida + " " + rsReporte!HoraRequerida, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            oWorkSheet.Cells(iFila, 22).Value = "'" & Format(rsReporte!FechaSolicitud, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
'            oWorkSheet.Cells(iFila, 24).Value = "'" & rsReporte!HoraRequerida
'            oWorkSheet.Cells(iFila, 25).Value = "'" & rsReporte!HoraSolicitud
            'Obtiene Historias de Otros Servicios
            If ml_Historicos Then
                Set rsSolicitudes = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorIdPaciente(rsReporte!idPaciente)
                Do While Not rsSolicitudes.EOF
                    If rsReporte!idServicio <> rsSolicitudes!idServicio Then
                        iFila = iFila + 1
                        oWorkSheet.Cells(iFila, 24) = rsSolicitudes!Nombre
                        oWorkSheet.Cells(iFila, 25) = rsSolicitudes!FechaRequerida
                    End If
                    rsSolicitudes.MoveNext
                Loop
                Set oRange = oExcel.Range("A" & iFila & ":" & "G" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                Set oRange = oExcel.Range("H" & iFila & ":" & "I" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                Set oRange = oExcel.Range("J" & iFila & ":" & "Q" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                Set oRange = oExcel.Range("R" & iFila & ":" & "S" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                Set oRange = oExcel.Range("T" & iFila & ":" & "U" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                Set oRange = oExcel.Range("V" & iFila & ":" & "W" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
            End If
            lIdServicio = rsReporte!idServicio
            iFila = iFila + 1
        End If
        rsReporte.MoveNext
    Loop
    
    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 8, 2, iFila - 1, 25
    
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

'***************daniel barrantes**************
'***************Se muestra o NO "datos históricos de Pacientes"
'***************Filtro por HORA de Cita
Sub CrearReporteHistoriaSolicitadasDeCEPorMedico()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim oRange As Range
Dim iFila As Integer
Dim lIdServicio As Integer
Dim lbContinuar As Boolean
Dim lcHora As String
Dim lIdServicioDestino As Long
On Error GoTo ManejadorError

    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    
    'Crea nueva hoja
    Set oWorkBook = oExcel.Workbooks.Add
    
    'Abre, copia y cierra la plantilla
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCHistoriasSolicitadasCE.xls")
    oWorkBookPlantilla.Worksheets("HistoriasSolicitadasCE").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    
    'Activa la primera hoja
    Set oWorkSheet = oWorkBook.Sheets(1)
    'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(ml_IdEmpleado)
    If Not oDOEmpleado Is Nothing Then
        oWorkSheet.Cells(4, 4).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    oWorkSheet.Cells(1, 18).Value = "'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oWorkSheet.Cells(1, 21).Value = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    
    
    Set rsReporte = mo_AdminReportes.ReporteHistoriasSolicitadasCEPorMedico(ml_IdEmpleado, mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_IdTipoServicio, ml_IncluyeHistoriasQueSalieron)
    
    iFila = 8
    lIdServicio = 0
    mo_ProgressRpt.Min = 0
    mo_ProgressRpt.Max = rsReporte.RecordCount
    mo_ProgressRpt.Value = 0
    Dim rsSolicitudes  As Recordset
    
    Do While Not rsReporte.EOF
        mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        
        lbContinuar = True
        If Not IsNull(rsReporte!FechaHoraReg) Then
            lcHora = Format(rsReporte!HoraRequerida, SIGHEntidades.DevuelveHoraSoloFormato_HM)
            If Not (lcHora >= ml_HoraReqIni And lcHora <= ml_HoraReqFin) Then
               lbContinuar = False
            End If
        End If
        If lIdServicio <> rsReporte!idMedico Then
            oWorkSheet.Cells(iFila, 1).Value = mo_ReporteUtil.ArmarNombreDeEmpleado(mo_ReporteUtil.NullToVacio(rsReporte!MedicoApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!MedicoApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!MedicoNombres)) & " (" & rsReporte!Servicio & ")"
        End If
        If lbContinuar Then
            If Not IsNull(rsReporte.Fields!FichaFamiliar) And Len(Trim(rsReporte.Fields!FichaFamiliar)) > 3 Then
               oWorkSheet.Cells(iFila, 8).Value = "'" & Trim(rsReporte.Fields!FichaFamiliar)
            Else
               oWorkSheet.Cells(iFila, 8).Value = "'" & rsReporte!NroHistoriaClinica
            End If
            oWorkSheet.Cells(iFila, 10).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
            oWorkSheet.Cells(iFila, 17).Value = "'" & rsReporte!TipoHistoria
            oWorkSheet.Cells(iFila, 18).Value = "'" & Format(rsReporte!FechaRequerida, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            oWorkSheet.Cells(iFila, 20).Value = "'" & Format(rsReporte!HoraRequerida, SIGHEntidades.DevuelveHoraSoloFormato_HM)
            oWorkSheet.Cells(iFila, 22).Value = "'" & Format(rsReporte!FechaSolicitud, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            
            'Obtiene Historias de Otros Servicios
            If ml_Historicos Then
                Set rsSolicitudes = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorIdPaciente(rsReporte!idPaciente)
                Do While Not rsSolicitudes.EOF
                    If rsReporte!idServicio <> rsSolicitudes!idServicio Then
                        iFila = iFila + 1
                        oWorkSheet.Cells(iFila, 24) = rsSolicitudes!Nombre
                        oWorkSheet.Cells(iFila, 25) = rsSolicitudes!FechaRequerida
                    End If
                    rsSolicitudes.MoveNext
                Loop
                Set oRange = oExcel.Range("A" & iFila & ":" & "G" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                Set oRange = oExcel.Range("H" & iFila & ":" & "I" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                Set oRange = oExcel.Range("J" & iFila & ":" & "Q" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                Set oRange = oExcel.Range("R" & iFila & ":" & "S" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                Set oRange = oExcel.Range("T" & iFila & ":" & "U" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                Set oRange = oExcel.Range("V" & iFila & ":" & "W" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
            End If
            lIdServicio = rsReporte!idMedico
            
            iFila = iFila + 1
        End If
        rsReporte.MoveNext
    Loop
    
    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 8, 2, iFila - 1, 25
    
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


Sub CrearReporteHistoriaSolicitadasDeCEPorMedicoConSaltoPaginaXconsultorio()
    Dim rsReporte As New Recordset
    Set rsReporte = mo_AdminReportes.ReporteHistoriasSolicitadasCEPorMedico(ml_IdEmpleado, mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_IdTipoServicio, ml_IncluyeHistoriasQueSalieron)
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        Set ACsolicitudHCporMedico.DataSource = rsReporte
        ACsolicitudHCporMedico.RightMargin = 10
        ACsolicitudHCporMedico.TopMargin = 10
        ACsolicitudHCporMedico.LeftMargin = 10
        ACsolicitudHCporMedico.BottomMargin = 10
        'ACsolicitudHCporMedico.Sections("cabecera").Controls("lblPc").Caption = ""
        'ACsolicitudHCporMedico.Sections("cabecera").Controls("lblUsuario").Caption = ""
        'ACsolicitudHCporMedico.Sections("cabecera").Controls("lblSubTitulo").Caption = ""
        Set ACsolicitudHCporMedico.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\logotipo.jpg")
        'ACsolicitudHCporMedico.Sections("pie").Controls("lblPie").Caption = ""
        ACsolicitudHCporMedico.Orientation = rptOrientLandscape
        ACsolicitudHCporMedico.Show 1
    End If
End Sub

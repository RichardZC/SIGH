VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEgresosHosp"
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

Sub CrearReporteEgresosHospitalarios()
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

    'Crea la cabecera
    Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        
        'Crea nueva hoja
        Set oWorkBook = oExcel.Workbooks.Add
        
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgreso_hosp.xls.xls")
        oWorkBookPlantilla.Worksheets("RPT_EGRESO_HOSP").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        oWorkSheet.PageSetup.CenterHeader = "Reporte de Egreso de Epicrisis" & Chr(13) & "Establecimiento: " & lcBuscaParametro.SeleccionaFilaParametro(205)
           
    
        mo_ProgressRpt.Min = 0
        mo_ProgressRpt.Max = rsReporte.RecordCount
        mo_ProgressRpt.Value = 0
        
        Dim lRecordCount As Long
        Dim I As Long
        Dim lCounter As Long
        Dim sRows As String
        
        lRecordCount = rsReporte.RecordCount
        
        I = 3
        lCounter = 0
        Do While Not rsReporte.EOF
            sRows = rsReporte.GetString(adClipString, 100, , , "")
            Clipboard.Clear
            Clipboard.SetText sRows
            oWorkSheet.Range("A" & I & ":A" & I & "").PasteSpecial
            I = I + 100
            
            lCounter = lCounter + 100
            If lRecordCount <> 0 Then
                mo_ProgressRpt.Value = IIf(10 + lCounter / lRecordCount * 90 > 100, 100, 10 + lCounter / lRecordCount * 90): DoEvents
            End If
        Loop
        
        If lRecordCount > 0 Then
            'Abrimos el archivo temporal en otra hoja
            oWorkSheet.Cells.Select
            oWorkSheet.Cells.EntireColumn.AutoFit
        End If
        
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 10, 1, iFila - 1, 17
        If oWorkSheet.PageSetup.PrintArea <> "" Then
           oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If

'        Clipboard.Clear
'        Clipboard.SetText Trim(Str(rsReporte.RecordCount))
'        oWorkSheet.Range("A" & i & ":A" & i & "").PasteSpecial
        
        'Falta que salga el nombre del responsable
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


'***************daniel barrantes**************
'***************Reporte EGRESOS de Pacientes, se considera EMERGENCIA
'***************
Sub CrearReporteEgresosHospitalariosII()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim rsServicio As New Recordset
Dim oRange As Range
Dim iFila As Integer
Dim lIdServicio As Integer
Dim lcCodServEgreso  As String: Dim lcServEgreso As String: Dim lnDiasEstancia As Integer
Dim lRecordCount As Long
Dim f As Long
Dim c As Integer
Dim rsDiagnosticos  As Recordset
Dim rsProcedimientos As Recordset
Dim iCantidadProcedimientos As Integer
Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcOpcs As String


'On Error GoTo ManejadorError

    Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    If ml_IdPlan > 0 Then
       rsReporte.Filter = "idFuenteFinanciamiento=" & ml_IdPlan
    End If
    lRecordCount = rsReporte.RecordCount
    If lRecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        'oExcel.Visible = True
    
        'Crea nueva hoja
        Set oWorkBook = oExcel.Workbooks.Add
    
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgreso_hosp.xls")
        oWorkBookPlantilla.Worksheets("RPT_EGRESO_HOSP").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
    
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        oWorkSheet.PageSetup.CenterHeader = "Reporte de Egreso de Epicrisis" & Chr(13) & "Establecimiento: " & lcBuscaParametro.SeleccionaFilaParametro(205)
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
        'Crea la cabecera
    
        mo_ProgressRpt.Min = 0
        mo_ProgressRpt.Max = lRecordCount
        mo_ProgressRpt.Value = 0
        mo_ProgressRpt.ShowText = True
        mo_ProgressRpt.Color = vbGreen
    
        If lRecordCount > 0 Then
            
            oWorkSheet.Cells(1, 1).Value = ml_TextoDelFiltro
            oWorkSheet.Cells(2, 39) = "Teléfono"
            oWorkSheet.Cells(2, 40) = "Dirección del Domicilio"
            f = 4
            rsReporte.MoveFirst
        
            Do While Not rsReporte.EOF
                lcCodServEgreso = ""
                lcServEgreso = ""
                If Not IsNull(rsReporte.Fields("idServicioEgreso").Value) Then
                    Set rsServicio = ServiciosSelecionarPorId(rsReporte.Fields("idServicioEgreso").Value)
                    lcCodServEgreso = rsServicio.Fields("codigo").Value
                    lcServEgreso = rsServicio.Fields("nombre").Value
                    rsServicio.Close
                End If
                'oWorkSheet.Cells(f, 1) = IIf(IsNull(rsReporte!IdAtencion), "", rsReporte!IdAtencion)
                oWorkSheet.Cells(f, 1) = IIf(IsNull(rsReporte!NroHistoriaClinica), "", rsReporte!NroHistoriaClinica)
                oWorkSheet.Cells(f, 2) = rsReporte!idCuentaAtencion
                oWorkSheet.Cells(f, 3) = IIf(IsNull(rsReporte!ApellidoPaterno), "", rsReporte!ApellidoPaterno)
                oWorkSheet.Cells(f, 4) = IIf(IsNull(rsReporte!ApellidoMaterno), "", rsReporte!ApellidoMaterno)
                oWorkSheet.Cells(f, 5) = IIf(IsNull(rsReporte!PrimerNombre), "", rsReporte!PrimerNombre)
                oWorkSheet.Cells(f, 6) = IIf(IsNull(rsReporte!SegundoNombre), "", rsReporte!SegundoNombre)
                oWorkSheet.Cells(f, 7) = IIf(IsNull(rsReporte!Sexo), "", rsReporte!Sexo)
                oWorkSheet.Cells(f, 8) = IIf(IsNull(rsReporte!Edad), "", rsReporte!Edad)
                oWorkSheet.Cells(f, 9) = IIf(IsNull(rsReporte!TipoEdad), "", rsReporte!TipoEdad)
                oWorkSheet.Cells(f, 10) = IIf(IsNull(rsReporte!Departamento), "", rsReporte!Departamento)
                oWorkSheet.Cells(f, 11) = IIf(IsNull(rsReporte!provincia), "", rsReporte!provincia)
                oWorkSheet.Cells(f, 12) = IIf(IsNull(rsReporte!Distrito), "", rsReporte!Distrito)
                oWorkSheet.Cells(f, 13) = IIf(IsNull(rsReporte!CentroPoblado), "", rsReporte!CentroPoblado)
                oWorkSheet.Cells(f, 14) = IIf(IsNull(rsReporte!IdFuenteFinanciamiento), "", rsReporte!IdFuenteFinanciamiento)
                oWorkSheet.Cells(f, 15) = IIf(IsNull(rsReporte!DescFuenteFinanciamiento), "", rsReporte!DescFuenteFinanciamiento)
                oWorkSheet.Cells(f, 16) = IIf(IsNull(rsReporte!FechaIngreso), "", rsReporte!FechaIngreso)
                oWorkSheet.Cells(f, 17) = IIf(IsNull(rsReporte!HoraIngreso), "", rsReporte!HoraIngreso)
                oWorkSheet.Cells(f, 18) = IIf(IsNull(rsReporte!FechaEgreso), "", rsReporte!FechaEgreso)
                oWorkSheet.Cells(f, 19) = IIf(IsNull(rsReporte!horaEgreso), "", rsReporte!horaEgreso)
                oWorkSheet.Cells(f, 20) = IIf(IsNull(rsReporte!CondicionAlta), "", rsReporte!CondicionAlta)
                oWorkSheet.Cells(f, 21) = IIf(IsNull(rsReporte!ServicioIngresoCodigo), "", rsReporte!ServicioIngresoCodigo)
                oWorkSheet.Cells(f, 22) = IIf(IsNull(rsReporte!ServicioIngresoNombre), "", rsReporte!ServicioIngresoNombre)
                oWorkSheet.Cells(f, 23) = lcCodServEgreso  'IIf(IsNull(rsReporte!ServicioEgresoCodigo), "", rsReporte!ServicioEgresoCodigo)
                oWorkSheet.Cells(f, 24) = lcServEgreso  'IIf(IsNull(rsReporte!ServicioEgresoNombre), "", rsReporte!ServicioEgresoNombre)
                oWorkSheet.Cells(f, 25) = IIf(IsNull(rsReporte!TipoAlta), "", rsReporte!TipoAlta)
                
'                lnDiasEstancia = DateDiff("d", rsReporte!FechaIngreso & " " & rsReporte!HoraIngreso, rsReporte!FechaEgreso & " " & rsReporte!HoraEgreso)
'                If lnDiasEstancia = 0 Then lnDiasEstancia = 1
                lnDiasEstancia = 0
                If ml_IdTipoEspecialidad <> 1 Then
                   lnDiasEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso)
                End If
                oWorkSheet.Cells(f, 26) = lnDiasEstancia
        
                'oWorkSheet.Cells(f, 26) = DateDiff("d", rsReporte!FechaIngreso & " " & rsReporte!HoraIngreso, rsReporte!FechaEgreso & " " & rsReporte!HoraEgreso)
        
                c = 27
                iCantidadProcedimientos = 1
                If ml_IdTipoEspecialidad <> 1 Then
                    Set rsDiagnosticos = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
                Else
                    Set rsDiagnosticos = BuscaCEAtencionesDx(rsReporte!idAtencion)
                End If
                Do While Not rsDiagnosticos.EOF
                    oWorkSheet.Cells(f, c) = IIf(IsNull(rsDiagnosticos!CodigoDx), "", rsDiagnosticos!CodigoDx)
        
                    'Si es Dx principal agrega descripcion
                    If iCantidadProcedimientos = 1 Then
                        oWorkSheet.Cells(f, c + 1) = IIf(IsNull(rsDiagnosticos!DescripcionDx), "", rsDiagnosticos!DescripcionDx)
                        c = c + 1
                    End If
        
                    c = c + 1
        
                    rsDiagnosticos.MoveNext
                    iCantidadProcedimientos = iCantidadProcedimientos + 1
        
                    'Muestra maximo 5 diagnosticos
                    If iCantidadProcedimientos > 6 Then
                        Exit Do
                    End If
                Loop
                rsDiagnosticos.Close
        
        
                iCantidadProcedimientos = 1
        
                Set rsProcedimientos = mo_AdminReportes.ReporteAtencionesProcedimientos(rsReporte!idAtencion)
                If Not (rsProcedimientos Is Nothing) Then
                    If Not (rsProcedimientos.BOF And rsProcedimientos.EOF) Then
                        Do While Not rsProcedimientos.EOF
                            If Not IsNull(rsProcedimientos!Codigo) Then
                                lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsProcedimientos!Codigo)
                                oWorkSheet.Cells(f, c) = rsProcedimientos!Codigo & IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                                c = c + 1
                                iCantidadProcedimientos = iCantidadProcedimientos + 1
                            End If
                            rsProcedimientos.MoveNext
                            'Muestra maximo 5 procedimientos
                            If iCantidadProcedimientos > 5 Then
                                Exit Do
                            End If
                        Loop
                    End If
                
                     rsProcedimientos.Close
                End If
                
                oWorkSheet.Cells(f, 39) = IIf(IsNull(rsReporte!Telefono), "", rsReporte!Telefono)
                oWorkSheet.Cells(f, 40) = IIf(IsNull(rsReporte!DireccionDomicilio), "", rsReporte!DireccionDomicilio)
                
                f = f + 1
                mo_ProgressRpt.Value = f
                rsReporte.MoveNext
            Loop
        End If
        iFila = f
        iFila = iFila + 1
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 22
        oWorkSheet.Cells(iFila, 1).Value = "Nº Historias Clínicas: " + Trim(Str(lRecordCount))
        If lRecordCount > 0 Then
            'Abrimos el archivo temporal en otra hoja
            oWorkSheet.Cells.Select
            oWorkSheet.Cells.EntireColumn.AutoFit
        End If
        If oWorkSheet.PageSetup.PrintArea <> "" Then
           oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
    
        'Falta que salga el nombre del responsable
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

Function BuscaCEAtencionesDx(idAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL1 As String: Dim sSQL2 As String
Dim oConexion As New ADODB.Connection
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set BuscaCEAtencionesDx = Nothing
        sSQL1 = "select SubClasificacionDiagnosticos.IdClasificacionDx,SubClasificacionDiagnosticos.IdSubClasificacionDx," & _
                      " Diagnosticos.CodigoCIE2004 as CodigoDx, Diagnosticos.descripcion as DescripcionDx" & _
                " from  (SubClasificacionDiagnosticos left join AtencionesDiagnosticos on SubClasificacionDiagnosticos.IdSubClasificacionDx = AtencionesDiagnosticos.IdSubClasificacionDx" & _
                        " and SubClasificacionDiagnosticos.IdClasificacionDx = AtencionesDiagnosticos.IdClasificacionDx" & _
                        " and AtencionesDiagnosticos.IdAtencion = " & idAtencion & ")left join Diagnosticos on AtencionesDiagnosticos.IdDiagnostico = Diagnosticos.IdDiagnostico" & _
                " WHERE     (dbo.SubclasificacionDiagnosticos.IdTipoServicio = 1) AND (dbo.SubclasificacionDiagnosticos.IdSubclasificacionDx = 102)" & _
                " order by SubClasificacionDiagnosticos.IdSubClasificacionDx"
        With oCommand
        .CommandType = adCmdText
        .CommandTimeout = 150
        Set .ActiveConnection = oConexion
        .CommandText = sSQL1
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set BuscaCEAtencionesDx = oRecordset
    oConexion.Close
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function




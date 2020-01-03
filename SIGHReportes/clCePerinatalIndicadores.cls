VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clCePerinatalIndicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Perinatal
'        Programado por: Garay M.
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim ml_IdDepartamento As Long, ml_IdProvincia As Long, ml_IdDistrito As Long
Dim md_FechaReporte As Date
Dim ms_NombreEstablecimiento As String


Private Sub Class_Initialize()
    ms_NombreEstablecimiento = lcBuscaParametro.SeleccionaFilaParametro(205)
End Sub

Property Let NombreEstablecimiento(sValue As String)
    ms_NombreEstablecimiento = sValue
End Property

Property Let FechaReporte(daValue As Date)
    md_FechaReporte = daValue
End Property

Property Let IdDepartamento(lValue As Long)
    ml_IdDepartamento = lValue
End Property

Property Let IdProvincia(lValue As Long)
    ml_IdProvincia = lValue
End Property

Property Let IdDistrito(lValue As Long)
    ml_IdDistrito = lValue
End Property

Sub EjecutaFormulario()
    MsgBox "Reporte en desarrollo...."
    Exit Sub

    Dim oFormulario As New CeRepPerinatalIndicadores
    oFormulario.Show 1
End Sub

Public Function reporteDeSuplementoDeHierro(sTitulo As String, sFiltroAplicados As String, lnHwnd As Long)
    'variables de accedo a datos
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporteTotal As New ADODB.Recordset
    
    'variables para crystal reports
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition

    Dim crReportObject As CRAXDRT.ReportObjects
    'ruta de los reportes
    Dim rutaRaiz As String
    
    On Error GoTo ManejadorErrorImprime
    
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGHReportes") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/" & getNombreArchivo(), 1)
    
    'datos de las recetas
    Set rsReporte = mo_ReglasReportes.ReporteCREDSuplementoHierro(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 6, 35)
    
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set rsReporteTotal = mo_ReglasReportes.ReporteCREDIndicadoresTotalPaciente(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 6, 35)

    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@TotalPacientes":
                crParameter.AddCurrentValue (rsReporteTotal!total)
            Case "@Establecimiento":
                crParameter.AddCurrentValue (ms_NombreEstablecimiento)
            Case "@Titulo":
                crParameter.AddCurrentValue (UCase(sTitulo))
            Case "@FiltrosAplicados":
                crParameter.AddCurrentValue (sFiltroAplicados)
        End Select
    Next
    
    
    'setear recorset al reporte
    crReport.Database.SetDataSource rsReporte

    
    crReport.PaperOrientation = crDefaultPaperOrientation
'    If lbImpresionIndirecta = False Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = sTitulo
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
'    Else
'       On Error Resume Next       '
'       crReport.PrintOut False
'       'OrdenMedica.PrintReport
'    End If
    Exit Function
    
ManejadorErrorImprime:
    MsgBox sighEntidades.ObtenerMensajeErrorParaImpresion(Err), vbExclamation, "Reporte Módulo Niño Sano"
End Function

Public Function reporteDeLactanciaMaternaExclusiva(sTitulo As String, sFiltroAplicados As String, lnHwnd As Long)
    'variables de accedo a datos
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporteTotal As New ADODB.Recordset
    
    'variables para crystal reports
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition

    Dim crReportObject As CRAXDRT.ReportObjects
    'ruta de los reportes
    Dim rutaRaiz As String
    
    On Error GoTo ManejadorErrorImprime
    
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGHReportes") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/" & getNombreArchivo(), 1)
    
    'datos de las recetas
    Set rsReporte = mo_ReglasReportes.ReporteCREDLactanciaMaternaExclusiva(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 6)
    
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set rsReporteTotal = mo_ReglasReportes.ReporteCREDIndicadoresTotalPaciente(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 6)

    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@TotalPacientes":
                crParameter.AddCurrentValue (rsReporteTotal!total)
            Case "@Establecimiento":
                crParameter.AddCurrentValue (ms_NombreEstablecimiento)
            Case "@Titulo":
                crParameter.AddCurrentValue (UCase(sTitulo))
            Case "@FiltrosAplicados":
                crParameter.AddCurrentValue (sFiltroAplicados)
        End Select
    Next
    
    
    'setear recorset al reporte
    crReport.Database.SetDataSource rsReporte

    
    crReport.PaperOrientation = crDefaultPaperOrientation
'    If lbImpresionIndirecta = False Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = sTitulo
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
'    Else
'       On Error Resume Next       '
'       crReport.PrintOut False
'       'OrdenMedica.PrintReport
'    End If
    Exit Function
    
ManejadorErrorImprime:
    MsgBox sighEntidades.ObtenerMensajeErrorParaImpresion(Err), vbExclamation, "Reporte Módulo Niño Sano"
End Function

Public Function reporteDeSesionesDemostrativas(sTitulo As String, sFiltroAplicados As String, lnHwnd As Long)
    'variables de accedo a datos
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporteTotal As New ADODB.Recordset
    
    'variables para crystal reports
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition

    Dim crReportObject As CRAXDRT.ReportObjects
    'ruta de los reportes
    Dim rutaRaiz As String
    
    On Error GoTo ManejadorErrorImprime
    
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGHReportes") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/" & getNombreArchivo(), 1)
    
    'datos de las recetas
    Set rsReporte = mo_ReglasReportes.ReporteCREDSesionesDemostrativas(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 12)
    
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set rsReporteTotal = mo_ReglasReportes.ReporteCREDIndicadoresTotalPaciente(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 12)

    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@TotalPacientes":
                crParameter.AddCurrentValue (rsReporteTotal!total)
            Case "@Establecimiento":
                crParameter.AddCurrentValue (ms_NombreEstablecimiento)
            Case "@Titulo":
                crParameter.AddCurrentValue (UCase(sTitulo))
            Case "@FiltrosAplicados":
                crParameter.AddCurrentValue (sFiltroAplicados)
        End Select
    Next
    
    
    'setear recorset al reporte
    crReport.Database.SetDataSource rsReporte

    
    crReport.PaperOrientation = crDefaultPaperOrientation
'    If lbImpresionIndirecta = False Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = sTitulo
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
'    Else
'       On Error Resume Next       '
'       crReport.PrintOut False
'       'OrdenMedica.PrintReport
'    End If
    Exit Function
    
ManejadorErrorImprime:
    MsgBox sighEntidades.ObtenerMensajeErrorParaImpresion(Err), vbExclamation, "Reporte Módulo Niño Sano"
End Function

Public Function reporteDeVacunaRotavirusNeumococo(sTitulo As String, sFiltroAplicados As String, lnHwnd As Long)
    'variables de accedo a datos
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporteTotal As New ADODB.Recordset
    
    'variables para crystal reports
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition

    Dim crReportObject As CRAXDRT.ReportObjects
    'ruta de los reportes
    Dim rutaRaiz As String
    
    On Error GoTo ManejadorErrorImprime
    
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGHReportes") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/" & getNombreArchivo(), 1)
    
    'datos de las recetas
    Set rsReporte = mo_ReglasReportes.ReporteCREDVacunaRatavirusNeumococo(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 24)
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set rsReporteTotal = mo_ReglasReportes.ReporteCREDIndicadoresTotalPaciente(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 24)
                                                
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If

    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@TotalPacientes":
                crParameter.AddCurrentValue (rsReporteTotal!total)
            Case "@Establecimiento":
                crParameter.AddCurrentValue (ms_NombreEstablecimiento)
            Case "@Titulo":
                crParameter.AddCurrentValue (UCase(sTitulo))
            Case "@FiltrosAplicados":
                crParameter.AddCurrentValue (sFiltroAplicados)
        End Select
    Next
    
    
    'setear recorset al reporte
    crReport.Database.SetDataSource rsReporte

    
    crReport.PaperOrientation = crDefaultPaperOrientation
'    If lbImpresionIndirecta = False Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = sTitulo
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
'    Else
'       On Error Resume Next       '
'       crReport.PrintOut False
'       'OrdenMedica.PrintReport
'    End If
    Exit Function
    
ManejadorErrorImprime:
    MsgBox sighEntidades.ObtenerMensajeErrorParaImpresion(Err), vbExclamation, "Reporte Módulo Niño Sano"
End Function


Public Function reporteDeCREDCompleto(sTitulo As String, sFiltroAplicados As String, lnHwnd As Long)
    'variables de accedo a datos
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporteTotal As New ADODB.Recordset
    
    'variables para crystal reports
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition

    Dim crReportObject As CRAXDRT.ReportObjects
    'ruta de los reportes
    Dim rutaRaiz As String
    
    On Error GoTo ManejadorErrorImprime
    
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGHReportes") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/" & getNombreArchivo(), 1)
    
    'datos de las recetas
    Set rsReporte = mo_ReglasReportes.ReporteCREDCompleto(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 36)
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If
    
    Set rsReporteTotal = mo_ReglasReportes.ReporteCREDIndicadoresTotalPaciente(md_FechaReporte, ml_IdDepartamento, _
                                                ml_IdProvincia, ml_IdDistrito, 0, 36)
                                                
    If ExisteErrorEnReporte(mo_ReglasReportes.MensajeError) = True Then
        Exit Function
    End If

    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@TotalPacientes":
                crParameter.AddCurrentValue (rsReporteTotal!total)
            Case "@Establecimiento":
                crParameter.AddCurrentValue (ms_NombreEstablecimiento)
            Case "@Titulo":
                crParameter.AddCurrentValue (UCase(sTitulo))
            Case "@FiltrosAplicados":
                crParameter.AddCurrentValue (sFiltroAplicados)
        End Select
    Next
    
    
    'setear recorset al reporte
    crReport.Database.SetDataSource rsReporte

    
    crReport.PaperOrientation = crDefaultPaperOrientation
'    If lbImpresionIndirecta = False Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = sTitulo
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
'    Else
'       On Error Resume Next       '
'       crReport.PrintOut False
'       'OrdenMedica.PrintReport
'    End If
    Exit Function
    
ManejadorErrorImprime:
    MsgBox sighEntidades.ObtenerMensajeErrorParaImpresion(Err), vbExclamation, "Reporte Módulo Niño Sano"
End Function

Private Function getNombreArchivo() As String
    Dim sNombre As String
    If ml_IdDepartamento > 0 And ml_IdProvincia > 0 And ml_IdDistrito > 0 Then
        sNombre = "RptCREDIndicadoresCentroPoblado.rpt"
    ElseIf ml_IdDepartamento > 0 And ml_IdProvincia > 0 Then
        sNombre = "RptCREDIndicadoresDistrito.rpt"
    ElseIf ml_IdDepartamento > 0 Then
        sNombre = "RptCREDIndicadoresProvincia.rpt"
    Else
        sNombre = "RptCREDIndicadoresDpto.rpt"
    End If
    getNombreArchivo = sNombre
End Function


Private Function ExisteErrorEnReporte(sMensajeError As String) As Boolean
    ExisteErrorEnReporte = True
    If sMensajeError <> "" Then
        MsgBox sMensajeError, vbInformation, "Reportes CRED"
        Exit Function
    End If
    ExisteErrorEnReporte = False
End Function

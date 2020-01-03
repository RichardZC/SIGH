VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clSolicitudHistorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para solicitud de historias
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdEmpleado As Long
Dim mda_FechaRequeridaDesde As Date
Dim mda_FechaRequeridaHasta As Date
Dim mda_FechaSolicitudDesde As Date
Dim mda_FechaSolicitudHasta As Date
Dim ml_idTipoServicio As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReporteUtil As New ReporteUtil
'Dim mo_ProgressRpt As XP_ProgressBar

Dim ml_HoraReqIni As String
Dim ml_HoraReqFin As String
Dim ml_HoraSoliIni As String
Dim ml_HoraSoliFin As String
Dim ml_Historicos As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IncluyeHistoriasQueSalieron As Boolean
Dim ml_SoloCitasPagadas As Boolean
Dim ms_TipoReporte  As String
Dim ml_idUsuario As Long
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim ms_UltimosDigitosHC As String
Property Let UltimosDigitosHC(sValue As String)
    ms_UltimosDigitosHC = sValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let TipoReporte(sValue As String)
    ms_TipoReporte = sValue
End Property


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
'Property Set progressRpt(oValue As XP_ProgressBar)
'    Set mo_ProgressRpt = oValue
'End Property
Property Let idTipoServicio(daValue As Date)
    ml_idTipoServicio = daValue
End Property

'***************daniel barrantes**************
'***************Se muestra o NO "datos históricos de Pacientes"
'***************Filtro por HORA de Cita
Sub CrearReporteHistoriaSolicitadas(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim oRsCitaPagada As New Recordset
Dim oRange As range
Dim iFila As Long
Dim lIdServicio As Integer
Dim lbContinuar As Boolean
Dim lcHora As String
Dim lIdServicioDestino As Long
Dim lcSql As String
Dim lcPago As String
Dim lbEsOpenOffice As Boolean

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError

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
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\HCHistoriasSolicitadas.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCHistoriasSolicitadas.xls")
        oWorkBookPlantilla.Worksheets("HistoriasSolicitadas").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(ml_IdEmpleado)
    If Not oDOEmpleado Is Nothing Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(3, 0).setFormula("Responsable: " & oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres)
        Else
            oWorkSheet.Cells(4, 1).Value = "Responsable: " & oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    If mda_FechaRequeridaDesde <> 0 Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(7, 1).setFormula("HISTORIAS CLINICAS SOLICITADAS  (" & IIf(ml_idTipoServicio = 1, "Consulta Externa", IIf(ml_idTipoServicio = 2, "Emergencia", "Hospitalizaciòn")) & ")")
            Call Feuille.getcellbyposition(9, 3).setFormula("   F.Requerida: " & mda_FechaRequeridaDesde & "  -  " & mda_FechaRequeridaHasta)
            Call Feuille.getcellbyposition(17, 6).setFormula(" ")
        Else
            oWorkSheet.Cells(2, 8).Value = "HISTORIAS CLINICAS SOLICITADAS  (" & IIf(ml_idTipoServicio = 1, "Consulta Externa", IIf(ml_idTipoServicio = 2, "Emergencia", "Hospitalizaciòn")) & ")"
            oWorkSheet.Cells(4, 10).Value = "   F.Requerida: " & mda_FechaRequeridaDesde & "  -  " & mda_FechaRequeridaHasta
            oWorkSheet.Cells(7, 18).Value = ""
        End If
    End If
    If lbEsOpenOffice = True Then
       Call Feuille.getcellbyposition(23, 3).setFormula("Hora Registro Cita: " & ml_HoraReqIni & " - " & ml_HoraReqFin)
    Else
       oWorkSheet.Cells(4, 24).Value = "Hora Registro Cita: " & ml_HoraReqIni & " - " & ml_HoraReqFin
    End If
    If mda_FechaSolicitudDesde <> 0 Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(11, 4).setFormula("F.Solicitada: " & mda_FechaSolicitudDesde & "  -  " & mda_FechaSolicitudHasta)
        Else
            oWorkSheet.Cells(5, 12).Value = "F.Solicitada: " & mda_FechaSolicitudDesde & "  -  " & mda_FechaSolicitudHasta
            
        End If
    End If
    If lbEsOpenOffice = True Then
       'Call Feuille.getcellbyposition(17, 0).setFormula(lcBuscaParametro.RetornaFechaServidorSQL)
       'Call Feuille.getcellbyposition(20, 0).setFormula(lcBuscaParametro.RetornaHoraServidorSQL)
    Else
        'oWorkSheet.Cells(1, 18).Value = lcBuscaParametro.RetornaFechaServidorSQL
        'oWorkSheet.Cells(1, 21).Value = lcBuscaParametro.RetornaHoraServidorSQL
    End If
    If ml_SoloCitasPagadas = True Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 4).setFormula("(sólo Pagadas para IAFA: Particular)")
        Else
            oWorkSheet.Cells(5, 2).Value = "(sólo Pagadas para IAFA: Particular)"
        End If
    End If
    
    Set rsReporte = mo_AdminReportes.FiltrarHistoriaSolicitadas(ml_IdEmpleado, mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_idTipoServicio, ml_IncluyeHistoriasQueSalieron)
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existen Historias para esos datos" & Chr(13) & _
              "                  ó                  " & Chr(13) & _
              "el USUARIO del Sistema debe trabajar en ARCHIVOS CLINICOS (Empleados->Area=Consultorios externos, SubArea=Archivos Clínicos)", vbInformation, "Historias Clínicas"
       Exit Sub
    End If
    
    iFila = 8
    lIdServicio = 0
'    mo_ProgressRpt.Min = 0
'    mo_ProgressRpt.Max = rsReporte.RecordCount
'    mo_ProgressRpt.Value = 0
    Dim rsSolicitudes  As Recordset
    rsReporte.MoveFirst
    Do While Not rsReporte.EOF
       ' mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        lcPago = "   "
        lbContinuar = True
        If Not IsNull(rsReporte!FechaHoraReg) Then
            lcHora = Format(rsReporte!HoraRequerida, sighentidades.DevuelveHoraSoloFormato_HM)
            If Not (lcHora >= ml_HoraReqIni And lcHora <= ml_HoraReqFin) Then
               lbContinuar = False
            End If
        End If
        If lbContinuar And rsReporte!idAtencion > 0 Then

             Set oRsCitaPagada = mo_ReglasAdmision.AtencionesSeleccionaCitasPagadas(rsReporte!idAtencion)
             oRsCitaPagada.Filter = "IdTipoServicio = " & ml_idTipoServicio
             '
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
        If lbContinuar And Val(ms_UltimosDigitosHC) > 0 Then
           If Right(Trim(Str(rsReporte!NroHistoriaClinica)), Len(ms_UltimosDigitosHC)) <> ms_UltimosDigitosHC Then
              lbContinuar = False
           End If
        End If
        If lbContinuar Then
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula(IIf(lIdServicio <> rsReporte!idServicio, rsReporte!Servicio, ""))
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula("'" & Trim(lcPago) & Trim(Str(rsReporte!NroHistoriaClinica)) & _
                                                                        IIf(Format(rsReporte!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY) = lcBuscaParametro.RetornaFechaServidorSQL, "(n)", ""))
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula("'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))
                'Call Feuille.getcellbyposition(17, iFila - 1).setFormula("'" & rsReporte!TipoHistoria)
                Call Feuille.getcellbyposition(19, iFila - 1).setFormula("'" & Format(rsReporte!FechaRequerida + " " + rsReporte!HoraRequerida, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
                Call Feuille.getcellbyposition(21, iFila - 1).setFormula("'" & Format(rsReporte!FechaSolicitud, sighentidades.DevuelveFechaSoloFormato_DMY))
            Else
                oWorkSheet.Cells(iFila, 1).Value = IIf(lIdServicio <> rsReporte!idServicio, rsReporte!Servicio, "")
                oWorkSheet.Cells(iFila, 8).Value = "'" & Trim(lcPago) & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), False) & _
                                                   IIf(Format(rsReporte!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY) = lcBuscaParametro.RetornaFechaServidorSQL, "(n)", "")
                oWorkSheet.Cells(iFila, 10).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
                'oWorkSheet.Cells(iFila, 18).Value = "'" & rsReporte!TipoHistoria
                oWorkSheet.Cells(iFila, 20).Value = "'" & Format(rsReporte!FechaRequerida + " " + rsReporte!HoraRequerida, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
                oWorkSheet.Cells(iFila, 22).Value = "'" & Format(rsReporte!FechaSolicitud, sighentidades.DevuelveFechaSoloFormato_DMY)
            End If
            'Obtiene Historias de Otros Servicios
            If ml_Historicos Then
                Set rsSolicitudes = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorIdPaciente(rsReporte!IdPaciente)
                Do While Not rsSolicitudes.EOF
                    If rsReporte!idServicio <> rsSolicitudes!idServicio Then
                        iFila = iFila + 1
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(23, iFila - 1).setFormula(rsSolicitudes!Nombre)
                            Call Feuille.getcellbyposition(24, iFila - 1).setFormula(rsSolicitudes!FechaRequerida)
                        Else
                            oWorkSheet.Cells(iFila, 24) = rsSolicitudes!Nombre
                            oWorkSheet.Cells(iFila, 25) = rsSolicitudes!FechaRequerida
                        End If
                    End If
                    rsSolicitudes.MoveNext
                Loop
                If lbEsOpenOffice = True Then
                Else
                    Set oRange = oExcel.range("A" & iFila & ":" & "G" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                    Set oRange = oExcel.range("H" & iFila & ":" & "I" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                    Set oRange = oExcel.range("J" & iFila & ":" & "Q" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                    Set oRange = oExcel.range("R" & iFila & ":" & "S" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                    Set oRange = oExcel.range("T" & iFila & ":" & "U" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                    Set oRange = oExcel.range("V" & iFila & ":" & "W" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                End If
            End If
            lIdServicio = rsReporte!idServicio
            iFila = iFila + 1
        End If
        rsReporte.MoveNext
    Loop
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Y" & CStr(iFila - 1))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
    Else
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 8, 2, iFila - 1, 25
    End If
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 25
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        'Falta que salga el nombre del responsable
        If oWorkSheet.PageSetup.PrintArea <> "" Then
           oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
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
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
    Resume
End Sub

'***************daniel barrantes**************
'***************Se muestra o NO "datos históricos de Pacientes"
'***************Filtro por HORA de Cita
Sub CrearReporteHistoriaSolicitadasDeCEPorMedico(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim oRange As range
Dim iFila As Long
Dim lIdServicio As Integer
Dim lbContinuar As Boolean
Dim lcHora As String
Dim lIdServicioDestino As Long
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError

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

    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\HCHistoriasSolicitadasCE.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCHistoriasSolicitadasCE.xls")
        oWorkBookPlantilla.Worksheets("HistoriasSolicitadasCE").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(ml_IdEmpleado)
    If Not oDOEmpleado Is Nothing Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(3, 3).setFormula(oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres)
        Else
            oWorkSheet.Cells(4, 4).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    If lbEsOpenOffice = True Then
       ' Call Feuille.getcellbyposition(17, 0).setFormula(lcBuscaParametro.RetornaFechaServidorSQL)
       ' Call Feuille.getcellbyposition(20, 0).setFormula(lcBuscaParametro.RetornaHoraServidorSQL)
        Call Feuille.getcellbyposition(9, 3).setFormula("Fecha requerida: " & _
                                        Format(mda_FechaRequeridaDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                        " " & ml_HoraReqIni & " hasta " & _
                                        Format(mda_FechaRequeridaHasta, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                        " " & ml_HoraReqFin)
    Else
        'oWorkSheet.Cells(1, 18).Value = lcBuscaParametro.RetornaFechaServidorSQL
       ' oWorkSheet.Cells(1, 21).Value = lcBuscaParametro.RetornaHoraServidorSQL
        oWorkSheet.Cells(4, 10).Value = "Fecha requerida: " & _
                                        Format(mda_FechaRequeridaDesde, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                        " " & ml_HoraReqIni & " hasta " & _
                                        Format(mda_FechaRequeridaHasta, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                        " " & ml_HoraReqFin
    End If
    
    Set rsReporte = mo_AdminReportes.ReporteHistoriasSolicitadasCEPorMedico(ml_IdEmpleado, mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_idTipoServicio, ml_IncluyeHistoriasQueSalieron)
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existen Historias para esos datos" & Chr(13) & _
              "                  ó                 " & Chr(13) & _
              "el USUARIO del Sistema debe trabajar en ARCHIVOS CLINICOS (Empleados->Area=Consultorios externos, SubArea=Archivos Clínicos)", vbInformation, "Historias Clínicas"
       Exit Sub
    End If
    
    iFila = 8
    lIdServicio = 0
'    mo_ProgressRpt.Min = 0
'    mo_ProgressRpt.Max = rsReporte.RecordCount
'    mo_ProgressRpt.Value = 0
    Dim rsSolicitudes  As Recordset
    
    Do While Not rsReporte.EOF
       ' mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        
        lbContinuar = True
        If Not IsNull(rsReporte!FechaHoraReg) Then
            lcHora = Format(rsReporte!HoraRequerida, sighentidades.DevuelveHoraSoloFormato_HM)
            If Not (lcHora >= ml_HoraReqIni And lcHora <= ml_HoraReqFin) Then
               lbContinuar = False
            End If
        End If
        If lIdServicio <> rsReporte!idMedico Then
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula(mo_ReporteUtil.ArmarNombreDeEmpleado(mo_ReporteUtil.NullToVacio(rsReporte!MedicoApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!MedicoApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!MedicoNombres)) & " (" & rsReporte!Servicio & ")")
            Else
                oWorkSheet.Cells(iFila, 1).Value = mo_ReporteUtil.ArmarNombreDeEmpleado(mo_ReporteUtil.NullToVacio(rsReporte!MedicoApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!MedicoApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!MedicoNombres)) & " (" & rsReporte!Servicio & ")"
            End If
        End If
        If lbContinuar And Val(ms_UltimosDigitosHC) > 0 Then
           If Right(Trim(Str(rsReporte!NroHistoriaClinica)), Len(ms_UltimosDigitosHC)) <> ms_UltimosDigitosHC Then
              lbContinuar = False
           End If
        End If
        If lbContinuar Then
            If Not IsNull(rsReporte.Fields!FichaFamiliar) And Len(Trim(rsReporte.Fields!FichaFamiliar)) > 3 Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula("'" & Trim(rsReporte.Fields!FichaFamiliar))
                Else
                    oWorkSheet.Cells(iFila, 8).Value = "'" & Trim(rsReporte.Fields!FichaFamiliar)
                End If
            Else
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula("'" & rsReporte!NroHistoriaClinica & _
                                                      IIf(Format(rsReporte!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY) = lcBuscaParametro.RetornaFechaServidorSQL, "(n)", ""))
                Else
                    oWorkSheet.Cells(iFila, 8).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True) & _
                                                      IIf(Format(rsReporte!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY) = lcBuscaParametro.RetornaFechaServidorSQL, "(n)", "")
                End If
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula("'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))
               ' Call Feuille.getcellbyposition(16, iFila - 1).setFormula("'" & rsReporte!TipoHistoria)
                Call Feuille.getcellbyposition(17, iFila - 1).setFormula("'" & Format(rsReporte!FechaRequerida, sighentidades.DevuelveFechaSoloFormato_DMY))
                Call Feuille.getcellbyposition(19, iFila - 1).setFormula("'" & Format(rsReporte!HoraRequerida, sighentidades.DevuelveHoraSoloFormato_HM))
                Call Feuille.getcellbyposition(21, iFila - 1).setFormula("'" & Format(rsReporte!FechaSolicitud, sighentidades.DevuelveFechaSoloFormato_DMY))
            Else
                oWorkSheet.Cells(iFila, 10).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
               ' oWorkSheet.Cells(iFila, 17).Value = "'" & rsReporte!TipoHistoria
                oWorkSheet.Cells(iFila, 18).Value = "'" & Format(rsReporte!FechaRequerida, sighentidades.DevuelveFechaSoloFormato_DMY)
                oWorkSheet.Cells(iFila, 20).Value = "'" & Format(rsReporte!HoraRequerida, sighentidades.DevuelveHoraSoloFormato_HM)
                oWorkSheet.Cells(iFila, 22).Value = "'" & Format(rsReporte!FechaSolicitud, sighentidades.DevuelveFechaSoloFormato_DMY)
            End If
            'Obtiene Historias de Otros Servicios
            If ml_Historicos Then
                Set rsSolicitudes = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorIdPaciente(rsReporte!IdPaciente)
                Do While Not rsSolicitudes.EOF
                    If rsReporte!idServicio <> rsSolicitudes!idServicio Then
                        iFila = iFila + 1
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(23, iFila - 1).setFormula(rsSolicitudes!Nombre)
                            Call Feuille.getcellbyposition(24, iFila - 1).setFormula(rsSolicitudes!FechaRequerida)
                        Else
                            oWorkSheet.Cells(iFila, 24) = rsSolicitudes!Nombre
                            oWorkSheet.Cells(iFila, 25) = rsSolicitudes!FechaRequerida
                        End If
                    End If
                    rsSolicitudes.MoveNext
                Loop
                If lbEsOpenOffice = True Then
                Else
                    Set oRange = oExcel.range("A" & iFila & ":" & "G" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                    Set oRange = oExcel.range("H" & iFila & ":" & "I" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                    Set oRange = oExcel.range("J" & iFila & ":" & "Q" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
                    Set oRange = oExcel.range("R" & iFila & ":" & "S" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                    Set oRange = oExcel.range("T" & iFila & ":" & "U" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                    Set oRange = oExcel.range("V" & iFila & ":" & "W" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
                End If
            End If
            lIdServicio = rsReporte!idMedico
            
            iFila = iFila + 1
        End If
        rsReporte.MoveNext
    Loop
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Y" & CStr(iFila - 1))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
    Else
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 8, 2, iFila - 1, 25
    End If

    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 6
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        'Falta que salga el nombre del responsable
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
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
    Resume
End Sub


Sub CrearReporteHistoriaSolicitadasDeCEPorMedicoConSaltoPaginaXconsultorio()
    Dim rsReporte As New Recordset
    Set rsReporte = mo_AdminReportes.ReporteHistoriasSolicitadasCEPorMedico(ml_IdEmpleado, mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_idTipoServicio, ml_IncluyeHistoriasQueSalieron)
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        Set ACsolicitudHCporMedico.DataSource = rsReporte
        ACsolicitudHCporMedico.RightMargin = 10
        ACsolicitudHCporMedico.TopMargin = 10
        ACsolicitudHCporMedico.LeftMargin = 10
        ACsolicitudHCporMedico.BottomMargin = 10
        ACsolicitudHCporMedico.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        ACsolicitudHCporMedico.Sections("cabecera").Controls("lblESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        ACsolicitudHCporMedico.Sections("cabecera").Controls("lblESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        Set ACsolicitudHCporMedico.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
        'ACsolicitudHCporMedico.Sections("pie").Controls("lblPie").Caption = ""
        ACsolicitudHCporMedico.Orientation = rptOrientLandscape
        ACsolicitudHCporMedico.Show 1
        'debb-27/05/2015
        Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
        mo_reglasComunes.grabaTablaAuditoria ("ACsolicitudHCporMedico: " & _
                                       mda_FechaRequeridaDesde & " " & mda_FechaRequeridaHasta)
        Set mo_reglasComunes = Nothing
    End If
End Sub



Sub EjecutaFormulario()
    Dim oFormulario As New SolicitudHistoriasReporte
    oFormulario.TipoReporte = ms_TipoReporte
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub



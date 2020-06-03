VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clReportesEgreHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Egresos Hospitalarios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdmiServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
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
Dim ml_IdTipoReporte As Long
Dim ml_idTipoServicio
Dim lnValorActualProgress As Long
Public Event ProgressActualizaValor(ByRef lnValorActual As Long, ByRef lnValorTotal As Long)
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let ValorActualProgress(lValue As Long)
   lnValorActualProgress = lValue
End Property
Property Get ValorActualProgress() As Long
   ValorActualProgress = lnValorActualProgress
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

Sub CrearReporteEgresosHospitalarios(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim oRange As Range
Dim iFila As Long
Dim lIdServicio As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
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
    'Crea la cabecera
    Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HEgreso_hosp.ods"
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
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgreso_hosp.xls.xls")
            oWorkBookPlantilla.Worksheets("RPT_EGRESO_HOSP").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
    
        'mo_ProgressRpt.Min = 0
        'mo_ProgressRpt.Max = rsReporte.RecordCount
        'mo_ProgressRpt.Value = 0
        
        Dim lRecordCount As Long
        Dim i As Long
        Dim lCounter As Long
        Dim sRows As String
        
        lRecordCount = rsReporte.RecordCount
        
        i = 3
        lCounter = 0
        Do While Not rsReporte.EOF
            sRows = rsReporte.GetString(adClipString, 100, , , "")
            Clipboard.Clear
            Clipboard.SetText sRows
            oWorkSheet.Range("A" & i & ":A" & i & "").PasteSpecial
            i = i + 100
            
            lCounter = lCounter + 100
            If lRecordCount <> 0 Then
                'mo_ProgressRpt.Value = IIf(10 + lCounter / lRecordCount * 90 > 100, 100, 10 + lCounter / lRecordCount * 90): DoEvents
            End If
        Loop
        
        If lRecordCount > 0 Then
            'Abrimos el archivo temporal en otra hoja
            oWorkSheet.Cells.Select
            oWorkSheet.Cells.EntireColumn.AutoFit
        End If
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("A" & 10 & ":Q" & CStr(iFila - 1))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 10, 1, iFila - 1, 17
        End If
'        Clipboard.Clear
'        Clipboard.SetText Trim(Str(rsReporte.RecordCount))
'        oWorkSheet.Range("A" & i & ":A" & i & "").PasteSpecial
        
        'Falta que salga el nombre del responsable
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 38
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
            'Set oExcel = Nothing
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
'FRANK
Sub CrearReporteEgresosHospitalariosII(lbConsiderarDX As Boolean, lbConsiderarCPT As Boolean, _
                                       lnHwnd As Long, lbEmergSoloLosQuePasaron24hr As Boolean, _
                                       lbMuestraHistoriasDuplicadas As Boolean, lcHistoria As String, _
                                       lbSoloCitados As Boolean)

Dim rsReporte As New Recordset
Dim rsServicio As New Recordset
Dim oRange As Range
Dim iFila As Long
Dim lIdServicio As Integer
Dim lcCodServEgreso  As String: Dim lcServEgreso As String: Dim lnDiasEstancia As Integer
Dim lRecordCount As Long
Dim f As Long
Dim c As Integer
Dim rsDiagnosticos  As New Recordset
Dim rsProcedimientos As New Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcOpcs As String, lcHoraEstanciaMax As String
Dim lcSql As String, lcRenaes As String
Dim iCantidadProcedimientos As Integer
Dim lbEsOpenOffice As Boolean
Dim lcReferenciaOrigen As String, lcReferenciaOrigenNro As String
Dim lcReferenciaDestino As String, lcReferenciaDestinoNro As String, lcNroEmergencia As String
Dim lcUsuario As String, ldUsuarioFecha As String
Dim lbSeguir As Boolean, lnIdCuentaAtencion As Long, lcFiltro As String
Dim oConexion As New Connection
Dim oRsErrores As New Recordset
Dim lbPrimerError As Boolean, lbNoExisteH As Boolean
Dim lcMensajeCelular As String, wxParametro205 As String
Dim oMensajeCelular As New SIGHProxies.Procesos
Dim mo_Pacientes As New DOPaciente
Dim oPaciente As New Pacientes
wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
lcRenaes = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 10)
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighentidades.CadenaConexion


Set oPaciente.Conexion = oConexion

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
'On Error GoTo ManejadorError
 
 
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
    Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(ml_IdDepartamento, ml_IdEspecialidad, ml_idServicio, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    lcFiltro = ""
    If ml_IdPlan > 0 Then
       lcFiltro = "idFuenteFinanciamiento=" & ml_IdPlan
    End If
    If lcHistoria <> "" Then
       If lcFiltro = "" Then
          lcFiltro = "nroHistoriaClinica=" & lcHistoria
       Else
          lcFiltro = lcFiltro & " and nroHistoriaClinica=" & lcHistoria
       End If
    End If
    If lbSoloCitados = True Then
       If lcFiltro = "" Then
          lcFiltro = "HoraEgreso=null"
       Else
          lcFiltro = lcFiltro & " and HoraEgreso=null"
       End If
    End If
    If lcFiltro <> "" Then
       rsReporte.Filter = lcFiltro
    End If
    lRecordCount = rsReporte.RecordCount
    If lRecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        '
        lbPrimerError = True
        With oRsErrores
              .Fields.Append "Historia", adInteger
              .Fields.Append "FechaIng", adDate
              .Fields.Append "Error", adVarChar, 50, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        '
        
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HEgreso_hosp.ods"
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
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        'Crea la cabecera
    
        'mo_ProgressRpt.Min = 0
        'mo_ProgressRpt.Max = lRecordCount
        'mo_ProgressRpt.Value = 0
        'mo_ProgressRpt.ShowText = True
        'mo_ProgressRpt.Color = vbGreen
    
        If lRecordCount > 0 Then
            lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, 0).setFormula(ml_TextoDelFiltro)
'                Call Feuille.getcellbyposition(38, 1).setFormula("Teléfono")
'                Call Feuille.getcellbyposition(39, 1).setFormula("Dirección del Domicilio")
'                Call Feuille.getcellbyposition(40, 1).setFormula("Médico Ingreso")
'                Call Feuille.getcellbyposition(41, 1).setFormula("Ref.Origen")
'                Call Feuille.getcellbyposition(42, 1).setFormula("N°Ref.Origen")
                If ml_IdTipoEspecialidad = 1 Then
                   Call Feuille.getcellbyposition(15, 1).setFormula("Fecha de Cita")
                   Call Feuille.getcellbyposition(16, 1).setFormula("Hora de Cita")
                   Call Feuille.getcellbyposition(17, 1).setFormula("F.Atención Cita")
                   Call Feuille.getcellbyposition(18, 1).setFormula("H.Atención Cita")
                   Call Feuille.getcellbyposition(19, 1).setFormula("")
                   Call Feuille.getcellbyposition(20, 1).setFormula("Consultorio")
                   Call Feuille.getcellbyposition(22, 1).setFormula("")
                   Call Feuille.getcellbyposition(22, 2).setFormula("")
                   Call Feuille.getcellbyposition(23, 2).setFormula("")
                   Call Feuille.getcellbyposition(24, 1).setFormula("")
                   Call Feuille.getcellbyposition(25, 1).setFormula("")
                Else
                   Call Feuille.getcellbyposition(40, 1).setFormula("Médico de Egreso")
                End If
                If ml_IdTipoEspecialidad = 1 Then
                    Call Feuille.getcellbyposition(29, 1).setFormula("Otros Dx")
                    Call Feuille.getcellbyposition(29, 2).setFormula("Dx3")
                    Call Feuille.getcellbyposition(30, 2).setFormula("Dx4")
                    Call Feuille.getcellbyposition(31, 2).setFormula("Dx5")
                    Call Feuille.getcellbyposition(32, 2).setFormula("Dx6")
                End If
                Call Feuille.getcellbyposition(43, 1).setFormula("Ref. Destino")
                Call Feuille.getcellbyposition(44, 1).setFormula("Ref. Destino Nro")
                Call Feuille.getcellbyposition(45, 1).setFormula("Usuario Digitador")
                Call Feuille.getcellbyposition(46, 1).setFormula("Fecha Digitador")
                If ml_IdTipoEspecialidad = 2 Then
                   Call Feuille.getcellbyposition(47, 1).setFormula("Emergencia N°")
                End If
            Else
                oWorkSheet.Cells(1, 1).Value = ml_TextoDelFiltro
'                oWorkSheet.Cells(2, 39) = "Teléfono"
'                oWorkSheet.Cells(2, 40) = "Dirección del Domicilio"
'                oWorkSheet.Cells(2, 41) = "Médico Ingreso"
'                oWorkSheet.Cells(2, 42) = "Ref.Origen"
'                oWorkSheet.Cells(2, 43) = "N°Ref.Origen"
                If ml_IdTipoEspecialidad = 1 Then
                   oWorkSheet.Cells(2, 16) = "Fecha de Cita"
                   oWorkSheet.Cells(2, 17) = "Hora de Cita"
                   oWorkSheet.Cells(2, 18) = "F.Atención Cita"
                   oWorkSheet.Cells(2, 19) = "H.Atención Cita"
                   oWorkSheet.Cells(2, 20) = ""
                   oWorkSheet.Cells(2, 21) = "Consultorio"
                   oWorkSheet.Cells(2, 23) = ""
                   oWorkSheet.Cells(3, 23) = ""
                   oWorkSheet.Cells(3, 24) = ""
                   oWorkSheet.Cells(2, 25) = ""
                   oWorkSheet.Cells(2, 26) = ""
                   If lbSoloCitados = True Then
                      oWorkSheet.Cells(2, 7) = "Teléfono"
                   End If
                Else
                   oWorkSheet.Cells(2, 41) = "Médico de Egreso"
                End If
                If ml_IdTipoEspecialidad = 1 Then
                    oWorkSheet.Cells(2, 30) = "Otros Dx"
                    oWorkSheet.Cells(3, 30) = "Dx3"
                    oWorkSheet.Cells(3, 31) = "Dx4"
                    oWorkSheet.Cells(3, 32) = "Dx5"
                    oWorkSheet.Cells(3, 33) = "Dx6"
                End If
                oWorkSheet.Cells(2, 44) = "Ref. Origen Dx"
                oWorkSheet.Cells(2, 45) = "Ref. Destino"
                oWorkSheet.Cells(2, 46) = "Ref. Destino Nro"
                oWorkSheet.Cells(2, 47) = "Usuario Digitador"
                oWorkSheet.Cells(2, 48) = "Fecha Digitador"
                If ml_IdTipoEspecialidad = 2 Then
                   oWorkSheet.Cells(2, 49) = "Emergencia N°"
                End If
                If ml_IdTipoEspecialidad = 1 Then
                   oWorkSheet.Cells(2, 50) = "Hr.Ini.Atención"
                End If
                oWorkSheet.Cells(2, 51) = "Dni Paciente"
                If ml_IdTipoEspecialidad = 1 Then
                   oWorkSheet.Cells(2, 52) = "Forma Genera Cita"
                End If
                
            End If
            f = 4
            '
            '
            '
            rsReporte.MoveFirst
        
            Do While Not rsReporte.EOF
                lbSeguir = True
                If lbEmergSoloLosQuePasaron24hr = True Then
                   If DateDiff("h", rsReporte!FechaIngreso & " " & rsReporte!HoraIngreso, rsReporte!FechaEgreso & " " & rsReporte!horaEgreso) <= 24 Then
                      lbSeguir = False
                   End If
                End If
                If lbSeguir = True Then
                    '
                    If lbMuestraHistoriasDuplicadas = True Then
                        If lbPrimerError = True Then
                           lbPrimerError = False
                           oRsErrores.AddNew
                           oRsErrores.Fields!historia = rsReporte!NroHistoriaClinica
                           oRsErrores.Fields!fechaIng = rsReporte!FechaIngreso
                           oRsErrores.Fields!Error = Space(1)
                           oRsErrores.Update
                        Else
                           lbNoExisteH = True
                           oRsErrores.MoveFirst
                           Do While Not oRsErrores.EOF
                              If oRsErrores!historia = rsReporte!NroHistoriaClinica And oRsErrores!fechaIng = rsReporte!FechaIngreso Then
                                 lbNoExisteH = False
                                 Exit Do
                              End If
                              oRsErrores.MoveNext
                           Loop
                           If lbNoExisteH = True Then
                                oRsErrores.AddNew
                                oRsErrores.Fields!historia = rsReporte!NroHistoriaClinica
                                oRsErrores.Fields!fechaIng = rsReporte!FechaIngreso
                                oRsErrores.Fields!Error = Space(1)
                           Else
                                oRsErrores.Fields!Error = "(misma F.Ingreso)"
                           End If
                           oRsErrores.Update
                        End If
                    End If
                    '
                
                
                    lnIdCuentaAtencion = rsReporte!idCuentaAtencion
                    lcCodServEgreso = ""
                    lcServEgreso = ""
                    If Not IsNull(rsReporte.Fields("idServicioEgreso").Value) Then
                        Set rsServicio = ServiciosSelecionarPorId(rsReporte.Fields("idServicioEgreso").Value)
                        lcCodServEgreso = rsServicio.Fields("codigo").Value
                        lcServEgreso = rsServicio.Fields("nombre").Value
                        rsServicio.Close
                    End If
                    '
                    lcReferenciaOrigen = ""
                    lcReferenciaOrigenNro = ""
                    lcReferenciaDestino = ""
                    lcReferenciaDestinoNro = ""
                    lcNroEmergencia = ""   'debb-13/07/2016
                    Set rsServicio = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarXidAtencion(rsReporte!idAtencion)
                    If rsServicio.RecordCount > 0 Then
                        lcReferenciaOrigen = IIf(IsNull(rsServicio!EstablecimientoOrigen), "", rsServicio!EstablecimientoOrigen)
                        lcReferenciaOrigenNro = IIf(IsNull(rsServicio!nroReferenciaOrigen), "", rsServicio!nroReferenciaOrigen)
                        lcReferenciaDestino = IIf(IsNull(rsServicio!EstablecimientoDestino), "", rsServicio!EstablecimientoDestino)
                        lcReferenciaDestinoNro = IIf(IsNull(rsServicio!nroReferenciaDestino), "", rsServicio!nroReferenciaDestino)
                        If ml_IdTipoEspecialidad = 2 Then
                           lcNroEmergencia = IIf(IsNull(rsServicio!emergenciaCorrelativo), "", rsServicio!emergenciaCorrelativo)
                        End If
                    End If
                    rsServicio.Close
                    '
                    lcUsuario = "": ldUsuarioFecha = ""
                    Set rsServicio = mo_ReglasDeSeguridad.AuditoriaFiltrarCitasPorIdAtencion(rsReporte!idAtencion, _
                                           IIf(ml_IdTipoEspecialidad = 2, sghAdmisionEmergencia, IIf(ml_IdTipoEspecialidad = 3, sghAdmisionHospitalizacion, sghRegistroCitaCE)), _
                                           oConexion)
                    If rsServicio.RecordCount > 0 Then
                       lcUsuario = IIf(IsNull(rsServicio!dUsuario), "", rsServicio!dUsuario)
                       ldUsuarioFecha = IIf(IsNull(rsServicio!fechaHora), "", Format(rsServicio!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
                    End If
                    rsServicio.Close
                    '
                    
                    If lbEsOpenOffice = True Then
                        If Val(lcRenaes) = 6918 Then      'hosp tarapoto
                           Call Feuille.getcellbyposition(0, f - 1).setFormula(IIf(IsNull(rsReporte!NroHistoriaClinica), "", Right("000000" & Trim(str(rsReporte!NroHistoriaClinica)), 2)))
                        Else
                           Call Feuille.getcellbyposition(0, f - 1).setFormula(IIf(IsNull(rsReporte!NroHistoriaClinica), "", rsReporte!NroHistoriaClinica))
                        End If
                        Call Feuille.getcellbyposition(1, f - 1).setFormula(rsReporte!idCuentaAtencion)
                        Call Feuille.getcellbyposition(2, f - 1).setFormula(IIf(IsNull(rsReporte!ApellidoPaterno), "", rsReporte!ApellidoPaterno))
                        Call Feuille.getcellbyposition(3, f - 1).setFormula(IIf(IsNull(rsReporte!ApellidoMaterno), "", rsReporte!ApellidoMaterno))
                        Call Feuille.getcellbyposition(4, f - 1).setFormula(IIf(IsNull(rsReporte!PrimerNombre), "", rsReporte!PrimerNombre))
                        Call Feuille.getcellbyposition(5, f - 1).setFormula(IIf(IsNull(rsReporte!SegundoNombre), "", rsReporte!SegundoNombre))
                        Call Feuille.getcellbyposition(6, f - 1).setFormula(IIf(IsNull(rsReporte!Sexo), "", rsReporte!Sexo))
                        Call Feuille.getcellbyposition(7, f - 1).setFormula(IIf(IsNull(rsReporte!Edad), "", rsReporte!Edad))
                        Call Feuille.getcellbyposition(8, f - 1).setFormula(IIf(IsNull(rsReporte!TipoEdad), "", rsReporte!TipoEdad))
                        Call Feuille.getcellbyposition(9, f - 1).setFormula(IIf(IsNull(rsReporte!Departamento), "", rsReporte!Departamento))
                        Call Feuille.getcellbyposition(10, f - 1).setFormula(IIf(IsNull(rsReporte!provincia), "", rsReporte!provincia))
                        Call Feuille.getcellbyposition(11, f - 1).setFormula(IIf(IsNull(rsReporte!distrito), "", rsReporte!distrito))
                        Call Feuille.getcellbyposition(12, f - 1).setFormula(IIf(IsNull(rsReporte!CentroPoblado), "", rsReporte!CentroPoblado))
                        Call Feuille.getcellbyposition(13, f - 1).setFormula(IIf(IsNull(rsReporte!idFuenteFinanciamiento), "", rsReporte!idFuenteFinanciamiento))
                        Call Feuille.getcellbyposition(14, f - 1).setFormula(IIf(IsNull(rsReporte!DescFuenteFinanciamiento), "", rsReporte!DescFuenteFinanciamiento))
                        Call Feuille.getcellbyposition(15, f - 1).setFormula(IIf(IsNull(rsReporte!FechaIngreso), "", rsReporte!FechaIngreso))
                        Call Feuille.getcellbyposition(16, f - 1).setFormula(IIf(IsNull(rsReporte!HoraIngreso), "", "'" & rsReporte!HoraIngreso))
                        Call Feuille.getcellbyposition(17, f - 1).setFormula(IIf(IsNull(rsReporte!FechaEgreso), "", rsReporte!FechaEgreso))
                        Call Feuille.getcellbyposition(18, f - 1).setFormula(IIf(IsNull(rsReporte!horaEgreso), "", "'" & rsReporte!horaEgreso))
                        Call Feuille.getcellbyposition(19, f - 1).setFormula(IIf(IsNull(rsReporte!CondicionAlta), "", rsReporte!CondicionAlta))
                        Call Feuille.getcellbyposition(20, f - 1).setFormula(IIf(IsNull(rsReporte!ServicioIngresoCodigo), "", rsReporte!ServicioIngresoCodigo))
                        Call Feuille.getcellbyposition(21, f - 1).setFormula(IIf(IsNull(rsReporte!ServicioIngresoNombre), "", rsReporte!ServicioIngresoNombre))
                        Call Feuille.getcellbyposition(22, f - 1).setFormula(lcCodServEgreso)
                        Call Feuille.getcellbyposition(23, f - 1).setFormula(lcServEgreso)
                        Call Feuille.getcellbyposition(24, f - 1).setFormula(IIf(IsNull(rsReporte!TipoAlta), "", rsReporte!TipoAlta))
                        
                    Else
                        If Val(lcRenaes) = 6918 Then      'hosp tarapoto
                           oWorkSheet.Cells(f, 1) = IIf(IsNull(rsReporte!NroHistoriaClinica), "", Right("000000" & Trim(str(rsReporte!NroHistoriaClinica)), 6))
                        Else
                           oWorkSheet.Cells(f, 1) = IIf(IsNull(rsReporte!NroHistoriaClinica), "", HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte!NroHistoriaClinica)), True))
                        End If
                        oWorkSheet.Cells(f, 2) = rsReporte!idCuentaAtencion
                        oWorkSheet.Cells(f, 3) = IIf(IsNull(rsReporte!ApellidoPaterno), "", rsReporte!ApellidoPaterno)
                        oWorkSheet.Cells(f, 4) = IIf(IsNull(rsReporte!ApellidoMaterno), "", rsReporte!ApellidoMaterno)
                        oWorkSheet.Cells(f, 5) = IIf(IsNull(rsReporte!PrimerNombre), "", rsReporte!PrimerNombre)
                        oWorkSheet.Cells(f, 6) = IIf(IsNull(rsReporte!SegundoNombre), "", rsReporte!SegundoNombre)
                        If lbSoloCitados = True Then
                           oWorkSheet.Cells(f, 7) = IIf(IsNull(rsReporte!telefono), "", rsReporte!telefono)
                           mo_Pacientes.idPaciente = rsReporte!idPaciente
                           If oPaciente.SeleccionarPorId(mo_Pacientes) = True Then
                           lcMensajeCelular = "RECUERDE QUE TIENE CITA para " & rsReporte!FechaIngreso & " " & rsReporte!HoraIngreso & _
                                               " (N° Cuenta: " & rsReporte!idCuentaAtencion & ") (Consultorio: " & _
                                                rsReporte!ServicioIngresoNombre & _
                                                ")  en " & wxParametro205
                           oMensajeCelular.MensajeCelularEnviar mo_Pacientes, rsReporte!idCuentaAtencion, lcMensajeCelular, _
                                            "RCITAS", oConexion
                           End If
                           
                        Else
                           oWorkSheet.Cells(f, 7) = IIf(IsNull(rsReporte!Sexo), "", rsReporte!Sexo)
                        End If
                        oWorkSheet.Cells(f, 8) = IIf(IsNull(rsReporte!Edad), "", rsReporte!Edad)
                        oWorkSheet.Cells(f, 9) = IIf(IsNull(rsReporte!TipoEdad), "", rsReporte!TipoEdad)
                        oWorkSheet.Cells(f, 10) = IIf(IsNull(rsReporte!Departamento), "", rsReporte!Departamento)
                        oWorkSheet.Cells(f, 11) = IIf(IsNull(rsReporte!provincia), "", rsReporte!provincia)
                        oWorkSheet.Cells(f, 12) = IIf(IsNull(rsReporte!distrito), "", rsReporte!distrito)
                        oWorkSheet.Cells(f, 13) = IIf(IsNull(rsReporte!CentroPoblado), "", rsReporte!CentroPoblado)
                        oWorkSheet.Cells(f, 14) = IIf(IsNull(rsReporte!idFuenteFinanciamiento), "", rsReporte!idFuenteFinanciamiento)
                        oWorkSheet.Cells(f, 15) = IIf(IsNull(rsReporte!DescFuenteFinanciamiento), "", rsReporte!DescFuenteFinanciamiento)
                        'SCCQ 29/05/2020 Cambio20 Inicio
                        oWorkSheet.Cells(f, 16) = IIf(IsNull(rsReporte!FechaIngreso), "", Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY))
                        'SCCQ 29/05/2020 Cambio20 Fin
                        oWorkSheet.Cells(f, 17) = IIf(IsNull(rsReporte!HoraIngreso), "", rsReporte!HoraIngreso)
                        'SCCQ 29/05/2020 Cambio20 Inicio
                        oWorkSheet.Cells(f, 18) = IIf(IsNull(rsReporte!FechaEgreso), "", Format(rsReporte!FechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY))
                        'SCCQ 29/05/2020 Cambio20 Fin
                        oWorkSheet.Cells(f, 19) = IIf(IsNull(rsReporte!horaEgreso), "", rsReporte!horaEgreso)
                        oWorkSheet.Cells(f, 20) = IIf(IsNull(rsReporte!CondicionAlta), "", rsReporte!CondicionAlta)
                        oWorkSheet.Cells(f, 21) = IIf(IsNull(rsReporte!ServicioIngresoCodigo), "", rsReporte!ServicioIngresoCodigo)
                        oWorkSheet.Cells(f, 22) = IIf(IsNull(rsReporte!ServicioIngresoNombre), "", rsReporte!ServicioIngresoNombre)
                        oWorkSheet.Cells(f, 23) = lcCodServEgreso  'IIf(IsNull(rsReporte!ServicioEgresoCodigo), "", rsReporte!ServicioEgresoCodigo)
                        oWorkSheet.Cells(f, 24) = lcServEgreso  'IIf(IsNull(rsReporte!ServicioEgresoNombre), "", rsReporte!ServicioEgresoNombre)
                        oWorkSheet.Cells(f, 25) = IIf(IsNull(rsReporte!TipoAlta), "", rsReporte!TipoAlta)
                        
                    End If
                    lnDiasEstancia = 0
                    If ml_IdTipoEspecialidad <> 1 Then
                       If Val(lcRenaes) = 6918 Then      'hosp tarapoto
                          lnDiasEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax, True)
                       Else
                          lnDiasEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax)
                       End If
                    End If
                    If lbEsOpenOffice = True Then
                       Call Feuille.getcellbyposition(25, f - 1).setFormula(Trim(str(lnDiasEstancia)))
                    Else
                       oWorkSheet.Cells(f, 26) = lnDiasEstancia
                    End If
                    If lbConsiderarDX = True Then
                        c = 27
                        iCantidadProcedimientos = 1
                        If ml_IdTipoEspecialidad <> 1 Then
                            Set rsDiagnosticos = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
                        Else
                            Set rsDiagnosticos = mo_ReglasAdmision.BuscaCEAtencionesDx(rsReporte!idAtencion)
                        End If
                        'rsDiagnosticos.Filter = "idAtencion=" & rsReporte!idAtencion
                        
                        If rsDiagnosticos.RecordCount > 0 Then
                           
                           rsDiagnosticos.MoveFirst
                           Do While Not rsDiagnosticos.EOF
                                
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(c - 1, f - 1).setFormula(IIf(IsNull(rsDiagnosticos!CodigoDx), "", rsDiagnosticos!CodigoDx))
                                Else
                                    oWorkSheet.Cells(f, c) = IIf(IsNull(rsDiagnosticos!CodigoDx), "", rsDiagnosticos!CodigoDx)
                                End If
                                'Si es Dx principal agrega descripcion
                                If iCantidadProcedimientos = 1 Then
                                    If lbEsOpenOffice = True Then
                                        Call Feuille.getcellbyposition(c, f - 1).setFormula(IIf(IsNull(rsDiagnosticos!DescripcionDx), "", rsDiagnosticos!DescripcionDx))
                                    Else
                                        oWorkSheet.Cells(f, c + 1) = IIf(IsNull(rsDiagnosticos!DescripcionDx), "", rsDiagnosticos!DescripcionDx)
                                    End If
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
                        End If
                        rsDiagnosticos.Close
                    End If
            
                    If lbConsiderarCPT = True Then
                        c = 34
                        iCantidadProcedimientos = 1
                        Set rsProcedimientos = mo_AdminReportes.ReporteAtencionesProcedimientos(rsReporte!idAtencion)
                        'rsProcedimientos.Filter = "idAtencion=" & rsReporte!idAtencion
                        If rsProcedimientos.RecordCount > 0 Then
                            rsProcedimientos.MoveFirst
                            Do While Not rsProcedimientos.EOF
                                If Not IsNull(rsProcedimientos!codigo) Then
                                    lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsProcedimientos!codigo)
                                    If lbEsOpenOffice = True Then
                                        Call Feuille.getcellbyposition(c - 1, f - 1).setFormula(rsProcedimientos!codigo & IIf(lcOpcs = "", "", "  (" & lcOpcs & ")"))
                                    Else
                                        oWorkSheet.Cells(f, c) = rsProcedimientos!codigo & IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                                    End If
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
                    
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(38, f - 1).setFormula(IIf(IsNull(rsReporte!telefono), "", rsReporte!telefono))
                        Call Feuille.getcellbyposition(39, f - 1).setFormula(IIf(IsNull(rsReporte!DireccionDomicilio), "", rsReporte!DireccionDomicilio))
                        Call Feuille.getcellbyposition(41, f - 1).setFormula(lcReferenciaOrigen)
                        Call Feuille.getcellbyposition(42, f - 1).setFormula(lcReferenciaOrigenNro)
                        Set rsDiagnosticos = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarDxRefOrigen(rsReporte!idAtencion, oConexion)
                        If rsDiagnosticos.RecordCount > 0 Then
                           If Not IsNull(rsDiagnosticos!CodigoCie2004) Then
                              Call Feuille.getcellbyposition(43, f - 1).setFormula(Trim(rsDiagnosticos!CodigoCie2004) & " " & rsDiagnosticos!Descripcion)
                           End If
                        End If
                        rsDiagnosticos.Close
                        Call Feuille.getcellbyposition(44, f - 1).setFormula(lcReferenciaDestino)
                        Call Feuille.getcellbyposition(45, f - 1).setFormula(lcReferenciaDestinoNro)
                        Call Feuille.getcellbyposition(46, f - 1).setFormula(lcUsuario)
                        Call Feuille.getcellbyposition(47, f - 1).setFormula(ldUsuarioFecha)
                        Call Feuille.getcellbyposition(48, f - 1).setFormula(lcNroEmergencia)
                    Else
                        oWorkSheet.Cells(f, 39) = IIf(IsNull(rsReporte!telefono), "", rsReporte!telefono)
                        oWorkSheet.Cells(f, 40) = IIf(IsNull(rsReporte!DireccionDomicilio), "", rsReporte!DireccionDomicilio)
                        oWorkSheet.Cells(f, 42) = lcReferenciaOrigen
                        oWorkSheet.Cells(f, 43) = lcReferenciaOrigenNro
                        If lcReferenciaOrigenNro <> "" Then
                            If ml_IdTipoEspecialidad <> 1 Then
                                Set rsDiagnosticos = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(rsReporte!idAtencion, oConexion)
                                rsDiagnosticos.Filter = "idSubClasificacionDx=null"
                                If rsDiagnosticos.RecordCount > 0 Then
                                   oWorkSheet.Cells(f, 44) = Trim(rsDiagnosticos!CodigoCie2004) & " " & rsDiagnosticos!Descripcion
                                End If
                                rsDiagnosticos.Close
                            Else
                                Set rsDiagnosticos = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarDxRefOrigen(rsReporte!idAtencion, oConexion)
                                If rsDiagnosticos.RecordCount > 0 Then
                                   If Not IsNull(rsDiagnosticos!CodigoCie2004) Then
                                      oWorkSheet.Cells(f, 44) = Trim(rsDiagnosticos!CodigoCie2004) & " " & rsDiagnosticos!Descripcion
                                   End If
                                End If
                                rsDiagnosticos.Close
                            End If
                        End If
                        oWorkSheet.Cells(f, 45) = lcReferenciaDestino
                        oWorkSheet.Cells(f, 46) = lcReferenciaDestinoNro
                        oWorkSheet.Cells(f, 47) = lcUsuario
                        oWorkSheet.Cells(f, 48) = ldUsuarioFecha
                        oWorkSheet.Cells(f, 49) = lcNroEmergencia
                        '
                        If ml_IdTipoEspecialidad = 1 And Not IsNull(rsReporte!horaInicioAtencion) Then
                           oWorkSheet.Cells(f, 50) = rsReporte!horaInicioAtencion
                        End If
                        '
                    End If
                    '
                    If ml_IdTipoEspecialidad = 1 Then
                        Set rsProcedimientos = mo_AdminServComunes.AtencionesSeleccionarMedicoPorCuenta(rsReporte!idCuentaAtencion)
                    Else
                        Set rsProcedimientos = mo_AdminServComunes.AtencionesSeleccionarMedicoEgresoPorCuenta(rsReporte!idCuentaAtencion)
                    End If
                    If rsProcedimientos.RecordCount > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(40, f - 1).setFormula(IIf(IsNull(rsProcedimientos!dmedico), "", rsProcedimientos!dmedico))
                        Else
                            oWorkSheet.Cells(f, 41) = IIf(IsNull(rsProcedimientos!dmedico), "", rsProcedimientos!dmedico)
                        End If
                    End If
                    '
                    oWorkSheet.Cells(f, 51) = IIf(IsNull(rsReporte!nroDocumento), "", rsReporte!nroDocumento)
                    '
                    If ml_IdTipoEspecialidad = 1 Then
                       oWorkSheet.Cells(f, 52) = mo_ReglasAdmision.CitasFormaCitaSeleccionaXcuenta(rsReporte!idCuentaAtencion, oConexion)
                    End If
                    
                    f = f + 1
                    '
                    'mo_ProgressRpt.Value = f
                    RaiseEvent ProgressActualizaValor(f, lRecordCount)
                End If
                rsReporte.MoveNext
            Loop
        End If
        
        iFila = f
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":AQ" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(0, iFila - 1).setFormula("Nº Historias Clínicas: " + Trim(str(lRecordCount)))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 43
            oWorkSheet.Cells(iFila, 1).Value = "Nº Historias Clínicas: " + Trim(str(lRecordCount))
        End If
        
        '
        f = iFila
        oRsErrores.Filter = "error<>' '"
        If oRsErrores.RecordCount > 0 Then
           f = f + 1
           oWorkSheet.Cells(f, 1) = "ERRORES"
           f = f + 1
           oWorkSheet.Cells(f, 1) = "======="
           f = f + 1
           oRsErrores.MoveFirst
           Do While Not oRsErrores.EOF
                lcSql = oRsErrores!Error & "  Historia: " & oRsErrores!historia & " , F.Ingreso:" & _
                                         oRsErrores!fechaIng & " solucionar en FACTURACION->ESTADO CUENTA"
                oWorkSheet.Cells(f, 1) = lcSql
                f = f + 1
                oRsErrores.MoveNext
           Loop
           iFila = f
        End If
        '
        
        
        If lRecordCount > 0 Then
            If lbEsOpenOffice = True Then
            Else
                'Abrimos el archivo temporal en otra hoja
                oWorkSheet.Cells.Select
                oWorkSheet.Cells.EntireColumn.AutoFit
            End If
        End If
        If lbEsOpenOffice = True Then
        Else
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
        End If
        'Falta que salga el nombre del responsable
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 38
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        Else
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
            'Set oExcel = Nothing
        End If
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
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
    Set rsDiagnosticos = Nothing
    Set rsProcedimientos = Nothing
    Set lcBuscaParametro = Nothing
    Set oRsErrores = Nothing
    Set oMensajeCelular = Nothing
    Set mo_Pacientes = Nothing
    Set oPaciente = Nothing
    
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description & Chr(13) & "N° Cuenta: " & lnIdCuentaAtencion
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
    Dim oFormulario As New ReportesEgresosHosp
    oFormulario.IdTipoReporte = ml_IdTipoReporte
    oFormulario.cmbConsiderar.ListIndex = ml_idTipoServicio
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHSolicPorServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historias solicitadas por consultorio
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_IdTipoHistoria As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdTipoHistoria(lValue As Long)
    ml_IdTipoHistoria = lValue
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
    Dim oFormulario As New AHCSolicPorServ
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

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
    Set rsReporte = HistoriasClinicasSeleccionarPorFechaRequerimiento(mda_FechaInicio, mda_FechaFin, ml_IdResponsable)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe informaci�n con esos Datos", vbInformation, "Resultado"
    Else
        GenerarRecordsetTemporal
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lnIdPaciente = rsReporte.Fields("IdPaciente").Value
            Call AgregaUltimosCuatroSolicitudesDeHistoria(rsReporte)
            Do While Not rsReporte.EOF And lnIdPaciente = rsReporte.Fields("IdPaciente").Value
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
        Loop
        If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existe informaci�n con esos Datos", vbInformation, "Resultado"
        Else
            mrs_Tmp.Sort = "Paciente asc,orden"
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\HCahSolicPorServicio.ods"
'                FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                Chemin = "file:///" & App.Path & "\Plantillas\"
'                Chemin = Replace(Chemin, "\", "/")
'                Fichier = Chemin & "/OpenOffice.ods"
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
                Set StyleFamilies = Document.StyleFamilies
                Set PageStyles = StyleFamilies.getByName("PageStyles")
                Set DefPage = PageStyles.getByName("Default")
                Set Hcontent = DefPage.RightPageHeaderContent
                Set Htext = Hcontent.centerText
                mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pas�ndole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCahSolicPorServicio.xls")
                oWorkBookPlantilla.Worksheets("AHSolicPorServicio").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
                
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(8, 1).setFormula("HISTORIAS CLINICAS SOLICITADAS POR SERVICIO")
                Call Feuille.getcellbyposition(4, 2).setFormula(ml_TextoDelFiltro)
                Call Feuille.getcellbyposition(17, 5).setFormula("")
                iFila = 7
                lnNumTotal = 0
            Else
                oWorkSheet.Cells(2, 9).Value = "HISTORIAS CLINICAS SOLICITADAS POR SERVICIO"
                oWorkSheet.Cells(3, 5).Value = ml_TextoDelFiltro
                oWorkSheet.Cells(6, 18).Value = ""
                iFila = 7
                lnNumTotal = 0
            End If
            
            Do While Not mrs_Tmp.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("Servicio").Value)
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(mrs_Tmp.Fields("nroHistoriaClinica").Value)
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(mrs_Tmp.Fields("Paciente").Value)
                'Call Feuille.getcellbyposition(17, iFila - 1).setFormula(mrs_Tmp.Fields("tipo").Value)
                lnNumTotal = lnNumTotal + 1
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula(lnNumTotal)
            Else
                oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("Servicio").Value
                oWorkSheet.Cells(iFila, 9).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp.Fields("nroHistoriaClinica").Value)), True)
                oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("Paciente").Value
                'oWorkSheet.Cells(iFila, 18).Value = mrs_Tmp.Fields("tipo").Value
                lnNumTotal = lnNumTotal + 1
                oWorkSheet.Cells(iFila, 1).Value = lnNumTotal
            End If
            
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":V" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 22
            End If
            lcPaciente = mrs_Tmp.Fields("Paciente").Value
               Do While Not mrs_Tmp.EOF And lcPaciente = mrs_Tmp.Fields("Paciente").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("Servicio").Value)
                        Call Feuille.getcellbyposition(18, iFila - 1).setFormula("'" + CStr(mrs_Tmp.Fields("fRequerida").Value))
                        Call Feuille.getcellbyposition(19, iFila - 1).setFormula("'" + CStr(mrs_Tmp.Fields("fsolicitud").Value))
                        Call Feuille.getcellbyposition(20, iFila - 1).setFormula(mrs_Tmp.Fields("servicio2").Value)
                        iFila = iFila + 1
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("Servicio").Value
                        oWorkSheet.Cells(iFila, 19).Value = mrs_Tmp.Fields("fRequerida").Value
                        oWorkSheet.Cells(iFila, 20).Value = mrs_Tmp.Fields("fsolicitud").Value
                        oWorkSheet.Cells(iFila, 21).Value = mrs_Tmp.Fields("servicio2").Value
                        iFila = iFila + 1
                    End If
                    mrs_Tmp.MoveNext
                    If mrs_Tmp.EOF Then
                       Exit Do
                    End If
               Loop
            Loop
            
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":V" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("N� Historias Cl�nicas: " + Trim(Str(lnNumTotal)))
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 22
                oWorkSheet.Cells(iFila, 2).Value = "N� Historias Cl�nicas: " + Trim(Str(lnNumTotal))
            End If
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 0
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 22
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
                MsgBox "El Reporte se gener� en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                'Falta que salga el nombre del responsable
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
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
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuraci�n en el men� Inicio de Windows, haga clic en Impresoras y despu�s haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia cl�nica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select
    Exit Sub
End Sub

Sub GenerarRecordsetTemporal()
    With mrs_Tmp
          .Fields.Append "Servicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "NroHistoriaClinica", adInteger, 4, adFldIsNullable
          .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
          .Fields.Append "tipo", adVarChar, 50, adFldIsNullable
          .Fields.Append "FRequerida", adDate, , adFldIsNullable
          .Fields.Append "fsolicitud", adDate, , adFldIsNullable
          .Fields.Append "Servicio2", adVarChar, 150, adFldIsNullable
          .Fields.Append "fecha2", adDate, , adFldIsNullable
          .Fields.Append "orden", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Function HistoriasClinicasSeleccionarPorFechaRequerimiento(ldFechaReqIni As Date, ldFechaReqFin As Date, lnIdEmpleado As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set HistoriasClinicasSeleccionarPorFechaRequerimiento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasSolicitadasSeleccionarPorFechaRequerimiento"
        Set oParameter = .CreateParameter("@FechaReqIni", adDate, adParamInput, 0, ldFechaReqIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaReqFin", adDate, adParamInput, 0, ldFechaReqFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set HistoriasClinicasSeleccionarPorFechaRequerimiento = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte t�cnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Sub AgregaUltimosCuatroSolicitudesDeHistoria(rsReporte1 As ADODB.Recordset)
On Error GoTo ManejadorDeError2
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim lnNumero As Integer
Dim ms_MensajeError As String
Dim lcPaciente As String
Dim lnNroHistoriaClinica As Long
Dim lcTipoHistoria As String
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    'Datos del Paciente
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, rsReporte1.Fields("IdPaciente").Value): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   lcPaciente = "": lnNroHistoriaClinica = 0: lcTipoHistoria = ""
   If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      lcPaciente = mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(oRecordset!ApellidoPaterno), mo_ReporteUtil.NullToVacio(oRecordset!ApellidoMaterno), mo_ReporteUtil.NullToVacio(oRecordset!PrimerNombre), mo_ReporteUtil.NullToVacio(oRecordset!SegundoNombre))
      lnNroHistoriaClinica = IIf(IsNull(oRecordset!NroHistoriaClinica), 0, oRecordset!NroHistoriaClinica)
      lcTipoHistoria = mo_ReporteUtil.NullToVacio(oRecordset!dTipoHistoria)
   End If
   oRecordset.Close
   mrs_Tmp.AddNew
   mrs_Tmp.Fields("Servicio").Value = rsReporte1!CodServ & " - " & rsReporte1!DServicio
   mrs_Tmp.Fields("nroHistoriaClinica").Value = lnNroHistoriaClinica
   mrs_Tmp.Fields("Paciente").Value = lcPaciente
   mrs_Tmp.Fields("tipo").Value = lcTipoHistoria
   mrs_Tmp.Fields("frequerida").Value = rsReporte1.Fields("fechaRequerida").Value
   mrs_Tmp.Fields("fsolicitud").Value = rsReporte1.Fields("fechaSolicitud").Value
   mrs_Tmp.Fields("orden").Value = 0
   mrs_Tmp.Update
   'Datos Historico Solicitudes

   With oCommand
        .Parameters.Delete (0)
        .CommandText = "HistoriasSolicitadasSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, rsReporte1.Fields("IdPaciente").Value): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
      lnNumero = 1
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
         If oRecordset.Fields("idHistoriaSolicitada").Value <> rsReporte1.Fields("idHistoriaSolicitada").Value Then
                   mrs_Tmp.AddNew
                   mrs_Tmp.Fields("Servicio2").Value = oRecordset!CodServ & " - " & oRecordset!DServicio
                   mrs_Tmp.Fields("nroHistoriaClinica").Value = lnNroHistoriaClinica
                   mrs_Tmp.Fields("Paciente").Value = lcPaciente
                   mrs_Tmp.Fields("tipo").Value = lcTipoHistoria
                   mrs_Tmp.Fields("frequerida").Value = oRecordset.Fields("fechaRequerida").Value
                   mrs_Tmp.Fields("fsolicitud").Value = oRecordset.Fields("fechaSolicitud").Value
                   mrs_Tmp.Fields("orden").Value = lnNumero
                   mrs_Tmp.Update
                   lnNumero = lnNumero + 1
          End If
          oRecordset.MoveNext
          If lnNumero > 4 Then
             Exit Do
          End If
      Loop
   End If
Exit Sub
ManejadorDeError2:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte t�cnico", vbInformation, "Error en la interface de acceso a datos"

Exit Sub
End Sub




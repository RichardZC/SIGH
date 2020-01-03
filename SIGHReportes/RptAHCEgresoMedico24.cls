VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHCEgresoMedico24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historias que no regresaron en 24 hr
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_IdResponsable As Long
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim ml_TipoServicio  As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Property Let TipoServicio(lValue As Integer)
    ml_TipoServicio = lValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property
Property Let IdResponsable(lValue As Long)
    ml_IdResponsable = lValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New AHCEgresoMedico24
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String


lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)

    
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
    Set rsReporte = mo_ReglasAdmision.HistoriasClinicasSeleccionarPorFechaEgresoMedico(mda_FechaInicio, mda_FechaFin, "Pacientes", ml_TipoServicio)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        'Filtra datos
        GenerarRecordsetTemporal
        Do While Not rsReporte.EOF
           If HistoriasClinicasQuePasan24HorasSinRetornar(rsReporte.Fields("idPaciente").Value, Now, ml_IdResponsable, rsReporte.Fields("fechaEgreso").Value) = True Then
              mrs_Tmp.AddNew
              mrs_Tmp.Fields("nroHistoriaClinica").Value = rsReporte.Fields("nroHistoriaClinica").Value
              mrs_Tmp.Fields("Paciente").Value = mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
              mrs_Tmp.Fields("Servicio").Value = mo_ReporteUtil.NullToVacio(rsReporte!CodServ) & " - " & mo_ReporteUtil.NullToVacio(rsReporte!DServicio)                ' "  << " & rsReporte.Fields("idpaciente").Value & "  << " & rsReporte.Fields("fechaEgreso").Value
              mrs_Tmp.Update
           End If
           rsReporte.MoveNext
        Loop
        If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            If ml_OrdenFiltro = "HC" Then
               mrs_Tmp.Sort = "nroHistoriaClinica"
            End If
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\AHCegrMed24.ods"
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
                mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\AHCegrMed24.xls")
                oWorkBookPlantilla.Worksheets("AHCegrMed24").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 2).setFormula("ml_TextoDelFiltro")
                iFila = 6
            Else
                oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
                iFila = 6
            End If
            
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("nroHistoriaClinica"))
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("paciente"))
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mrs_Tmp.Fields("servicio"))
                    iFila = iFila + 1
                Else
                    oWorkSheet.Cells(iFila, 2).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp.Fields("nroHistoriaClinica").Value)), True)
                    oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("paciente").Value
                    oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("servicio").Value
                    iFila = iFila + 1
                End If
               mrs_Tmp.MoveNext
            Loop
            
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Historias Clínicas: " + Trim(Str(mrs_Tmp.RecordCount)))
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 8
                oWorkSheet.Cells(iFila, 2).Value = "Nº Historias Clínicas: " + Trim(Str(mrs_Tmp.RecordCount))
            End If
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 8
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
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
        End If
    End If
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
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

Sub GenerarRecordsetTemporal()
    With mrs_Tmp
          .Fields.Append "NroHistoriaClinica", adInteger, 4, adFldIsNullable
          .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
          .Fields.Append "Servicio", adVarChar, 150, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub


Function HistoriasClinicasQuePasan24HorasSinRetornar(lnIdPaciente As Long, ldFechaHoy As Date, lnIdResponsable As Long, ldFechaEgreso As Date) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim lnDateDiff As Integer
Dim ms_MensajeError As String
    HistoriasClinicasQuePasan24HorasSinRetornar = False
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "MovimientosHistoriasClinicasParaDevolverPacienteHospitalizado"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaHoy", adDate, adParamInput, 0, ldFechaHoy): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEmpleadoArchivo", adInteger, adParamInput, 0, lnIdResponsable): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      lnDateDiff = Date - ldFechaEgreso
      If lnDateDiff >= 2 And oRecordset.Fields("idMotivo").Value <> 9 Then
         HistoriasClinicasQuePasan24HorasSinRetornar = True
      End If
   End If
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



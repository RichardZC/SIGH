VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEAtencConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para atenciones por convenio
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Convenio MINSA-ESSALUD
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New EAtencConv
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim mrs_Tmp As New Recordset
Dim iFila As Long: Dim lnNroDias As Integer
Dim lnMes As Integer: Dim lnTotDias As Long: Dim lnAnio As Long
Dim lnTotal As Double: Dim lnPrecio As Double
Dim lcPaciente As String: Dim lcDias As String
Dim lnIdPaciente As Long: Dim lnIdServicio As Long: Dim lnNroHistorias As Long
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

    On Error GoTo ManejadorErrorExcel
    'Filtra los Datos
    Set mrs_Tmp = AtencionesConvenioSeleccionarPorFechaSesion(mda_FechaInicio, mda_FechaFin)
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\EatencConv.ods"
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
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EatencConv.xls")
            oWorkBookPlantilla.Worksheets("EAtencConv").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, 2).setFormula(ml_TextoDelFiltro)
            Else
                oWorkSheet.Cells(3, 2).Value = ml_TextoDelFiltro
            End If
        iFila = 7: lnTotal = 0: lnNroHistorias = 0: lnTotDias = 0
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
           lnMes = Month(mrs_Tmp.Fields!FechaSesion)
           lnAnio = Year(mrs_Tmp.Fields!FechaSesion)
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(SIGHEntidades.DevuelveNombreMes(lnMes))
            Else
                oWorkSheet.Cells(iFila, 3).Value = SIGHEntidades.DevuelveNombreMes(lnMes)
            End If
           iFila = iFila + 1
           Do While Not mrs_Tmp.EOF And lnMes = Month(mrs_Tmp.Fields!FechaSesion) And lnAnio = Year(mrs_Tmp.Fields!FechaSesion)
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("dPaciente").Value)
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("nroOficio").Value)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("nroHistoriaClinica").Value)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("dPaciente").Value
                    oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("nroOficio").Value
                    oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
                End If
                lcPaciente = mrs_Tmp.Fields!dPaciente
                lnIdPaciente = mrs_Tmp.Fields!IdPaciente
                lnNroHistorias = lnNroHistorias + 1
                Do While Not mrs_Tmp.EOF And lnMes = Month(mrs_Tmp.Fields!FechaSesion) And lnAnio = Year(mrs_Tmp.Fields!FechaSesion) And lcPaciente = mrs_Tmp.Fields!dPaciente And lnIdPaciente = mrs_Tmp.Fields!IdPaciente
                   lnPrecio = mrs_Tmp.Fields!ImporteSesion
                   lnNroDias = 0: lcDias = ""
                   lnIdServicio = mrs_Tmp.Fields!idProducto
                   Do While Not mrs_Tmp.EOF And lnMes = Month(mrs_Tmp.Fields!FechaSesion) And lnAnio = Year(mrs_Tmp.Fields!FechaSesion) And lcPaciente = mrs_Tmp.Fields!dPaciente And lnIdPaciente = mrs_Tmp.Fields!IdPaciente And lnIdServicio = mrs_Tmp.Fields!idProducto
                      lnNroDias = lnNroDias + 1
                      lcDias = lcDias & Trim(Str(Day(mrs_Tmp.Fields!FechaSesion))) + ", "
                      mrs_Tmp.MoveNext
                      If mrs_Tmp.EOF Then
                        Exit Do
                      End If
                   Loop
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lcDias)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnNroDias)
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnPrecio)
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(Round(lnNroDias * lnPrecio, 2))
                    Else
                        oWorkSheet.Cells(iFila, 5).Value = lcDias
                        oWorkSheet.Cells(iFila, 6).Value = lnNroDias
                        oWorkSheet.Cells(iFila, 7).Value = lnPrecio
                        oWorkSheet.Cells(iFila, 8).Value = Round(lnNroDias * lnPrecio, 2)
                    End If
                   lnTotal = lnTotal + Round(lnNroDias * lnPrecio, 2)
                   lnTotDias = lnTotDias + lnNroDias
                   iFila = iFila + 1
                   If mrs_Tmp.EOF Then
                      Exit Do
                   End If
                Loop
                iFila = iFila + 1
              If mrs_Tmp.EOF Then
                 Exit Do
              End If
           Loop
        Loop
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Historias Clínicas: " + Trim(Str(lnNroHistorias)))
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnTotDias)
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnTotal)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 8
                oWorkSheet.Cells(iFila, 2).Value = "Nº Historias Clínicas: " + Trim(Str(lnNroHistorias))
                oWorkSheet.Cells(iFila, 6).Value = lnTotDias
                oWorkSheet.Cells(iFila, 8).Value = lnTotal
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
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
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

Function AtencionesConvenioSeleccionarPorFechaSesion(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesConvenioSeleccionarPorFechaSesion = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesConvenioSeleccionarPorFechaSesion"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDate, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDate, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesConvenioSeleccionarPorFechaSesion = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clRptModMaterno2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Materno
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_IdDepartamento As Long
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim mo_ReporteUtil As New ReporteUtil
'Dim mo_ProgressRpt As XP_ProgressBar
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let IdDepartamento(lValue As Long)
    ml_IdDepartamento = lValue
End Property
Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Sub CrearReporteProgramacionMedica(ByVal TotalControles As Integer, ml_lnHWnd As Long)
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnNumFila As Integer
Dim lnEdadSemanas As Integer
Dim lnDiasProtocolo As Integer
Dim ldFechaReqVisita As Date
Dim lbEsOpenOffice As Boolean
Dim lcSql As String

Dim oConexion As New Connection
oConexion.CommandTimeout = 300
oConexion.Open SIGHEntidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    
    
    
Dim oRsTmp As New Recordset
On Error GoTo ManejadorError
    
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
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If

    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\ELiquidacion.ods"
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
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\RptModMaterno2.xls")
        oWorkBookPlantilla.Worksheets("CEatencionPaciente").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    '
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(3, 0).setFormula("REPORTE DE SEGUIMIENTO DE GESTANTES CON MENOS DE " & TotalControles & " CONTROLES")
        Call Feuille.getcellbyposition(9, 1).setFormula(lcBuscaParametro.RetornaFechaServidorSQL)
        Call Feuille.getcellbyposition(9, 2).setFormula(lcBuscaParametro.RetornaFechaServidorSQL)
    Else
        oWorkSheet.Cells(1, 4).Value = "REPORTE DE SEGUIMIENTO DE GESTANTES CON MENOS DE " & TotalControles & " CONTROLES"
        oWorkSheet.Cells(2, 10).Value = lcBuscaParametro.RetornaFechaServidorSQL
        oWorkSheet.Cells(3, 10).Value = lcBuscaParametro.RetornaHoraServidorSQL
    End If
    Set oRsTmp = mo_ReglasComunes.ProgramaControlMaternoConsultarTotalControles(TotalControles, oConexion)
    If oRsTmp.RecordCount > 0 Then
        lnNumFila = 10
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
            lnEdadSemanas = DevuelveEdadGestacional(CDate(oRsTmp.Fields!FUM), CDate(lcBuscaParametro.RetornaFechaServidorSQL))
            lnDiasProtocolo = DevuelveProtocoloDias(lnEdadSemanas)
            ldFechaReqVisita = CDate(oRsTmp.Fields!FechaUltimoControl) + lnDiasProtocolo
            
            If CDate(lcBuscaParametro.RetornaFechaServidorSQL) >= ldFechaReqVisita Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, lnNumFila - 1).setFormula(oRsTmp.Fields!FPP)
                    Call Feuille.getcellbyposition(2, lnNumFila - 1).setFormula(oRsTmp.Fields!nrohistoriaclinica)
                    Call Feuille.getcellbyposition(3, lnNumFila - 1).setFormula(Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & ", " & Trim(oRsTmp.Fields!PrimerNombre) & " " & Trim(oRsTmp.Fields!SegundoNombre) & " " & Trim(oRsTmp.Fields!TercerNombre))
                    Call Feuille.getcellbyposition(4, lnNumFila - 1).setFormula(lnEdadSemanas)
                    Call Feuille.getcellbyposition(5, lnNumFila - 1).setFormula(oRsTmp.Fields!distrito)
                    Call Feuille.getcellbyposition(6, lnNumFila - 1).setFormula(oRsTmp.Fields!telefono)
                    Call Feuille.getcellbyposition(7, lnNumFila - 1).setFormula(oRsTmp.Fields!email)
                    Call Feuille.getcellbyposition(8, lnNumFila - 1).setFormula(oRsTmp.Fields!FechaUltimoControl)
                    Call Feuille.getcellbyposition(9, lnNumFila - 1).setFormula(CStr(ldFechaReqVisita))
                Else
                    oWorkSheet.Cells(lnNumFila, 2).Value = oRsTmp.Fields!FPP
                    oWorkSheet.Cells(lnNumFila, 3).Value = oRsTmp.Fields!nrohistoriaclinica
                    oWorkSheet.Cells(lnNumFila, 4).Value = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & ", " & Trim(oRsTmp.Fields!PrimerNombre) & " " & Trim(oRsTmp.Fields!SegundoNombre) & " " & Trim(oRsTmp.Fields!TercerNombre)
                    oWorkSheet.Cells(lnNumFila, 5).Value = lnEdadSemanas
                    oWorkSheet.Cells(lnNumFila, 6).Value = oRsTmp.Fields!distrito
                    oWorkSheet.Cells(lnNumFila, 7).Value = oRsTmp.Fields!telefono
                    oWorkSheet.Cells(lnNumFila, 8).Value = oRsTmp.Fields!email
                    oWorkSheet.Cells(lnNumFila, 9).Value = oRsTmp.Fields!FechaUltimoControl
                    oWorkSheet.Cells(lnNumFila, 10).Value = CStr(ldFechaReqVisita)
                End If
                lnNumFila = lnNumFila + 1
            End If
            oRsTmp.MoveNext
        Loop
    End If
    '
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 11
        PrintArea(0).EndRow = lnNumFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
        'oWorkBook.Close SaveChanges:=False
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
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    
    Set oRsTmp = Nothing
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

Public Function DevuelveProtocoloDias(lnEdadSemanas As Integer) As Integer
    DevuelveProtocoloDias = 0
    If lnEdadSemanas >= 1 And lnEdadSemanas <= 29 Then
        DevuelveProtocoloDias = 30
    End If
    If lnEdadSemanas >= 30 And lnEdadSemanas <= 36 Then
        DevuelveProtocoloDias = 15
    End If
    If lnEdadSemanas >= 37 And lnEdadSemanas <= 42 Then
        DevuelveProtocoloDias = 7
    End If
    If lnEdadSemanas >= 43 Then
        DevuelveProtocoloDias = 2
    End If
End Function

Sub EjecutaFormulario()
    Dim oFormulario As New RptModMaterno2
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

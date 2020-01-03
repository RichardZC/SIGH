VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clAtencionesCenso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: MINISTERIO DE SALUD
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para reporte de atenciones censo de hospitalizados
'        Programado por: Cachay F
'        Fecha: Febrero 2015
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

Sub CrearReporte(ByVal OrsDatosHosp As Recordset, ml_lnHWnd As Long)
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnNumFila As Long
Dim lbEsOpenOffice As Boolean
Dim lcSql As String

Dim oConexion As New Connection
oConexion.CommandTimeout = 300
oConexion.Open sighEntidades.CadenaConexion
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
        Dim range_row As Excel.range
        Dim borders As Excel.borders
    End If

    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\RepCensoHospitalario.ods"
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
        
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\RepCensoHospitalario.xls")
        oWorkBookPlantilla.Worksheets("Reporte").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    '
'    Set oRsTmp = mo_reglasComunes.ProgramaControlMaternoConsultarTotalControles(2, oConexion)
    If OrsDatosHosp.RecordCount > 0 Then
        lnNumFila = 2
        OrsDatosHosp.MoveFirst
        Do While Not OrsDatosHosp.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!NroDocumento)
                Call Feuille.getcellbyposition(1, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!Apellido_Paterno)
                Call Feuille.getcellbyposition(2, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!Apellido_Materno)
                Call Feuille.getcellbyposition(3, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!Nombres)
                Call Feuille.getcellbyposition(4, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!Fec_Nac)
                Call Feuille.getcellbyposition(5, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!NroCama)
                Call Feuille.getcellbyposition(6, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!Servicio_Ingreso_Origen)
                Call Feuille.getcellbyposition(7, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!FechaHora_IngresoEstabl)
                Call Feuille.getcellbyposition(8, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!Servicio_Actual_Atencion)
                Call Feuille.getcellbyposition(9, lnNumFila - 1).setFormula(OrsDatosHosp.Fields!IdCuenta_Atencion)
                Call Feuille.getcellbyposition(10, lnNumFila - 1).setFormula(IIf(IsNull(OrsDatosHosp.Fields!TotalPorPagar), "", OrsDatosHosp.Fields!TotalPorPagar))
            Else
                oWorkSheet.Cells(lnNumFila, 1).Value = OrsDatosHosp.Fields!NroDocumento
                oWorkSheet.Cells(lnNumFila, 2).Value = OrsDatosHosp.Fields!Apellido_Paterno
                oWorkSheet.Cells(lnNumFila, 3).Value = OrsDatosHosp.Fields!Apellido_Materno
                oWorkSheet.Cells(lnNumFila, 4).Value = OrsDatosHosp.Fields!Nombres
                oWorkSheet.Cells(lnNumFila, 5).Value = OrsDatosHosp.Fields!Fec_Nac
                oWorkSheet.Cells(lnNumFila, 6).Value = OrsDatosHosp.Fields!NroCama
                oWorkSheet.Cells(lnNumFila, 7).Value = OrsDatosHosp.Fields!Servicio_Ingreso_Origen
                oWorkSheet.Cells(lnNumFila, 8).Value = OrsDatosHosp.Fields!FechaHora_IngresoEstabl
                oWorkSheet.Cells(lnNumFila, 9).Value = OrsDatosHosp.Fields!Servicio_Actual_Atencion
                oWorkSheet.Cells(lnNumFila, 10).Value = OrsDatosHosp.Fields!IdCuenta_Atencion
                oWorkSheet.Cells(lnNumFila, 11).Value = IIf(IsNull(OrsDatosHosp.Fields!TotalPorPagar), "", OrsDatosHosp.Fields!TotalPorPagar)
                
                Set range_row = oWorkSheet.range("A" & CStr(lnNumFila), "K" & CStr(lnNumFila))
                range_row.Interior.Color = RGB(Val(OrsDatosHosp.Fields!RGBROJO), Val(OrsDatosHosp.Fields!RGBVERDE), Val(OrsDatosHosp.Fields!RGBAZUL))
                Set borders = range_row.borders
                borders.LineStyle = Excel.XlLineStyle.xlContinuous
            End If
            lnNumFila = lnNumFila + 1
            OrsDatosHosp.MoveNext
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

Public Function AtencionesCensoEstanciaHospitalariaPacientes() As Recordset
    Dim oRsTmp As New Recordset
    Set AtencionesCensoEstanciaHospitalariaPacientes = Nothing
    Set oRsTmp = mo_AdminReportes.AtencionesCensoEstanciaHospitalariaPacientes()
    Set AtencionesCensoEstanciaHospitalariaPacientes = oRsTmp
'    Set oRsTmp = Nothing
End Function

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
    Dim oFormulario As New AtencionesCenso
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub



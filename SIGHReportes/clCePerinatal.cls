VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clCePerinatal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Perinatal
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Sub EjecutaFormulario()
    Dim oFormulario As New CErepPerinatal
    oFormulario.Show 1
End Sub

Sub ProcesarYemitirReporteMenoresAnnControles(ldFechaReporte As Date, lnNroControl As Integer, _
                                              lnHwnd As Long)
    Dim rsReporte As New Recordset
    Set rsReporte = mo_reglasComunes.PerinatalAtencionSeleccionarTodos
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        Dim iFila As Long
        Dim lnTotal As Long, lnEdadEnSemanas As Long
        Dim mo_ReporteUtil As New ReporteUtil
        Dim lbContinua As Boolean, ldFechaActualServidor As Date, lnCantControles As Long
        Dim lcPaciente As String, lcApellidoPaterno As String, lcApellidoMaterno As String
        Dim lcPrimerNombre As String, lnNroHistoriaClinica As Long, ldFechaNacimiento As Date
        Dim lnIdPaciente As Long, lcTelefono As String, lcEmail As String, lcDistrito As String
        Dim ldFechaAtencion As Date, ldFechaCitaProxima As Date
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
            Dim range As Excel.range
            Dim borders As Excel.borders
        End If
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\CePerinatalMenoresAnnControl.ods"
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
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CePerinatalMenoresAnnControl.xls")
            oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 1).setFormula("REPORTE DE SEGUIMIENTO DE NIÑOS SANOS CON MENOS DE " & _
                                                            Trim(Str(lnNroControl)) & " CONTROLES A LA FECHA " & _
                                                            ldFechaReporte)
        Else
            oWorkSheet.Cells(2, 3).Value = "REPORTE DE SEGUIMIENTO DE NIÑOS SANOS CON MENOS DE " & _
                                           Trim(Str(lnNroControl)) & " CONTROLES A LA FECHA  " & _
                                           ldFechaReporte
        End If
        
        iFila = 6
        lnTotal = 0
        ldFechaActualServidor = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lcPaciente = Trim(rsReporte!ApellidoPaterno) & " " & Trim(rsReporte!ApellidoMaterno) & " " & Trim(rsReporte!PrimerNombre)
            lcApellidoPaterno = rsReporte!ApellidoPaterno
            lcApellidoMaterno = rsReporte!ApellidoMaterno
            lcPrimerNombre = rsReporte!PrimerNombre
            lnNroHistoriaClinica = rsReporte!nrohistoriaclinica
            ldFechaNacimiento = rsReporte!FechaNacimiento
            lcTelefono = IIf(IsNull(rsReporte!telefono), "", rsReporte!telefono)
            lcEmail = IIf(IsNull(rsReporte!email), "", rsReporte!email)
            lcDistrito = ""
            If Not IsNull(rsReporte!distrito) Then
               lcDistrito = Trim(rsReporte!distrito) & "/" & Trim(rsReporte!provincia) & "/" & Trim(rsReporte!dpto)
            End If
            lnIdPaciente = rsReporte!IdPaciente
            lnCantControles = 0
            ldFechaAtencion = CDate("01/01/1980")
            Do While Not rsReporte.EOF And lcApellidoPaterno = rsReporte!ApellidoPaterno And _
                                           lcApellidoMaterno = rsReporte!ApellidoMaterno And _
                                           lcPrimerNombre = rsReporte!PrimerNombre And _
                                           ldFechaNacimiento = rsReporte!FechaNacimiento And _
                                           lnIdPaciente = rsReporte!IdPaciente
               If ldFechaAtencion <> rsReporte!fechaAtencion Then
                  lnCantControles = lnCantControles + 1
               End If
               ldFechaAtencion = rsReporte!fechaAtencion
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            lbContinua = False
            lnEdadEnSemanas = sighEntidades.DevuelveEdadEnSemanas(ldFechaNacimiento, ldFechaReporte)
            If lnEdadEnSemanas <= 29 Then
               ldFechaCitaProxima = ldFechaAtencion + 30
            ElseIf lnEdadEnSemanas > 29 And lnEdadEnSemanas <= 36 Then
               ldFechaCitaProxima = ldFechaAtencion + 15
            ElseIf lnEdadEnSemanas > 36 And lnEdadEnSemanas <= 42 Then
               ldFechaCitaProxima = ldFechaAtencion + 7
            Else
               ldFechaCitaProxima = ldFechaAtencion + 2
            End If
            If ldFechaCitaProxima <= ldFechaReporte And lnCantControles < lnNroControl Then
                  lbContinua = True
            End If
            If lbContinua = True Then
                If lbEsOpenOffice = True Then
                   Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lnNroHistoriaClinica)
                   Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcPaciente)
                   Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcDistrito)
                   Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lcTelefono)
                   Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lcEmail)
                   Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnCantControles)
                   Call Feuille.getcellbyposition(7, iFila - 1).setFormula(ldFechaAtencion)
                   Call Feuille.getcellbyposition(8, iFila - 1).setFormula(lnEdadEnSemanas)
                   Call Feuille.getcellbyposition(9, iFila - 1).setFormula(ldFechaCitaProxima)
                Else
                   oWorkSheet.Cells(iFila, 2).Value = lnNroHistoriaClinica
                   oWorkSheet.Cells(iFila, 3).Value = lcPaciente
                   oWorkSheet.Cells(iFila, 4).Value = lcDistrito
                   oWorkSheet.Cells(iFila, 5).Value = lcTelefono
                   oWorkSheet.Cells(iFila, 6).Value = lcEmail
                   oWorkSheet.Cells(iFila, 7).Value = lnCantControles
                   oWorkSheet.Cells(iFila, 8).Value = ldFechaAtencion
                   oWorkSheet.Cells(iFila, 9).Value = lnEdadEnSemanas
                   oWorkSheet.Cells(iFila, 10).Value = ldFechaCitaProxima
                End If
                iFila = iFila + 1
                lnTotal = lnTotal + 1
            End If
            If rsReporte.EOF Then
               Exit Do
            End If
        Loop
        
        If lbEsOpenOffice = True Then
    '        Set Plage = Feuille.getCellRangeByName("B10:D15")
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":J" & CStr(iFila))
'            Plage.Merge (True)
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Pacientes: " & Format(lnTotal, "####,###"))
            iFila = iFila + 2
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("INTERVALO DE CITAS SEGUN PROTOCOLO: 1-29 SEM 30 DIAS, 30-36 SEM 15 DIAS, 37-42 SEM 7 DIAS, MAS DE 43 SEM 2 VECES X SEMANA")
            iFila = iFila + 1
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 10
            oWorkSheet.Cells(iFila, 2).Value = "Nro Pacientes: " & Format(lnTotal, "####,###")
            iFila = iFila + 2
            oWorkSheet.Cells(iFila, 2).Value = "INTERVALO DE CITAS SEGUN PROTOCOLO: 1-29 SEM 30 DIAS, 30-36 SEM 15 DIAS, 37-42 SEM 7 DIAS, MAS DE 43 SEM 2 VECES X SEMANA"
            iFila = iFila + 1
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 10
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
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
        Set rsReporte = Nothing
    End If
End Sub

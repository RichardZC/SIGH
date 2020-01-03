VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCEatenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para atenciones en Consultorios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_IdServicioCE As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim ml_TextoDelFiltro1 As String
Dim mda_FechaInicio As String
Dim mda_FechaFin As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Property Let FechaInicio(daValue As String)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As String)
    mda_FechaFin = daValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let TextoDelFiltro1(lValue As String)
    ml_TextoDelFiltro1 = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdServicioCE(lValue As Long)
    ml_IdServicioCE = lValue
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
    Dim oFormulario As New CEatenciones
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim iFila As Long
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim lnAtencionT As Long: Dim lnAtendidoT As Long
Dim lnAtencionD As Long: Dim lnAtendidoD As Long
Dim lnAtencionE As Long: Dim lnAtendidoE As Long
Dim lnAtencionS As Long: Dim lnAtendidoS As Long
Dim lcDpto As String: Dim lcEspecialidad As String: Dim lcServicio As String
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
    Set mrs_Tmp = mo_AdminAdmision.BuscaAtencionesCE(mda_FechaInicio, mda_FechaFin, ml_IdResponsable)
    If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
           
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\CEatenciones.ods"
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
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEatenciones.xls")
            oWorkBookPlantilla.Worksheets("CEatenciones").Copy Before:=oWorkBook.Sheets(1)
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
        iFila = 6
        lnAtencionT = 0
        lnAtendidoT = 0
                   
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lcDpto = mrs_Tmp.Fields!dpto
               lnAtencionD = 0: lnAtendidoD = 0
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lcDpto)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = lcDpto
                End If
                Do While Not mrs_Tmp.EOF And lcDpto = mrs_Tmp.Fields!dpto
                   lcEspecialidad = mrs_Tmp.Fields!especialidad
                   lnAtencionE = 0: lnAtendidoE = 0
                   If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcEspecialidad)
                   Else
                        oWorkSheet.Cells(iFila, 3).Value = lcEspecialidad
                   End If
                   Do While Not mrs_Tmp.EOF And lcDpto = mrs_Tmp.Fields!dpto And lcEspecialidad = mrs_Tmp.Fields!especialidad
                     lcServicio = mrs_Tmp.Fields!Servicio
                     lnAtencionS = 0: lnAtendidoS = 0
                     If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcServicio)
                     Else
                        oWorkSheet.Cells(iFila, 4).Value = lcServicio
                     End If
                     Do While Not mrs_Tmp.EOF And lcDpto = mrs_Tmp.Fields!dpto And lcEspecialidad = mrs_Tmp.Fields!especialidad And lcServicio = mrs_Tmp.Fields!Servicio
                        lnAtencionS = lnAtencionS + 1
                        If Not IsNull(mrs_Tmp.Fields!horaEgreso) Then
                           lnAtendidoS = lnAtendidoS + 1
                        End If
                        mrs_Tmp.MoveNext
                        If mrs_Tmp.EOF Then
                           Exit Do
                        End If
                     Loop
                     If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnAtencionS)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnAtendidoS)
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnAtencionS - lnAtendidoS)
                     Else
                        oWorkSheet.Cells(iFila, 5).Value = lnAtencionS
                        oWorkSheet.Cells(iFila, 6).Value = lnAtendidoS
                        oWorkSheet.Cells(iFila, 7).Value = lnAtencionS - lnAtendidoS
                     End If
                     iFila = iFila + 1
                     lnAtencionE = lnAtencionE + lnAtencionS
                     lnAtendidoE = lnAtendidoE + lnAtendidoS
                     If mrs_Tmp.EOF Then
                        Exit Do
                     End If
                  Loop
                  
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":G" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnAtencionE)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnAtendidoE)
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnAtencionE - lnAtendidoE)
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 7
                    oWorkSheet.Cells(iFila, 5).Value = lnAtencionE
                    oWorkSheet.Cells(iFila, 6).Value = lnAtendidoE
                    oWorkSheet.Cells(iFila, 7).Value = lnAtencionE - lnAtendidoE
                End If
                iFila = iFila + 1
                lnAtencionD = lnAtencionD + lnAtencionE
                lnAtendidoD = lnAtendidoD + lnAtendidoE
                If mrs_Tmp.EOF Then
                   Exit Do
                End If
               Loop
               If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":G" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnAtencionD)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnAtendidoD)
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnAtencionD - lnAtendidoD)
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
                    oWorkSheet.Cells(iFila, 5).Value = lnAtencionD
                    oWorkSheet.Cells(iFila, 6).Value = lnAtendidoD
                    oWorkSheet.Cells(iFila, 7).Value = lnAtencionD - lnAtendidoD
                End If
               iFila = iFila + 1
               lnAtencionT = lnAtencionT + lnAtencionD
               lnAtendidoT = lnAtendidoT + lnAtendidoD
            Loop
           
               If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":G" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnAtencionT)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnAtendidoT)
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnAtencionT - lnAtendidoT)
                    iFila = iFila + 1
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
                    oWorkSheet.Cells(iFila, 5).Value = lnAtencionT
                    oWorkSheet.Cells(iFila, 6).Value = lnAtendidoT
                    oWorkSheet.Cells(iFila, 7).Value = lnAtencionT - lnAtendidoT
                    iFila = iFila + 1
                End If
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 7
                    PrintArea(0).EndRow = iFila
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                    Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
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







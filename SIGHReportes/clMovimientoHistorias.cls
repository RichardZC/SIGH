VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clMovimientoHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Movimientos de Historia
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdGrupoMovimiento As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdmiServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReporteUtil As New ReporteUtil
Dim lcBuscaParametro As New SIGHDatos.Parametros
'Dim mo_ProgressRpt As XP_ProgressBar
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let IdGrupoMovimiento(lValue As Long)
    ml_IdGrupoMovimiento = lValue
End Property
'Property Set progressRpt(oValue As XP_ProgressBar)
'    Set mo_ProgressRpt = oValue
'End Property

Sub CrearReporteMovimientoHistoria(lnHwnd As Long)

Dim rsReporte As New Recordset
Dim oRange As range
Dim iFila As Long
Dim lIdServicio As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lIdServicioDestino As Long
Dim lbEsOpenOffice As Boolean
lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError
Dim lcNombre As String, lcSql As String

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
        lcArchivoExcel = App.Path + "\Plantillas\HCMovimientoHistorias.ods"
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

        'Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oExcel = New Excel.Application
        
        'Crea nueva hoja
        Set oWorkBook = oExcel.Workbooks.Add
        
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HCMovimientoHistorias.xls")
        oWorkBookPlantilla.Worksheets("MovimientoHistorias").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
       
    'Crea la cabecera
    Set rsReporte = mo_AdminReportes.ObtenerCabeceraDeMovimiento(ml_IdGrupoMovimiento)
    
    Dim oDOEmpleado As dOEmpleado
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(rsReporte!IdEmpleadoArchivo)
    If Not oDOEmpleado Is Nothing Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(4, 4).setFormula(oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres)
        Else
            oWorkSheet.Cells(5, 5).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(rsReporte!IdEmpleadoTransporte)
    If Not oDOEmpleado Is Nothing Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(4, 5).setFormula(oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres)
        Else
            oWorkSheet.Cells(6, 5).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    Set oDOEmpleado = mo_AdminServComunes.EmpleadosSeleccionarPorId(rsReporte!IdEmpleadoRecepcion)
    If Not oDOEmpleado Is Nothing Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(4, 5).setFormula(oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres)
        Else
            oWorkSheet.Cells(7, 5).Value = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    
    If lbEsOpenOffice = True Then
        'Fecha y hora del movimiento
        Call Feuille.getcellbyposition(14, 3).setFormula("'" & Format(rsReporte!FechaMovimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
        Call Feuille.getcellbyposition(14, 4).setFormula("'" & Format(rsReporte!FechaMovimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
        'Fecha hora de reporte
        Call Feuille.getcellbyposition(13, 0).setFormula("'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
        Call Feuille.getcellbyposition(16, 0).setFormula("'" & Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM))

    Else
        'Fecha y hora del movimiento
        oWorkSheet.Cells(4, 15).Value = "'" & Format(rsReporte!FechaMovimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        oWorkSheet.Cells(5, 15).Value = "'" & Format(rsReporte!FechaMovimiento, SIGHEntidades.DevuelveHoraSoloFormato_HM)
        
        'Fecha hora de reporte
        oWorkSheet.Cells(1, 14).Value = "'" & Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        oWorkSheet.Cells(1, 17).Value = "'" & Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    End If
    
    Set rsReporte = mo_AdminReportes.ObtenerHistoriasClinicasDeEsteMovimiento(ml_IdGrupoMovimiento)
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(14, 5).setFormula("'" & rsReporte.Fields!motivo)
        Else
            oWorkSheet.Cells(6, 15).Value = "'" & rsReporte.Fields!motivo
        End If
       
    
    iFila = 10
    lIdServicioDestino = 0
'    mo_ProgressRpt.Min = 0
'    mo_ProgressRpt.Max = rsReporte.RecordCount
'    mo_ProgressRpt.Value = 0
    
    Do While Not rsReporte.EOF
        'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(0, iFila).setFormula("'" & IIf(lIdServicioDestino <> rsReporte!idServicioDestino, rsReporte!ServicioDestino, ""))
            Call Feuille.getcellbyposition(5, iFila).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!nrohistoriaclinica))
            Call Feuille.getcellbyposition(7, iFila).setFormula("'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))
            Call Feuille.getcellbyposition(13, iFila).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!ServicioOrigen))
            Call Feuille.getcellbyposition(18, iFila).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!NroFolios))
            Call Feuille.getcellbyposition(20, iFila).setFormula("'" & Format(mo_ReporteUtil.NullToVacio(rsReporte!FechaRequerida), SIGHEntidades.DevuelveFechaSoloFormato_DMY))

            Set Plage = Feuille.getCellRangeByName(("A" & iFila & ":" & "E" & iFila))
            Plage.Merge (True)
            Set Plage = Feuille.getCellRangeByName(("F" & iFila & ":" & "G" & iFila))
            Plage.Merge (True)
            Set Plage = Feuille.getCellRangeByName(("H" & iFila & ":" & "M" & iFila))
            Plage.Merge (True)
            Set Plage = Feuille.getCellRangeByName(("N" & iFila & ":" & "R" & iFila))
            Plage.Merge (True)
            Set Plage = Feuille.getCellRangeByName(("S" & iFila & ":" & "T" & iFila))
            Plage.Merge (True)
            Set Plage = Feuille.getCellRangeByName(("U" & iFila & ":" & "V" & iFila))
            Plage.Merge (True)
            ' Modification apportée à la plage
            Set Border = Plage.TableBorder
            Set Style = Border.TopLine
            Style.OuterLineWidth = 25
            Style.Color = RGB(0, 0, 0)
            Border.TopLine = Style
            Plage.TableBorder = Border
        Else
            oWorkSheet.Cells(iFila, 1).Value = "'" & IIf(lIdServicioDestino <> rsReporte!idServicioDestino, rsReporte!ServicioDestino, "")
            oWorkSheet.Cells(iFila, 6).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!nrohistoriaclinica)
            oWorkSheet.Cells(iFila, 8).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
            oWorkSheet.Cells(iFila, 14).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ServicioOrigen)
            oWorkSheet.Cells(iFila, 19).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroFolios)
            oWorkSheet.Cells(iFila, 21).Value = "'" & Format(mo_ReporteUtil.NullToVacio(rsReporte!FechaRequerida), SIGHEntidades.DevuelveFechaSoloFormato_DMY)
            
            Set oRange = oExcel.range("A" & iFila & ":" & "E" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
            Set oRange = oExcel.range("F" & iFila & ":" & "G" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
            Set oRange = oExcel.range("H" & iFila & ":" & "M" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
            Set oRange = oExcel.range("N" & iFila & ":" & "R" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
            Set oRange = oExcel.range("S" & iFila & ":" & "T" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
            Set oRange = oExcel.range("U" & iFila & ":" & "V" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlRight
        End If
        
        lIdServicioDestino = rsReporte!idServicioDestino
        
        iFila = iFila + 1
        rsReporte.MoveNext
    Loop
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName(("A10:" & "U" & Str(iFila - 1)))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
    Else
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 10, 1, iFila - 1, 21
    End If
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 22
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        If oWorkSheet.PageSetup.PrintArea <> "" Then
         oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
    End If
    
    'Falta que salga el nombre del responsable
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



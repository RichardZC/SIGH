VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clCuposAsignadosRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para cupos asignados
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
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_ProgressRpt As XP_ProgressBar
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim rsReporte1 As New Recordset
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

Sub EjecutaFormulario()
    Dim oFormulario As New CuposAsignadosReporte
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

Sub CrearReporteCuposAsignados(lnHwnd As Long)
Dim rsReporte As New Recordset
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
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEcuposAsignados.ods"
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
        Set oWorkBook = oExcel.Workbooks.Add
        
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEcuposAsignados.xls")
        oWorkBookPlantilla.Worksheets("CuposAsignados").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    
    If lbEsOpenOffice = True Then
        'Call Feuille.getcellbyposition(6, 0).setFormula("'" & Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY))
       ' Call Feuille.getcellbyposition(7, 0).setFormula("'" & Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM))
        Call Feuille.getcellbyposition(2, 4).setFormula("'" & mda_FechaRequeridaDesde)
        Call Feuille.getcellbyposition(5, 4).setFormula("'" & mda_FechaRequeridaHasta)
        Call Feuille.getcellbyposition(1, 6).setFormula("Consultorio")
    Else
       ' oWorkSheet.Cells(1, 7).Value = "'" & Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
       ' oWorkSheet.Cells(1, 8).Value = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        oWorkSheet.Cells(5, 3).Value = mda_FechaRequeridaDesde
        oWorkSheet.Cells(5, 6).Value = mda_FechaRequeridaHasta
        oWorkSheet.Cells(7, 2).Value = "Consultorio"
    End If
    
    Set rsReporte = mo_AdminReportes.CantidadHistoriaSolicitadasPorServicio(mda_FechaRequeridaDesde, mda_FechaRequeridaHasta, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, ml_idTipoServicio)
    rsReporte.Filter = "idTipoServicio=1"
    iFila = 8
   ' mo_ProgressRpt.Min = 0
   ' mo_ProgressRpt.Max = rsReporte.RecordCount
   ' mo_ProgressRpt.Value = 0
    Dim rsSolicitudes  As Recordset
    
    Do While Not rsReporte.EOF
        'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte!Codigo)
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("'" & rsReporte!Servicio)
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte!Cantidad)
            Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":F" & CStr(iFila))
            Plage.Merge (True)
            Set Plage = Feuille.getCellRangeByName("B" & CStr(8) & ":G" & CStr(iFila))
            
        Else
            oWorkSheet.Cells(iFila, 2).Value = rsReporte!Codigo
            oWorkSheet.Cells(iFila, 3).Value = "'" & rsReporte!Servicio
            oWorkSheet.Cells(iFila, 7).Value = rsReporte!Cantidad
            Set oRange = oExcel.range("C" & iFila & ":" & "F" & iFila): oRange.Merge: oRange.HorizontalAlignment = xlLeft
            
        End If
        iFila = iFila + 1
        rsReporte.MoveNext
    Loop
    If lbEsOpenOffice = True Then
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
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
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 8, 2, iFila - 1, 7
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
    Resume
End Sub

Sub CrearReporteCuposAsignadosVaciosPorMedico(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsReporte2 As New ADODB.Recordset
Dim iFila As Long
Dim lIdServicio As Integer: Dim lnCuposLlenos As Integer
Dim lnIdMedico As Long: Dim lnCuposTotal As Integer: Dim lnCuposAsignados As Integer
Dim lnTotal As Integer: Dim lnLlenos As Integer: Dim lnVacios As Integer
Dim lnTTotal As Integer: Dim lnTLlenos As Integer: Dim lnTVacios As Integer
Dim lnCitasAdicionales As Integer, lnAdicionales As Integer, lnTAdicionales As Integer
Dim lnIdServicio As Long
Dim lcServicio As String: Dim lcMedico As String
Dim lHoraInicio As Long
Dim lHoraFin  As Long
Dim lTiempoPromedio As Long
Dim lHoraSiguiente As Long
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
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEcuposPorMedico.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEcuposPorMedico.xls")
        oWorkBookPlantilla.Worksheets("CuposPorMedico").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    
    If lbEsOpenOffice = True Then
       ' Call Feuille.getcellbyposition(6, 0).setFormula("'" & Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY))
       ' Call Feuille.getcellbyposition(7, 0).setFormula("'" & Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM))
        Call Feuille.getcellbyposition(2, 4).setFormula("'" & mda_FechaSolicitudDesde)
        Call Feuille.getcellbyposition(5, 4).setFormula("'" & mda_FechaSolicitudHasta)
        Call Feuille.getcellbyposition(1, 6).setFormula("Especialidad")
    Else
      '  oWorkSheet.Cells(1, 7).Value = "'" & Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
      '  oWorkSheet.Cells(1, 8).Value = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        oWorkSheet.Cells(5, 3).Value = mda_FechaSolicitudDesde
        oWorkSheet.Cells(5, 6).Value = mda_FechaSolicitudHasta
        oWorkSheet.Cells(7, 2).Value = "Especialidad"
    End If
    Set rsReporte = mo_ReglasDeProgMedica.CuposProgramadosPorMedicosPorFechas(mda_FechaSolicitudDesde, mda_FechaSolicitudHasta)
        If rsReporte.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            GenerarRecordsetTemporal
    '        mo_ProgressRpt.Min = 0
    '        mo_ProgressRpt.Max = rsReporte.RecordCount
    '        mo_ProgressRpt.Value = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lIdServicio = rsReporte.Fields!idServicio
                lnIdMedico = rsReporte.Fields!idMedico
                lnIdServicio = rsReporte.Fields!IdEspecialidad
                lcServicio = rsReporte.Fields!Nombre
                lcMedico = Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!Nombres)
                lnCuposTotal = 0
                Do While Not rsReporte.EOF And lnIdMedico = rsReporte.Fields!idMedico
                  lHoraInicio = ConvertirAMinutos(rsReporte.Fields!HoraInicio)
                  lHoraFin = ConvertirAMinutos(rsReporte.Fields!HoraFin)
                  lTiempoPromedio = rsReporte.Fields!TiempoPromedioAtencion
                  lHoraSiguiente = lHoraInicio
                  
                  Do While lHoraSiguiente < lHoraFin
                        lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
                        lnCuposTotal = lnCuposTotal + 1
                        lHoraInicio = lHoraSiguiente
                  Loop
                  rsReporte.MoveNext
                  If rsReporte.EOF Then
                     Exit Do
                  End If
                Loop
                Set rsReporte2 = mo_ReglasDeProgMedica.DevuelveCantidadCitasYaAsignadas(mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, lnIdMedico)
                rsReporte2.Filter = "idEstadoAtencion<>0 and idServicioIngreso=" & lIdServicio
                lnCuposLlenos = rsReporte2.RecordCount
                lnCitasAdicionales = mo_ReglasDeProgMedica.DevuelveCantidadDeCitasAdicionales(mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, lnIdMedico)
                
                rsReporte1.AddNew
                rsReporte1.Fields!DServicio = lcServicio
                rsReporte1.Fields!idServicio = lnIdServicio
                rsReporte1.Fields!idMedico = lnIdMedico
                rsReporte1.Fields!dmedico = lcMedico
                rsReporte1.Fields!cuposTotal = lnCuposTotal - lnCitasAdicionales
                rsReporte1.Fields!cuposLlenos = lnCuposLlenos - lnCitasAdicionales
                rsReporte1.Fields!cuposVacios = (lnCuposTotal - lnCitasAdicionales) - (lnCuposLlenos - lnCitasAdicionales)
                rsReporte1.Fields!cuposAdicionales = lnCitasAdicionales
                rsReporte1.Update
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
            iFila = 8
            rsReporte1.Sort = "dServicio,dMedico"
            rsReporte1.MoveFirst
            lnTTotal = 0: lnTLlenos = 0: lnTVacios = 0: lnTAdicionales = 0
            Do While Not rsReporte1.EOF
               lcServicio = rsReporte1.Fields!DServicio
               lnIdServicio = rsReporte1.Fields!idServicio
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lcServicio)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = lcServicio
                End If
               lnTotal = 0: lnLlenos = 0: lnVacios = 0: lnAdicionales = 0
               Do While Not rsReporte1.EOF And lcServicio = rsReporte1.Fields!DServicio And lnIdServicio = rsReporte1.Fields!idServicio
                  If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(3, CInt(iFila - 1)).setFormula(rsReporte1.Fields!dmedico)
                        Call Feuille.getcellbyposition(6, CInt(iFila - 1)).setFormula(rsReporte1.Fields!cuposTotal)
                        Call Feuille.getcellbyposition(7, CInt(iFila - 1)).setFormula(rsReporte1.Fields!cuposAdicionales)
                        Call Feuille.getcellbyposition(8, CInt(iFila - 1)).setFormula(rsReporte1.Fields!cuposLlenos)
                        Call Feuille.getcellbyposition(9, CInt(iFila - 1)).setFormula(rsReporte1.Fields!cuposVacios)
                  Else
                        oWorkSheet.Cells(iFila, 4).Value = "'" & rsReporte1.Fields!dmedico
                        oWorkSheet.Cells(iFila, 7).Value = rsReporte1.Fields!cuposTotal
                        oWorkSheet.Cells(iFila, 8).Value = rsReporte1.Fields!cuposAdicionales
                        oWorkSheet.Cells(iFila, 9).Value = rsReporte1.Fields!cuposLlenos
                        oWorkSheet.Cells(iFila, 10).Value = rsReporte1.Fields!cuposVacios
                  End If
                  lnTotal = lnTotal + rsReporte1.Fields!cuposTotal
                  lnLlenos = lnLlenos + rsReporte1.Fields!cuposLlenos
                  lnVacios = lnVacios + rsReporte1.Fields!cuposVacios
                  lnAdicionales = lnAdicionales + rsReporte1.Fields!cuposAdicionales
                  lnTTotal = lnTTotal + rsReporte1.Fields!cuposTotal
                  lnTLlenos = lnTLlenos + rsReporte1.Fields!cuposLlenos
                  lnTVacios = lnTVacios + rsReporte1.Fields!cuposVacios
                  lnTAdicionales = lnTAdicionales + rsReporte1.Fields!cuposAdicionales
                  iFila = iFila + 1
                  rsReporte1.MoveNext
                  If rsReporte1.EOF Then
                     Exit Do
                  End If
               Loop
               If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(6, CInt(iFila - 1)).setFormula(lnTotal)
                    Call Feuille.getcellbyposition(7, CInt(iFila - 1)).setFormula(lnAdicionales)
                    Call Feuille.getcellbyposition(8, CInt(iFila - 1)).setFormula(lnLlenos)
                    Call Feuille.getcellbyposition(9, CInt(iFila - 1)).setFormula(lnVacios)
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":J" & CStr(iFila))
                     mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
               Else
                    oWorkSheet.Cells(iFila, 7).Value = lnTotal
                    oWorkSheet.Cells(iFila, 8).Value = lnAdicionales
                    oWorkSheet.Cells(iFila, 9).Value = lnLlenos
                    oWorkSheet.Cells(iFila, 10).Value = lnVacios
                     mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 10
               End If
               iFila = iFila + 2
               If rsReporte1.EOF Then
                   Exit Do
               End If
            Loop
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(6, CInt(iFila - 1)).setFormula(lnTotal)
                Call Feuille.getcellbyposition(7, CInt(iFila - 1)).setFormula(lnTAdicionales)
                Call Feuille.getcellbyposition(8, CInt(iFila - 1)).setFormula(lnTLlenos)
                Call Feuille.getcellbyposition(9, CInt(iFila - 1)).setFormula(lnTVacios)
            Else
                oWorkSheet.Cells(iFila, 7).Value = lnTTotal
                oWorkSheet.Cells(iFila, 8).Value = lnTAdicionales
                oWorkSheet.Cells(iFila, 9).Value = lnTLlenos
                oWorkSheet.Cells(iFila, 10).Value = lnTVacios
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
                Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                    oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
        
        If lbEsOpenOffice = True Then
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
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
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




Sub GenerarRecordsetTemporal()
    With rsReporte1
          .Fields.Append "dServicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "IdMedico", adInteger, 4, adFldIsNullable
          .Fields.Append "dMedico", adVarChar, 150, adFldIsNullable
          .Fields.Append "CuposTotal", adInteger
          .Fields.Append "cuposLlenos", adInteger
          .Fields.Append "cuposVacios", adInteger
          .Fields.Append "cuposAdicionales", adInteger
          .LockType = adLockOptimistic
          .Open
          .Sort = "dServicio,dMedico"
    End With
End Sub


Function ConvertirAMinutos(sHora As String) As Long
Dim sHoras() As String
        
        sHoras = Split(sHora, ":")
        ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))
        
End Function
Function ConvertirAHora(lMinutos As Long) As String
Dim sHoras() As String
Dim lMin As Long
Dim lHora As Long
        
        lMin = lMinutos Mod 60
        lHora = (lMinutos - lMin) / 60
        
        ConvertirAHora = Right("00" & lHora, 2) + ":" + Right("00" & lMin, 2)
        
End Function





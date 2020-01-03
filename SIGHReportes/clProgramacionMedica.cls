VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clProgramMedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para programación médica
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim ml_IdDepartamento As Long
Dim ml_IdEspecialidad As Long
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim mo_ReporteUtil As New ReporteUtil
'Dim mo_ProgressRpt As XP_ProgressBar
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let IdDepartamento(lValue As Long)
    ml_IdDepartamento = lValue
End Property
'yamill palomino
Property Let IdEspecialidad(lValue As Long)
    ml_IdEspecialidad = lValue
End Property
Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Sub CrearReporteProgramacionMedica(bMostrarHoras As Boolean, lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsDiasProgramados  As New Recordset
Dim iNroDeDias As Integer
Dim iColumnaFin As Integer
Dim lIdServicioDestino As Long
Dim lnTotalDias As Long
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String, lnCantidadHrs As Integer, lcTurnos As String


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

    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEprogMedica.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEprogMedica.xls")
        oWorkBookPlantilla.Worksheets("ProgMedica").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close

        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If

    'Crea la cabecera
    Dim iColumnaIni As Integer
    Dim iFila As Long
    Dim daFechaProg  As Date
    iFila = 4:  iColumnaIni = 30
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iColumnaIni - 1, iFila - 1).setFormula("'" & Format(mda_FechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY))
        Else
            oWorkSheet.Cells(iFila, iColumnaIni).Value = "'" & Format(mda_FechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY)
        End If
    iFila = 4:  iColumnaIni = 38
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iColumnaIni - 1, iFila - 1).setFormula("'" + Format(mda_FechaFin, sighentidades.DevuelveFechaSoloFormato_DMY))
        Else
            oWorkSheet.Cells(iFila, iColumnaIni).Value = "'" + Format(mda_FechaFin, sighentidades.DevuelveFechaSoloFormato_DMY)
        End If
    'yamill palomino
    Dim oDoDepartamento As New DODepartamentoHospital
    Dim oDoEspecialidad As New DOEspecialidades
    Set oDoDepartamento = mo_AdminServHosp.DepartamentosSeleccionarPorId(ml_IdDepartamento)
    Set oDoEspecialidad = mo_AdminServHosp.EspecialidadesSeleccionarPorId(ml_IdEspecialidad)

    'yamill palomino
    If lbEsOpenOffice = True Then
        If ml_IdEspecialidad > 0 Then
            Call Feuille.getcellbyposition(14, 4).setFormula("DEPARTAMENTO DE " + UCase(oDoDepartamento.Nombre) + UCase(oDoEspecialidad.Nombre))
        Else
            Call Feuille.getcellbyposition(14, 4).setFormula("DEPARTAMENTO DE " + UCase(oDoDepartamento.Nombre))
        End If
    Else
        If ml_IdEspecialidad > 0 Then
            oWorkSheet.Cells(5, 15).Value = "DEPARTAMENTO DE " + UCase(oDoDepartamento.Nombre) + " (ESPECIALIDAD: " + UCase(Trim(oDoEspecialidad.Nombre)) + ")"
        Else
            oWorkSheet.Cells(5, 15).Value = "DEPARTAMENTO DE " + UCase(oDoDepartamento.Nombre)
        End If
    End If
    iFila = 7
    iColumnaIni = 16
    daFechaProg = mda_FechaInicio
    Do While daFechaProg <= mda_FechaFin
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iColumnaIni - 1, iFila - 1).setFormula(Left(DiaDeLASemana(daFechaProg), 1))
            Call Feuille.getcellbyposition(iColumnaIni - 1, iFila + 0).setFormula(DatePart("d", daFechaProg))
        Else
            oWorkSheet.Cells(iFila, iColumnaIni).Value = Left(DiaDeLASemana(daFechaProg), 1)
            oWorkSheet.Cells(iFila + 1, iColumnaIni).Value = DatePart("d", daFechaProg)
        End If
    iColumnaIni = iColumnaIni + 1
    daFechaProg = DateAdd("d", 1, daFechaProg)
    Loop
    iColumnaFin = iColumnaIni
    'Crea el contenido
    iFila = 9
    iColumnaIni = 16

    'yamill palomino
    If ml_IdEspecialidad > 0 Then
        Set rsReporte = mo_AdminReportes.FiltrarMedicosPorDepartamentoyEspecialidad(ml_IdDepartamento, ml_IdEspecialidad)
    Else
        Set rsReporte = mo_AdminReportes.FiltrarMedicosPorDepartamento(ml_IdDepartamento)
    End If
'    mo_ProgressRpt.Min = 0
'    mo_ProgressRpt.Max = rsReporte.RecordCount
'    mo_ProgressRpt.Value = 0
    lcTurnos = ""
    Do While Not rsReporte.EOF
        If ml_IdEspecialidad > 0 Then
            Set rsDiasProgramados = mo_AdminReportes.ReporteProgramacionMedicaPorMedicoyEspecialidad(rsReporte!idMedico, mda_FechaInicio, mda_FechaFin, ml_IdEspecialidad)
        Else
            Set rsDiasProgramados = mo_AdminReportes.ReporteProgramacionMedicaPorMedico(rsReporte!idMedico, mda_FechaInicio, mda_FechaFin)
        End If
        If rsDiasProgramados.RecordCount > 0 Then

            'mo_ProgressRpt.Value = mo_ProgressRpt.Value + 1: DoEvents: DoEvents
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula(mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!Nombres)))
                Call Feuille.getcellbyposition(13, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!CodigoPlanilla))
            Else
                oWorkSheet.Cells(iFila, 1).Value = mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!Nombres))
                oWorkSheet.Cells(iFila, 14).Value = mo_ReporteUtil.NullToVacio(rsReporte!CodigoPlanilla)
            End If
            lnTotalDias = 0

            Do While Not rsDiasProgramados.EOF
                'lnTotalDias = lnTotalDias + CantidadDeHoras(rsDiasProgramados!HoraInicio, rsDiasProgramados!HoraFin)
                daFechaProg = rsDiasProgramados!Fecha
                iNroDeDias = DateDiff("d", mda_FechaInicio, daFechaProg)
                lcSql = "": lnCantidadHrs = 0
                Do While Not rsDiasProgramados.EOF And daFechaProg = rsDiasProgramados!Fecha
                    If lcTurnos = "" Then
                       lcTurnos = Trim(Str(rsDiasProgramados!idTurno)) & "/"
                    Else
                    If InStr(lcTurnos, Trim(Str(rsDiasProgramados!idTurno)) & "/") = 0 Then
                       lcTurnos = lcTurnos & Trim(Str(rsDiasProgramados!idTurno)) & "/"
                    End If
                    End If
                    lnCantidadHrs = lnCantidadHrs + CantidadDeHoras(rsDiasProgramados!HoraInicio, rsDiasProgramados!HoraFin)
                    If bMostrarHoras = True Then
                    Else
                       lcSql = lcSql & Trim(rsDiasProgramados!CodigoTurno) & "/"
                    End If
                    rsDiasProgramados.MoveNext
                    If rsDiasProgramados.EOF Then
                       Exit Do
                    End If
                Loop
                lnTotalDias = lnTotalDias + lnCantidadHrs
                If bMostrarHoras = True Then
                   lcSql = Trim(Str(lnCantidadHrs))
                Else
                   lcSql = Left(lcSql, Len(lcSql) - 1)
                End If
                If lbEsOpenOffice = True Then
                    'Call Feuille.getcellbyposition(iColumnaIni + iNroDeDias - 1, iFila - 1).setFormula(IIf(bMostrarHoras, CantidadDeHoras(rsDiasProgramados!HoraInicio, rsDiasProgramados!HoraFin), rsDiasProgramados!CodigoTurno))
                    Call Feuille.getcellbyposition(iColumnaIni + iNroDeDias - 1, iFila - 1).setFormula(lcSql)
                Else
                    'oWorkSheet.Cells(iFila, iColumnaIni + iNroDeDias).Value = IIf(bMostrarHoras, CantidadDeHoras(rsDiasProgramados!HoraInicio, rsDiasProgramados!HoraFin), rsDiasProgramados!CodigoTurno)
                     oWorkSheet.Cells(iFila, iColumnaIni + iNroDeDias).Value = lcSql
                End If
                'rsDiasProgramados.MoveNext
            Loop

            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("A" & CStr(iFila) & ":M" & CStr(iFila))
                Plage.Merge (True)
                Call Feuille.getcellbyposition(14, iFila - 1).setFormula(lnTotalDias)
            Else
                Dim oRange As range
                Set oRange = oExcel.range("A" & iFila & ":" & "M" & iFila)
                oRange.Merge: oRange.HorizontalAlignment = xlLeft
                oWorkSheet.Cells(iFila, 15).Value = lnTotalDias
             End If
            iFila = iFila + 1
        End If
        rsReporte.MoveNext

    Loop
    Dim iIndice As Integer
    iIndice = 9 - iFila

    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("O" & CStr(iFila) & ":O" & CStr(iFila))
        Plage.Merge (True)
        'Call Feuille.getcellbyposition(15, iFila - 1).setFormula("=SUM(R[" & iIndice & "]C:R[-1]C)")
        Set Plage = Feuille.getCellRangeByName(mo_AdminReportes.BuscaNombreColumna(1) & CStr(7) & ":" & mo_AdminReportes.BuscaNombreColumna(iColumnaFin - 1) & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
    Else
        Set oRange = oExcel.range("O" & iFila & ":" & "O" & iFila)
        oRange.Merge: oRange.HorizontalAlignment = xlRight
        oRange.FormulaR1C1 = "=SUM(R[" & iIndice & "]C:R[-1]C)"
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 7, 1, iFila, iColumnaFin - 1
    End If

    Dim i As Integer
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("A6:AT6")
        Plage.Merge (True)
    Else
        Set oRange = oExcel.range(oWorkSheet.Cells(7, 1), oWorkSheet.Cells(7, iColumnaFin - 1))
        oRange.Interior.ColorIndex = 5
    End If
    For i = 15 To iColumnaFin - 1
        If lbEsOpenOffice = True Then
            If Feuille.getcellbyposition(i - 1, 6).GetString = "D" Then
                Set Plage = Feuille.getCellRangeByName("I" & CStr(7) & ":I" & CStr(iFila - 1))
                Plage.Merge (True)
            End If
        Else
            If oWorkSheet.Cells(7, i).Value = "D" Then
                Set oRange = oExcel.range(oWorkSheet.Cells(7, i), oWorkSheet.Cells(iFila - 1, i))
                oRange.Interior.ColorIndex = 6
            End If
        End If
    Next
    'leyenda
    If bMostrarHoras = False Then
        iFila = iFila + 2
        Set rsDiasProgramados = mo_ReglasDeProgMedica.TurnosSeleccionarTodos
        If rsDiasProgramados.RecordCount > 0 Then
           rsDiasProgramados.MoveFirst
           Do While Not rsDiasProgramados.EOF
                lcSql = Trim(Str(rsDiasProgramados!idTurno)) & "/"
                If InStr(lcTurnos, lcSql) > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsDiasProgramados!Codigo & " " & rsDiasProgramados!descripcion)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = "(" & rsDiasProgramados!Codigo & ") " & rsDiasProgramados!descripcion
                    End If
                    iFila = iFila + 1
                End If
                rsDiasProgramados.MoveNext
           Loop
        End If
    End If
    
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 36
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
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
Resume
End Sub
Sub EjecutaFormulario()
    Dim oFormulario As New ProgMedicaReporte
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

Function CantidadDeHoras(sHoraInicio As String, sHoraFin As String) As Integer
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer

    sHoras = Split(sHoraInicio, ":")
    iHoraIni = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60
    
    sHoras = Split(sHoraFin, ":")
    iHoraFin = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60
        
    CantidadDeHoras = iHoraFin - iHoraIni

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rlRepProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reporte de Producción
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario  As Long
Public Event ProgressActualizaValor(ByRef lnValorActual As Long, ByRef lnValorTotal As Long)
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New LabRepProduccion
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub



Sub ReportePorPacientes(lcFechaInicio As String, lcHoraInicio As String, lcFechaFinal As String, _
                        lcHoraFinal As String, lbConRecalculo As Boolean, lcTitulo As String, lnHwnd As Long, _
                        lbSoloGestantes As Boolean)
    Dim rsReporte As New ADODB.Recordset
    Dim oRsTmp1 As New ADODB.Recordset
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim oConexion As New Connection
    Dim mo_ReporteUtil As New ReporteUtil
    Dim ret As Long
    Dim mda_FechaInicio As Date
    Dim mda_FechaFin As Date, lRecordCount As Long, f As Long, lbCont2 As Boolean
    mda_FechaInicio = Format(lcFechaInicio & " " & lcHoraInicio, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
    mda_FechaFin = Format(lcFechaFinal & " " & lcHoraFinal, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '************PROCEDIMIENTO ALMACENADO sin recalculo***************************************
    Set rsReporte = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorFechas(mda_FechaInicio, mda_FechaFin, _
                                            sghPorFechaYhora)
    If lbSoloGestantes = True Then
       rsReporte.Filter = "Eo_EG>0"
    End If
    lRecordCount = rsReporte.RecordCount
    If lRecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        Dim iFila As Long, iColumna As Integer, lnTotal As Double, lnLineas As Integer, lcPaciente As String
        Dim ldFecha As Date, lnIdPaciente As Long, lnTotalGeneral As Double, lnCantidadItems As Long
        Dim lbEsOpenOffice As Boolean, lcTieneResultado As String
        
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
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
        End If
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\LabProductivadXpaciente.ods"
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
            Set oExcel = GalenhosExcelApplication()
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LabProductivadXpaciente.xls")
            oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 2).setFormula(lcTitulo)
            Call Feuille.getcellbyposition(8, 0).Width(3)
        Else
            oWorkSheet.Cells(3, 2).Value = lcTitulo
            oWorkSheet.Cells(1, 9).ColumnWidth = 3
        End If
        iFila = 5: lnTotalGeneral = 0: lnCantidadItems = 0
        rsReporte.MoveFirst
        f = 0
        Do While Not rsReporte.EOF
           f = f + 1
           RaiseEvent ProgressActualizaValor(f, lRecordCount)
           '
           lnCantidadItems = lnCantidadItems + 1
           ldFecha = rsReporte.Fields!fecha
           lcPaciente = rsReporte.Fields!Paciente
           If lbEsOpenOffice = True Then
               Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields!fecha)
           Else
               oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!fecha
           End If
           If IsNull(rsReporte.Fields!idPaciente) Then
            If lbEsOpenOffice = True Then
               Call Feuille.getcellbyposition(2, iFila - 1).setFormula(IIf(IsNull(lcPaciente), "", lcPaciente))
            Else
               oWorkSheet.Cells(iFila, 3).Value = lcPaciente
                'oWorkSheet.Cells(iFila, 3).Value = "(" & Trim(Str(rsReporte.Fields!NroHistoriaClinica)) & ") " & _
                                                   lcPaciente
            End If
           Else
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("(" & Trim(Str(rsReporte.Fields!NroHistoriaClinica)) & ") " & IIf(IsNull(lcPaciente), "", lcPaciente))
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields!dni), "", rsReporte.Fields!dni))
            Else
                oWorkSheet.Cells(iFila, 3).Value = "(" & Trim(Str(rsReporte.Fields!NroHistoriaClinica)) & ") " & lcPaciente
                oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields!dni
            End If
           End If
           '
           lbCont2 = True
           If lbConRecalculo = True Then
                If IsNull(rsReporte!idCuentaAtencion) Then
                   lbCont2 = False
                ElseIf rsReporte!idCuentaAtencion = 0 Then
                   lbCont2 = False
                End If
           Else
                lbCont2 = False
           End If
           If lbCont2 = True Then
                'con RECALCULO, busca FUENTE FINANCIAMIENTO
                Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
                If oRsTmp1.RecordCount > 0 Then
                   If lbEsOpenOffice = True Then
                      Call Feuille.getcellbyposition(4, iFila - 1).setFormula(oRsTmp1!dFuenteFinanciamiento)
                   Else
                      oWorkSheet.Cells(iFila, 5).Value = oRsTmp1!dFuenteFinanciamiento
                   End If
                End If
                oRsTmp1.Close
           Else
                'sin RECALCULO
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields!nombrePlan), "", rsReporte.Fields!nombrePlan))
                Else
                    oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields!nombrePlan
                End If
           End If
           '
           lnTotal = 0: lnLineas = 0
           Do While Not rsReporte.EOF And ldFecha = rsReporte.Fields!fecha And lcPaciente = rsReporte.Fields!Paciente
                lnTotal = lnTotal + rsReporte.Fields!Total
                lnTotalGeneral = lnTotalGeneral + rsReporte.Fields!Total
                lnLineas = lnLineas + 1
                '
                lcTieneResultado = " "
                Set oRsTmp1 = mo_ReglasLaboratorio.LabResultadosPorItemsSeleccionarXfiltro("idOrden=" & rsReporte!idOrden & _
                                                                                           " and idProductoCpt=" & rsReporte!idProductoCPT)
                If oRsTmp1.RecordCount > 0 Then
                   lcTieneResultado = "R"
                Else
                   If mo_ReglasLaboratorio.PruebaTieneResultado(rsReporte!Codigo, rsReporte!idOrden) = True Then
                      lcTieneResultado = "R"
                   End If
                End If
                oRsTmp1.Close
                '
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Trim(rsReporte.Fields!Nombre) & " (" & Trim(rsReporte.Fields!Codigo) & ")")
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte.Fields!Total)
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(lcTieneResultado)
                Else
                    oWorkSheet.Cells(iFila, 6).Value = Trim(rsReporte.Fields!Nombre) & " (" & Trim(rsReporte.Fields!Codigo) & ")"
                    oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields!Total
                    oWorkSheet.Cells(iFila, 9).Value = lcTieneResultado
                End If
              iFila = iFila + 1
              rsReporte.MoveNext
              If rsReporte.EOF Then
                 Exit Do
              End If
           Loop
           iFila = iFila - 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnTotal)
            Else
                oWorkSheet.Cells(iFila, 8).Value = lnTotal
            End If
           iFila = iFila + 1
        Loop
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("----------------------------------------------")
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula("------------------")
        Else
            oWorkSheet.Cells(iFila, 3).Value = "----------------------------------------------"
            oWorkSheet.Cells(iFila, 8).Value = "------------------"
        End If
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("N° Pacientes: " & Trim(Str(lnCantidadItems)))
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnTotalGeneral)
        Else
            oWorkSheet.Cells(iFila, 3).Value = "N° Pacientes: " & Trim(Str(lnCantidadItems))
            oWorkSheet.Cells(iFila, 8).Value = lnTotalGeneral
        End If
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":H" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 100
        Else
            oWorkSheet.Range(oWorkSheet.Cells(iFila, 3), oWorkSheet.Cells(iFila + 2, 100)).Select
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
    oConexion.Close
    
    Set rsReporte = Nothing
    Set oRsTmp1 = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oConexion = Nothing
    
    Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select


End Sub



'debb-12/09/2016
Sub ReportePorPacientesResultado(lcFechaInicio As String, lcHoraInicio As String, lcFechaFinal As String, _
                                 lcHoraFinal As String, lbConRecalculo As Boolean, ml_TextoDelFiltro As String, _
                                 lnHwnd As Long, lbSoloGestantes As Boolean)
    Dim rsReporte As New ADODB.Recordset
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mda_FechaInicio As Date
    Dim mda_FechaFin As Date, lRecordCount As Long, f As Long
    mda_FechaInicio = Format(lcFechaInicio & " " & lcHoraInicio, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
    mda_FechaFin = Format(lcFechaFinal & " " & lcHoraFinal, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
    If lbConRecalculo = True Then
        Dim mo_ReporteUtil As New ReporteUtil
        Dim ret As Long
        Dim lcSql As String
       '*********************************************************** con recalculos***************************************
       MsgBox "Se está trabajando el reporte CON RECALCULO, por ahora quite el CHECK DE RECALCULO"
       Exit Sub   'falta
    Else
       '*********************************************************** sin recalculos***************************************
       Set rsReporte = mo_ReglasLaboratorio.labResultadoXfechas(mda_FechaInicio, mda_FechaFin, lbSoloGestantes)
    End If
    lRecordCount = rsReporte.RecordCount
    If lRecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
       mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Resultados de Laboratorio", ml_TextoDelFiltro, "", lnHwnd, True, True
    End If
    Set mo_ReglasLaboratorio = Nothing
    Set rsReporte = Nothing
    Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select


End Sub



Sub ReportePorGrupoExamen(lcFechaInicio As String, lcHoraInicio As String, lcFechaFinal As String, _
                        lcHoraFinal As String, lbConRecalculo As Boolean, lcTitulo As String, lnHwnd As Long, _
                        lbSoloGestantes As Boolean)
    On Error GoTo ManejadorError
    Dim mrs_Tmp As New Recordset
    Dim rsReporte As New ADODB.Recordset
    Dim oRsTmp1 As New ADODB.Recordset
    Dim oRsTmp2 As New ADODB.Recordset
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    Dim mo_ReglasConfiguarcionReslab As New ReglasConfiguarcionReslab
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim oConexion As New Connection
    Dim mo_ReporteUtil As New ReporteUtil
    Dim oEdad As Edad
    Dim ret As Long
    Dim mda_FechaInicio As Date, lcGrupo As String, lnEdad As Integer, lnTedad As Long

    Dim mda_FechaFin As Date, lRecordCount As Long, f As Long, lbCont2 As Boolean
    Dim lnIdFuenteFinanciamiento As Long, lnIdTipoServicio As Long, lcSql As String, lcTieneResultado As String
    
    mda_FechaInicio = Format(lcFechaInicio & " " & lcHoraInicio, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
    mda_FechaFin = Format(lcFechaFinal & " " & lcHoraFinal, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    '************PROCEDIMIENTO ALMACENADO sin recalculo***************************************
    Set rsReporte = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorFechas(mda_FechaInicio, mda_FechaFin, _
                                            sghPorFechaYhora)
    If lbSoloGestantes = True Then
       rsReporte.Filter = "Eo_EG>0"
    End If
    lRecordCount = rsReporte.RecordCount
    If lRecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        Set oRsTmp2 = mo_ReglasConfiguarcionReslab.LabGruposSeleccionarTodos
        With mrs_Tmp
              .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable          'cpt (descripcion+codigo)
              .Fields.Append "Ene", adInteger, 4              'total-hospi
              .Fields.Append "Feb", adInteger, 4               'total-emerg
              .Fields.Append "Mar", adInteger, 4              'total-ce
              .Fields.Append "Abr", adInteger, 4              'soat-hospi
              .Fields.Append "May", adInteger, 4              'soat-emerg
              .Fields.Append "Jun", adInteger, 4               'soat-ce
              .Fields.Append "Jul", adInteger, 4             '0-29d
              .Fields.Append "Ago", adInteger, 4              '1-11mese
              .Fields.Append "Set", adInteger, 4               '1-4 años
              .Fields.Append "Oct", adInteger, 4              '5-9 años
              .Fields.Append "Nov", adInteger, 4              '10-19 años
              .Fields.Append "Dic", adInteger, 4              '20-64 años
              .Fields.Append "total", adInteger, 4             '65 años a mas
              .Fields.Append "totalAlm", adInteger, 4         'total-otros
              .Fields.Append "Promedio", adDouble                               'soat-otros
              .Fields.Append "mesesExistencia", adDouble
              .Fields.Append "Estado", adVarChar, 50, adFldIsNullable           'grupo examen
              .LockType = adLockOptimistic
              .Open
        End With
'rsReporte.Filter = "codigo='82465'"
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lcTieneResultado = " "
            Set oRsTmp1 = mo_ReglasLaboratorio.LabResultadosPorItemsSeleccionarXfiltro("idOrden=" & rsReporte!idOrden & _
                                                                                       " and idProductoCpt=" & rsReporte!idProductoCPT)
            If oRsTmp1.RecordCount > 0 Then
               lcTieneResultado = "R"
            Else
               If mo_ReglasLaboratorio.PruebaTieneResultado(rsReporte!Codigo, rsReporte!idOrden) = True Then
                  lcTieneResultado = "R"
               End If
            End If
            oRsTmp1.Close
            If lcTieneResultado = "R" Then
                
                '
                lnEdad = 0
                lnTedad = 0
                lnIdFuenteFinanciamiento = 0
                lnIdTipoServicio = 0
                lbCont2 = True
                If lbConRecalculo = True Then
                     If IsNull(rsReporte!idCuentaAtencion) Then
                        lbCont2 = False
                     ElseIf rsReporte!idCuentaAtencion = 0 Then
                        lbCont2 = False
                     End If
                Else
                     lbCont2 = False
                End If
                If lbCont2 = True Then
                     'con RECALCULO, busca FUENTE FINANCIAMIENTO
                     Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
                     If oRsTmp1.RecordCount > 0 Then
                        lnIdFuenteFinanciamiento = oRsTmp1!idFuenteFinanciamiento
                        lnIdTipoServicio = oRsTmp1!idTipoServicio
                        lnEdad = oRsTmp1!Edad
                        lnTedad = oRsTmp1!idTipoEdad
                     End If
                     oRsTmp1.Close
                Else
                    'sin RECALCULO
                    lnIdFuenteFinanciamiento = rsReporte!idFuenteFinanciamiento
                    If Not IsNull(rsReporte!idTipoServicio) Then
                       lnIdTipoServicio = rsReporte!idTipoServicio
                    End If
                    If Not IsNull(rsReporte!FechaNacimiento) Then
                        oEdad = sighentidades.CalcularEdad(rsReporte!FechaNacimiento, rsReporte!fecha)
                        lnEdad = oEdad.Edad
                        lnTedad = oEdad.TipoEdad
                    End If
                    '
                End If
                '
                lcSql = Left(rsReporte!Nombre, 120) & " (" & Trim(rsReporte!Codigo) & ")"
                lbCont2 = True
                If mrs_Tmp.RecordCount > 0 Then
                   mrs_Tmp.MoveFirst
                   mrs_Tmp.Find "nombre='" & lcSql & "'"
                   If Not mrs_Tmp.EOF Then
                      lbCont2 = False
                   End If
                End If
                If lbCont2 = True Then
                   '
                   lcGrupo = ""
                   Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro(" dbo.FactCatalogoServicios.Codigo='" & rsReporte!Codigo & "' ")
                   If oRsTmp1.RecordCount > 0 Then
                       oRsTmp2.MoveFirst
                       oRsTmp2.Find "idGrupo=" & oRsTmp1!idGrupo
                       If Not oRsTmp2.EOF Then
                          lcGrupo = oRsTmp2!NombreGrupo
                       End If
                   Else
                        oRsTmp1.Close
                        Set oRsTmp1 = mo_ReglasLaboratorio.LabPruebasSeleccionarPorCpt(rsReporte!Codigo, oConexion)
                        If oRsTmp1.RecordCount > 0 Then
                           lcGrupo = oRsTmp1!dGrupoNombre
                        End If
                   End If
                   oRsTmp1.Close
                   '
                   mrs_Tmp.AddNew
                   mrs_Tmp!Nombre = lcSql
                   mrs_Tmp!Estado = Left(lcGrupo, 50)
                End If
                Select Case lnIdTipoServicio
                Case sghTipoServicio.sghConsultaExterna
                     If lbCont2 = True Then
                        mrs_Tmp!mar = 1
                     Else
                        mrs_Tmp!mar = mrs_Tmp!mar + 1
                     End If
                     If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSoat Then
                         If lbCont2 = True Then
                            mrs_Tmp!jun = 1
                         Else
                            mrs_Tmp!jun = mrs_Tmp!jun + 1
                         End If
                     End If
                Case sghTipoServicio.sghEmergenciaConsultorios
                     If lbCont2 = True Then
                        mrs_Tmp!feb = 1
                     Else
                        mrs_Tmp!feb = mrs_Tmp!feb + 1
                     End If
                     If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSoat Then
                        If lbCont2 = True Then
                            mrs_Tmp!may = 1
                        Else
                            mrs_Tmp!may = mrs_Tmp!may + 1
                        End If
                     End If
                Case sghTipoServicio.sghHospitalizacion
                     If lbCont2 = True Then
                         mrs_Tmp!ene = 1
                     Else
                         mrs_Tmp!ene = mrs_Tmp!ene + 1
                     End If
                     If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSoat Then
                        If lbCont2 = True Then
                            mrs_Tmp!abr = 1
                        Else
                            mrs_Tmp!abr = mrs_Tmp!abr + 1
                        End If
                     End If
                Case Else
                     If lbCont2 = True Then
                        mrs_Tmp!totalAlm = 1
                     Else
                        mrs_Tmp!totalAlm = mrs_Tmp!totalAlm + 1
                     End If
                     If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSoat Then
                        If lbCont2 = True Then
                           mrs_Tmp!promedio = 1
                        Else
                            mrs_Tmp!promedio = mrs_Tmp!promedio + 1
                        End If
                     End If
                End Select
                '
                If lnEdad = 0 And lnTedad = 0 Then
                   If lbCont2 = True Then
                      mrs_Tmp!dic = 1
                   Else
                       mrs_Tmp!dic = mrs_Tmp!dic + 1
                   End If
                Else
                    Select Case sighentidades.EdadDevuelveTipo(lnTedad)
                    Case "H", "D"
                         If lbCont2 = True Then
                            mrs_Tmp!jul = 1
                         Else
                            mrs_Tmp!jul = mrs_Tmp!jul + 1
                         End If
                    Case "M"
                         If lbCont2 = True Then
                            mrs_Tmp!ago = 1
                         Else
                            mrs_Tmp!ago = mrs_Tmp!ago + 1
                         End If
                    Case Else
                         If lnEdad >= 1 And lnEdad <= 4 Then
                            If lbCont2 = True Then
                               mrs_Tmp!Set = 1
                            Else
                                mrs_Tmp!Set = mrs_Tmp!Set + 1
                            End If
                         ElseIf lnEdad >= 5 And lnEdad <= 9 Then
                            If lbCont2 = True Then
                                mrs_Tmp!Oct = 1
                            Else
                                mrs_Tmp!Oct = mrs_Tmp!Oct + 1
                            End If
                         ElseIf lnEdad >= 10 And lnEdad <= 19 Then
                            If lbCont2 = True Then
                               mrs_Tmp!nov = 1
                            Else
                               mrs_Tmp!nov = mrs_Tmp!nov + 1
                            End If
                         ElseIf lnEdad >= 20 And lnEdad <= 64 Then
                            If lbCont2 = True Then
                               mrs_Tmp!dic = 1
                            Else
                                mrs_Tmp!dic = mrs_Tmp!dic + 1
                            End If
                         Else
                            If lbCont2 = True Then
                               mrs_Tmp!Total = 1
                            Else
                                mrs_Tmp!Total = mrs_Tmp!Total + 1
                            End If
                         End If
                    End Select
                End If
                '
                mrs_Tmp.Update
            End If
            rsReporte.MoveNext
        Loop
    End If
    oConexion.Close
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        Dim oRptClaseCry1 As New frmCrystalR
        oRptClaseCry1.TextoDelFiltro = lcTitulo
        oRptClaseCry1.Excel = False
        oRptClaseCry1.Archivo = "LabPorGrupoExamen"
        oRptClaseCry1.Tabla = mrs_Tmp
        oRptClaseCry1.Show vbModal
        Set oRptClaseCry1 = Nothing
    End If
    
    Set rsReporte = Nothing
    Set oRsTmp1 = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oConexion = Nothing
    Set mo_ReporteUtil = Nothing
    Set oRsTmp2 = Nothing
    Set mo_ReglasConfiguarcionReslab = Nothing
    
    Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select

End Sub



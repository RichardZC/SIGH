VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "laboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Mantenimiento de Patología Clínica, Anatomía Patológica, Banco de Sangre
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_IdMovimiento As New frmLaboratorio
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReporteUtil As New ReporteUtil
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdMovimiento As Long
Dim ml_puntoCarga As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_areaTrabajo As Long
Dim ml_SeEligioGridBoleta As Boolean

Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property
Property Let AreaTrabajo(lValue As Long)
    ml_areaTrabajo = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
    mo_IdMovimiento.Opcion = mi_Opcion
    mo_IdMovimiento.idUsuario = ml_idUsuario
    mo_IdMovimiento.IdMovimiento = ml_IdMovimiento
    mo_IdMovimiento.puntoCarga = ml_puntoCarga
    mo_IdMovimiento.AreaTrabajo = ml_areaTrabajo
    mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
    mo_IdMovimiento.SeEligioGridBoleta = ml_SeEligioGridBoleta
    mo_IdMovimiento.Show 1
    Set mo_IdMovimiento = Nothing
End Sub

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Let Opcion(lValue As sghOpciones)
  mi_Opcion = lValue
End Property

Property Let IdMovimiento(lValue As Long)
  ml_IdMovimiento = lValue
End Property

Property Let puntoCarga(lValue As Long)
  ml_puntoCarga = lValue
End Property



Sub ImpresionDeItems(ml_IdMovimiento As Long, lcResponsableImg As String, lcFechaRealizaCpt As String, _
                     lcServicio As String, lnHwnd As Long, lcMedico As String)
    Dim mrs_Tmp As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oConexion As New Connection
    Dim mo_Diagnostico As New DODiagnostico
    Dim mo_reglasComunes As New ReglasComunes
    Dim mo_reglasCaja As New ReglasCaja
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim mo_ReglasLaboratorio As New ReglasLaboratorio
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReglasConfiguarcionReslab As New ReglasConfiguarcionReslab
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim lcDx As String, lnIdComprobantePago As Long, lcBoleta As String, ml_IdDiagnostico As Long
    Dim lcPaciente As String, lcEdad As String, lcFfinanciamiento As String, lnIdCuenta As Long
    Dim lbEsOpenOffice As Boolean, iFila As Long, lcGrupo As String, lcSql As String
    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With mrs_Tmp
        .Fields.Append "cpt", adVarChar, 250, adFldIsNullable
        .Fields.Append "cptCantidad", adInteger
        .Fields.Append "grupo", adVarChar, 100, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
    End With
    Set oRsTmp2 = mo_ReglasConfiguarcionReslab.LabGruposSeleccionarTodos
    Set oRsTmp = mo_ReglasLaboratorio.LabMovimientoCPTSeleccionarPorIdMovimiento(ml_IdMovimiento)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lcGrupo = "-"
            Set oRsTmp1 = mo_ReglasLaboratorio.LabPruebasSeleccionarPorCpt(oRsTmp!Codigo, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               lcGrupo = oRsTmp1!dGrupoNombre
            Else
               oRsTmp1.Close
               Set oRsTmp1 = mo_ReglasConfiguarcionReslab.LabItemsCPTSeleccionarPorIdRecordset(oRsTmp!idProductoCPT)
               If oRsTmp1.RecordCount > 0 Then
                  oRsTmp2.Filter = "idGrupo=" & oRsTmp1!idGrupo
                  If oRsTmp2.RecordCount > 0 Then
                     lcGrupo = oRsTmp2!NombreGrupo
                  End If
               End If
            End If
            oRsTmp1.Close
            
            mrs_Tmp.AddNew
            mrs_Tmp!cpt = Left(oRsTmp!Codigo & " " & oRsTmp!Nombre, 250)
            mrs_Tmp!cptCantidad = oRsTmp!Cantidad
            mrs_Tmp!Grupo = lcGrupo
            mrs_Tmp.Update
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorIdMovimiento(ml_IdMovimiento)
    If oRsTmp.RecordCount > 0 Then
            
            '
            lcDx = ""
            ml_IdDiagnostico = IIf(IsNull(oRsTmp!idDiagnostico), 0, oRsTmp!idDiagnostico)
            If ml_IdDiagnostico > 0 Then
                Set mo_Diagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(ml_IdDiagnostico)
                lcDx = mo_Diagnostico.CodigoCIE2004 & " " & mo_Diagnostico.descripcion
            End If
            '
            lcPaciente = ""
            lcBoleta = ""
            lcFfinanciamiento = ""
            lnIdComprobantePago = IIf(IsNull(oRsTmp.Fields!IdComprobantePago), 0, oRsTmp.Fields!IdComprobantePago)
            If lnIdComprobantePago > 0 Then
               Set oRsTmp1 = mo_reglasCaja.CajaComprobantesSeleccionarPorId(lnIdComprobantePago, oConexion)
               If oRsTmp1.RecordCount > 0 Then
                  lcBoleta = oRsTmp1!NroSerie & "-" & oRsTmp1!nroDocumento
                  lcFfinanciamiento = "Particular"
                  lcPaciente = oRsTmp1!razonSocial
                  If Not IsNull(oRsTmp1!idCuentaAtencion) Then
                     lnIdCuenta = oRsTmp1!idCuentaAtencion
                     oRsTmp1.Close
                     Set oRsTmp1 = mo_ReglasAdmision.AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio(0, lnIdCuenta, "", "", 0, "")
                     If oRsTmp1.RecordCount > 0 Then
                        lcPaciente = Trim(Str(oRsTmp1!NroHistoriaClinica)) & " " & lcPaciente
                        lcServicio = IIf(IsNull(oRsTmp1!ServicioIngreso), "", oRsTmp1!ServicioIngreso)
                     End If
                  End If
               End If
               oRsTmp1.Close
            End If
            '
            lcEdad = ""
            lnIdCuenta = IIf(IsNull(oRsTmp.Fields!idCuentaAtencion), 0, oRsTmp.Fields!idCuentaAtencion)
            If lnIdCuenta > 0 Then
                Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(lnIdCuenta, oConexion)
                If oRsTmp1.RecordCount > 0 Then
                   lcPaciente = Trim(Str(oRsTmp1!NroHistoriaClinica)) & " " & Trim(oRsTmp1!ApellidoPaterno) & " " & _
                              Trim(oRsTmp1!ApellidoMaterno) & " " & oRsTmp1!PrimerNombre
                   lcFfinanciamiento = oRsTmp1!dFuenteFinanciamiento
                   lcEdad = Trim(Str(oRsTmp1!Edad)) & " " & oRsTmp1!tEdad
'                   iif(orstmp1!idTipoEdad= sghTipoEdades.sghAño," Años", _
'                   iif(orstmp1!idTipoEdad=sghTipoEdades.sghAño," Meses", _
'                   iif(orstmp1!idTipoEdad=sghTipoEdades.sghDias ," Días", "Horas")))
                End If
                oRsTmp1.Close
            End If
            '
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
                lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
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
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
                oWorkBookPlantilla.Worksheets("Hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            iFila = 1
            mrs_Tmp.Sort = "grupo"
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                lcGrupo = mrs_Tmp!Grupo
                If lbEsOpenOffice = True Then
                   ' Call Feuille.getcellbyposition(iFila, 3).setFormula(oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres)
                Else
                    oWorkSheet.Cells(iFila, 1).Value = "Movimiento N° " & Trim(Str(ml_IdMovimiento))
                    oWorkSheet.Cells(iFila, 3).Value = "Paciente: " & lcPaciente
                    iFila = iFila + 1
                    oWorkSheet.Cells(iFila, 1).Value = "F.Financ: " & lcFfinanciamiento
                    oWorkSheet.Cells(iFila, 3).Value = "Servicio: " & lcServicio
                    iFila = iFila + 1
                    oWorkSheet.Cells(iFila, 1).Value = "F. realiza CPT: " & lcFechaRealizaCpt
                    oWorkSheet.Cells(iFila, 4).Value = "Grupo: " & lcGrupo
                    iFila = iFila + 1
                    oWorkSheet.Cells(iFila, 1).Value = "CPT"
                    oWorkSheet.Cells(iFila, 6).Value = "Cantidad "
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 6
                End If
                iFila = iFila + 1
                
                Do While Not mrs_Tmp.EOF And lcGrupo = mrs_Tmp!Grupo
                   If lbEsOpenOffice = True Then
                   Else
                        lcSql = mrs_Tmp!cpt
                        oWorkSheet.Cells(iFila, 1).Value = lcSql
                        lcSql = Trim(Str(mrs_Tmp!cptCantidad))
                        oWorkSheet.Cells(iFila, 6).Value = lcSql
                   End If
                   iFila = iFila + 1
                   mrs_Tmp.MoveNext
                   If mrs_Tmp.EOF Then
                      Exit Do
                   End If
                Loop
                iFila = iFila + 3
            Loop
            If lbEsOpenOffice = True Then
            Else
                 oWorkSheet.Cells(iFila, 1).Value = "Prescribe Receta: " & lcMedico
            End If
            iFila = iFila + 1
            
            
'            If lbEsOpenOffice = True Then
'                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Y" & CStr(iFila - 1))
'                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
'            Else
'                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila - 1, 6
'            End If
        
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 6
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
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
            
            
            
'            Set repLaboratorio.DataSource = mrs_Tmp
'            repLaboratorio.RightMargin = 10
'            repLaboratorio.TopMargin = 10
'            repLaboratorio.LeftMargin = 10
'            repLaboratorio.BottomMargin = 10
'            repLaboratorio.Sections("cabecera").Controls("lblTitulo").Caption = "MOVIMIENTO N° " & Trim(Str(ml_IdMovimiento))
'            repLaboratorio.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
'            repLaboratorio.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
'            repLaboratorio.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
'            repLaboratorio.Sections("cabecera").Controls("lblFmovimiento").Caption = lcFechaRealizaCpt
'            repLaboratorio.Sections("cabecera").Controls("lblPaciente").Caption = lcPaciente
'            repLaboratorio.Sections("cabecera").Controls("lblEdad").Caption = lcEdad
'            repLaboratorio.Sections("cabecera").Controls("lblFfinanciamiento").Caption = lcFfinanciamiento
'            repLaboratorio.Sections("cabecera").Controls("lblCuenta").Caption = lnIdCuenta
'            repLaboratorio.Sections("cabecera").Controls("lblBoleta").Caption = lcBoleta
'            repLaboratorio.Sections("cabecera").Controls("lblServicio").Caption = lcServicio
'            repLaboratorio.Orientation = rptOrientPortrait
'            repLaboratorio.Show 1
    End If
    Set mrs_Tmp = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set mo_Diagnostico = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_reglasCaja = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasConfiguarcionReslab = Nothing
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHCconVIH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historias de Pacientes con VIH
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Historias Clinicas VIH
Option Explicit
Dim ml_IdTipoHistoria As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim ml_tipoFechaNacimiento As Integer
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdTipoHistoria(lValue As Long)
    ml_IdTipoHistoria = lValue
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

Property Let TipoFechaNacimiento(lValue As String)
    ml_tipoFechaNacimiento = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New AHCconVIH
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
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

    Set rsReporte = HistoriasClinicasSeleccionarPorTipo()
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\AHCegrMed24.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\AHCegrMed24.xls")
        oWorkBookPlantilla.Worksheets("AHCegrMed24").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
        'Filtra los Datos
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 1).setFormula("HISTORIAS CLINICAS SOLICITADAS POR SERVICIO")
            Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
            Call Feuille.getcellbyposition(6, 4).setFormula("")
        Else
            oWorkSheet.Cells(2, 3).Value = "HISTORIAS CLINICAS SOLICITADAS POR SERVICIO"
            oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
            oWorkSheet.Cells(5, 7).Value = ""
        End If
        iFila = 6
        Do While Not rsReporte.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("nroHistoriaClinica").Value)
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("lcPaciente").Value)
            Else
                oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("nroHistoriaClinica").Value
                oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("lcPaciente").Value
            End If
           iFila = iFila + 1
           rsReporte.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":H" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Historias: " & Trim(Str(rsReporte.RecordCount)))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 8
            oWorkSheet.Cells(iFila, 2).Value = "Nº Historias: " & Trim(Str(rsReporte.RecordCount))
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
            Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
        End If
         
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

Sub CrearReporte(lbEnExcel As Boolean, lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    On Error GoTo ManejadorError
    'Filtra los Datos
    Set rsReporte = HistoriasClinicasSeleccionarPorTipo()
    If lbEnExcel = True Then
        mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Historias por Tipo", ml_TextoDelFiltro, "Nro historias: " & Trim(Str(rsReporte.RecordCount)), lnHwnd
    Else
        Set RpAHCconVIH.DataSource = rsReporte
        RpAHCconVIH.RightMargin = 10
        RpAHCconVIH.TopMargin = 10
        RpAHCconVIH.LeftMargin = 10
        RpAHCconVIH.BottomMargin = 10
        RpAHCconVIH.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        RpAHCconVIH.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        RpAHCconVIH.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        RpAHCconVIH.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
        RpAHCconVIH.Sections("cabecera").Controls("lblFiltro").Caption = "(" & UCase(Trim(lcBuscaParametro.SeleccionaFilaParametro(205))) & ") " & ml_TextoDelFiltro
        Set RpAHCconVIH.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
        RpAHCconVIH.Sections("pie").Controls("lblNhistorias").Caption = rsReporte.RecordCount
        RpAHCconVIH.Orientation = rptOrientLandscape
        RpAHCconVIH.Show 1
    End If
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Resume
    Exit Sub
End Sub




Function HistoriasClinicasSeleccionarPorTipo() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set HistoriasClinicasSeleccionarPorTipo = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HistoriasClinicasSeleccionarPorTipo"
        Set oParameter = .CreateParameter("@IdtipoHistoria", adInteger, adParamInput, 0, ml_IdTipoHistoria): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@OrdenFiltro", adVarChar, adParamInput, 20, ml_OrdenFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoFechaNacimiento", adInteger, adParamInput, 0, ml_tipoFechaNacimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set HistoriasClinicasSeleccionarPorTipo = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Sub AfiliadosXsexoGrupoEdad(lcANIO As String)
     Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
     Dim lcBuscaParametro As New SIGHDatos.Parametros
     Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
     Dim oRsTmp As New Recordset
     Dim oConexion As New Connection
     Dim lcSql As String
     Dim lnTotalAfiliados As Long, lnTotalMasculino As Long, lnTotalFemenino As Long, lnTotalMenorasAanios As Long
     Dim lnPorFemenino As Double
     Dim lnMujeresEdadFertil As Long
     Dim lnGErn As Long, lnGE1anio As Long, lnGE1a4anios As Long, lnGE5a9anios As Long
     Dim lnGE10a11anios As Long, lnGENinos As Long
     Dim lnGEadolescente As Long, lnGEjoven As Long, lnGEadulto As Long, lnGEadultoMayor As Long
     Dim ldHoy As Date, lcFiliacionInicio As String, lcFiliacionFinal As String
     Dim lnGErnP As Double, lnGE1anioP As Double, lnGE1a4aniosP As Double, lnGE10a11aniosP As Double
     Dim lnGE5a9aniosP    As Double, lnGENinosP As Double, lnGEadolescenteP As Double
     Dim lnGEjovenP    As Double, lnGEadultoP As Double, lnGEadultoMayorP As Double, lnPorFertil As Double
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighentidades.CadenaConexion
     '
     ldHoy = lcBuscaParametro.RetornaFechaHoraServidorSQL
     '
     Set oRsTmp = mo_ReglasArchivoClinico.HistoriasClinicasSeleccionarTodosFcreacion(lcANIO, oConexion)
     If oRsTmp.RecordCount = 0 Then
        MsgBox "No existe ningun Paciente", vbInformation, "Pacientes"
     Else
        oRsTmp.MoveFirst
        lcFiliacionInicio = Format(oRsTmp.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
        oRsTmp.MoveLast
        lcFiliacionFinal = Format(oRsTmp.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
        
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarTotalMasculino(lcANIO, oConexion)
        lnTotalMasculino = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarTotalFemenino(lcANIO, oConexion)
        lnTotalFemenino = oRsTmp.RecordCount
        '
        lnTotalAfiliados = lnTotalFemenino + lnTotalMasculino
        lnPorFemenino = Round((lnTotalFemenino * 100 / lnTotalAfiliados), 1)
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMujeresEdadFertil(lcANIO, oConexion)
        lnMujeresEdadFertil = oRsTmp.RecordCount
        lnPorFertil = Round(lnMujeresEdadFertil * 100 / lnTotalFemenino, 1)
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarRn(lcANIO, oConexion)
        lnGErn = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMenores1anio(lcANIO, oConexion)
        lnGE1anio = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMayor1menor5(lcANIO, oConexion)
        lnGE1a4anios = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMayor5Menor10(lcANIO, oConexion)
        lnGE5a9anios = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSelecionarMayor10Menor12(lcANIO, oConexion)
        lnGE10a11anios = oRsTmp.RecordCount
        lnGENinos = lnGErn + lnGE1anio + lnGE1a4anios + lnGE5a9anios + lnGE10a11anios
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMayor12Menor18(lcANIO, oConexion)
        lnGEadolescente = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMayor18Menor30(lcANIO, oConexion)
        lnGEjoven = oRsTmp.RecordCount
        '
        Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarMayor30Menor60(lcANIO, oConexion)
        lnGEadulto = oRsTmp.RecordCount
        '
        lnGEadultoMayor = lnTotalAfiliados - (lnGENinos + lnGEadolescente + lnGEjoven + lnGEadulto)
        lnGErnP = Round(lnGErn * 100 / lnTotalAfiliados, 1)
        lnGE1anioP = Round(lnGE1anio * 100 / lnTotalAfiliados, 1)
        lnGE1a4aniosP = Round(lnGE1a4anios * 100 / lnTotalAfiliados, 1)
        lnGE5a9aniosP = Round(lnGE5a9anios * 100 / lnTotalAfiliados, 1)
        lnGE10a11aniosP = Round(lnGE10a11anios * 100 / lnTotalAfiliados, 1)
        lnGENinosP = lnGErnP + lnGE1anioP + lnGE1a4aniosP + lnGE5a9aniosP + lnGE10a11aniosP
        lnGEadolescenteP = Round(lnGEadolescente * 100 / lnTotalAfiliados, 1)
        lnGEjovenP = Round(lnGEjoven * 100 / lnTotalAfiliados, 1)
        lnGEadultoP = Round(lnGEadulto * 100 / lnTotalAfiliados, 1)
        lnGEadultoMayorP = 100 - (lnGENinosP + lnGEadolescenteP + lnGEjovenP + lnGEadultoP)
        
        '
        Set FormPacientes.DataSource = oRsTmp
       ' If Val(lcAnio) > 0 Then
       '    FormPacientes.Sections("cabecera").Controls("lblTitulo").Caption = "Indicadores de Pacientes en el Establecimiento (Año de inscripción HC: " & lcAnio & ")"
       ' End If
        FormPacientes.RightMargin = 10
        FormPacientes.TopMargin = 10
        FormPacientes.LeftMargin = 10
        FormPacientes.BottomMargin = 10
        
        FormPacientes.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        FormPacientes.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        FormPacientes.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        FormPacientes.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
        FormPacientes.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
        FormPacientes.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
        If Val(lcANIO) > 0 Then
           FormPacientes.Sections("cabecera").Controls("lblTitulo").Caption = "Indicadores de Pacientes en el Establecimiento : " & lcBuscaParametro.SeleccionaFilaParametro(205) & " (Año de inscripción HC: " & lcANIO & ")"
        Else
           FormPacientes.Sections("cabecera").Controls("lblTitulo").Caption = "Indicadores de Pacientes en el Establecimiento : " & lcBuscaParametro.SeleccionaFilaParametro(205)
        End If
        FormPacientes.Sections("cabecera").Controls("lblFinicial").Caption = lcFiliacionInicio
        FormPacientes.Sections("cabecera").Controls("lblFfinal").Caption = lcFiliacionFinal
        FormPacientes.Sections("cabecera").Controls("lblTotalPacientes").Caption = Format(lnTotalAfiliados, "###,###,###")
        FormPacientes.Sections("cabecera").Controls("lblTotalMujeres").Caption = Format(lnTotalFemenino, "###,###,###")
        FormPacientes.Sections("cabecera").Controls("lblTotalMujeresP").Caption = Format(lnPorFemenino, "###0.0")
        FormPacientes.Sections("cabecera").Controls("lblTotalHombres").Caption = Format(lnTotalMasculino, "###,###,###")
        FormPacientes.Sections("cabecera").Controls("lblTotalHombresP").Caption = Format(100 - lnPorFemenino, "###0.0")
        FormPacientes.Sections("cabecera").Controls("lblFertil").Caption = IIf(lnMujeresEdadFertil = 0, "0", Format(lnMujeresEdadFertil, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblFertilP").Caption = IIf(lnPorFertil = 0, "0", Format(lnPorFertil, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lblNinos").Caption = IIf(lnGENinos = 0, "0", Format(lnGENinos, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblNinosP").Caption = IIf(lnGENinosP = 0, "0", Format(lnGENinosP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lblRN").Caption = IIf(lnGErn = 0, "0", Format(lnGErn, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblRNp").Caption = IIf(lnGErnP = 0, "0", Format(lnGErnP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lbl1anio").Caption = IIf(lnGE1anio = 0, "0", Format(lnGE1anio, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lbl1anioP").Caption = IIf(lnGE1anioP = 0, "0", Format(lnGE1anioP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lbl1a4").Caption = IIf(lnGE1a4anios = 0, "0", Format(lnGE1a4anios, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lbl1a4P").Caption = IIf(lnGE1a4aniosP = 0, "0", Format(lnGE1a4aniosP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lbl5a9").Caption = IIf(lnGE5a9anios = 0, "0", Format(lnGE5a9anios, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lbl5a9p").Caption = IIf(lnGE5a9aniosP = 0, "0", Format(lnGE5a9aniosP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lbl10a11").Caption = IIf(lnGE10a11anios = 0, "0", Format(lnGE10a11anios, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lbl10a11P").Caption = IIf(lnGE10a11aniosP = 0, "0", Format(lnGE10a11aniosP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lblAdolecente").Caption = IIf(lnGEadolescente = 0, "0", Format(lnGEadolescente, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblAdolecenteP").Caption = IIf(lnGEadolescenteP = 0, "0", Format(lnGEadolescenteP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lblJoven").Caption = IIf(lnGEjoven = 0, "0", Format(lnGEjoven, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblJovenP").Caption = IIf(lnGEjovenP = 0, "0", Format(lnGEjovenP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lblAdulto").Caption = IIf(lnGEadulto = 0, "0", Format(lnGEadulto, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblAdultoP").Caption = IIf(lnGEadultoP = 0, "0", Format(lnGEadultoP, "###0.0"))
        FormPacientes.Sections("cabecera").Controls("lblAdultoM").Caption = IIf(lnGEadultoMayor = 0, "0", Format(lnGEadultoMayor, "###,###,###"))
        FormPacientes.Sections("cabecera").Controls("lblAdultoMp").Caption = IIf(lnGEadultoMayorP = 0, "0", Format(lnGEadultoMayorP, "###0.0"))
        Set FormPacientes.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")

        FormPacientes.Show 1
        '
        'debb-27/05/2015
        Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
        mo_reglasComunes.grabaTablaAuditoria ("FormPacientes: " & _
                                       lcFiliacionInicio & " " & lcFiliacionFinal)
        Set mo_reglasComunes = Nothing
        '
    End If
    Set mo_ReglasAdmision = Nothing
    Set lcBuscaParametro = Nothing
    Set oRsTmp = Nothing
    Set mo_ReglasArchivoClinico = Nothing
    Set oConexion = Nothing

End Sub


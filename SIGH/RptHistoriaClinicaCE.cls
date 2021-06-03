VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHistoriaClinicaCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reportes de Historia Clinica
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idCuentaAtencion As Long
Dim ml_idOrden As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReporteUtil As New ReporteUtil
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let IdOrden(lValue As Long)
    ml_idOrden = lValue
End Property
Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property
Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property




Sub CrearReporteHistoriaClinicaDeLaAtencionCE(ml_lnHwnd As Long) ' scrafet reporte 3

Dim rsreporte As New Recordset
Dim NroHistoria As Long
Dim lnTotalRegistro As Integer
Dim lbEsOpenOffice As Boolean
Dim lcSql As String

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
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEHistoriaClinicaCE.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEHistoriaClinicaCE.xls")
        oWorkBookPlantilla.Worksheets("HistoriaClinicaCE").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    
    Set rsreporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    If mo_AdminReportes.MensajeError <> "" Then
        MsgBox mo_AdminReportes.MensajeError, vbInformation, "Reporte"
    Else
        NroHistoria = rsreporte!NroHistoriaClinica
            If lbEsOpenOffice = True Then
                
'<(Inicio) Añadido Por: WABG el: 13/05/2021-09:39:47 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
                Call Feuille.getcellbyposition(22, 4).setFormula("'" & rsreporte!idCuentaAtencion) 'agregando para Nro Cuenta WABG
'</(Fin) Añadido Por: WABG el: 13/05/2021-09:39:47 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-74>
                                                                
                Call Feuille.getcellbyposition(84, 6).setFormula("'" & rsreporte!NroHistoriaClinica)
                Call Feuille.getcellbyposition(27, 9).setFormula("'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsreporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsreporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsreporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsreporte!SegundoNombre)))
                Call Feuille.getcellbyposition(27, 11).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!FechaIngreso))
                Call Feuille.getcellbyposition(27, 12).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!HoraIngreso))
                Call Feuille.getcellbyposition(27, 13).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!Servicio))
                Call Feuille.getcellbyposition(27, 14).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!EstadoCivil))
                Call Feuille.getcellbyposition(27, 15).setFormula("'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsreporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsreporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsreporte!NombresMedico)))
                Call Feuille.getcellbyposition(27, 16).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!EspecialidadMedico))
                Call Feuille.getcellbyposition(27, 17).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!DireccionDomicilio))
                Call Feuille.getcellbyposition(27, 18).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!DepartamentoDomicilio))
                Call Feuille.getcellbyposition(27, 19).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!DistritoDomicilio))
                
                Call Feuille.getcellbyposition(84, 11).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!Edad))
                Call Feuille.getcellbyposition(84, 12).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!FechaNacimiento))
                Call Feuille.getcellbyposition(84, 13).setFormula("'" & IIf(mo_ReporteUtil.NullToVacio(rsreporte!idTipoSexo) = "1", "Masculino", "Femenino"))
                Call Feuille.getcellbyposition(84, 14).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!nrodocumento))
                Call Feuille.getcellbyposition(84, 15).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!TipoDocumento))
                Call Feuille.getcellbyposition(84, 18).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!ProvinciaDomicilio))
                Call Feuille.getcellbyposition(84, 19).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!CentroPobladoDomicilio))
            Else
            
                           
'<(Inicio) Añadido Por: WABG el: 13/05/2021-09:40:30 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
                oWorkSheet.Cells(6, 21).Value = "'" & mo_ReporteUtil.NullToVacio(Str(rsreporte!idCuentaAtencion))
'</(Fin) Añadido Por: WABG el: 13/05/2021-09:40:30 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-74>
                           
                oWorkSheet.Cells(6, 85).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte!NroHistoriaClinica)), True)
                oWorkSheet.Cells(10, 28).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsreporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsreporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsreporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsreporte!SegundoNombre))
                oWorkSheet.Cells(12, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!FechaIngreso)
                oWorkSheet.Cells(13, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!HoraIngreso)
                oWorkSheet.Cells(14, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!Servicio)
                oWorkSheet.Cells(15, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!EstadoCivil)
                oWorkSheet.Cells(16, 28).Value = "'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsreporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsreporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsreporte!NombresMedico))
                oWorkSheet.Cells(17, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!EspecialidadMedico)
                oWorkSheet.Cells(18, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!DireccionDomicilio)
                oWorkSheet.Cells(19, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!DepartamentoDomicilio)
                oWorkSheet.Cells(20, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!DistritoDomicilio)
                
                oWorkSheet.Cells(12, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!Edad)
                oWorkSheet.Cells(13, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!FechaNacimiento)
                oWorkSheet.Cells(14, 85).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsreporte!idTipoSexo) = "1", "Masculino", "Femenino")
                oWorkSheet.Cells(15, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!nrodocumento)
                oWorkSheet.Cells(16, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!TipoDocumento)
                oWorkSheet.Cells(19, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!ProvinciaDomicilio)
                oWorkSheet.Cells(20, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!CentroPobladoDomicilio)
            End If
        '***************daniel barrantes**************
        '***************datos ADICIONALES DEL PACIENTE
        '***************
        rsreporte.Close
        Set rsreporte = PacientesSeleccionarPoNroHistoriaDevuelveProcedencia(NroHistoria)
        If rsreporte.RecordCount > 0 Then
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(84, 15).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!docupacion))
                
                Call Feuille.getcellbyposition(27, 24).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!dDptoProc))
                Call Feuille.getcellbyposition(27, 25).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!dprovinciaProc))
                Call Feuille.getcellbyposition(27, 26).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!ddistritoProc))
                Call Feuille.getcellbyposition(27, 27).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!dCentroPobladoProc))
            Else
'<(Inicio)Comentado Por: WABG el: 14/05/2021-09:14:36 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
'                oWorkSheet.Cells(16, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!docupacion)
'
'                oWorkSheet.Cells(25, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!dDptoProc)
'                oWorkSheet.Cells(26, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!dprovinciaProc)
'                oWorkSheet.Cells(27, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!ddistritoProc)
'                oWorkSheet.Cells(28, 28).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!dCentroPobladoProc)
'</(Fin)Comentado por: WABG el: 14/05/2021-09:14:36 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
            End If
        End If
        rsreporte.Close
        Set rsreporte = PacientesSeleccionarPoNroHistoriaDevuelveNacimiento(NroHistoria)
        If rsreporte.RecordCount > 0 Then
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(84, 24).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!ddptoNacim))
                Call Feuille.getcellbyposition(84, 25).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!dProvinciaNacim))
                Call Feuille.getcellbyposition(84, 26).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!ddistritoNacim))
                Call Feuille.getcellbyposition(84, 27).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!dCentroPobladoNacim))
            Else
'<(Inicio)Comentado Por: WABG el: 14/05/2021-09:14:51 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
'                oWorkSheet.Cells(25, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!ddptoNacim)
'                oWorkSheet.Cells(26, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!dProvinciaNacim)
'                oWorkSheet.Cells(27, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!ddistritoNacim)
'                oWorkSheet.Cells(28, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!dCentroPobladoNacim)
'</(Fin)Comentado por: WABG el: 14/05/2021-09:14:51 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
            End If
        End If
'<(Inicio) Añadido Por: WABG el: 13/05/2021-09:47:38 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-74>
        'rsreporte.Close
        Set rsreporte = mo_AdminReportes.ObtenerDatosAdicionalesDeAtencionPorIdAtencion(ml_idAtencion)
        If rsreporte.RecordCount > 0 Then
            If lbEsOpenOffice = True Then
            
                Call Feuille.getcellbyposition(16, 14).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!NroReferenciaOrigen))
                
            Else
            
                oWorkSheet.Cells(7, 84).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!NroReferenciaOrigen)
                oWorkSheet.Cells(7, 16).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!IdEstablecimientoOrigen)
                 oWorkSheet.Cells(7, 25).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!NombreEstablecimiento)
                 oWorkSheet.Cells(3, 40).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!FuenteFinanciamiento)
                
                
            End If
        End If
        ' Set rsreporte = Nothing
         rsreporte.Close
'</(Fin) Añadido Por: WABG el: 13/05/2021-09:47:38 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-74>

       ' If lbEsOpenOffice = True Then
       '    Call Feuille.getcellbyposition(27, 30).setFormula("'" & sighentidades.NombreUsuario)
       '    Call Feuille.getcellbyposition(84, 30).setFormula("'" & sighentidades.RetornaNombrePC)
       ' Else
       '     oWorkSheet.Cells(31, 28).Value = "'" & sighentidades.NombreUsuario
       '     oWorkSheet.Cells(31, 85).Value = "'" & sighentidades.RetornaNombrePC
       ' End If
        '******************* daniel barrantes *****************************

'        If oWorkSheet.PageSetup.PrintArea <> "" Then
'           oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 31)
'        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 0
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 85
            PrintArea(0).EndRow = 34
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
    End If
'    If lbEsOpenOffice = True Then
'    Else
'        oWorkBook.Close SaveChanges:=False
'    End If
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
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub

End Sub


'***********************daniel barrantes - inicio********************************
Function PacientesSeleccionarPoNroHistoriaDevuelveNacimiento(lnNroHistoriaClinica As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set PacientesSeleccionarPoNroHistoriaDevuelveNacimiento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PacientesSeleccionarPoNroHistoriaDevuelveNacimiento"
        Set oParameter = .CreateParameter("@nroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesSeleccionarPoNroHistoriaDevuelveNacimiento = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function PacientesSeleccionarPoNroHistoriaDevuelveProcedencia(lnNroHistoriaClinica As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set PacientesSeleccionarPoNroHistoriaDevuelveProcedencia = Nothing
    ms_MensajeError = ""
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "PacientesSeleccionarPoNroHistoriaDevuelveProcedencia"
        Set oParameter = .CreateParameter("@nroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesSeleccionarPoNroHistoriaDevuelveProcedencia = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function





'***********************daniel barrantes - fin********************************

'***************daniel barrantes**************
'***************imprime datos de PACIENTES en el "FORMATO SIS"
'***************
Sub ImprimeFormatoSIS(lnIdAtencion As Long, lnIdUsuario As Long, lnCabeceraDetalle As Integer)
'    Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
'    Dim mo_ReporteUtil As New ReporteUtil
'    Dim oExcel As Excel.Application
'    Dim oWorkBookPlantilla As Workbook
'    Dim oWorkBook As Workbook
'    Dim oWorkSheet As Worksheet
'    Dim oConexion As New Connection
'    Dim rsReporte As New Recordset
'    Dim rsDiagnosticos As New Recordset
'    Dim rsReferido As New Recordset
'    Dim sSQL As String: Dim lccodreferido As String: Dim lcdesreferido As String
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Dim oRsLaboraEn As Recordset
'    On Error GoTo ManejadorError
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    Set rsReporte = mo_ReglasAdmision.AtencionesFiltraDatosParaFormatoSIS(lnIdAtencion, oConexion)
'    'Crea nueva hoja
'    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'    Set oWorkBook = oExcel.Workbooks.Add
'    'Abre, copia y cierra la plantilla
'    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEsis.xls")
'    oWorkBookPlantilla.Worksheets("sis").Copy Before:=oWorkBook.Sheets(1)
'    oWorkBookPlantilla.Close
'    'Activa la primera hoja
'    Set oWorkSheet = oWorkBook.Sheets(1)
'    If lnCabeceraDetalle = 1 Or lnCabeceraDetalle = 3 Then
'        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        oWorkSheet.Cells(3, 16).Value = "'" & lcBuscaParametro.SeleccionaFilaParametro(239)
'        oWorkSheet.Cells(6, 1).Value = "'" & rsReporte!ApellidoPaterno
'        oWorkSheet.Cells(6, 35).Value = "'" & rsReporte!ApellidoMaterno
'        oWorkSheet.Cells(9, 1).Value = "'" & rsReporte!PrimerNombre
'        oWorkSheet.Cells(9, 35).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)
'        If rsReporte.Fields!idTipoSexo = 1 Then     'Masculino
'           oWorkSheet.Cells(13, 16).Value = "X"
'        End If
'        If Not IsNull(rsReporte!FechaNacimiento) Then
'           oWorkSheet.Cells(14, 1).Value = "'" & Day(rsReporte!FechaNacimiento) & "   " & Month(rsReporte!FechaNacimiento) & "     " & Year(rsReporte!FechaNacimiento)
'        End If
'        If rsReporte.Fields!idTipoSexo <> 1 Then     'Femenino
'           oWorkSheet.Cells(14, 16).Value = "X"
'        End If
'        lccodreferido = "": lcdesreferido = ""
'        If Not IsNull(rsReporte.Fields!IdTipoReferenciaOrigen) Then
'            If rsReporte.Fields!IdTipoReferenciaOrigen = 1 Then
'               'Referido MINSA
'               Set rsReferido = mo_reglasComunes.EstablecimientosXidentificador(rsReporte.Fields!idEstablecimientoOrigen, oConexion)
'               If rsReferido.RecordCount > 0 Then
'                  lccodreferido = rsReferido.Fields!Codigo
'                  lcdesreferido = rsReferido.Fields!Nombre
'               End If
'            ElseIf rsReporte.Fields!IdTipoReferenciaOrigen = 2 Then
'               'Referido Externo
'               Set rsReferido = mo_reglasComunes.EstablecimientosNoMinsaXidentificador(rsReporte.Fields!IdEstablecimientoNoMinsaOrigen, oConexion)
'               If rsReferido.RecordCount > 0 Then
'                  lccodreferido = ""
'                  lcdesreferido = rsReferido.Fields!Nombre
'               End If
'            End If
'        End If
'        oWorkSheet.Cells(19, 1).Value = "'" & lccodreferido
'        oWorkSheet.Cells(19, 20).Value = "'" & lcdesreferido
'        oWorkSheet.Cells(24, 1).Value = "'" & Day(rsReporte!FechaIngreso) & "     " & Month(rsReporte!FechaIngreso)
'        oWorkSheet.Cells(24, 30).Value = "'" & rsReporte!NroHistoriaClinica
'    End If
'    If lnCabeceraDetalle = 2 Or lnCabeceraDetalle = 3 Then
'        '******Triaje
'        Dim oAtencionesCE As New AtencionesCE
'        Dim oConexionJamo As New Connection
'        Dim mo_DOAtencionesCE As New DOAtencionesCE
'        oConexionJamo.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'        mo_DOAtencionesCE.idAtencion = lnIdAtencion
'        Set oAtencionesCE.Conexion = oConexionJamo
'        If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = True Then
'            If Not mo_DOAtencionesCE Is Nothing Then
'                 With mo_DOAtencionesCE
'                      oWorkSheet.Cells(38, 15).Value = "'" & mo_DOAtencionesCE.TriajePeso
'                      oWorkSheet.Cells(38, 23).Value = "'" & mo_DOAtencionesCE.TriajeTalla
'                 End With
'            End If
'        End If
'        Set oConexionJamo = Nothing
'        Set oAtencionesCE = Nothing
'        Set mo_DOAtencionesCE = Nothing
'        '******Dx
'        Dim iFila As Integer, lnLineas As Integer
'        Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
'        iFila = 52
'        lnLineas = 1
'        Set rsDiagnosticos = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(lnIdAtencion)
'        If rsDiagnosticos.RecordCount > 0 Then
'             rsDiagnosticos.MoveFirst
'             Do While Not rsDiagnosticos.EOF
'                If lnLineas <= 6 Then
'                      oWorkSheet.Cells(iFila + lnLineas - 1, 1).Value = rsDiagnosticos.Fields("dDiagnostico").Value
'                      oWorkSheet.Cells(iFila + lnLineas - 1, 46).Value = IIf(rsDiagnosticos.Fields("idSubClasificacionDx").Value = 101, "X", "") 'Presuntivo
'                      oWorkSheet.Cells(iFila + lnLineas - 1, 47).Value = IIf(rsDiagnosticos.Fields("idSubClasificacionDx").Value = 102, "X", "") 'Definitivo
'                      oWorkSheet.Cells(iFila + lnLineas - 1, 48).Value = IIf(rsDiagnosticos.Fields("idSubClasificacionDx").Value = 103, "X", "") 'Repetido
'                      'oWorkSheet.Cells(iFila + lnLineas - 1, 21).Value = Trim(rsDiagnosticos.Fields("labConfHIS").Value)
'                      oWorkSheet.Cells(iFila + lnLineas - 1, 50).Value = Trim(rsDiagnosticos.Fields("codigoCie10").Value)
'                End If
'                lnLineas = lnLineas + 1
'                rsDiagnosticos.MoveNext
'             Loop
'        End If
'        '******Medico
'        oWorkSheet.Cells(62, 1).Value = "'" & Trim(rsReporte!MedicoDni)
'        oWorkSheet.Cells(62, 12).Value = "'" & Trim(rsReporte!MedicoPaterno) & " " & Trim(rsReporte!medicomaterno) & " " & Trim(rsReporte!MedicoNombres)
'        oWorkSheet.Cells(62, 50).Value = "'" & Trim(rsReporte!medicoColegiatura)
'
'    End If
'
'    If oWorkSheet.PageSetup.PrintArea <> "" Then
'       oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 24)
'    End If
'
'    'oExcel.Visible = True
'    'oWorkSheet.PrintPreview
'    oWorkSheet.PrintOut
'
'    oWorkBook.Close SaveChanges:=False
'
'    Set oWorkSheet = Nothing
'    Set oExcel = Nothing
'    Set oConexion = Nothing
'    Set rsReporte = Nothing
'    Set rsDiagnosticos = Nothing
'    Set rsReferido = Nothing
'Exit Sub
'ManejadorError:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'    End Select
'    Exit Sub
    Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReporteUtil As New ReporteUtil
    Dim oConexion As New Connection
    Dim rsreporte As New Recordset
    Dim rsDiagnosticos As New Recordset
    Dim rsReferido As New Recordset
    Dim sSQL As String: Dim lccodreferido As String: Dim lcdesreferido As String
    Dim oRsLaboraEn As Recordset
    Dim lnTotalRegistro As Integer
    Dim lbEsOpenOffice As Boolean

    
    
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
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set rsreporte = mo_ReglasAdmision.AtencionesFiltraDatosParaFormatoSIS(lnIdAtencion, oConexion)
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEsis.ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/OpenOffice.ods"
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEsis.xls")
        oWorkBookPlantilla.Worksheets("sis").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    
        If lnCabeceraDetalle = 1 Or lnCabeceraDetalle = 3 Then
'            'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'            oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(15, 2).setFormula("'" & lcBuscaParametro.SeleccionaFilaParametro(239))
                Call Feuille.getcellbyposition(0, 5).setFormula("'" & rsreporte!ApellidoPaterno)
                Call Feuille.getcellbyposition(34, 6).setFormula("'" & rsreporte!ApellidoMaterno)
                Call Feuille.getcellbyposition(0, 8).setFormula("'" & rsreporte!PrimerNombre)
                Call Feuille.getcellbyposition(34, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsreporte!SegundoNombre))
            Else
                oWorkSheet.Cells(3, 16).Value = "'" & lcBuscaParametro.SeleccionaFilaParametro(239)
                oWorkSheet.Cells(6, 1).Value = "'" & rsreporte!ApellidoPaterno
                oWorkSheet.Cells(6, 35).Value = "'" & rsreporte!ApellidoMaterno
                oWorkSheet.Cells(9, 1).Value = "'" & rsreporte!PrimerNombre
                oWorkSheet.Cells(9, 35).Value = "'" & mo_ReporteUtil.NullToVacio(rsreporte!SegundoNombre)
            End If
            If rsreporte.Fields!idTipoSexo = 1 Then     'Masculino
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(15, 12).setFormula("X")
                Else
                    oWorkSheet.Cells(13, 16).Value = "X"
                End If
            End If
            If Not IsNull(rsreporte!FechaNacimiento) Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(0, 13).setFormula("'" & Day(rsreporte!FechaNacimiento) & "   " & Month(rsreporte!FechaNacimiento) & "     " & Year(rsreporte!FechaNacimiento))
                Else
                    oWorkSheet.Cells(14, 1).Value = "'" & Day(rsreporte!FechaNacimiento) & "   " & Month(rsreporte!FechaNacimiento) & "     " & Year(rsreporte!FechaNacimiento)
                End If
            End If
            If rsreporte.Fields!idTipoSexo <> 1 Then     'Femenino
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(15, 13).setFormula("X")
                Else
                    oWorkSheet.Cells(14, 16).Value = "X"
                End If
            End If
            lccodreferido = "": lcdesreferido = ""
            If Not IsNull(rsreporte.Fields!IdTipoReferenciaOrigen) Then
                If rsreporte.Fields!IdTipoReferenciaOrigen = 1 Then
                   'Referido MINSA
                   Set rsReferido = mo_ReglasComunes.EstablecimientosXidentificador(rsreporte.Fields!IdEstablecimientoOrigen, oConexion)
                   If rsReferido.RecordCount > 0 Then
                      lccodreferido = rsReferido.Fields!Codigo
                      lcdesreferido = rsReferido.Fields!nombre
                   End If
                ElseIf rsreporte.Fields!IdTipoReferenciaOrigen = 2 Then
                   'Referido Externo
                   Set rsReferido = mo_ReglasComunes.EstablecimientosNoMinsaXidentificador(rsreporte.Fields!IdEstablecimientoNoMinsaOrigen, oConexion)
                   If rsReferido.RecordCount > 0 Then
                      lccodreferido = ""
                      lcdesreferido = rsReferido.Fields!nombre
                   End If
                End If
            End If
        End If
        
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(0, 18).setFormula("'" & lccodreferido)
        Call Feuille.getcellbyposition(19, 18).setFormula("'" & lcdesreferido)
        Call Feuille.getcellbyposition(0, 23).setFormula("'" & Day(rsreporte!FechaIngreso) & "     " & Month(rsreporte!FechaIngreso))
        Call Feuille.getcellbyposition(29, 23).setFormula("'" & rsreporte!NroHistoriaClinica)
    Else
        oWorkSheet.Cells(19, 1).Value = "'" & lccodreferido
        oWorkSheet.Cells(19, 20).Value = "'" & lcdesreferido
        oWorkSheet.Cells(24, 1).Value = "'" & Day(rsreporte!FechaIngreso) & "     " & Month(rsreporte!FechaIngreso)
        oWorkSheet.Cells(24, 30).Value = "'" & rsreporte!NroHistoriaClinica
    End If
    
    If lnCabeceraDetalle = 2 Or lnCabeceraDetalle = 3 Then
        '******Triaje
        Dim oAtencionesCE As New AtencionesCE
        Dim oConexionJamo As New Connection
        Dim mo_DOAtencionesCE As New DOAtencionesCE
        oConexionJamo.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        mo_DOAtencionesCE.idAtencion = lnIdAtencion
        Set oAtencionesCE.Conexion = oConexionJamo
        If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = True Then
            If Not mo_DOAtencionesCE Is Nothing Then
                 With mo_DOAtencionesCE
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(14, 37).setFormula("'" & mo_DOAtencionesCE.triajePeso)
                        Call Feuille.getcellbyposition(22, 37).setFormula("'" & mo_DOAtencionesCE.triajeTalla)
                    Else
                        oWorkSheet.Cells(38, 15).Value = "'" & mo_DOAtencionesCE.triajePeso
                        oWorkSheet.Cells(38, 23).Value = "'" & mo_DOAtencionesCE.triajeTalla
                    End If
                 End With
            End If
        End If
        Set oConexionJamo = Nothing
        Set oAtencionesCE = Nothing
        Set mo_DOAtencionesCE = Nothing
        '******Dx
        Dim iFila As Integer, lnLineas As Integer
        Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
        iFila = 52
        lnLineas = 1
        Set rsDiagnosticos = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(lnIdAtencion)
        If rsDiagnosticos.RecordCount > 0 Then
             rsDiagnosticos.MoveFirst
             Do While Not rsDiagnosticos.EOF
                If lnLineas <= 6 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(0, iFila + lnLineas - 2).setFormula(rsDiagnosticos.Fields("dDiagnostico").Value)
                        Call Feuille.getcellbyposition(45, iFila + lnLineas - 2).setFormula(rsDiagnosticos.Fields("dDiagnostico").Value)
                        Call Feuille.getcellbyposition(46, iFila + lnLineas - 2).setFormula(rsDiagnosticos.Fields("dDiagnostico").Value)
                        Call Feuille.getcellbyposition(47, iFila + lnLineas - 2).setFormula(rsDiagnosticos.Fields("dDiagnostico").Value)
                        Call Feuille.getcellbyposition(49, iFila + lnLineas - 2).setFormula(rsDiagnosticos.Fields("dDiagnostico").Value)
                    Else
                        oWorkSheet.Cells(iFila + lnLineas - 1, 1).Value = rsDiagnosticos.Fields("dDiagnostico").Value
                        oWorkSheet.Cells(iFila + lnLineas - 1, 46).Value = IIf(rsDiagnosticos.Fields("idSubClasificacionDx").Value = 101, "X", "") 'Presuntivo
                        oWorkSheet.Cells(iFila + lnLineas - 1, 47).Value = IIf(rsDiagnosticos.Fields("idSubClasificacionDx").Value = 102, "X", "") 'Definitivo
                        oWorkSheet.Cells(iFila + lnLineas - 1, 48).Value = IIf(rsDiagnosticos.Fields("idSubClasificacionDx").Value = 103, "X", "") 'Repetido
                        'oWorkSheet.Cells(iFila + lnLineas - 1, 21).Value = Trim(rsDiagnosticos.Fields("labConfHIS").Value)
                        oWorkSheet.Cells(iFila + lnLineas - 1, 50).Value = Trim(rsDiagnosticos.Fields("codigoCie10").Value)
                    End If
                End If
                lnLineas = lnLineas + 1
                rsDiagnosticos.MoveNext
             Loop
        End If
        '******Medico
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(0, 61).setFormula("'" & Trim(rsreporte!MedicoDni))
            Call Feuille.getcellbyposition(11, 61).setFormula("'" & Trim(rsreporte!MedicoPaterno) & " " & Trim(rsreporte!medicomaterno) & " " & Trim(rsreporte!MedicoNombres))
            Call Feuille.getcellbyposition(49, 61).setFormula("'" & Trim(rsreporte!medicoColegiatura))
        Else
            oWorkSheet.Cells(62, 1).Value = "'" & Trim(rsreporte!MedicoDni)
            oWorkSheet.Cells(62, 12).Value = "'" & Trim(rsreporte!MedicoPaterno) & " " & Trim(rsreporte!medicomaterno) & " " & Trim(rsreporte!MedicoNombres)
            oWorkSheet.Cells(62, 50).Value = "'" & Trim(rsreporte!medicoColegiatura)
        End If
    End If
    If lbEsOpenOffice = True Then
    Else
        If oWorkSheet.PageSetup.PrintArea <> "" Then
           oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 24)
        End If
    End If

    If lbEsOpenOffice = True Then
        Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
    Else
        'oWorkSheet.PrintOut
        oExcel.Visible = True
        oWorkSheet.PrintPreview
    End If
    If lbEsOpenOffice = True Then
    Else
        oWorkBook.Close SaveChanges:=False
    End If
    
    If lbEsOpenOffice = True Then
    Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
    Else
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
        Set oConexion = Nothing
        Set rsreporte = Nothing
        Set rsDiagnosticos = Nothing
        Set rsReferido = Nothing
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



Sub CrearReporteCeAtencionPaciente(ml_lnHwnd As Long, ml_idAtencion As Long, lcPaciente As String, _
                                   lnIdCuentaAtencion As Long, lcConsultorio As String, _
                                   ldFechaCita As Date, lcMedico As String, _
                                   lcTriajePresion As String, lcTriajeTalla As String, _
                                   lcTriajeTemperatura As String, lcTriajePeso As String, _
                                   lcCitaMotivo As String, lcCitaExamenClinico As String, _
                                   lcDxMedico As String, lcCitaDiagMed As String, _
                                   lcCitaTratamiento As String, lcCitaExClinicos As String, _
                                   lcCitaObservaciones As String, lbReporteEsEnPDF As Boolean, lnIdPaciente As Long, _
                                   lbVistaPreviaPDF As Boolean, _
                                   Optional oRsVacunas As Recordset, _
                                   Optional oRsOtrosCpt As Recordset, Optional oRsDxDesarrollo As Recordset, _
                                   Optional oRsDxMorbilidad As Recordset, Optional oRsFarmacia As Recordset, _
                                   Optional lcCitaReceta As String = "", Optional lcOtrosCptDx As String)
Dim oAtencionesCE As New AtencionesCE
Dim oConexion As New Connection
Dim lnTotalRegistro As Integer
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
Dim lcTexto As String
Const lcSeparador As String = "¨"
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
        Dim lnHwnd As Long
        Dim PrintArgs(2)
        Dim PathFileOpenOffice As String
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    
    Dim wxParametro542_a  As String, wxParametro352_a As String
    wxParametro542_a = wxParametro542
    wxParametro352_a = lcBuscaParametro.SeleccionaFilaParametro(352)
    Dim lbSePuedeImprimirPDF As Boolean, lcArchivoPDF As String
    If lbReporteEsEnPDF = True Then
       lcArchivoPDF = sighEntidades.DevuelveRutaConSlashInvertida(lcBuscaParametro.SeleccionaFilaParametro(237)) & _
       Trim(Str(lnIdPaciente)) & "-" & Format(ldFechaCita, "DDMMYYYY") & _
       "-CTA" & Trim(Str(lnIdCuentaAtencion)) & _
       IIf(lbVistaPreviaPDF = False, "-GRABAATENCIONCE-" & Format(Now, "hhmmss") & ".PDF", "-IMPRESIONATENCIONCE.pdf")
       
       If SePuedeImprimirPDF(lcArchivoPDF, lbVistaPreviaPDF) = True Then
          lbSePuedeImprimirPDF = True
          wxParametro542_a = "S"
          wxParametro352_a = "S"
       End If
    End If
    
    
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEatencionPaciente.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEatencionPaciente.xls")
        oWorkBookPlantilla.Worksheets("CEatencionPaciente").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    If ml_idAtencion > 0 Then
       lcTriajePresion = lcTriajePresion & " Sistólica/Diastólica"
       lcTriajeTemperatura = lcTriajeTemperatura & " °C"
       lcTriajePeso = lcTriajePeso & " Kg."
       lcTriajeTalla = lcTriajeTalla & " cm."
    End If
    
    If lcBuscaParametro.SeleccionaFilaParametro(582) = "S" And lbSePuedeImprimirPDF = True Then
       oWorkSheet.PageSetup.LeftHeaderPicture.FileName = ""
    End If
    
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(2, 3).setFormula(lcMedico)
        Call Feuille.getcellbyposition(4, 3).setFormula(ldFechaCita)
        Call Feuille.getcellbyposition(2, 4).setFormula(lcPaciente)
        Call Feuille.getcellbyposition(2, 5).setFormula(lcConsultorio)
        Call Feuille.getcellbyposition(4, 5).setFormula("'" & lnIdCuentaAtencion)
        Call Feuille.getcellbyposition(2, 8).setFormula(lcTriajePresion)
        Call Feuille.getcellbyposition(4, 8).setFormula(lcTriajeTalla)
        Call Feuille.getcellbyposition(2, 9).setFormula(lcTriajeTemperatura)
        Call Feuille.getcellbyposition(4, 9).setFormula(lcTriajePeso)
        Call Feuille.getcellbyposition(2, 12).setFormula(lcCitaMotivo)
        Call Feuille.getcellbyposition(2, 14).setFormula(lcCitaExamenClinico)
        Call Feuille.getcellbyposition(2, 16).setFormula(lcDxMedico)
        Call Feuille.getcellbyposition(2, 18).setFormula(lcCitaDiagMed)
        Call Feuille.getcellbyposition(2, 20).setFormula(lcOtrosCptDx)
        'debb-09/09/2015 (inicio)
        Call Feuille.getcellbyposition(2, 22).setFormula("")
        Call Feuille.getcellbyposition(2, 24).setFormula("Ord.Mèdicas")
        Call Feuille.getcellbyposition(2, 26).setFormula("")
        'debb-09/09/2015 (fin)
        
        Call Feuille.getcellbyposition(2, 22).setFormula(lcCitaTratamiento)
        Call Feuille.getcellbyposition(2, 24).setFormula(lcCitaExClinicos)
        'mgaray201410d
        Call Feuille.getcellbyposition(2, 26).setFormula(lcCitaReceta)
        Call Feuille.getcellbyposition(2, 28).setFormula(lcCitaObservaciones)
        On Error Resume Next
        If Not oRsVacunas Is Nothing Then
            lcTexto = ""
            If oRsVacunas.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsVacunas.MoveFirst
               Do While Not oRsVacunas.EOF
                  lcTexto = lcTexto & Trim(oRsVacunas!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                  oRsVacunas.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 30).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsOtrosCpt Is Nothing) Then
            lcTexto = ""
            If oRsOtrosCpt.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsOtrosCpt.MoveFirst
               Do While Not oRsOtrosCpt.EOF
                  lcTexto = lcTexto & Trim(oRsOtrosCpt!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                  oRsOtrosCpt.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 32).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsDxDesarrollo Is Nothing) Then
            lcTexto = ""
            If oRsDxDesarrollo.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxDesarrollo.MoveFirst
               Do While Not oRsDxDesarrollo.EOF
                  lcTexto = lcTexto & Trim(oRsDxDesarrollo!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxDesarrollo.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 34).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsDxMorbilidad Is Nothing) Then
            lcTexto = ""
            If oRsDxMorbilidad.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxMorbilidad.MoveFirst
               Do While Not oRsDxMorbilidad.EOF
                  lcTexto = lcTexto & Trim(oRsDxMorbilidad!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxMorbilidad.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 36).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsFarmacia Is Nothing) Then
            lcTexto = ""
            If oRsFarmacia.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsFarmacia.MoveFirst
               Do While Not oRsFarmacia.EOF
                  If oRsFarmacia!seleccionar = True Then
                     lcTexto = lcTexto & Trim(oRsFarmacia!Medicamento) & lcSeparador & Chr(13) & Chr(10)
                  End If
                  oRsFarmacia.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 38).setFormula(lcTexto)
            End If
        End If
    Else
        oWorkSheet.Cells(4, 3).Value = lcMedico
        oWorkSheet.Cells(4, 5).Value = ldFechaCita
        oWorkSheet.Cells(5, 3).Value = lcPaciente
        oWorkSheet.Cells(6, 3).Value = lcConsultorio
        oWorkSheet.Cells(6, 5).Value = "'" & lnIdCuentaAtencion
        oWorkSheet.Cells(9, 3).Value = lcTriajePresion
        oWorkSheet.Cells(9, 5).Value = lcTriajeTalla
        oWorkSheet.Cells(10, 3).Value = lcTriajeTemperatura
        oWorkSheet.Cells(10, 5).Value = lcTriajePeso
        oWorkSheet.Cells(13, 3).Value = lcCitaMotivo
        oWorkSheet.Cells(15, 3).Value = lcCitaExamenClinico
        oWorkSheet.Cells(17, 3).Value = lcDxMedico
        oWorkSheet.Cells(19, 3).Value = lcCitaDiagMed
        oWorkSheet.Cells(21, 3).Value = lcOtrosCptDx
        'debb-09/09/2015 (inicio)
        oWorkSheet.Cells(23, 2).Value = ""
        oWorkSheet.Cells(25, 2).Value = "Ord.Mèdicas"
        oWorkSheet.Cells(27, 2).Value = ""
        'debb-09/09/2015 (fin)
        If lcBuscaParametro.SeleccionaFilaParametro(362) = "S" Then
             oWorkSheet.Cells(23, 2).Value = "Tratamiento"
             oWorkSheet.Cells(23, 3).Value = lcCitaTratamiento
        End If
        oWorkSheet.Cells(25, 3).Value = lcCitaExClinicos
        'mgaray201410d
        oWorkSheet.Cells(27, 3).Value = lcCitaReceta
        oWorkSheet.Cells(29, 3).Value = lcCitaObservaciones
        On Error Resume Next
        If Not (oRsVacunas Is Nothing) Then
            lcTexto = ""
            If oRsVacunas.RecordCount > 0 Then
               If Err.Number = 0 Then
                oRsVacunas.MoveFirst
                Do While Not oRsVacunas.EOF
                   lcTexto = lcTexto & Trim(oRsVacunas!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                   oRsVacunas.MoveNext
                Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(31, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsOtrosCpt Is Nothing) Then
            lcTexto = ""
            If oRsOtrosCpt.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsOtrosCpt.MoveFirst
               Do While Not oRsOtrosCpt.EOF
                  lcTexto = lcTexto & Trim(oRsOtrosCpt!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                  oRsOtrosCpt.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(33, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsDxDesarrollo Is Nothing) Then
            lcTexto = ""
            If oRsDxDesarrollo.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxDesarrollo.MoveFirst
               Do While Not oRsDxDesarrollo.EOF
                  lcTexto = lcTexto & Trim(oRsDxDesarrollo!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxDesarrollo.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(35, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsDxMorbilidad Is Nothing) Then
            lcTexto = ""
            If oRsDxMorbilidad.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxMorbilidad.MoveFirst
               Do While Not oRsDxMorbilidad.EOF
                  lcTexto = lcTexto & Trim(oRsDxMorbilidad!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxMorbilidad.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(37, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsFarmacia Is Nothing) Then
            lcTexto = ""
            If oRsFarmacia.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsFarmacia.MoveFirst
               Do While Not oRsFarmacia.EOF
                  If oRsFarmacia!seleccionar = True Then
                     lcTexto = lcTexto & Trim(oRsFarmacia!Medicamento) & lcSeparador & Chr(13) & Chr(10)
                  End If
                  oRsFarmacia.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(39, 3).Value = lcTexto
            End If
        End If
    
    End If
'    If lbEsOpenOffice = True Then
'        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
'        PrintArea(0).Sheet = 0
'        PrintArea(0).startcolumn = 0
'        PrintArea(0).StartRow = 0
'        PrintArea(0).EndColumn = 6
'        'mgaray201410d
'        PrintArea(0).EndRow = 39
'        Call Feuille.SetPrintAreas(PrintArea())
'        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
'        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
'    Else
'        oExcel.Visible = True
'        oWorkSheet.PrintPreview
'        'oWorkSheet.PrintOut
'        'oWorkBook.Close SaveChanges:=False
'    End If
    
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 6
        PrintArea(0).EndRow = 39
        Call Feuille.SetPrintAreas(PrintArea())
        If lcBuscaParametro.SeleccionaFilaParametro(352) <> "S" Then
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             CallByName Document, "print", VbMethod, PrintArgs()
             Sleep (1000)
             Call Document.Close(True)
             If Dir$(PathFileOpenOffice, vbArchive) <> "" Then
                 Kill PathFileOpenOffice
             End If
         End If
    Else
        If wxParametro352_a <> "S" Then
          oExcel.Visible = True
          oWorkSheet.PrintPreview
        Else
          oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 32)
          oWorkSheet.PrintOut
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
        If wxParametro542_a <> "S" Then
            oExcel.DisplayAlerts = False
            oExcel.Quit
        End If
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    
    SeteaOtraImpresoraDefault sighEntidades.ImpresoraDefaultDeEstaPC

    
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

Sub CrearReporteCeAtencionPacienteCRED(ml_lnHwnd As Long, ml_idAtencion As Long, lcPaciente As String, _
                                   lnIdCuentaAtencion As Long, lcConsultorio As String, _
                                   ldFechaCita As Date, lcMedico As String, _
                                   lcTriajePresion As String, lcTriajeTalla As String, _
                                   lcTriajeTemperatura As String, lcTriajePeso As String, _
                                   lcCitaMotivo As String, lcCitaExamenClinico As String, _
                                   lcDxMedico As String, lcCitaDiagMed As String, _
                                   lcCitaTratamiento As String, lcCitaExClinicos As String, _
                                   lcCitaObservaciones As String, Optional oRsVacunas As Recordset, _
                                   Optional oRsOtrosCpt As Recordset, Optional oRsDxDesarrollo As Recordset, _
                                   Optional oRsDxMorbilidad As Recordset, Optional oRsFarmacia As Recordset, _
                                   Optional lcCitaReceta As String = "", Optional lcOtrosCptDx As String)
Dim oAtencionesCE As New AtencionesCE
Dim oConexion As New Connection
Dim lnTotalRegistro As Integer
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
Dim lcTexto As String
Const lcSeparador As String = "¨"
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
        Dim lnHwnd As Long
        Dim PrintArgs(2)
        Dim PathFileOpenOffice As String
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEatencionPacienteCRED.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEatencionPacienteCRED.xls")
        oWorkBookPlantilla.Worksheets("CEatencionPaciente").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    If ml_idAtencion > 0 Then
       lcTriajePresion = lcTriajePresion & " Sistóloca/Diastólica"
       lcTriajeTemperatura = lcTriajeTemperatura & " °C"
       lcTriajePeso = lcTriajePeso & " Kg."
       lcTriajeTalla = lcTriajeTalla & " cm."
    End If
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(2, 3).setFormula(lcMedico)
        Call Feuille.getcellbyposition(4, 3).setFormula(ldFechaCita)
        Call Feuille.getcellbyposition(2, 4).setFormula(lcPaciente)
        Call Feuille.getcellbyposition(2, 5).setFormula(lcConsultorio)
        Call Feuille.getcellbyposition(4, 5).setFormula("'" & lnIdCuentaAtencion)
        Call Feuille.getcellbyposition(2, 8).setFormula(lcTriajePresion)
        Call Feuille.getcellbyposition(4, 8).setFormula(lcTriajeTalla)
        Call Feuille.getcellbyposition(2, 9).setFormula(lcTriajeTemperatura)
        Call Feuille.getcellbyposition(4, 9).setFormula(lcTriajePeso)
        Call Feuille.getcellbyposition(2, 12).setFormula(lcCitaMotivo)
        Call Feuille.getcellbyposition(2, 14).setFormula(lcCitaExamenClinico)
        Call Feuille.getcellbyposition(2, 16).setFormula(lcDxMedico)
        Call Feuille.getcellbyposition(2, 18).setFormula(lcCitaDiagMed)
        Call Feuille.getcellbyposition(2, 20).setFormula(lcOtrosCptDx)
        
        Call Feuille.getcellbyposition(2, 22).setFormula(lcCitaTratamiento)
        Call Feuille.getcellbyposition(2, 24).setFormula(lcCitaExClinicos)
        'mgaray201410d
        Call Feuille.getcellbyposition(2, 26).setFormula(lcCitaReceta)
        Call Feuille.getcellbyposition(2, 28).setFormula(lcCitaObservaciones)
        On Error Resume Next
        If Not oRsVacunas Is Nothing Then
            lcTexto = ""
            If oRsVacunas.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsVacunas.MoveFirst
               Do While Not oRsVacunas.EOF
                  lcTexto = lcTexto & Trim(oRsVacunas!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                  oRsVacunas.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 30).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsOtrosCpt Is Nothing) Then
            lcTexto = ""
            If oRsOtrosCpt.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsOtrosCpt.MoveFirst
               Do While Not oRsOtrosCpt.EOF
                  lcTexto = lcTexto & Trim(oRsOtrosCpt!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                  oRsOtrosCpt.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 32).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsDxDesarrollo Is Nothing) Then
            lcTexto = ""
            If oRsDxDesarrollo.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxDesarrollo.MoveFirst
               Do While Not oRsDxDesarrollo.EOF
                  lcTexto = lcTexto & Trim(oRsDxDesarrollo!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxDesarrollo.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 34).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsDxMorbilidad Is Nothing) Then
            lcTexto = ""
            If oRsDxMorbilidad.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxMorbilidad.MoveFirst
               Do While Not oRsDxMorbilidad.EOF
                  lcTexto = lcTexto & Trim(oRsDxMorbilidad!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxMorbilidad.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 36).setFormula(lcTexto)
            End If
        End If
        '
        If Not (oRsFarmacia Is Nothing) Then
            lcTexto = ""
            If oRsFarmacia.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsFarmacia.MoveFirst
               Do While Not oRsFarmacia.EOF
                  If oRsFarmacia!seleccionar = True Then
                     lcTexto = lcTexto & Trim(oRsFarmacia!Medicamento) & lcSeparador & Chr(13) & Chr(10)
                  End If
                  oRsFarmacia.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               Call Feuille.getcellbyposition(2, 38).setFormula(lcTexto)
            End If
        End If
    Else
        oWorkSheet.Cells(4, 3).Value = lcMedico
        oWorkSheet.Cells(4, 5).Value = ldFechaCita
        oWorkSheet.Cells(5, 3).Value = lcPaciente
        oWorkSheet.Cells(6, 3).Value = lcConsultorio
        oWorkSheet.Cells(6, 5).Value = "'" & lnIdCuentaAtencion
        oWorkSheet.Cells(9, 3).Value = lcTriajePresion
        oWorkSheet.Cells(9, 5).Value = lcTriajeTalla
        oWorkSheet.Cells(10, 3).Value = lcTriajeTemperatura
        oWorkSheet.Cells(10, 5).Value = lcTriajePeso
        oWorkSheet.Cells(13, 3).Value = lcCitaMotivo
        oWorkSheet.Cells(15, 3).Value = lcCitaExamenClinico
        oWorkSheet.Cells(17, 3).Value = lcDxMedico
        oWorkSheet.Cells(19, 3).Value = lcCitaDiagMed
        oWorkSheet.Cells(21, 3).Value = lcOtrosCptDx
        
        oWorkSheet.Cells(23, 3).Value = lcCitaTratamiento
        oWorkSheet.Cells(25, 3).Value = lcCitaExClinicos
        'mgaray201410d
        oWorkSheet.Cells(27, 3).Value = lcCitaReceta
        oWorkSheet.Cells(29, 3).Value = lcCitaObservaciones
        On Error Resume Next
        If Not (oRsVacunas Is Nothing) Then
            lcTexto = ""
            If oRsVacunas.RecordCount > 0 Then
               If Err.Number = 0 Then
                oRsVacunas.MoveFirst
                Do While Not oRsVacunas.EOF
                   lcTexto = lcTexto & Trim(oRsVacunas!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                   oRsVacunas.MoveNext
                Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(31, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsOtrosCpt Is Nothing) Then
            lcTexto = ""
            If oRsOtrosCpt.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsOtrosCpt.MoveFirst
               Do While Not oRsOtrosCpt.EOF
                  lcTexto = lcTexto & Trim(oRsOtrosCpt!procedimiento) & lcSeparador & Chr(13) & Chr(10)
                  oRsOtrosCpt.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(33, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsDxDesarrollo Is Nothing) Then
            lcTexto = ""
            If oRsDxDesarrollo.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxDesarrollo.MoveFirst
               Do While Not oRsDxDesarrollo.EOF
                  lcTexto = lcTexto & Trim(oRsDxDesarrollo!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxDesarrollo.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(35, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsDxMorbilidad Is Nothing) Then
            lcTexto = ""
            If oRsDxMorbilidad.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsDxMorbilidad.MoveFirst
               Do While Not oRsDxMorbilidad.EOF
                  lcTexto = lcTexto & Trim(oRsDxMorbilidad!DIAGNOSTICO) & lcSeparador & Chr(13) & Chr(10)
                  oRsDxMorbilidad.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(37, 3).Value = lcTexto
            End If
        End If
        '
        If Not (oRsFarmacia Is Nothing) Then
            lcTexto = ""
            If oRsFarmacia.RecordCount > 0 Then
               If Err.Number = 0 Then
               oRsFarmacia.MoveFirst
               Do While Not oRsFarmacia.EOF
                  If oRsFarmacia!seleccionar = True Then
                     lcTexto = lcTexto & Trim(oRsFarmacia!Medicamento) & lcSeparador & Chr(13) & Chr(10)
                  End If
                  oRsFarmacia.MoveNext
               Loop
               End If
            End If
            If lcTexto <> "" Then
                'mgaray201410d
               oWorkSheet.Cells(39, 3).Value = lcTexto
            End If
        End If
    
    End If
    
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 6
        PrintArea(0).EndRow = 39
        Call Feuille.SetPrintAreas(PrintArea())
        If lcBuscaParametro.SeleccionaFilaParametro(352) <> "S" Then
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             CallByName Document, "print", VbMethod, PrintArgs()
             Sleep (1000)
             Call Document.Close(True)
             If Dir$(PathFileOpenOffice, vbArchive) <> "" Then
                 Kill PathFileOpenOffice
             End If
         End If
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(352) <> "S" Then
          oExcel.Visible = True
          oWorkSheet.PrintPreview
        Else
          oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 32)
          oWorkSheet.PrintOut
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

Sub CrearReporteCeAtencionPacienteConPrograma(ml_idAtencion As Long, lcPaciente As String, _
                                   lnIdCuentaAtencion As Long, lcConsultorio As String, _
                                   ldFechaCita As Date, lcMedico As String, _
                                   lcTriajePresion As String, lcTriajeTalla As String, _
                                   lcTriajeTemperatura As String, lcTriajePeso As String, _
                                   lcCitaMotivo As String, lcCitaExamenClinico As String, _
                                   lcDxMedico As String, lcCitaDiagMed As String, _
                                   lcCitaTratamiento As String, lcCitaExClinicos As String, _
                                   lcCitaObservaciones As String, lbTabProgramaAbierto As Boolean, _
                                   lnIdProgramaActual As Long, lbControlActual As Boolean, _
                                   lcNumControl As String, lcFechaControl As String, _
                                   lcNombreProgramaAbierto As String, oRsProCabeceraDatos As Recordset, _
                                   oRsProControlDatos As Recordset, oRsProDiagnosticos As Recordset, _
                                   oRsProProcedimientos As Recordset, oRsProTratamientos As Recordset, _
                                   ml_lnHwnd As Long, _
                                   Optional lcCitaReceta As String = "", Optional lcOtrosCptDx As String)
Dim oAtencionesCE As New AtencionesCE
Dim oConexion As New Connection
'''''''''''FRANK '''''''''''''''''''''''''''''''
Dim oRsTmp1 As New Recordset
Dim oRsTmp2 As New Recordset
Dim lnPosColumna As Integer
Dim lnPosFila As Integer
Dim lnNumCabecera As Integer
Dim lcNombDatoCabecera As String
Dim lnNumControl As Integer
Dim lcNombDatoControl As String
Dim lcDIagPrograma As String
Dim lcProcPrograma As String
Dim lcTratPrograma As String
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
''''''''''''''''''''''''''''''''''''''
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
        Dim lnHwnd As Long
        Dim PrintArgs(2)
        Dim PathFileOpenOffice As String
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        If lbTabProgramaAbierto Then
            lcArchivoExcel = App.Path + "\Plantillas\CEAtencionPacientePrograma.ods"
        Else
            lcArchivoExcel = App.Path + "\Plantillas\CEatencionPaciente.ods"
        End If
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
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
        
        '''''Modificado por Frank 18032014 - selecciona excel
        If lbTabProgramaAbierto Then
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEAtencionPacientePrograma.xls")
        Else
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEatencionPaciente.xls")
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        oWorkBookPlantilla.Worksheets("CEatencionPaciente").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    If ml_idAtencion > 0 Then
       lcTriajePresion = lcTriajePresion & " Sistóloca/Diastólica"
       lcTriajeTemperatura = lcTriajeTemperatura & " °C"
       lcTriajePeso = lcTriajePeso & " Kg."
       lcTriajeTalla = lcTriajeTalla & " cm."
    End If
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(2, 3).setFormula(lcMedico)
        Call Feuille.getcellbyposition(4, 3).setFormula(ldFechaCita)
        Call Feuille.getcellbyposition(2, 4).setFormula(lcPaciente)
        Call Feuille.getcellbyposition(2, 5).setFormula(lcConsultorio)
        Call Feuille.getcellbyposition(4, 5).setFormula("'" & lnIdCuentaAtencion)
        Call Feuille.getcellbyposition(2, 8).setFormula(lcTriajePresion)
        Call Feuille.getcellbyposition(4, 8).setFormula(lcTriajeTalla)
        Call Feuille.getcellbyposition(2, 9).setFormula(lcTriajeTemperatura)
        Call Feuille.getcellbyposition(4, 9).setFormula(lcTriajePeso)
        Call Feuille.getcellbyposition(2, 12).setFormula(lcCitaMotivo)
        Call Feuille.getcellbyposition(2, 14).setFormula(lcCitaExamenClinico)
        Call Feuille.getcellbyposition(2, 16).setFormula(lcDxMedico)
        Call Feuille.getcellbyposition(2, 18).setFormula(lcCitaDiagMed)
        
        Call Feuille.getcellbyposition(2, 20).setFormula(lcOtrosCptDx)
        Call Feuille.getcellbyposition(2, 22).setFormula(lcCitaTratamiento)
        Call Feuille.getcellbyposition(2, 24).setFormula(lcCitaExClinicos)
        'mgaray201410d
        Call Feuille.getcellbyposition(2, 26).setFormula(lcCitaReceta)
        Call Feuille.getcellbyposition(2, 28).setFormula(lcCitaObservaciones)
    Else
        oWorkSheet.Cells(4, 3).Value = lcMedico
        oWorkSheet.Cells(4, 5).Value = ldFechaCita
        oWorkSheet.Cells(5, 3).Value = lcPaciente
        oWorkSheet.Cells(6, 3).Value = lcConsultorio
        oWorkSheet.Cells(6, 5).Value = "'" & lnIdCuentaAtencion
        oWorkSheet.Cells(9, 3).Value = lcTriajePresion
        oWorkSheet.Cells(9, 5).Value = lcTriajeTalla
        oWorkSheet.Cells(10, 3).Value = lcTriajeTemperatura
        oWorkSheet.Cells(10, 5).Value = lcTriajePeso
        oWorkSheet.Cells(13, 3).Value = lcCitaMotivo
        oWorkSheet.Cells(15, 3).Value = lcCitaExamenClinico
        oWorkSheet.Cells(17, 3).Value = lcDxMedico
        oWorkSheet.Cells(19, 3).Value = lcCitaDiagMed
        
        oWorkSheet.Cells(21, 3).Value = lcOtrosCptDx
        oWorkSheet.Cells(23, 3).Value = lcCitaTratamiento
        oWorkSheet.Cells(25, 3).Value = lcCitaExClinicos
        'mgaray201410d
        oWorkSheet.Cells(27, 3).Value = lcCitaReceta
        oWorkSheet.Cells(29, 3).Value = lcCitaObservaciones
    End If
    ''''' frank'''
    If lbTabProgramaAbierto Then
            oConexion.CommandTimeout = 300
            oConexion.Open sighEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            'mgaray201410d
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 30).setFormula(lcNombreProgramaAbierto)
            Else
                oWorkSheet.Cells(31, 3).Value = lcNombreProgramaAbierto
            End If
    '        lbControlActual lnIdProgramaActual
            Set oRsTmp1 = oRsProCabeceraDatos
            If oRsTmp1.RecordCount > 0 Then
                Set oRsTmp2 = mo_ReglasComunes.ProCabeceraConfigDatosSeleccionarPorIdPrograma(lnIdProgramaActual, oConexion)
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    lnNumCabecera = oRsTmp1.Fields!IdCabDato
                    lnPosColumna = IIf(Round(lnNumCabecera / 4) < lnNumCabecera / 4, Round(lnNumCabecera / 4) + 1, Round(lnNumCabecera / 4))
                    lnPosFila = IIf(lnNumCabecera Mod 4 = 0, 4, lnNumCabecera Mod 4)
                    lcNombDatoCabecera = ""
                    If oRsTmp2.RecordCount > 0 Then
                        oRsTmp2.MoveFirst
                        Do While Not oRsTmp2.EOF
                            If oRsTmp1.Fields!IdCabDato = oRsTmp2.Fields!IdCabDato Then
                                lcNombDatoCabecera = oRsTmp2.Fields!Cab_Texto
                                Exit Do
                            End If
                            oRsTmp2.MoveNext
                        Loop
                    End If
                    'mgaray201410d
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(2 * lnPosColumna - 1, 30 + lnPosFila).setFormula(lcNombDatoCabecera)
                        Call Feuille.getcellbyposition(1 + 2 * lnPosColumna - 1, 30 + lnPosFila).setFormula(Format(oRsTmp1.Fields!CabDato, sighEntidades.DevuelveFechaSoloFormato_DMY))
                    Else
                        oWorkSheet.Cells(31 + lnPosFila, 2 * lnPosColumna).Value = lcNombDatoCabecera
                        oWorkSheet.Cells(31 + lnPosFila, 1 + 2 * lnPosColumna).Value = oRsTmp1.Fields!CabDato
                    End If
                    oRsTmp1.MoveNext
                Loop
            End If
            'mgaray201410d
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, 37).setFormula(lcNumControl)
                Call Feuille.getcellbyposition(2, 37).setFormula(lcFechaControl)
            Else
                oWorkSheet.Cells(38, 2).Value = lcNumControl
                oWorkSheet.Cells(38, 3).Value = lcFechaControl
            End If
            If oRsTmp1.State = 1 Then oRsTmp1.Close
            Set oRsTmp1 = oRsProControlDatos
            If oRsTmp1.RecordCount > 0 Then
                Set oRsTmp2 = mo_ReglasComunes.ProControlConfigDatoSeleccionarPorIdPrograma(lnIdProgramaActual, oConexion)
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    lcNombDatoControl = ""
                    If oRsTmp2.RecordCount > 0 Then
                       oRsTmp2.MoveFirst
                       Do While Not oRsTmp2.EOF
                            If oRsTmp2.Fields!Control_EsDatoGrafico = False Then
                                If oRsTmp1.Fields!idcontroldato = oRsTmp2.Fields!idcontroldato Then
                                    lcNombDatoControl = oRsTmp2.Fields!control_texto
                                End If
                            End If
                           oRsTmp2.MoveNext
                       Loop
                    End If
                    If lcNombDatoControl <> "" Then
                        lnNumControl = oRsTmp1.Fields!idcontroldato
                        lnPosColumna = IIf(Round(lnNumControl / 4) < lnNumControl / 4, Round(lnNumControl / 4) + 1, Round(lnNumControl / 4))
                        lnPosFila = IIf(lnNumControl Mod 4 = 0, 4, lnNumControl Mod 4)
                        
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(2 * lnPosColumna - 1, 2 * (lnPosFila + 18) - 1).setFormula(lcNombDatoControl)
                            Call Feuille.getcellbyposition(1 + 2 * lnPosColumna - 1, 2 * (lnPosFila + 18) - 1).setFormula(Trim(Str(Val(oRsTmp1.Fields!controldato))))
                        Else
                            oWorkSheet.Cells(2 * (lnPosFila + 18), 2 * lnPosColumna).Value = lcNombDatoControl
                            oWorkSheet.Cells(2 * (lnPosFila + 18), 1 + 2 * lnPosColumna).Value = oRsTmp1.Fields!controldato
                        End If
                    End If
                    oRsTmp1.MoveNext
                Loop
             End If
             lcDIagPrograma = ""
             lcProcPrograma = ""
             lcTratPrograma = ""
             'Cargar Diagnosticos
             If oRsTmp1.State = 1 Then oRsTmp1.Close
             Set oRsTmp1 = oRsProDiagnosticos
             If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    lcDIagPrograma = lcDIagPrograma & ConsultaCodNombDiagnostico(oRsTmp1.Fields!idDiagnostico) & ", "
                    oRsTmp1.MoveNext
                Loop
                'mgaray201410d
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, 47).setFormula(lcDIagPrograma)
                Else
                    oWorkSheet.Cells(48, 3).Value = lcDIagPrograma
                End If
             End If
            'Cargar Procedimientos
            If oRsTmp1.State = 1 Then oRsTmp1.Close
            Set oRsTmp1 = oRsProProcedimientos
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    lcProcPrograma = lcProcPrograma & ConsultaCodNombProcedimiento(oRsTmp1.Fields!idProducto) & ", "
                    oRsTmp1.MoveNext
                Loop
                'mgaray201410d
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, 49).setFormula(lcProcPrograma)
                Else
                    oWorkSheet.Cells(50, 3).Value = lcProcPrograma
                End If
            End If
            'Cargar Tratamientos
            If oRsTmp1.State = 1 Then oRsTmp1.Close
            Set oRsTmp1 = oRsProTratamientos
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    lcTratPrograma = lcTratPrograma & ConsultaCodNombTratamiento(oRsTmp1.Fields!idProducto) & ", "
                    oRsTmp1.MoveNext
                Loop
                'mgaray201410d
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, 51).setFormula(lcTratPrograma)
                Else
                    oWorkSheet.Cells(52, 3).Value = lcTratPrograma
                End If
            End If

            oConexion.Close
            Set oConexion = Nothing
    End If
    '''''''''''''''
    If lbEsOpenOffice = True Then
        Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 6
        'mgaray201410d
        PrintArea(0).EndRow = 51
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
        'oWorkBook.Close SaveChanges:=False
    End If
    
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 6
        PrintArea(0).EndRow = 51
        Call Feuille.SetPrintAreas(PrintArea())
        If lcBuscaParametro.SeleccionaFilaParametro(352) <> "S" Then
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             CallByName Document, "print", VbMethod, PrintArgs()
             Sleep (1000)
             Call Document.Close(True)
             If Dir$(PathFileOpenOffice, vbArchive) <> "" Then
                 Kill PathFileOpenOffice
             End If
         End If
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(352) <> "S" Then
          oExcel.Visible = True
          oWorkSheet.PrintPreview
        Else
          oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 52)
          oWorkSheet.PrintOut
        End If
    End If

    '''Frank'''''''''''''
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    '''''''''''''''''''''
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
ManejadorError:

    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub

End Sub

'Frank
Function ConsultaCodNombDiagnostico(lnIdDiagnostico As Long) As String
    Dim oDODiagnostico As New DODiagnostico
    Set oDODiagnostico = mo_ReglasComunes.DiagnosticosSeleccionarPorId(lnIdDiagnostico)
    ConsultaCodNombDiagnostico = oDODiagnostico.CodigoCIE2004 & "-" & oDODiagnostico.descripcion
End Function

Function ConsultaCodNombProcedimiento(lnIdProcedimiento As Long) As String
    Dim oDOCatalogoServicio As New DOCatalogoServicio
    Set oDOCatalogoServicio = mo_ReglasComunes.CatalogoServiciosSeleccionarPorId(lnIdProcedimiento)
    ConsultaCodNombProcedimiento = oDOCatalogoServicio.nombre
    Set oDOCatalogoServicio = Nothing
End Function

Function ConsultaCodNombTratamiento(lnIdTratamiento As Long) As String
    Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
    Set oDOCatalogoBienesInsumos = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarPorId(lnIdTratamiento)
    ConsultaCodNombTratamiento = oDOCatalogoBienesInsumos.nombre
    Set oDOCatalogoBienesInsumos = Nothing
End Function
'''''''''''''''


'debb-13/05/2016
Sub ImprimeEnFormatoDeFiliacionParaHistoriaClinica(lnIdPaciente As Long, lnEdad As Integer, lnTipoEdad As Integer, ml_lnHwnd As Long)
    Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReporteUtil As New ReporteUtil
    Dim oConexion As New Connection
    Dim rsreporte As New Recordset
    Dim rsDiagnosticos As New Recordset
    Dim rsReferido As New Recordset
    Dim sSQL As String: Dim lccodreferido As String: Dim lcdesreferido As String
    Dim oRsLaboraEn As Recordset
    Dim lbEsOpenOffice As Boolean
    Dim lcSql As String
    Dim lbFichaFam As Boolean
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    lbFichaFam = IIf(lcBuscaParametro.SeleccionaFilaParametro(277) = "S", True, False)
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
    Set rsreporte = mo_ReglasAdmision.PacientesDatosParaImpresionFiliacion(lnIdPaciente, oConexion)
    If rsreporte.RecordCount = 0 Then
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
            Dim lnHwnd As Long
            '
            Dim PrintArgs(2)
            Dim PathFileOpenOffice As String
        Else
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
        
        Set oConexion = Nothing
        Set rsreporte = Nothing
        Set rsDiagnosticos = Nothing
        Set rsReferido = Nothing
    
        Exit Sub
    End If
    Set rsDiagnosticos = mo_ReglasArchivoClinico.HistoriasClinicasXIdPaciente(lnIdPaciente, oConexion)
    Dim NOMBREEESS As String
    Dim DIRECCIONEESS As String
    Dim TELEFONO As String

    NOMBREEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    DIRECCIONEESS = lcBuscaParametro.SeleccionaFilaParametro(206)
    TELEFONO = lcBuscaParametro.SeleccionaFilaParametro(207)
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\CEfiliacion.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        PathFileOpenOffice = App.Path + "\Plantillas\" & lcArchivoExcel
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        'Encabezado de Pagina
'        Set StyleFamilies = Document.StyleFamilies
'        Set PageStyles = StyleFamilies.getByName("PageStyles")
'        Set DefPage = PageStyles.getByName("Default")
'        Set Hcontent = DefPage.RightPageHeaderContent
'        Set Htext = Hcontent.centerText
        mo_CabeceraReportes.CabeceraReportes Document, True
'        DefPage.RightPageHeaderContent = Hcontent
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEfiliacion.xls")
        oWorkBookPlantilla.Worksheets("sis").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        
        
        
        Dim lcRutaImg As String
        lcRutaImg = lcBuscaParametro.SeleccionaFilaParametro(237) & Trim(Str(rsreporte!NroHistoriaClinica)) & ".jpg"
        If sighEntidades.ArchivoExiste(lcRutaImg) Then
           'oWorkSheet.PageSetup.LeftFooterPicture.FileName = lcRutaImg
           mo_CabeceraReportes.CabeceraReportes oWorkSheet, False, , lcRutaImg
        Else
           mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
    End If
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(5, 2).setFormula("HOJA DE FILIACION")
    Else
        oWorkSheet.Cells(3, 6).Value = "HOJA DE FILIACION"
    End If
    '
'    If rsDiagnosticos.RecordCount > 0 Then
'        If lbEsOpenOffice = True Then
'            Call Feuille.getcellbyposition(1, 3).setFormula("F.Creación: " & rsDiagnosticos!fechacreacion)
'        Else
'           oWorkSheet.Cells(4, 2).Value = "F.Creación: " & rsDiagnosticos!fechacreacion
'        End If
'    Else
'        If lbEsOpenOffice = True Then
'            Call Feuille.getcellbyposition(1, 3).setFormula("Fecha: " & Date)
'        Else
'           oWorkSheet.Cells(4, 2).Value = "Fecha: " & Date
'        End If
'    End If
    If lbEsOpenOffice = True Then
       'Call Feuille.getcellbyposition(14, 2).setFormula(Trim(Str(rsReporte.Fields!NroHistoriaClinica)))
        
        If lbFichaFam = True Then
            If Trim(rsreporte.Fields!fichafam) <> "" Or Not IsNull(rsreporte.Fields!fichafam) Then
                Call Feuille.getcellbyposition(14, 1).setFormula("N° Ficha Familiar")
                Call Feuille.getcellbyposition(14, 2).setFormula("'" & rsreporte.Fields!fichafam)
            Else
                Call Feuille.getcellbyposition(14, 1).setFormula("N° de Historia Clinica")
                Call Feuille.getcellbyposition(14, 2).setFormula(Trim(Str(rsreporte.Fields!NroHistoriaClinica)))
            End If
        Else
           Call Feuille.getcellbyposition(14, 1).setFormula("N° de Historia Clinica")
           Call Feuille.getcellbyposition(14, 2).setFormula(Trim(Str(rsreporte.Fields!NroHistoriaClinica)))
        End If
        
        
        
        Call Feuille.getcellbyposition(4, 5).setFormula("'" & rsreporte.Fields!ApellidoPaterno)
        Call Feuille.getcellbyposition(8, 5).setFormula("'" & rsreporte.Fields!ApellidoMaterno)
        Call Feuille.getcellbyposition(12, 5).setFormula("'" & rsreporte.Fields!PrimerNombre & " " & IIf(IsNull(rsreporte.Fields!SegundoNombre), "", rsreporte.Fields!SegundoNombre))
        
        Call Feuille.getcellbyposition(4, 7).setFormula("'" & rsreporte.Fields!FechaNacimiento)
        Call Feuille.getcellbyposition(4, 8).setFormula("'" & descripcionFechaNaciemiento(rsreporte.Fields!FNacimientoCalculada))
        Call Feuille.getcellbyposition(12, 21).setFormula(IIf(IsNull(rsreporte.Fields!NacDistrito), "", Trim(rsreporte.Fields!NacDistrito)))
        Call Feuille.getcellbyposition(8, 21).setFormula(IIf(IsNull(rsreporte.Fields!NacProvincia), "", Trim(rsreporte.Fields!NacProvincia)))
        Call Feuille.getcellbyposition(4, 21).setFormula(IIf(IsNull(rsreporte.Fields!NacDpto), "", Trim(rsreporte.Fields!NacDpto)))

    Else
        'carga Imagen..........si demora mucho al cargar, cambiar en parametros la ruta

        If lbFichaFam = True Then
            If Trim(rsreporte.Fields!fichafam) <> "" Or Not IsNull(rsreporte.Fields!fichafam) Then
                oWorkSheet.Cells(2, 15).Value = "N° Ficha Familiar"
                oWorkSheet.Cells(3, 15).Value = "'" & rsreporte.Fields!fichafam
            Else
                oWorkSheet.Cells(2, 15).Value = "N° de Historia Clinica"
                'oWorkSheet.Cells(3, 15).Value = Trim(Str(rsreporte.Fields!NroHistoriaClinica))
                oWorkSheet.Cells(3, 15).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte!NroHistoriaClinica)), True)
            End If
        Else
            oWorkSheet.Cells(2, 15).Value = "N° de Historia Clinica"
            'oWorkSheet.Cells(3, 15).Value = Trim(Str(rsreporte.Fields!NroHistoriaClinica))
            oWorkSheet.Cells(3, 15).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte!NroHistoriaClinica)), True)
        End If
        oWorkSheet.Cells(6, 5).Value = "'" & rsreporte.Fields!ApellidoPaterno
        oWorkSheet.Cells(6, 9).Value = "'" & rsreporte.Fields!ApellidoMaterno
        oWorkSheet.Cells(6, 13).Value = "'" & rsreporte.Fields!PrimerNombre & " " & IIf(IsNull(rsreporte.Fields!SegundoNombre), "", rsreporte.Fields!SegundoNombre)

        oWorkSheet.Cells(8, 5).Value = "'" & rsreporte.Fields!FechaNacimiento
        oWorkSheet.Cells(9, 5).Value = "'" & descripcionFechaNaciemiento(rsreporte.Fields!FNacimientoCalculada)
        oWorkSheet.Cells(22, 13).Value = IIf(IsNull(rsreporte.Fields!NacDistrito), "", Trim(rsreporte.Fields!NacDistrito))
        oWorkSheet.Cells(22, 9).Value = IIf(IsNull(rsreporte.Fields!NacProvincia), "", Trim(rsreporte.Fields!NacProvincia))
        oWorkSheet.Cells(22, 5).Value = IIf(IsNull(rsreporte.Fields!NacDpto), "", Trim(rsreporte.Fields!NacDpto))

    
    End If
    If rsreporte.Fields!idTipoSexo = 1 Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(7, 7).setFormula("Masculino")
        Else
            oWorkSheet.Cells(8, 8).Value = "Masculino"
        End If
    Else
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(7, 7).setFormula("Femenino")
        Else
            oWorkSheet.Cells(8, 8).Value = "Femenino"
        End If
    End If
    If lbEsOpenOffice = True Then
'        Call Feuille.getcellbyposition(42, 11).setFormula("'" & Trim(Str(lnEdad)) & " (en inscripción HC)")
'        Call Feuille.getcellbyposition(2, 12).setFormula("________________________________________________________")
'        Call Feuille.getcellbyposition(29, 12).setFormula("_____________")
'        Call Feuille.getcellbyposition(42, 12).setFormula("_____________")
'        Call Feuille.getcellbyposition(2, 13).setFormula("Fecha Nacimiento   Distrito/Provincia/Departamento")
'        Call Feuille.getcellbyposition(29, 13).setFormula("Sexo")
'    Else
'        oWorkSheet.Cells(12, 43).Value = "'" & Trim(Str(lnEdad)) & " (en inscripción HC)"
'        oWorkSheet.Cells(13, 3).Value = "________________________________________________________"
'        oWorkSheet.Cells(13, 30).Value = "_____________"
'        oWorkSheet.Cells(13, 43).Value = "_____________"
'        oWorkSheet.Cells(14, 3).Value = "Fecha Nacimiento   Distrito/Provincia/Departamento"
'        oWorkSheet.Cells(14, 30).Value = "Sexo"
    End If
'    If lnTipoEdad = 1 Then
'        If lbEsOpenOffice = True Then
'            Call Feuille.getcellbyposition(42, 13).setFormula("Edad (años)")
'        Else
'            oWorkSheet.Cells(14, 43).Value = "Edad (años)"
'        End If
'    ElseIf lnTipoEdad = 2 Then
'        If lbEsOpenOffice = True Then
'            Call Feuille.getcellbyposition(42, 13).setFormula("Edad (meses)")
'        Else
'            oWorkSheet.Cells(14, 43).Value = "Edad (meses)"
'        End If
'    Else
'        If lbEsOpenOffice = True Then
'            Call Feuille.getcellbyposition(42, 13).setFormula("Edad (días)")
'        Else
'            oWorkSheet.Cells(14, 43).Value = "Edad (días)"
'        End If
'    End If
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(9, 9).setFormula("'" & IIf(IsNull(rsreporte.Fields!EstadoCivil), "", rsreporte.Fields!EstadoCivil))
        Call Feuille.getcellbyposition(9, 11).setFormula("'" & IIf(IsNull(rsreporte.Fields!TELEFONO), "", rsreporte.Fields!TELEFONO))
        Call Feuille.getcellbyposition(2, 11).setFormula("'" & IIf(IsNull(rsreporte.Fields!Email), "", rsreporte.Fields!Email))
        Call Feuille.getcellbyposition(14, 7).setFormula("'" & IIf(IsNull(rsreporte.Fields!nrodocumento), "", rsreporte.Fields!nrodocumento))
        Call Feuille.getcellbyposition(9, 7).setFormula("'" & IIf(IsNull(rsreporte.Fields!TiposDocIdentidad), "", rsreporte.Fields!TiposDocIdentidad))
        
        Call Feuille.getcellbyposition(4, 18).setFormula(IIf(IsNull(rsreporte.Fields!DireccionDomicilio), "", Trim(rsreporte.Fields!DireccionDomicilio)))
        Call Feuille.getcellbyposition(12, 18).setFormula(IIf(IsNull(rsreporte.Fields!CentroPobladoDomicilio), "", Trim(rsreporte.Fields!CentroPobladoDomicilio)))
        Call Feuille.getcellbyposition(12, 16).setFormula(IIf(IsNull(rsreporte.Fields!DomicilioDistrito), "", rsreporte.Fields!DomicilioDistrito))
        Call Feuille.getcellbyposition(7, 16).setFormula(IIf(IsNull(rsreporte!domProvincia), "", Trim(rsreporte!domProvincia)))
        Call Feuille.getcellbyposition(4, 16).setFormula(IIf(IsNull(rsreporte!domDpto), "", Trim(rsreporte!domDpto)))
        

        'Call Feuille.getcellbyposition(34, 20).setFormula(IIf(IsNull(rsReporte!FactorRh), "", rsReporte!FactorRh))
        Call Feuille.getcellbyposition(14, 11).setFormula(IIf(IsNull(rsreporte!GrupoSanguineo), "", rsreporte!GrupoSanguineo) & " " & IIf(IsNull(rsreporte!FactorRh), "", rsreporte!FactorRh))

        Call Feuille.getcellbyposition(9, 29).setFormula("'" & IIf(IsNull(rsreporte.Fields!NombrePadre), "", rsreporte.Fields!NombrePadre))
        '**** Programa: se puso el DNI de la madre en open oficce
        '**** Programado por:Eder Yamill Palomino Espinoza
        '**** Fecha: 06102014
        Call Feuille.getcellbyposition(13, 32).setFormula("'" & IIf(IsNull(rsreporte.Fields!madreDocumento), "", rsreporte.Fields!madreDocumento))
        Call Feuille.getcellbyposition(4, 32).setFormula("'" & IIf(IsNull(rsreporte.Fields!ApellidosMadre), "", rsreporte.Fields!ApellidosMadre))
        Call Feuille.getcellbyposition(9, 32).setFormula("'" & IIf(IsNull(rsreporte.Fields!Nombremadre), "", rsreporte.Fields!Nombremadre)) 'Actualizado 23092014
        Call Feuille.getcellbyposition(12, 9).setFormula("'" & IIf(IsNull(rsreporte.Fields!etnia), "", Trim(rsreporte.Fields!etnia))) 'Actualizado 230930214
        Call Feuille.getcellbyposition(15, 9).setFormula("'" & IIf(IsNull(rsreporte.Fields!dIdioma), "", Trim(rsreporte.Fields!dIdioma))) 'Actualizado 230930214
        Call Feuille.getcellbyposition(4, 9).setFormula("'" & IIf(IsNull(rsreporte.Fields!dGradoInst), "", rsreporte.Fields!dGradoInst))
        
        Call Feuille.getcellbyposition(12, 24).setFormula(IIf(IsNull(rsreporte.Fields!ProcedDistrito), "", rsreporte.Fields!ProcedDistrito))
        Call Feuille.getcellbyposition(8, 24).setFormula(IIf(IsNull(rsreporte!ProcedProvincia), "", Trim(rsreporte!ProcedProvincia)))
        Call Feuille.getcellbyposition(4, 24).setFormula(IIf(IsNull(rsreporte!ProcedDpto), "", Trim(rsreporte!ProcedDpto)))
        Call Feuille.getcellbyposition(1, 37).setFormula("'" & IIf(IsNull(rsreporte.Fields!Observacion), "", rsreporte.Fields!Observacion))
        Call Feuille.getcellbyposition(7, 43).setFormula(mo_ReglasCaja.SeleccionaDatosCajero(sighEntidades.Usuario, sghUsuario))
    Else
        oWorkSheet.Cells(10, 10).Value = "'" & IIf(IsNull(rsreporte.Fields!EstadoCivil), "", rsreporte.Fields!EstadoCivil)
        oWorkSheet.Cells(12, 10).Value = "'" & IIf(IsNull(rsreporte.Fields!TELEFONO), "", rsreporte.Fields!TELEFONO)
        oWorkSheet.Cells(12, 3).Value = "'" & IIf(IsNull(rsreporte.Fields!Email), "", rsreporte.Fields!Email)
        oWorkSheet.Cells(8, 15).Value = "'" & IIf(IsNull(rsreporte.Fields!nrodocumento), "", rsreporte.Fields!nrodocumento)
        oWorkSheet.Cells(8, 10).Value = "'" & IIf(IsNull(rsreporte.Fields!TiposDocIdentidad), "", rsreporte.Fields!TiposDocIdentidad)

        oWorkSheet.Cells(19, 5).Value = IIf(IsNull(rsreporte.Fields!DireccionDomicilio), "", Trim(rsreporte.Fields!DireccionDomicilio))
        oWorkSheet.Cells(19, 13).Value = IIf(IsNull(rsreporte.Fields!CentroPobladoDomicilio), "", Trim(rsreporte.Fields!CentroPobladoDomicilio))
        oWorkSheet.Cells(17, 13).Value = IIf(IsNull(rsreporte.Fields!DomicilioDistrito), "", rsreporte.Fields!DomicilioDistrito)
        oWorkSheet.Cells(17, 8).Value = IIf(IsNull(rsreporte!domProvincia), "", Trim(rsreporte!domProvincia))
        oWorkSheet.Cells(17, 5).Value = IIf(IsNull(rsreporte!domDpto), "", Trim(rsreporte!domDpto))
        
        'oWorkSheet.Cells(21, 35).Value = IIf(IsNull(rsReporte!FactorRh), "", rsReporte!FactorRh)
        oWorkSheet.Cells(12, 15).Value = IIf(IsNull(rsreporte!GrupoSanguineo), "", rsreporte!GrupoSanguineo) & " " & IIf(IsNull(rsreporte!FactorRh), "", rsreporte!FactorRh)
        '
        oWorkSheet.Cells(30, 10).Value = "'" & IIf(IsNull(rsreporte.Fields!NombrePadre), "", rsreporte.Fields!NombrePadre)
        oWorkSheet.Cells(33, 14).Value = "'" & IIf(IsNull(rsreporte.Fields!madreDocumento), "", rsreporte.Fields!madreDocumento)
        oWorkSheet.Cells(33, 5).Value = "'" & IIf(IsNull(rsreporte.Fields!ApellidosMadre), "", rsreporte.Fields!ApellidosMadre)
        oWorkSheet.Cells(33, 10).Value = "'" & IIf(IsNull(rsreporte.Fields!Nombremadre), "", rsreporte.Fields!Nombremadre) 'Actualizado 230930214
        oWorkSheet.Cells(10, 13).Value = "'" & IIf(IsNull(rsreporte.Fields!etnia), "", Trim(rsreporte.Fields!etnia)) 'Actualizado 230930214
        oWorkSheet.Cells(10, 16).Value = "'" & IIf(IsNull(rsreporte.Fields!dIdioma), "", Trim(rsreporte.Fields!dIdioma)) 'Actualizado 230930214
        oWorkSheet.Cells(10, 5).Value = "'" & IIf(IsNull(rsreporte.Fields!dGradoInst), "", rsreporte.Fields!dGradoInst)

        oWorkSheet.Cells(25, 13).Value = IIf(IsNull(rsreporte.Fields!ProcedDistrito), "", rsreporte.Fields!ProcedDistrito)
        oWorkSheet.Cells(25, 9).Value = IIf(IsNull(rsreporte!ProcedProvincia), "", Trim(rsreporte!ProcedProvincia))
        oWorkSheet.Cells(25, 5).Value = IIf(IsNull(rsreporte!ProcedDpto), "", Trim(rsreporte!ProcedDpto))
        oWorkSheet.Cells(38, 2).Value = "'" & IIf(IsNull(rsreporte.Fields!Observacion), "", rsreporte.Fields!Observacion)
        oWorkSheet.Cells(44, 8).Value = mo_ReglasCaja.SeleccionaDatosCajero(sighEntidades.Usuario, sghUsuario)
    End If
    
    
    'If oWorkSheet.PageSetup.PrintArea <> "" Then
    '   oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 33)
    'End If
    

    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 16
        PrintArea(0).EndRow = 40
        Call Feuille.SetPrintAreas(PrintArea())
        If lcBuscaParametro.SeleccionaFilaParametro(341) <> "S" Then
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
             Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
             CallByName Document, "print", VbMethod, PrintArgs()
             Sleep (1000)
             Call Document.Close(True)
             If Dir$(PathFileOpenOffice, vbArchive) <> "" Then
                 Kill PathFileOpenOffice
             End If
         End If
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(341) <> "S" Then
          oExcel.Visible = True
          oWorkSheet.PrintPreview
        Else
          oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, 32)
          oWorkSheet.PrintOut 'lcBuscaParametro.SeleccionaFilaParametro(572)
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
        oWorkBook.Close False
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    
    Set oConexion = Nothing
    Set rsreporte = Nothing
    Set rsDiagnosticos = Nothing
    Set rsReferido = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_ReglasArchivoClinico = Nothing
    Set mo_ReglasCaja = Nothing
    Set mo_ReporteUtil = Nothing
    
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        'Resume
    End Select
    
    Exit Sub
Resume
End Sub

Public Function descripcionFechaNaciemiento(lbfNacionientoCalculada As Boolean) As String
    Dim oCadena As New cadena
    Dim lcDescripcion As String
    If lbfNacionientoCalculada = True Then
        lcDescripcion = oCadena.RetornaDescFechaNacimientoCalculada(lbfNacionientoCalculada, False)
    Else
        lcDescripcion = "D/M/A"
    End If
    descripcionFechaNaciemiento = lcDescripcion
End Function

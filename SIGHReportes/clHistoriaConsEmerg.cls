VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clHistoriaConsEmerg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historia Emergencia
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idCuentaAtencion As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_DOAtencionesCE As New DOAtencionesCE
Dim ml_Plan As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Property Let Plan(lValue As String)
    ml_Plan = lValue
End Property
Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property
Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property
Sub CrearReporteHistoriaClinicaConsultorioEmergCab(lnHwnd As Long, lcTipoEdad As String, lcGravedad As String, _
                                                   lcEmergenciaCorrelativo As String, lnFilaInicio As Long, _
                                                   lnColumnaInicio As Long, lcReferencia As String, _
                                                   lcDNIacompaniante As String, lbFormato As Integer, _
                                                   lcAfiliacionSIS As String)
Dim rsReporte As New Recordset
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim iFila As Long, iColumna As Long
Dim oCadena As New Cadena

iFila = lnFilaInicio            '1 en parametro 7 valortexto='1-1'
iColumna = lnColumnaInicio      '1 en parametro 7 valortexto='1-1'

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
        'mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
    Else
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        'Crea nueva hoja
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        If lbFormato = 1 Then
           Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Eadmision.xls")
        Else
           If iColumna = 0 Then
              Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Eadmision2_hp.xls")
           Else
              Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Eadmision2_epson.xls")
           End If
        End If
        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        'mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    'Izquierdo
    CargaAtencionCEJamo
    Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    If lbEsOpenOffice = True Then
    Else
       If lbFormato = 1 Then
            With oWorkSheet.PageSetup
               ' .LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
               ' .CenterHeader = "&""Arial Narrow,Normal""&8" & NOMBREEESS & Chr(13) & DIRECCIONEESS & Chr(13) & "TELEFONO:  " & TELEFONO
               ' If lbNoMuestraFechaHoraImpresion = False Then
                   .RightHeader = Chr(13) & Chr(13) & lcEmergenciaCorrelativo & Chr(13)
              '  End If
            End With
            oWorkSheet.Cells(iFila, iColumna + 17).Value = "' N° Cuenta: " & Trim(Str(ml_idCuentaAtencion))
            iFila = iFila + 1
            If rsReporte!IdTipoNumeracion <= 3 Then
               oWorkSheet.Cells(iFila, iColumna + 2).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True)
            End If
            oWorkSheet.Cells(iFila, iColumna + 7).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
            oWorkSheet.Cells(iFila, iColumna + 11).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
            iFila = iFila + 3
            oWorkSheet.Cells(iFila, iColumna + 6).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
            iFila = iFila + 1
            'oWorkSheet.Cells(iFila, iColumna + 1).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio) & " / " & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
            oWorkSheet.Cells(iFila, iColumna + 16).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento)
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, iColumna + 1).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
            oWorkSheet.Cells(iFila, iColumna + 7).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
            If Not IsNull(rsReporte!DireccionDomicilio) Then
               oWorkSheet.Cells(iFila, iColumna + 15).Value = "'" & mo_ReporteUtil.NullToVacio(Left(rsReporte!DireccionDomicilio, 40)) & " / " & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
            End If
            iFila = iFila + 3
            oWorkSheet.Cells(iFila, iColumna + 6).Value = "'" & lcReferencia
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, iColumna + 9).Value = "'" & mo_ReporteUtil.NullToVacio(UCase(rsReporte!NombreAcompaniante)) & "    DNI: " & lcDNIacompaniante
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, iColumna + 9).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!telefono)
       Else
            With oWorkSheet.PageSetup
               ' .LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
               ' .CenterHeader = "&""Arial Narrow,Normal""&8" & NOMBREEESS & Chr(13) & DIRECCIONEESS & Chr(13) & "TELEFONO:  " & TELEFONO
               ' If lbNoMuestraFechaHoraImpresion = False Then
                   .RightHeader = Chr(13) & Chr(13) & Space(8) & lcEmergenciaCorrelativo & Chr(13) & Chr(13) & "N°Cta: " & Trim(Str(ml_idCuentaAtencion))
              '  End If
            End With
            If iColumna = 0 Then
               '****** impresora HP
                iFila = iFila + 1
                If rsReporte!IdTipoNumeracion <= 3 Then
                   oWorkSheet.Cells(iFila, iColumna + 2).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True)
                End If
                oWorkSheet.Cells(iFila, iColumna + 7).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
                oWorkSheet.Cells(iFila, iColumna + 11).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
                iFila = iFila + 3
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
                oWorkSheet.Cells(iFila, iColumna + 18).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento)
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 1).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
                oWorkSheet.Cells(iFila, iColumna + 7).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
                If Not IsNull(rsReporte!DireccionDomicilio) Then
                   oWorkSheet.Cells(iFila, iColumna + 14).Value = "'" & mo_ReporteUtil.NullToVacio(Left(Trim(rsReporte!DireccionDomicilio), 55))
                End If
                iFila = iFila + 2
                oWorkSheet.Cells(iFila, iColumna + IIf(Len(lcReferencia) > 3, 1, 2)).Value = "X"
                oWorkSheet.Cells(iFila, iColumna + IIf(UCase(Left(ml_Plan, 3)) <> "SIS", 6, 9)).Value = "X"
                oWorkSheet.Cells(iFila, iColumna + 12).Value = lcAfiliacionSIS
                oWorkSheet.Cells(iFila, iColumna + 20).Value = IIf(UCase(Left(ml_Plan, 3)) <> "SIS", Left(ml_Plan, 8), "")
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & lcReferencia
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 8).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante) & "    DNI: " & lcDNIacompaniante
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 7).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!telefono)
           ElseIf iColumna = 1 Then
                '**** impresora EPSON *****
                iFila = iFila + 1
                If rsReporte!IdTipoNumeracion <= 3 Then
                   oWorkSheet.Cells(iFila, iColumna + 2).Value = "'" & Trim(Str(rsReporte!NroHistoriaClinica))
                End If
                oWorkSheet.Cells(iFila, iColumna + 7).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
                oWorkSheet.Cells(iFila, iColumna + 11).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
                iFila = iFila + 3
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
                oWorkSheet.Cells(iFila, iColumna + 18).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento)
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 1).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
                If Not IsNull(rsReporte!DireccionDomicilio) Then
                   oWorkSheet.Cells(iFila, iColumna + 10).Value = "'" & mo_ReporteUtil.NullToVacio(Left(Trim(rsReporte!DireccionDomicilio), 55))
                End If
                iFila = iFila + 2
                oWorkSheet.Cells(iFila, iColumna + IIf(Len(lcReferencia) > 3, 1, 2)).Value = "X"
                oWorkSheet.Cells(iFila, iColumna + IIf(UCase(Left(ml_Plan, 3)) <> "SIS", 6, 8)).Value = "X"
                oWorkSheet.Cells(iFila, iColumna + 10).Value = lcAfiliacionSIS
                oWorkSheet.Cells(iFila, iColumna + 18).Value = IIf(UCase(Left(ml_Plan, 3)) <> "SIS", Left(ml_Plan, 10), "")
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 4).Value = "'" & lcReferencia
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 6).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante) & "    DNI: " & lcDNIacompaniante
                iFila = iFila + 1
                oWorkSheet.Cells(iFila, iColumna + 6).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!telefono)
                
           End If
       End If
    End If
    

    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 100
        PrintArea(0).EndRow = 117
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
            oExcel.Visible = True
             
            
            oWorkSheet.PrintPreview
        Else
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

        If lcBuscaParametro.SeleccionaFilaParametro(216) = "1" Then
            oExcel.DisplayAlerts = False
            oExcel.Quit
        End If
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    'mgaray
    Call mo_AdminReportes.GrabarImpresionFichaAtencion(ml_idAtencion)
Exit Sub
ManejadorError:

    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
'Resume
    Exit Sub

End Sub
Sub CrearReporteHistoriaClinicaEmergScan(lnHwnd As Long, lcTipoEdad As String, lcGravedad As String, _
                                                   lcEmergenciaCorrelativo As String, lcParametro7 As String, _
                                                   lcReferencia As String, lcDNIacompaniante As String, _
                                                   lbFormato As Integer, lcAfiliacionSIS As String)
    Dim rsReporte As New Recordset
    Dim lcNombre As String, lcSql As String
    Dim iFila As Long, iColumna As Long
    Dim oCadena As New Cadena
    
'    Dim rsRsTmp As New Recordset
'    With rsRsTmp
'      .Fields.Append "FechaCobranza", adDouble
'      .LockType = adLockOptimistic
'      .Open
'    End With
    Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    
    
    
    Set EmergAdmision1.DataSource = rsReporte
    EmergAdmision1.RightMargin = 0
    EmergAdmision1.TopMargin = 0
    EmergAdmision1.LeftMargin = 0
    EmergAdmision1.BottomMargin = 10
    
    'EmergAdmision1.Sections("cabecera").Height = 16260  '14445
    Set EmergAdmision1.Sections("cabecera").Controls("ImgDISA").Picture = LoadPicture(App.Path & "\imagenes\FORMATO-EMERGENCIA111 DIRESA.jpg")
    EmergAdmision1.Sections("cabecera").Controls("etEstablecimiento").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
    EmergAdmision1.Sections("cabecera").Controls("etEmergenciaNro").Caption = lcEmergenciaCorrelativo
    EmergAdmision1.Sections("cabecera").Controls("etCuenta").Caption = "N°Cta: " & Trim(Str(ml_idCuentaAtencion))
    EmergAdmision1.Sections("cabecera").Controls("etHistoria").Caption = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), False)
    EmergAdmision1.Sections("cabecera").Controls("etHora").Caption = mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
    EmergAdmision1.Sections("cabecera").Controls("etFecha").Caption = mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
    EmergAdmision1.Sections("cabecera").Controls("etPaciente").Caption = mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
    EmergAdmision1.Sections("cabecera").Controls("etProcedencia").Caption = mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
    EmergAdmision1.Sections("cabecera").Controls("etDNI").Caption = mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento)
    EmergAdmision1.Sections("cabecera").Controls("etEdad").Caption = mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
    EmergAdmision1.Sections("cabecera").Controls("etSexo").Caption = IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
    EmergAdmision1.Sections("cabecera").Controls("etDomicilio").Caption = mo_ReporteUtil.NullToVacio(Left(Trim(rsReporte!DireccionDomicilio), 65))
    EmergAdmision1.Sections("cabecera").Controls("etReferidoSI").Caption = IIf(Len(lcReferencia) > 3, "X", "")
    EmergAdmision1.Sections("cabecera").Controls("etReferidoNO").Caption = IIf(Len(lcReferencia) > 3, "", "X")
    EmergAdmision1.Sections("cabecera").Controls("etSISno").Caption = IIf(UCase(Left(ml_Plan, 3)) <> "SIS", "X", "")
    EmergAdmision1.Sections("cabecera").Controls("etSISsi").Caption = IIf(UCase(Left(ml_Plan, 3)) <> "SIS", "", "X")
    EmergAdmision1.Sections("cabecera").Controls("etSIS").Caption = IIf(UCase(Left(ml_Plan, 3)) <> "SIS", "", lcAfiliacionSIS)
    EmergAdmision1.Sections("cabecera").Controls("etConveniosSoat").Caption = IIf(UCase(Left(ml_Plan, 3)) <> "SIS", Left(ml_Plan, 15), "")
    EmergAdmision1.Sections("cabecera").Controls("etReferencia").Caption = lcReferencia
    EmergAdmision1.Sections("cabecera").Controls("etAcompaniante").Caption = mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante) & "    DNI: " & lcDNIacompaniante
    EmergAdmision1.Sections("cabecera").Controls("etTelefono").Caption = mo_ReporteUtil.NullToVacio(rsReporte!telefono)
    If Right(lcGravedad, 2) = "1)" Then
       EmergAdmision1.Sections("cabecera").Controls("etp1").Caption = "X"
    ElseIf Right(lcGravedad, 2) = "2)" Then
       EmergAdmision1.Sections("cabecera").Controls("etp2").Caption = "X"
    ElseIf Right(lcGravedad, 2) = "3)" Then
       EmergAdmision1.Sections("cabecera").Controls("etp3").Caption = "X"
    Else
       EmergAdmision1.Sections("cabecera").Controls("etp4").Caption = "X"
    End If
    EmergAdmision1.Orientation = rptOrientPortrait
    If lcParametro7 = "1" Then
       EmergAdmision1.PrintReport False, rptRangeFromTo, 1, 1
    ElseIf lcParametro7 = "2" Then
       EmergAdmision1.Show 1
    ElseIf lcParametro7 = "3" Then
       EmergAdmision1.PrintReport False
    ElseIf lcParametro7 = "4" Then
       EmergAdmision1.PrintReport False, rptRangeAllPages
    ElseIf lcParametro7 = "5" Then
       EmergAdmision1.PrintReport False, rptRangeFromTo, 1, 2
    ElseIf lcParametro7 = "6" Then
       EmergAdmision1.PrintReport
    End If
End Sub

Sub CrearReporteHistoriaClinicaConsultorioEmerg(lnHwnd As Long, lcTipoEdad As String, lcGravedad As String, _
                                                lcEmergenciaCorrelativo As String, lcReferencia As String, _
                                                lcDNIacompaniante As String, lbFormato As Integer, _
                                                lcAfiliacionSIS As String, lcComoLlego As String, _
                                                cmbComoLlego As Integer, cmbTipoAtencion As Integer, _
                                                cmbIdTipoGravedad As Integer)
Dim lcParametro7 As String
lcParametro7 = lcBuscaParametro.SeleccionaFilaParametro(510)
If lcParametro7 = "S" Then
   lcParametro7 = IIf(Left(lcGravedad, 1) = "1" Or Left(lcGravedad, 1) = "2", "2", "")
Else
   lcParametro7 = lcBuscaParametro.SeleccionaFilaParametro(505)
End If
'lcParametro7 = "11"
If lcParametro7 <> "" Then
   If lcParametro7 = "1" Or lcParametro7 = "2" Or lcParametro7 = "3" Or _
                                     lcParametro7 = "4" Or lcParametro7 = "5" Or lcParametro7 = "6" Then
        CrearReporteHistoriaClinicaEmergScan lnHwnd, lcTipoEdad, lcGravedad, _
                                                        lcEmergenciaCorrelativo, lcParametro7, _
                                                        lcReferencia, lcDNIacompaniante, _
                                                        lbFormato, lcAfiliacionSIS
   Else
        Dim lnFila1 As Long, lnColumna1 As Long
        lnFila1 = Left(lcParametro7, InStr(lcParametro7, "-") - 1)
        lnColumna1 = Mid(lcParametro7, InStr(lcParametro7, "-") + 1, 10)
        CrearReporteHistoriaClinicaConsultorioEmergCab lnHwnd, lcTipoEdad, lcGravedad, _
                                                        lcEmergenciaCorrelativo, lnFila1, _
                                                        lnColumna1, lcReferencia, _
                                                        lcDNIacompaniante, lbFormato, lcAfiliacionSIS
   End If
   Exit Sub
End If

If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 1910 Then                               'sullana
   CrearReporteHistoriaClinicaConsultorioEmSULL lnHwnd, lcTipoEdad, lcGravedad, _
                                                lcEmergenciaCorrelativo, lcReferencia, _
                                                lcDNIacompaniante, lbFormato, _
                                                lcAfiliacionSIS, lcComoLlego, _
                                                cmbComoLlego, cmbTipoAtencion, _
                                                cmbIdTipoGravedad

   Exit Sub
ElseIf Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3299 Then                               'juliaca
   CrearReporteHistoriaClinicaConsultorioEmSICUANI lnHwnd, lcTipoEdad, lcGravedad, _
                                                lcEmergenciaCorrelativo, lcReferencia, _
                                                lcDNIacompaniante, lbFormato, _
                                                lcAfiliacionSIS, lcComoLlego, _
                                                cmbComoLlego, cmbTipoAtencion, _
                                                cmbIdTipoGravedad

   Exit Sub
End If

Dim rsReporte As New Recordset
Dim oRsTmp98 As New Recordset
Dim oConexion As New Connection
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String, lcSql1 As String
Dim iFila As Long, iPos As Long
Dim oCadena As New Cadena

oConexion.CommandTimeout = 900
oConexion.CursorLocation = adUseClient
oConexion.Open SIGHEntidades.CadenaConexion

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
        lcArchivoExcel = App.Path + "\Plantillas\EHistoriaConsultorioEmerg.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EHistoriaConsultorioEmerg.xls")
        oWorkBookPlantilla.Worksheets("HistoriaConsultorioEmerg").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    'Izquierdo
    CargaAtencionCEJamo
    Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    '
    If IsNull(rsReporte!telefono) Then
       lcSql = mo_ReporteUtil.NullToVacio(rsReporte!telefono) & Space(50) & " GRAVEDAD: " & lcGravedad
    Else
       lcSql = mo_ReporteUtil.NullToVacio(rsReporte!telefono) & Space(38) & " GRAVEDAD: " & lcGravedad
    End If
    '
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(79, 3).setFormula(lcEmergenciaCorrelativo)
        Call Feuille.getcellbyposition(28, 5).setFormula("'" & Trim(Str(rsReporte!NroHistoriaClinica)) & " - " & rsReporte!tipoNumeracion)
        Call Feuille.getcellbyposition(28, 6).setFormula(mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))
        Call Feuille.getcellbyposition(28, 7).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso))
        Call Feuille.getcellbyposition(28, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso))
        Call Feuille.getcellbyposition(28, 9).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!Servicio))
        Call Feuille.getcellbyposition(28, 10).setFormula(UCase(mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico))))
        Call Feuille.getcellbyposition(28, 11).setFormula(lcSql)
        Call Feuille.getcellbyposition(28, 12).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio))
        Call Feuille.getcellbyposition(28, 13).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio))
        Call Feuille.getcellbyposition(28, 14).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio))
        Call Feuille.getcellbyposition(28, 15).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante))

        Call Feuille.getcellbyposition(79, 5).setFormula("'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan)
        Call Feuille.getcellbyposition(84, 7).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad)
        Call Feuille.getcellbyposition(84, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada))
        Call Feuille.getcellbyposition(84, 9).setFormula(IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino"))
        Call Feuille.getcellbyposition(84, 13).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio))
        Call Feuille.getcellbyposition(84, 14).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio))
        
    Else
        oWorkSheet.Cells(4, 80).Value = "'" & lcEmergenciaCorrelativo
        oWorkSheet.Cells(4, 80).Font.Bold = True
        
        
        oWorkSheet.Cells(6, 29).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True) & _
                                            " - " & rsReporte!tipoNumeracion
        oWorkSheet.Cells(7, 29).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
        oWorkSheet.Cells(8, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
        oWorkSheet.Cells(9, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
        oWorkSheet.Cells(10, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio)
        oWorkSheet.Cells(11, 29).Value = UCase("'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico)))
        oWorkSheet.Cells(12, 29).Value = "'" & lcSql
        oWorkSheet.Cells(12, 29).Font.Bold = True     'debb-15/04/2016
        
        If Not IsNull(rsReporte!NroDocumento) Then
           oWorkSheet.Cells(11, 63).Value = " DNI: " & rsReporte!NroDocumento
           oWorkSheet.Cells(11, 63).Font.Bold = True
        End If
        

        oWorkSheet.Cells(13, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio)
        oWorkSheet.Cells(14, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio)
        oWorkSheet.Cells(15, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
        oWorkSheet.Cells(16, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante)
        
        'Derecho
        oWorkSheet.Cells(6, 80).Value = "'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan
        oWorkSheet.Cells(8, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
        oWorkSheet.Cells(9, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada)
        oWorkSheet.Cells(10, 85).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
        oWorkSheet.Cells(14, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio)
        oWorkSheet.Cells(15, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio)
        oWorkSheet.Cells(16, 63).Value = "USUARIO: " & oReglasCaja.SeleccionaDatosCajero(SIGHEntidades.Usuario, sghUsuario)
        If Not IsNull(rsReporte!obsPaciente) Then
           oWorkSheet.Cells(20, 3).Font.Name = "Arial Narrow"
           oWorkSheet.Cells(20, 3).Font.Size = "9"
           oWorkSheet.Cells(20, 3).Value = "Observaciones del Paciente: " & mo_ReporteUtil.NullToVacio(rsReporte!obsPaciente)
        End If
    End If
    If mo_DOAtencionesCE.idAtencion > 0 Then
        If lbEsOpenOffice = True Then
            
            Call Feuille.getcellbyposition(3, 22).setFormula(mo_DOAtencionesCE.TriajePresion)
            Call Feuille.getcellbyposition(24, 22).setFormula("'" & mo_DOAtencionesCE.TriajeFrecRespiratoria)
            Call Feuille.getcellbyposition(49, 22).setFormula("'" & mo_DOAtencionesCE.TriajePulso)
            Call Feuille.getcellbyposition(72, 22).setFormula("'" & mo_DOAtencionesCE.TriajeTemperatura)
            
        Else
            oWorkSheet.Cells(23, 4).Value = mo_DOAtencionesCE.TriajePresion
            oWorkSheet.Cells(23, 25).Value = "'" & mo_DOAtencionesCE.TriajeFrecRespiratoria
            oWorkSheet.Cells(23, 50).Value = "'" & mo_DOAtencionesCE.TriajePulso
            oWorkSheet.Cells(23, 73).Value = "'" & mo_DOAtencionesCE.TriajeTemperatura
            
            oWorkSheet.Cells(23, 90).Value = "'" & mo_DOAtencionesCE.TriajeFrecCardiaca
            
            oWorkSheet.Cells(24, 4).Value = "Peso: " & IIf(mo_DOAtencionesCE.TriajePeso = "", Space(10), mo_DOAtencionesCE.TriajePeso) & " (Kg)"
            oWorkSheet.Cells(24, 30).Value = "Talla: " & IIf(mo_DOAtencionesCE.TriajeTalla = "", Space(10), mo_DOAtencionesCE.TriajeTalla) & " (cm)"
            oWorkSheet.Cells(24, 60).Value = "Saturación de Oxígeno: " & IIf(mo_DOAtencionesCE.TriajeSaturacionOxigeno = "", Space(5), mo_DOAtencionesCE.TriajeSaturacionOxigeno)
            
        End If
    Else

            oWorkSheet.Cells(24, 4).Value = "Peso: " & Space(10) & " (Kg)"
            oWorkSheet.Cells(24, 30).Value = "Talla: " & Space(10) & " (cm)"
            oWorkSheet.Cells(24, 60).Value = "Saturación de Oxígeno: "
    End If
    
    
    'dx
    Set oRsTmp98 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(ml_idAtencion, oConexion)
    If oRsTmp98.RecordCount > 0 Then
       iFila = 44
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!descripcion
                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!CodigoCIE2004
                oWorkSheet.Cells(iFila, IIf(oRsTmp98!idSubClasificacionDx = 301 Or oRsTmp98!idSubClasificacionDx = 303, 85, 82)).Value = "X"
            End If
            iFila = iFila + 1
            If iFila = 51 Then
               Exit Do
            End If
            oRsTmp98.MoveNext
       Loop
       If iFila < 51 Then
          Do While iFila < 51
             If lbEsOpenOffice = True Then
             Else
                oWorkSheet.Cells(iFila, 4).Value = "'"
                oWorkSheet.Cells(iFila, 92).Value = "'"
             End If
             iFila = iFila + 1
          Loop
       End If
    End If
    oRsTmp98.Close
    'tratamiento
    If Not IsNull(mo_DOAtencionesCE.CitaTratamiento) Then
        iFila = 53
        iPos = 70
        lcSql = Left(mo_DOAtencionesCE.CitaTratamiento, iPos)
        If lbEsOpenOffice = True Then
        Else
            oWorkSheet.Cells(iFila, 23).Value = lcSql
        End If
        iPos = 71
        iFila = iFila + 1
        Do While True
           lcSql = Mid(mo_DOAtencionesCE.CitaTratamiento, iPos, 90)
           iPos = iPos + 90
           If lbEsOpenOffice = True Then
           Else
                oWorkSheet.Cells(iFila, 3).Value = "'" & lcSql
                iFila = iFila + 1
                If iPos >= Len(mo_DOAtencionesCE.CitaTratamiento) Then
                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
                If iFila = 61 Then
                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
           End If
        Loop
        
        
    End If
    'Servicios Intermedios
    Set oRsTmp98 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
    oRsTmp98.Filter = "idEstadoFacturacion<>9 and idPuntoCarga<>1"
    If oRsTmp98.RecordCount > 0 Then
        lcSql1 = "/"
        oRsTmp98.MoveFirst
        Do While Not oRsTmp98.EOF
            lcSql1 = lcSql1 & Trim(oRsTmp98!Codigo) & "-" & Trim(oRsTmp98!Nombre) & "/"
            oRsTmp98.MoveNext
        Loop
        oRsTmp98.Close
        Set oRsTmp98 = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
        If oRsTmp98.RecordCount > 0 Then
            lcSql1 = lcSql1 & "**/ "
            oRsTmp98.MoveFirst
            Do While Not oRsTmp98.EOF
                lcSql1 = lcSql1 & Trim(oRsTmp98!Codigo) & "-" & Trim(oRsTmp98!Nombre) & "/"
                oRsTmp98.MoveNext
            Loop
        End If
        
        iFila = 67
        iPos = 60
        lcSql = Left(lcSql1, iPos)
        If lbEsOpenOffice = True Then
        Else
            oWorkSheet.Cells(iFila, 31).Value = lcSql
        End If
        iPos = 61
        iFila = iFila + 1
        Do While True
           lcSql = Mid(lcSql1, iPos, 80)
           iPos = iPos + 80
           If lbEsOpenOffice = True Then
           Else
                oWorkSheet.Cells(iFila, 3).Value = "'" & lcSql
                iFila = iFila + 1
                If iPos >= Len(lcSql1) Then
                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
                If iFila = 72 Then
                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
           End If
        Loop
    End If
    oRsTmp98.Close
    'CPT realizados
    Set oRsTmp98 = mo_ReglasAdmision.BuscaAtencionesCptCEparaFormatoHIS(rsReporte!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
    oRsTmp98.Filter = "idPuntoCarga=1"
    If oRsTmp98.RecordCount > 0 Then
       iFila = 94
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!Nombre
                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!Codigo
            End If
            iFila = iFila + 1
            If iFila = 98 Then
               Exit Do
            End If
            oRsTmp98.MoveNext
       Loop
       If iFila < 98 Then
          Do While iFila < 98
             If lbEsOpenOffice = True Then
             Else
                oWorkSheet.Cells(iFila, 4).Value = "'"
                oWorkSheet.Cells(iFila, 92).Value = "'"
             End If
             iFila = iFila + 1
          Loop
       End If
    End If
    oRsTmp98.Close
    '
    'destino de atencion
    If Not IsNull(rsReporte!IdDestinoAtencion) Then
        If lbEsOpenOffice = True Then
        Else
            Select Case rsReporte!IdDestinoAtencion
            Case 20  'Su casa
                oWorkSheet.Cells(102, 3).Value = "*"
            Case 21   'Hospitalizacióm
                oWorkSheet.Cells(104, 3).Value = "*"
            Case 23  'Referencia
                oWorkSheet.Cells(107, 3).Value = "*"
            Case 24  'contrareferencia
                oWorkSheet.Cells(107, 49).Value = "*"
            Case 25      'morgue
                oWorkSheet.Cells(106, 3).Value = "*"
            Case 56      'fuga
                oWorkSheet.Cells(108, 3).Value = "*"
            Case 26      'otros
                oWorkSheet.Cells(109, 3).Value = "*"
            Case 65      'citado
                oWorkSheet.Cells(102, 30).Value = "*"
            End Select
        End If
    End If
    '
    Set oRsTmp98 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaListaMedicos(ml_idAtencion, oConexion)
    If oRsTmp98.RecordCount > 1 Then
       iFila = 116
       iPos = 1
       oWorkSheet.Cells(iFila, 20).Font.Bold = True
       oWorkSheet.Cells(iFila, 20).Font.Name = "Arial Narrow"
       oWorkSheet.Cells(iFila, 2).Font.Underline = True
       oWorkSheet.Cells(iFila, 20).Value = "Lista de Médicos que atendieron al Paciente:"
       iFila = iFila + 1
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
          oWorkSheet.Cells(iFila, IIf(iPos <= 3, 3, 60)).Value = Left(oRsTmp98!secuencia & ") " & oRsTmp98!ApellidoPaterno & " " & oRsTmp98!ApellidoMaterno & " " & oRsTmp98!Nombres & "_______________________", 45)
          iFila = iFila + 1
          iPos = iPos + 1
          If iPos = 4 Then
             iFila = 117
          End If
          oRsTmp98.MoveNext
       Loop
    End If
    oRsTmp98.Close
    '
    
    '
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 100
        PrintArea(0).EndRow = 117
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
            oExcel.Visible = True
             
            
            oWorkSheet.PrintPreview
        Else
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
        'oExcel.ActiveWindow.Close False
        'oExcel.Workbooks.Close
        'oExcel.Quit
        
        oExcel.DisplayAlerts = False
        oExcel.Quit
        
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    Set oRsTmp98 = Nothing
    Set rsReporte = Nothing
    Set oConexion = Nothing
    'mgaray
    Call mo_AdminReportes.GrabarImpresionFichaAtencion(ml_idAtencion)
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




Sub CargaAtencionCEJamo()
            On Error GoTo ErrJamo
            Dim oConexion As New Connection
            Dim lcBuscaParametro As New SIGHDatos.Parametros
            Dim oAtencionesCE As New AtencionesCE
            oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            mo_DOAtencionesCE.idAtencion = ml_idAtencion
            Set oAtencionesCE.Conexion = oConexion
            If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
                mo_DOAtencionesCE.idAtencion = 0
                Exit Sub
            End If
            oConexion.Close
            Set oConexion = Nothing
            Set oAtencionesCE = Nothing
            Exit Sub
ErrJamo:
End Sub


Sub CrearReporteHistoriaClinicaConsultorioEmSICUANI(lnHwnd As Long, lcTipoEdad As String, lcGravedad As String, _
                                                lcEmergenciaCorrelativo As String, lcReferencia As String, _
                                                lcDNIacompaniante As String, lbFormato As Integer, _
                                                lcAfiliacionSIS As String, lcComoLlego As String, _
                                                cmbComoLlego As Integer, cmbTipoAtencion As Integer, _
                                                cmbIdTipoGravedad As Integer)

Dim rsReporte As New Recordset
Dim oRsTmp98 As New Recordset
Dim oConexion As New Connection
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String, lcSql1 As String
Dim iFila As Long, iPos As Long
Dim oCadena As New Cadena

oConexion.CommandTimeout = 900
oConexion.CursorLocation = adUseClient
oConexion.Open SIGHEntidades.CadenaConexion

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
        lcArchivoExcel = App.Path + "\Plantillas\EHistoriaConsultorioEmerg.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EHistoriaConsultorioEmSICU.xls")
        oWorkBookPlantilla.Worksheets("HistoriaConsultorioEmerg").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    'Izquierdo
    CargaAtencionCEJamo
    Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    '
    If IsNull(rsReporte!telefono) Then
       lcSql = mo_ReporteUtil.NullToVacio(rsReporte!telefono) & Space(50) & " GRAVEDAD: " & lcGravedad
    Else
       lcSql = mo_ReporteUtil.NullToVacio(rsReporte!telefono) & Space(38) & " GRAVEDAD: " & lcGravedad
    End If
    '
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(79, 3).setFormula(lcEmergenciaCorrelativo)
        Call Feuille.getcellbyposition(28, 5).setFormula("'" & Trim(Str(rsReporte!NroHistoriaClinica)) & " - " & rsReporte!tipoNumeracion)
        Call Feuille.getcellbyposition(28, 6).setFormula(mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))
        Call Feuille.getcellbyposition(28, 7).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso))
        Call Feuille.getcellbyposition(28, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso))
        Call Feuille.getcellbyposition(28, 9).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!Servicio))
        Call Feuille.getcellbyposition(28, 10).setFormula(UCase(mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico))))
        Call Feuille.getcellbyposition(28, 11).setFormula(lcSql)
        Call Feuille.getcellbyposition(28, 12).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio))
        Call Feuille.getcellbyposition(28, 13).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio))
        Call Feuille.getcellbyposition(28, 14).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio))
        Call Feuille.getcellbyposition(28, 15).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante))

        Call Feuille.getcellbyposition(79, 5).setFormula("'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan)
        Call Feuille.getcellbyposition(84, 7).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad)
        Call Feuille.getcellbyposition(84, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada))
        Call Feuille.getcellbyposition(84, 9).setFormula(IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino"))
        Call Feuille.getcellbyposition(84, 13).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio))
        Call Feuille.getcellbyposition(84, 14).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio))
        
    Else
        oWorkSheet.Cells(4, 80).Value = "'" & lcEmergenciaCorrelativo
        oWorkSheet.Cells(4, 80).Font.Bold = True
        
        
        oWorkSheet.Cells(6, 29).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True) & _
                                            " - " & rsReporte!tipoNumeracion
        oWorkSheet.Cells(7, 29).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
        oWorkSheet.Cells(8, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
        oWorkSheet.Cells(9, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
        oWorkSheet.Cells(10, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio)
        oWorkSheet.Cells(11, 29).Value = UCase("'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico)))
        oWorkSheet.Cells(12, 29).Value = "'" & lcSql
        oWorkSheet.Cells(12, 29).Font.Bold = True     'debb-15/04/2016
        
        If Not IsNull(rsReporte!NroDocumento) Then
           oWorkSheet.Cells(11, 63).Value = " DNI: " & rsReporte!NroDocumento
           oWorkSheet.Cells(11, 63).Font.Bold = True
        End If
        

        oWorkSheet.Cells(13, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio)
        oWorkSheet.Cells(14, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio)
        oWorkSheet.Cells(15, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
        oWorkSheet.Cells(16, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante)
        
        'Derecho
        oWorkSheet.Cells(6, 80).Value = "'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan
        oWorkSheet.Cells(8, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
        oWorkSheet.Cells(9, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada)
        oWorkSheet.Cells(10, 85).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
        oWorkSheet.Cells(14, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio)
        oWorkSheet.Cells(15, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio)
        oWorkSheet.Cells(16, 63).Value = "USUARIO: " & oReglasCaja.SeleccionaDatosCajero(SIGHEntidades.Usuario, sghUsuario)
        If Not IsNull(rsReporte!obsPaciente) Then
           oWorkSheet.Cells(20, 3).Font.Name = "Arial Narrow"
           oWorkSheet.Cells(20, 3).Font.Size = "9"
           oWorkSheet.Cells(20, 3).Value = "Observaciones del Paciente: " & mo_ReporteUtil.NullToVacio(rsReporte!obsPaciente)
        End If
    End If
    If mo_DOAtencionesCE.idAtencion > 0 Then
        If lbEsOpenOffice = True Then
            
            Call Feuille.getcellbyposition(3, 22).setFormula(mo_DOAtencionesCE.TriajePresion)
            Call Feuille.getcellbyposition(24, 22).setFormula("'" & mo_DOAtencionesCE.TriajeFrecRespiratoria)
            Call Feuille.getcellbyposition(49, 22).setFormula("'" & mo_DOAtencionesCE.TriajePulso)
            Call Feuille.getcellbyposition(72, 22).setFormula("'" & mo_DOAtencionesCE.TriajeTemperatura)
            
        Else
            oWorkSheet.Cells(32, 92).Value = mo_DOAtencionesCE.TriajePresion
            oWorkSheet.Cells(33, 92).Value = "'" & mo_DOAtencionesCE.TriajeFrecCardiaca
            oWorkSheet.Cells(34, 92).Value = "'" & mo_DOAtencionesCE.TriajeTemperatura
            oWorkSheet.Cells(35, 92).Value = "'" & mo_DOAtencionesCE.TriajeFrecRespiratoria
            oWorkSheet.Cells(37, 92).Value = "'" & mo_DOAtencionesCE.TriajePeso
            oWorkSheet.Cells(38, 92).Value = "'" & mo_DOAtencionesCE.TriajeTalla
            
        End If
    End If
    
    
    'dx - ingresos
    Set oRsTmp98 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(ml_idAtencion, oConexion)
    oRsTmp98.Filter = "idSubClasificacionDx=null"
    If oRsTmp98.RecordCount > 0 Then
       iFila = 43
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!descripcion
                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!CodigoCIE2004
                oWorkSheet.Cells(iFila, IIf(oRsTmp98!idSubClasificacionDx = 301 Or oRsTmp98!idSubClasificacionDx = 303, 85, 82)).Value = "X"
            End If
            iFila = iFila + 1
            If iFila = 46 Then
               Exit Do
            End If
            oRsTmp98.MoveNext
       Loop
       If iFila < 46 Then
          Do While iFila < 46
             If lbEsOpenOffice = True Then
             Else
                oWorkSheet.Cells(iFila, 4).Value = "'"
                oWorkSheet.Cells(iFila, 92).Value = "'"
             End If
             iFila = iFila + 1
          Loop
       End If
    End If
    oRsTmp98.Close
'    'tratamiento
'    If Not IsNull(mo_DOAtencionesCE.CitaTratamiento) Then
'        iFila = 53
'        iPos = 70
'        lcSql = Left(mo_DOAtencionesCE.CitaTratamiento, iPos)
'        If lbEsOpenOffice = True Then
'        Else
'            oWorkSheet.Cells(iFila, 23).Value = lcSql
'        End If
'        iPos = 71
'        iFila = iFila + 1
'        Do While True
'           lcSql = Mid(mo_DOAtencionesCE.CitaTratamiento, iPos, 90)
'           iPos = iPos + 90
'           If lbEsOpenOffice = True Then
'           Else
'                oWorkSheet.Cells(iFila, 3).Value = "'" & lcSql
'                iFila = iFila + 1
'                If iPos >= Len(mo_DOAtencionesCE.CitaTratamiento) Then
'                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
'                   Exit Do
'                End If
'                If iFila = 61 Then
'                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
'                   Exit Do
'                End If
'           End If
'        Loop
'
'
'    End If
'    'Servicios Intermedios
'    Set oRsTmp98 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
'    oRsTmp98.Filter = "idEstadoFacturacion<>9 and idPuntoCarga<>1"
'    If oRsTmp98.RecordCount > 0 Then
'        lcSql1 = "/"
'        oRsTmp98.MoveFirst
'        Do While Not oRsTmp98.EOF
'            lcSql1 = lcSql1 & Trim(oRsTmp98!Codigo) & "-" & Trim(oRsTmp98!Nombre) & "/"
'            oRsTmp98.MoveNext
'        Loop
'        oRsTmp98.Close
'        Set oRsTmp98 = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
'        If oRsTmp98.RecordCount > 0 Then
'            lcSql1 = lcSql1 & "**/ "
'            oRsTmp98.MoveFirst
'            Do While Not oRsTmp98.EOF
'                lcSql1 = lcSql1 & Trim(oRsTmp98!Codigo) & "-" & Trim(oRsTmp98!Nombre) & "/"
'                oRsTmp98.MoveNext
'            Loop
'        End If
'
'        iFila = 67
'        iPos = 60
'        lcSql = Left(lcSql1, iPos)
'        If lbEsOpenOffice = True Then
'        Else
'            oWorkSheet.Cells(iFila, 31).Value = lcSql
'        End If
'        iPos = 61
'        iFila = iFila + 1
'        Do While True
'           lcSql = Mid(lcSql1, iPos, 80)
'           iPos = iPos + 80
'           If lbEsOpenOffice = True Then
'           Else
'                oWorkSheet.Cells(iFila, 3).Value = "'" & lcSql
'                iFila = iFila + 1
'                If iPos >= Len(lcSql1) Then
'                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
'                   Exit Do
'                End If
'                If iFila = 72 Then
'                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
'                   Exit Do
'                End If
'           End If
'        Loop
'    End If
'    oRsTmp98.Close
'    'CPT realizados
'    Set oRsTmp98 = mo_ReglasAdmision.BuscaAtencionesCptCEparaFormatoHIS(rsReporte!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
'    oRsTmp98.Filter = "idPuntoCarga=1"
'    If oRsTmp98.RecordCount > 0 Then
'       iFila = 94
'       oRsTmp98.MoveFirst
'       Do While Not oRsTmp98.EOF
'            If lbEsOpenOffice = True Then
'            Else
'                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!Nombre
'                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!Codigo
'            End If
'            iFila = iFila + 1
'            If iFila = 98 Then
'               Exit Do
'            End If
'            oRsTmp98.MoveNext
'       Loop
'       If iFila < 98 Then
'          Do While iFila < 98
'             If lbEsOpenOffice = True Then
'             Else
'                oWorkSheet.Cells(iFila, 4).Value = "'"
'                oWorkSheet.Cells(iFila, 92).Value = "'"
'             End If
'             iFila = iFila + 1
'          Loop
'       End If
'    End If
'    oRsTmp98.Close
    '
    'destino de atencion
    If Not IsNull(rsReporte!IdDestinoAtencion) Then
        If lbEsOpenOffice = True Then
        Else
            Select Case rsReporte!IdDestinoAtencion
            Case 20  'Su casa
                oWorkSheet.Cells(96, 3).Value = "*"
            Case 21   'Hospitalizacióm
                oWorkSheet.Cells(98, 3).Value = "*"
            Case 23  'Referencia
                oWorkSheet.Cells(101, 3).Value = "*"
            Case 24  'contrareferencia
                oWorkSheet.Cells(101, 49).Value = "*"
            Case 25      'morgue
                oWorkSheet.Cells(100, 3).Value = "*"
            Case 56      'fuga
                oWorkSheet.Cells(102, 3).Value = "*"
            Case 26      'otros
                oWorkSheet.Cells(103, 3).Value = "*"
            Case 65      'citado
                oWorkSheet.Cells(96, 30).Value = "*"
            End Select
        End If
    End If
    '
'    Set oRsTmp98 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaListaMedicos(ml_idAtencion, oConexion)
'    If oRsTmp98.RecordCount > 1 Then
'       iFila = 116
'       iPos = 1
'       oWorkSheet.Cells(iFila, 20).Font.Bold = True
'       oWorkSheet.Cells(iFila, 20).Font.Name = "Arial Narrow"
'       oWorkSheet.Cells(iFila, 2).Font.Underline = True
'       oWorkSheet.Cells(iFila, 20).Value = "Lista de Médicos que atendieron al Paciente:"
'       iFila = iFila + 1
'       oRsTmp98.MoveFirst
'       Do While Not oRsTmp98.EOF
'          oWorkSheet.Cells(iFila, IIf(iPos <= 3, 3, 60)).Value = Left(oRsTmp98!secuencia & ") " & oRsTmp98!ApellidoPaterno & " " & oRsTmp98!apellidoMaterno & " " & oRsTmp98!nombres & "_______________________", 45)
'          iFila = iFila + 1
'          iPos = iPos + 1
'          If iPos = 4 Then
'             iFila = 117
'          End If
'          oRsTmp98.MoveNext
'       Loop
'    End If
'    oRsTmp98.Close
    '
    'dx - salidas
    Set oRsTmp98 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(ml_idAtencion, oConexion)
    oRsTmp98.Filter = "idSubClasificacionDx<>null"
    If oRsTmp98.RecordCount > 0 Then
       iFila = 109
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!descripcion
                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!CodigoCIE2004
                oWorkSheet.Cells(iFila, IIf(oRsTmp98!idSubClasificacionDx = 301 Or oRsTmp98!idSubClasificacionDx = 303, 85, 82)).Value = "X"
            End If
            iFila = iFila + 1
            If iFila = 112 Then
               Exit Do
            End If
            oRsTmp98.MoveNext
       Loop
       If iFila < 112 Then
          Do While iFila < 112
             If lbEsOpenOffice = True Then
             Else
                oWorkSheet.Cells(iFila, 4).Value = "'"
                oWorkSheet.Cells(iFila, 92).Value = "'"
             End If
             iFila = iFila + 1
          Loop
       End If
    End If
    oRsTmp98.Close
    
    '
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 100
        PrintArea(0).EndRow = 117
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
            oExcel.Visible = True
             
            
            oWorkSheet.PrintPreview
        Else
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
        'oExcel.ActiveWindow.Close False
        'oExcel.Workbooks.Close
        'oExcel.Quit
        
        oExcel.DisplayAlerts = False
        oExcel.Quit
        
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    Set oRsTmp98 = Nothing
    Set rsReporte = Nothing
    Set oConexion = Nothing
    'mgaray
    Call mo_AdminReportes.GrabarImpresionFichaAtencion(ml_idAtencion)
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

Sub CrearReporteHistoriaClinicaConsultorioEmSULL(lnHwnd As Long, lcTipoEdad As String, lcGravedad As String, _
                                                lcEmergenciaCorrelativo As String, lcReferencia As String, _
                                                lcDNIacompaniante As String, lbFormato As Integer, _
                                                lcAfiliacionSIS As String, lcComoLlego As String, _
                                                cmbComoLlego As Integer, cmbTipoAtencion As Integer, _
                                                cmbIdTipoGravedad As Integer)
                                                
Dim rsReporte As New Recordset
Dim oRsTmp98 As New Recordset
Dim oConexion As New Connection
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String, lcSql1 As String
Dim iFila As Long, iPos As Long
Dim oCadena As New Cadena

oConexion.CommandTimeout = 900
oConexion.CursorLocation = adUseClient
oConexion.Open SIGHEntidades.CadenaConexion

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
        lcArchivoExcel = App.Path + "\Plantillas\EHistoriaConsultorioEmerg.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EHistoriaConsultorioEmSULL.xls")
        oWorkBookPlantilla.Worksheets("HistoriaConsultorioEmerg").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    'Izquierdo
    CargaAtencionCEJamo
    Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    '
    If IsNull(rsReporte!telefono) Then
       lcSql = mo_ReporteUtil.NullToVacio(rsReporte!telefono) & Space(56) & lcComoLlego
    Else
       lcSql = mo_ReporteUtil.NullToVacio(rsReporte!telefono) & Space(41) & lcComoLlego
    End If
    '
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(79, 3).setFormula(lcEmergenciaCorrelativo)
        Call Feuille.getcellbyposition(28, 5).setFormula("'" & Trim(Str(rsReporte!NroHistoriaClinica)) & " - " & rsReporte!tipoNumeracion)
        Call Feuille.getcellbyposition(28, 6).setFormula(mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))
        Call Feuille.getcellbyposition(28, 7).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso))
        Call Feuille.getcellbyposition(28, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso))
        Call Feuille.getcellbyposition(28, 9).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!Servicio))
        Call Feuille.getcellbyposition(28, 10).setFormula(UCase(mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico))))
        Call Feuille.getcellbyposition(28, 11).setFormula(lcSql)
        Call Feuille.getcellbyposition(28, 12).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio))
        Call Feuille.getcellbyposition(28, 13).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio))
        Call Feuille.getcellbyposition(28, 14).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio))
        Call Feuille.getcellbyposition(28, 15).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante))

        Call Feuille.getcellbyposition(79, 5).setFormula("'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan)
        Call Feuille.getcellbyposition(84, 7).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad)
        Call Feuille.getcellbyposition(84, 8).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada))
        Call Feuille.getcellbyposition(84, 9).setFormula(IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino"))
        Call Feuille.getcellbyposition(84, 13).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio))
        Call Feuille.getcellbyposition(84, 14).setFormula(mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio))
        
    Else
        lcGravedad = Mid(lcGravedad, InStr(lcGravedad, "sem=") + 4, 1)
         
        oWorkSheet.Cells(4, 80).Value = "'" & lcEmergenciaCorrelativo
        oWorkSheet.Cells(4, 80).Font.Bold = True
        
        
        
        oWorkSheet.Cells(6, 29).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True) & _
                                              " - " & rsReporte!tipoNumeracion
        oWorkSheet.Cells(7, 29).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
        oWorkSheet.Cells(8, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
        oWorkSheet.Cells(9, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
        oWorkSheet.Cells(10, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio)
        oWorkSheet.Cells(11, 29).Value = UCase("'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico)))
       ' oWorkSheet.Cells(12, 29).Value = "'" & lcSql
        'oWorkSheet.Cells(12, 29).Font.Bold = True     'debb-15/04/2016

        oWorkSheet.Cells(13, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio)
        oWorkSheet.Cells(14, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio)
        oWorkSheet.Cells(15, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
        oWorkSheet.Cells(16, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante)
        
        'Derecho
        If Not IsNull(rsReporte!IdDestinoAtencion) Then
           oWorkSheet.Cells(5, IIf(lcGravedad = "1", 79, IIf(lcGravedad = 2, 87, IIf(lcGravedad = 3, 95, 101)))).Value = "x"
        End If
        oWorkSheet.Cells(6, 80).Value = "'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan
        oWorkSheet.Cells(8, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
        oWorkSheet.Cells(9, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada)
        oWorkSheet.Cells(10, 85).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
        If Not IsNull(rsReporte!IdDestinoAtencion) Then
           oWorkSheet.Cells(12, IIf(cmbComoLlego = 0, 41, IIf(cmbComoLlego = 1, 62, IIf(cmbComoLlego = 2, 77, 88)))).Value = "x"
        End If
        oWorkSheet.Cells(14, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio)
        oWorkSheet.Cells(15, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio)
        If Not IsNull(rsReporte!IdDestinoAtencion) Then
           oWorkSheet.Cells(16, IIf(cmbTipoAtencion = 0, 83, 99)).Value = "x"
        End If
    End If
    If mo_DOAtencionesCE.idAtencion > 0 Then
        If lbEsOpenOffice = True Then
            
            Call Feuille.getcellbyposition(3, 22).setFormula(mo_DOAtencionesCE.TriajePresion)
            Call Feuille.getcellbyposition(24, 22).setFormula("'" & mo_DOAtencionesCE.TriajeFrecRespiratoria)
            Call Feuille.getcellbyposition(49, 22).setFormula("'" & mo_DOAtencionesCE.TriajePulso)
            Call Feuille.getcellbyposition(72, 22).setFormula("'" & mo_DOAtencionesCE.TriajeTemperatura)
            
        Else
            oWorkSheet.Cells(23, 4).Value = mo_DOAtencionesCE.TriajePresion
            oWorkSheet.Cells(23, 25).Value = "'" & mo_DOAtencionesCE.TriajeFrecRespiratoria
            oWorkSheet.Cells(23, 50).Value = "'" & mo_DOAtencionesCE.TriajePulso
            oWorkSheet.Cells(23, 73).Value = "'" & mo_DOAtencionesCE.TriajeTemperatura
            
            oWorkSheet.Cells(23, 90).Value = "'" & mo_DOAtencionesCE.TriajeFrecCardiaca
        End If
    End If
    
    
    'dx
    Set oRsTmp98 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(ml_idAtencion, oConexion)
    If oRsTmp98.RecordCount > 0 Then
       iFila = 84
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!descripcion
                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!CodigoCIE2004
                oWorkSheet.Cells(iFila, IIf(oRsTmp98!idSubClasificacionDx = 301 Or oRsTmp98!idSubClasificacionDx = 303, 85, 82)).Value = "X"
            End If
            iFila = iFila + 1
            If iFila = 91 Then
               Exit Do
            End If
            oRsTmp98.MoveNext
       Loop
       If iFila < 91 Then
          Do While iFila < 91
             If lbEsOpenOffice = True Then
             Else
                oWorkSheet.Cells(iFila, 4).Value = "'"
                oWorkSheet.Cells(iFila, 92).Value = "'"
             End If
             iFila = iFila + 1
          Loop
       End If
    End If
    oRsTmp98.Close
    'tratamiento
    If Not IsNull(mo_DOAtencionesCE.CitaTratamiento) Then
        iFila = 89
        iPos = 70
        lcSql = Left(mo_DOAtencionesCE.CitaTratamiento, iPos)
        If lbEsOpenOffice = True Then
        Else
            oWorkSheet.Cells(iFila, 23).Value = lcSql
        End If
        iPos = 71
        iFila = iFila + 1
        Do While True
           lcSql = Mid(mo_DOAtencionesCE.CitaTratamiento, iPos, 90)
           iPos = iPos + 90
           If lbEsOpenOffice = True Then
           Else
                oWorkSheet.Cells(iFila, 3).Value = "'" & lcSql
                iFila = iFila + 1
                If iPos >= Len(mo_DOAtencionesCE.CitaTratamiento) Then
                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
                If iFila = 61 Then
                   oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
           End If
        Loop
        
        
    End If
    'Servicios Intermedios
    Set oRsTmp98 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
    oRsTmp98.Filter = "idEstadoFacturacion<>9 and idPuntoCarga<>1"
    If oRsTmp98.RecordCount > 0 Then
        lcSql1 = "/"
        oRsTmp98.MoveFirst
        Do While Not oRsTmp98.EOF
            lcSql1 = lcSql1 & Trim(oRsTmp98!Codigo) & "-" & Trim(oRsTmp98!Nombre) & "/"
            oRsTmp98.MoveNext
        Loop
        oRsTmp98.Close
        Set oRsTmp98 = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
        If oRsTmp98.RecordCount > 0 Then
            lcSql1 = lcSql1 & "**/ "
            oRsTmp98.MoveFirst
            Do While Not oRsTmp98.EOF
                lcSql1 = lcSql1 & Trim(oRsTmp98!Codigo) & "-" & Trim(oRsTmp98!Nombre) & "/"
                oRsTmp98.MoveNext
            Loop
        End If
        
        iFila = 67
        iPos = 60
        lcSql = Left(lcSql1, iPos)
        If lbEsOpenOffice = True Then
        Else
           ' oWorkSheet.Cells(iFila, 31).Value = lcSql
        End If
        iPos = 61
        iFila = iFila + 1
        Do While True
           lcSql = Mid(lcSql1, iPos, 80)
           iPos = iPos + 80
           If lbEsOpenOffice = True Then
           Else
              '  oWorkSheet.Cells(iFila, 3).Value = "'" & lcSql
                iFila = iFila + 1
                If iPos >= Len(lcSql1) Then
                  ' oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
                If iFila = 72 Then
                  ' oWorkSheet.Cells(iFila - 1, 3).Value = "'" & lcSql & "..."
                   Exit Do
                End If
           End If
        Loop
    End If
    oRsTmp98.Close
    'CPT realizados
    Set oRsTmp98 = mo_ReglasAdmision.BuscaAtencionesCptCEparaFormatoHIS(rsReporte!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
    oRsTmp98.Filter = "idPuntoCarga=1"
    If oRsTmp98.RecordCount > 0 Then
       iFila = 149
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(iFila, 4).Value = "'" & oRsTmp98!Nombre
                oWorkSheet.Cells(iFila, 92).Value = "'" & oRsTmp98!Codigo
            End If
            iFila = iFila + 1
            If iFila = 152 Then
               Exit Do
            End If
            oRsTmp98.MoveNext
       Loop
       If iFila < 152 Then
          Do While iFila < 152
             If lbEsOpenOffice = True Then
             Else
                oWorkSheet.Cells(iFila, 4).Value = "'"
                oWorkSheet.Cells(iFila, 92).Value = "'"
             End If
             iFila = iFila + 1
          Loop
       End If
    End If
    oRsTmp98.Close
    '
    'destino de atencion
    If Not IsNull(rsReporte!IdDestinoAtencion) Then
        If lbEsOpenOffice = True Then
        Else
            Select Case rsReporte!IdDestinoAtencion
            Case 20  'Su casa
                oWorkSheet.Cells(156, 3).Value = "*"
            Case 21   'Hospitalizacióm
                oWorkSheet.Cells(158, 3).Value = "*"
            Case 23  'Referencia
                oWorkSheet.Cells(161, 3).Value = "*"
            Case 24  'contrareferencia
                oWorkSheet.Cells(161, 49).Value = "*"
            Case 25      'morgue
                oWorkSheet.Cells(160, 3).Value = "*"
            Case 56      'fuga
                oWorkSheet.Cells(162, 3).Value = "*"
            Case 26      'otros
                oWorkSheet.Cells(163, 3).Value = "*"
            Case 65      'citado
                oWorkSheet.Cells(156, 30).Value = "*"
            End Select
        End If
    End If
    '
    Set oRsTmp98 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaListaMedicos(ml_idAtencion, oConexion)
    If oRsTmp98.RecordCount > 1 Then
       iFila = 169
       iPos = 1
       oWorkSheet.Cells(iFila, 20).Font.Bold = True
       oWorkSheet.Cells(iFila, 20).Font.Name = "Arial Narrow"
       oWorkSheet.Cells(iFila, 2).Font.Underline = True
       oWorkSheet.Cells(iFila, 20).Value = "Lista de Médicos que atendieron al Paciente:"
       iFila = iFila + 1
       oRsTmp98.MoveFirst
       Do While Not oRsTmp98.EOF
          oWorkSheet.Cells(iFila, IIf(iPos <= 3, 3, 60)).Value = Left(oRsTmp98!secuencia & ") " & oRsTmp98!ApellidoPaterno & " " & oRsTmp98!ApellidoMaterno & " " & oRsTmp98!Nombres & "_______________________", 45)
          iFila = iFila + 1
          iPos = iPos + 1
          If iPos = 4 Then
             iFila = 117
          End If
          oRsTmp98.MoveNext
       Loop
    End If
    oRsTmp98.Close
    '
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 0
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 100
        PrintArea(0).EndRow = 117
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
            oExcel.Visible = True
             
            
            oWorkSheet.PrintPreview
        Else
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
        'oExcel.ActiveWindow.Close False
        'oExcel.Workbooks.Close
        'oExcel.Quit
        
        oExcel.DisplayAlerts = False
        oExcel.Quit
        
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    Set oRsTmp98 = Nothing
    Set rsReporte = Nothing
    Set oConexion = Nothing
    'mgaray
    Call mo_AdminReportes.GrabarImpresionFichaAtencion(ml_idAtencion)
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


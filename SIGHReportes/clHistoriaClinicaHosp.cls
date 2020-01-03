VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clHistoriaClinicaHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historia clinica
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idCuentaAtencion As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReporteUtil As New ReporteUtil
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property

Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property

Sub CrearReporteHistoriaClinicaDeLaAtencion(lcPlan As String, lcTransferencia As String, _
                     lcDx As String, lcCamaIng As String, lnUsuario As Long, lnHwnd As Long, lcTipoEdad As String)
  Dim rsReporte As New Recordset
  Dim lbEsOpenOffice As Boolean
  Dim lcNombre As String, lcSql As String
  Dim iFila As Long
  Dim oCadena As New Cadena
  
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
        lcArchivoExcel = App.Path + "\Plantillas\HHistoriaClinicaHosp.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HHistoriaClinicaHosp.xls")
        oWorkBookPlantilla.Worksheets("HistoriaClinicaHosp").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
          
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If

  Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
  If mo_AdminReportes.MensajeError <> "" Then
    MsgBox mo_AdminReportes.MensajeError, vbInformation, "Reporte"
  Else
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(84, 7).setFormula("'" & rsReporte!NroHistoriaClinica)
        Call Feuille.getcellbyposition(84, 8).setFormula("'" & Trim(Str(ml_idCuentaAtencion)))
        Call Feuille.getcellbyposition(32, 9).setFormula("'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre)))

        Call Feuille.getcellbyposition(23, 11).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad)
        Call Feuille.getcellbyposition(23, 13).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!EstadoCivil))
        Call Feuille.getcellbyposition(23, 14).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio))
        Call Feuille.getcellbyposition(23, 15).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio))
        Call Feuille.getcellbyposition(23, 16).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio))
        Call Feuille.getcellbyposition(23, 17).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio))
        Call Feuille.getcellbyposition(23, 18).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso))
        Call Feuille.getcellbyposition(23, 19).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso))
        Call Feuille.getcellbyposition(23, 20).setFormula("'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico)))
        Call Feuille.getcellbyposition(23, 21).setFormula(lcTransferencia)

        Call Feuille.getcellbyposition(80, 11).setFormula("'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino"))
        Call Feuille.getcellbyposition(23, 12).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada))
        Call Feuille.getcellbyposition(80, 13).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!telefono))
        Call Feuille.getcellbyposition(80, 14).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento) & " " & mo_ReporteUtil.NullToVacio(rsReporte!TipoDocumento))
        Call Feuille.getcellbyposition(80, 15).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio))
        Call Feuille.getcellbyposition(80, 16).setFormula("'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio))
        Call Feuille.getcellbyposition(80, 17).setFormula(lcCamaIng)
        Call Feuille.getcellbyposition(80, 20).setFormula(lcPlan)
    Else
        
    
        oWorkSheet.Cells(8, 85).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte!NroHistoriaClinica)), True)
        oWorkSheet.Cells(9, 85).Value = "'" & Trim(Str(ml_idCuentaAtencion))
        oWorkSheet.Cells(10, 33).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
        
        oWorkSheet.Cells(12, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad) & " " & lcTipoEdad
        oWorkSheet.Cells(14, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!EstadoCivil)
        oWorkSheet.Cells(15, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio)
        oWorkSheet.Cells(16, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio)
        oWorkSheet.Cells(17, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
        oWorkSheet.Cells(18, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio)
        oWorkSheet.Cells(19, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
        oWorkSheet.Cells(20, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
        oWorkSheet.Cells(21, 24).Value = "'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico))
        oWorkSheet.Cells(22, 24).Value = "'" & lcTransferencia
        '
        oWorkSheet.Cells(12, 81).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
        oWorkSheet.Cells(13, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento) & oCadena.RetornaDescFechaNacimientoCalculada(rsReporte!FNacimientoCalculada)
        oWorkSheet.Cells(14, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!telefono)
        oWorkSheet.Cells(15, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento) & " " & mo_ReporteUtil.NullToVacio(rsReporte!TipoDocumento)
        oWorkSheet.Cells(16, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio)
        oWorkSheet.Cells(17, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio)
        oWorkSheet.Cells(18, 81).Value = "'" & lcCamaIng
        oWorkSheet.Cells(21, 81).Value = "'" & lcPlan
    End If
    '
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(27, 22).setFormula(lcDx)
    Else
        oWorkSheet.Cells(23, 28).Value = "'" & lcDx
    End If
    '
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(3, 52).setFormula("'Usuario: " & oReglasCaja.SeleccionaDatosCajero(lnUsuario, sghApellidosYnombres))
    Else
        oWorkSheet.Cells(53, 4).Value = "'Usuario: " & oReglasCaja.SeleccionaDatosCajero(lnUsuario, sghApellidosYnombres)
    End If
    '
    
    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 101
        PrintArea(0).EndRow = 52
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
          oExcel.Visible = True
          oWorkSheet.PrintPreview
        Else
          oWorkSheet.PageSetup.PrintArea = "$A$1:$DB$54"
          oWorkSheet.PrintOut
        End If
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
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptMN_Cantidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para indicadores de Neonatología
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ReporteUtil As New ReporteUtil
Dim ml_lnAnio As Long
Dim ml_FechaAltaMedica As Boolean
Dim ml_TextoDelFiltro As String
Dim ml_idDepartamento1 As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes


Property Let idDepartamento1(lValue As Long)
  ml_idDepartamento1 = lValue
End Property

Property Let TextoDelFiltro(lValue As String)
  ml_TextoDelFiltro = lValue
End Property

Property Let FechaAltaMedica(lValue As Boolean)
  ml_FechaAltaMedica = lValue
End Property

Property Let Anio(lValue As Long)
  ml_lnAnio = lValue
End Property

Sub EjecutaFormulario()
  Dim oFormulario As New MN_Cantidades
  oFormulario.Show 1
End Sub

Sub CrearReporte(lnHwnd As Long)
  CrearReporteDetallado lnHwnd
End Sub

Sub CrearReporteDetallado(lnHwnd As Long)
  Dim rsReporte As New Recordset, rsTmpReporte As New Recordset
  Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
  Dim iFila As Long, iCol As Integer, lnMesDato As Integer
  Dim lnTSTot As Long, lnTSEne As Long, lnTSFeb As Long, lnTSMar As Long, lnTSAbr As Long, lnTSMay As Long, lnTSJun As Long, lnTSJul As Long, lnTSAgo As Long, lnTSSet As Long, lnTSOct As Long, lnTSNov As Long, lnTSDic As Long, lnTS1Tri As Long, lnTS2Tri As Long, lnTS3Tri As Long, lnTS4Tri As Long
  Dim lnTETot As Long, lnTEEne As Long, lnTEFeb As Long, lnTEMar As Long, lnTEAbr As Long, lnTEMay As Long, lnTEJun As Long, lnTEJul As Long, lnTEAgo As Long, lnTESet As Long, lnTEOct As Long, lnTENov As Long, lnTEDic As Long, lnTE1Tri As Long, lnTE2Tri As Long, lnTE3Tri As Long, lnTE4Tri As Long
  Dim lnTDTot As Long, lnTDEne As Long, lnTDFeb As Long, lnTDMar As Long, lnTDAbr As Long, lnTDMay As Long, lnTDJun As Long, lnTDJul As Long, lnTDAgo As Long, lnTDSet As Long, lnTDOct As Long, lnTDNov As Long, lnTDDic As Long, lnTD1Tri As Long, lnTD2Tri As Long, lnTD3Tri As Long, lnTD4Tri As Long
  Dim lnTTot As Long, lnTEne As Long, lnTFeb As Long, lnTMar As Long, lnTAbr As Long, lnTMay As Long, lnTJun As Long, lnTJul As Long, lnTAgo As Long, lnTSet As Long, lnTOct As Long, lnTNov As Long, lnTDic As Long, lnT1Tri As Long, lnT2Tri As Long, lnT3Tri As Long, lnT4Tri As Long
  Dim lcDpto As String, lcEspecialidad As String, lcServicio As String, lcFiltro As String, Cadena As String
  Dim lnIdDpto As Long, lnIdEspecialidad As Long, lnIdServicio As Long
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
    End If

        'Abre el archivo ExcelOpenOffice
        Cadena = SIGHEntidades.CadenaConexion
        'On Error GoTo ManejadorError
        'Filtra los Datos
        lnTSEne = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 1, Cadena, ml_idDepartamento1)
        lnTSFeb = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 2, Cadena, ml_idDepartamento1)
        lnTSMar = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 3, Cadena, ml_idDepartamento1)
        lnTS1Tri = lnTSEne + lnTSFeb + lnTSMar
        
        lnTSAbr = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 4, Cadena, ml_idDepartamento1)
        lnTSMay = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 5, Cadena, ml_idDepartamento1)
        lnTSJun = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 6, Cadena, ml_idDepartamento1)
        lnTS2Tri = lnTSAbr + lnTSMay + lnTSJun
        
        lnTSJul = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 7, Cadena, ml_idDepartamento1)
        lnTSAgo = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 8, Cadena, ml_idDepartamento1)
        lnTSSet = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 9, Cadena, ml_idDepartamento1)
        lnTS3Tri = lnTSJul + lnTSAgo + lnTSSet
        
        lnTSOct = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 10, Cadena, ml_idDepartamento1)
        lnTSNov = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 11, Cadena, ml_idDepartamento1)
        lnTSDic = mo_ReglasReportes.MuerteFetalSeleccionarPorAnioYMes(ml_lnAnio, 12, Cadena, ml_idDepartamento1)
        lnTS4Tri = lnTSOct + lnTSNov + lnTSDic
        lnTSTot = lnTS1Tri + lnTS2Tri + lnTS3Tri + lnTS4Tri
        
        
        lnTEEne = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 1, Cadena, ml_idDepartamento1)
        lnTEFeb = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 2, Cadena, ml_idDepartamento1)
        lnTEMar = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 3, Cadena, ml_idDepartamento1)
        lnTE1Tri = lnTEEne + lnTEFeb + lnTEMar
        
        lnTEAbr = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 4, Cadena, ml_idDepartamento1)
        lnTEMay = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 5, Cadena, ml_idDepartamento1)
        lnTEJun = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 6, Cadena, ml_idDepartamento1)
        lnTE2Tri = lnTEAbr + lnTEMay + lnTEJun
        
        lnTEJul = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 7, Cadena, ml_idDepartamento1)
        lnTEAgo = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 8, Cadena, ml_idDepartamento1)
        lnTESet = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 9, Cadena, ml_idDepartamento1)
        lnTE3Tri = lnTEJul + lnTEAgo + lnTESet
        
        lnTEOct = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 10, Cadena, ml_idDepartamento1)
        lnTENov = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 11, Cadena, ml_idDepartamento1)
        lnTEDic = mo_ReglasReportes.MuerteNeoNatalSeleccionarPorAnioYMes(ml_lnAnio, 12, Cadena, ml_idDepartamento1)
        lnTE4Tri = lnTEOct + lnTENov + lnTEDic
        lnTETot = lnTE1Tri + lnTE2Tri + lnTE3Tri + lnTE4Tri
        
        
        lnTDEne = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 1, Cadena, ml_idDepartamento1)
        lnTDFeb = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 2, Cadena, ml_idDepartamento1)
        lnTDMar = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 3, Cadena, ml_idDepartamento1)
        lnTD1Tri = lnTDEne + lnTDFeb + lnTDMar
        
        lnTDAbr = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 4, Cadena, ml_idDepartamento1)
        lnTDMay = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 5, Cadena, ml_idDepartamento1)
        lnTDJun = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 6, Cadena, ml_idDepartamento1)
        lnTD2Tri = lnTDAbr + lnTDMay + lnTDJun
        
        lnTDJul = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 7, Cadena, ml_idDepartamento1)
        lnTDAgo = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 8, Cadena, ml_idDepartamento1)
        lnTDSet = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 9, Cadena, ml_idDepartamento1)
        lnTD3Tri = lnTDJul + lnTDAgo + lnTDSet
        
        lnTDOct = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 10, Cadena, ml_idDepartamento1)
        lnTDNov = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 11, Cadena, ml_idDepartamento1)
        lnTDDic = mo_ReglasReportes.NacidosVivosSeleccionarPorAnioYMes(ml_lnAnio, 12, Cadena, ml_idDepartamento1)
        lnTD4Tri = lnTDOct + lnTDNov + lnTDDic
        lnTDTot = lnTD1Tri + lnTD2Tri + lnTD3Tri + lnTD4Tri
        
        lnTEne = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 1, Cadena, ml_idDepartamento1)
        lnTFeb = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 2, Cadena, ml_idDepartamento1)
        lnTMar = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 3, Cadena, ml_idDepartamento1)
        lnT1Tri = lnTEne + lnTFeb + lnTMar
        
        lnTAbr = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 4, Cadena, ml_idDepartamento1)
        lnTMay = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 5, Cadena, ml_idDepartamento1)
        lnTJun = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 6, Cadena, ml_idDepartamento1)
        lnT2Tri = lnTAbr + lnTMay + lnTJun
        
        lnTJul = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 7, Cadena, ml_idDepartamento1)
        lnTAgo = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 8, Cadena, ml_idDepartamento1)
        lnTSet = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 9, Cadena, ml_idDepartamento1)
        lnT3Tri = lnTJul + lnTAgo + lnTSet
        
        lnTOct = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 10, Cadena, ml_idDepartamento1)
        lnTNov = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 11, Cadena, ml_idDepartamento1)
        lnTDic = mo_ReglasReportes.RecienNacidosVivosHospitalizadosSeleccionarPorAnioYMes(ml_lnAnio, 12, Cadena, ml_idDepartamento1)
        lnT4Tri = lnTOct + lnTNov + lnTDic
        lnTTot = lnT1Tri + lnT2Tri + lnT3Tri + lnT4Tri
    
    If lbEsOpenOffice = True Then
        lcArchivoExcel = App.Path + "\Plantillas\MN_CantidadMortalidad.ods"
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
        Call Feuille.getcellbyposition(1, 2).setFormula(ml_TextoDelFiltro)
        iFila = 7
        iCol = 3
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\MN_CantidadMortalidad.xls")
        oWorkBookPlantilla.Worksheets("Mortalidad").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        'Inicio de Impresion
        oWorkSheet.Cells(3, 2).Value = ml_TextoDelFiltro
        iFila = 7
        iCol = 3
    End If

    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTSEne)
        Call Feuille.getcellbyposition(iCol, iFila - 1).setFormula(lnTSFeb)
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTSMar)
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTS1Tri)
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTSAbr)
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTSMay)
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTSJun)
        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTS2Tri)
        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTSJul)
        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTSAgo)
        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTSSet)
        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTS3Tri)
        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTSOct)
        Call Feuille.getcellbyposition(iCol + 12, iFila - 1).setFormula(lnTSNov)
        Call Feuille.getcellbyposition(iCol + 13, iFila - 1).setFormula(lnTSDic)
        Call Feuille.getcellbyposition(iCol + 14, iFila - 1).setFormula(lnTS4Tri)
        Call Feuille.getcellbyposition(iCol + 15, iFila - 1).setFormula(lnTSTot)
        iFila = iFila + 1

        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTEEne)
        Call Feuille.getcellbyposition(iCol, iFila - 1).setFormula(lnTEFeb)
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTEMar)
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTE1Tri)
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTEAbr)
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTEMay)
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTEJun)
        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTE2Tri)
        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTEJul)
        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTEAgo)
        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTESet)
        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTE3Tri)
        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTEOct)
        Call Feuille.getcellbyposition(iCol + 12, iFila - 1).setFormula(lnTENov)
        Call Feuille.getcellbyposition(iCol + 13, iFila - 1).setFormula(lnTEDic)
        Call Feuille.getcellbyposition(iCol + 14, iFila - 1).setFormula(lnTE4Tri)
        Call Feuille.getcellbyposition(iCol + 15, iFila - 1).setFormula(lnTETot)
        iFila = iFila + 1
        
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTDEne)
        Call Feuille.getcellbyposition(iCol, iFila - 1).setFormula(lnTDFeb)
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTDMar)
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTD1Tri)
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTDAbr)
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTDMay)
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTDJun)
        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTD2Tri)
        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTDJul)
        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTDAgo)
        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTDSet)
        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTD3Tri)
        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTDOct)
        Call Feuille.getcellbyposition(iCol + 12, iFila - 1).setFormula(lnTDNov)
        Call Feuille.getcellbyposition(iCol + 13, iFila - 1).setFormula(lnTDDic)
        Call Feuille.getcellbyposition(iCol + 14, iFila - 1).setFormula(lnTD4Tri)
        Call Feuille.getcellbyposition(iCol + 15, iFila - 1).setFormula(lnTDTot)
        iFila = iFila + 1
        
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTEne)
        Call Feuille.getcellbyposition(iCol, iFila - 1).setFormula(lnTFeb)
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTMar)
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnT1Tri)
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTAbr)
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTMay)
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTJun)
        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnT2Tri)
        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTJul)
        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTAgo)
        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTSet)
        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnT3Tri)
        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTSOct)
        Call Feuille.getcellbyposition(iCol + 12, iFila - 1).setFormula(lnTNov)
        Call Feuille.getcellbyposition(iCol + 13, iFila - 1).setFormula(lnTDic)
        Call Feuille.getcellbyposition(iCol + 14, iFila - 1).setFormula(lnT4Tri)
        Call Feuille.getcellbyposition(iCol + 15, iFila - 1).setFormula(lnTTot)
        iFila = iFila + 1
    Else

        oWorkSheet.Cells(iFila, iCol).Value = lnTSEne
        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTSFeb
        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTSMar
        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTS1Tri
        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTSAbr
        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTSMay
        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTSJun
        oWorkSheet.Cells(iFila, iCol + 7).Value = lnTS2Tri
        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTSJul
        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTSAgo
        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTSSet
        oWorkSheet.Cells(iFila, iCol + 11).Value = lnTS3Tri
        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTSOct
        oWorkSheet.Cells(iFila, iCol + 13).Value = lnTSNov
        oWorkSheet.Cells(iFila, iCol + 14).Value = lnTSDic
        oWorkSheet.Cells(iFila, iCol + 15).Value = lnTS4Tri
        oWorkSheet.Cells(iFila, iCol + 16).Value = lnTSTot
        iFila = iFila + 1
        
        oWorkSheet.Cells(iFila, iCol).Value = lnTEEne
        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTEFeb
        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTEMar
        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTE1Tri
        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTEAbr
        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTEMay
        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTEJun
        oWorkSheet.Cells(iFila, iCol + 7).Value = lnTE2Tri
        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTEJul
        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTEAgo
        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTESet
        oWorkSheet.Cells(iFila, iCol + 11).Value = lnTE3Tri
        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTEOct
        oWorkSheet.Cells(iFila, iCol + 13).Value = lnTENov
        oWorkSheet.Cells(iFila, iCol + 14).Value = lnTEDic
        oWorkSheet.Cells(iFila, iCol + 15).Value = lnTE4Tri
        oWorkSheet.Cells(iFila, iCol + 16).Value = lnTETot
        iFila = iFila + 1
        
        oWorkSheet.Cells(iFila, iCol).Value = lnTDEne
        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTDFeb
        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTDMar
        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTD1Tri
        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTDAbr
        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTDMay
        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTDJun
        oWorkSheet.Cells(iFila, iCol + 7).Value = lnTD2Tri
        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTDJul
        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTDAgo
        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTDSet
        oWorkSheet.Cells(iFila, iCol + 11).Value = lnTD3Tri
        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTDOct
        oWorkSheet.Cells(iFila, iCol + 13).Value = lnTDNov
        oWorkSheet.Cells(iFila, iCol + 14).Value = lnTDDic
        oWorkSheet.Cells(iFila, iCol + 15).Value = lnTD4Tri
        oWorkSheet.Cells(iFila, iCol + 16).Value = lnTDTot
        iFila = iFila + 1
        
        oWorkSheet.Cells(iFila, iCol).Value = lnTEne
        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTFeb
        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTMar
        oWorkSheet.Cells(iFila, iCol + 3).Value = lnT1Tri
        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTAbr
        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTMay
        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTJun
        oWorkSheet.Cells(iFila, iCol + 7).Value = lnT2Tri
        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTJul
        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTAgo
        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTSet
        oWorkSheet.Cells(iFila, iCol + 11).Value = lnT3Tri
        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTOct
        oWorkSheet.Cells(iFila, iCol + 13).Value = lnTNov
        oWorkSheet.Cells(iFila, iCol + 14).Value = lnTDic
        oWorkSheet.Cells(iFila, iCol + 15).Value = lnT4Tri
        oWorkSheet.Cells(iFila, iCol + 16).Value = lnTTot
        iFila = iFila + 1
    End If

    If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 19
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
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
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        Set oWorkSheet = Nothing
        Set oExcel = Nothing
    End If
    Set mo_ReglasReportes = Nothing
  Exit Sub
  
ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
      MsgBox Err.Description
  End Select
End Sub

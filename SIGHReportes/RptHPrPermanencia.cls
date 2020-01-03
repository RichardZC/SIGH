VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHPrPermanencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Promedio de permanencia
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_lnAnio As Long
Dim ml_FechaAltaMedica As Boolean
Dim ml_TipoReporte As Integer
Dim ml_Titulo As String
Dim ml_TextoDelFiltro As String
Dim ml_idDepartamento1 As Long
Dim ml_idEspecialidad1 As Long
Dim ml_idServicio1 As Long
Dim ml_idDepartamento2 As Long
Dim ml_idEspecialidad2 As Long
Dim ml_idServicio2 As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let idDepartamento1(lValue As Long)
    ml_idDepartamento1 = lValue
End Property
Property Let idServicio1(lValue As Long)
    ml_idServicio1 = lValue
End Property
Property Let idEspecialidad1(lValue As Long)
    ml_idEspecialidad1 = lValue
End Property
Property Let idDepartamento2(lValue As Long)
    ml_idDepartamento2 = lValue
End Property
Property Let idServicio2(lValue As Long)
    ml_idServicio2 = lValue
End Property
Property Let idEspecialidad2(lValue As Long)
    ml_idEspecialidad2 = lValue
End Property
Property Let Titulo(lValue As String)
    ml_Titulo = lValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let TipoReporte(lValue As Integer)
    ml_TipoReporte = lValue
End Property
Property Let FechaAltaMedica(lValue As Boolean)
    ml_FechaAltaMedica = lValue
End Property
Property Let Anio(lValue As Long)
    ml_lnAnio = lValue
End Property


Sub EjecutaFormulario()
    Dim oFormulario As New HPrPermanencia
    oFormulario.Show 1
End Sub

Sub CrearReporte(lnHwnd As Long)
    Select Case ml_TipoReporte
    Case 0
         CrearReporteDetallado lnHwnd
    Case 1
         CrearReportePorDepartamentos lnHwnd
    Case 2
         CrearReportePorServicios lnHwnd
    Case 3
         CrearReportePorDosEspecialidades lnHwnd
    End Select
End Sub

Sub CrearReporteDetallado(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsTmpReporte As New Recordset
Dim iFila As Long: Dim iCol As Integer

Dim lnTSTot As Long: Dim lnTSEne As Long: Dim lnTSFeb As Long: Dim lnTSMar As Long: Dim lnTSAbr As Long: Dim lnTSMay As Long: Dim lnTSJun As Long: Dim lnTSJul As Long: Dim lnTSAgo As Long: Dim lnTSSep As Long: Dim lnTSOct As Long: Dim lnTSNov As Long: Dim lnTSDic As Long
Dim lnTETot As Long: Dim lnTEEne As Long: Dim lnTEFeb As Long: Dim lnTEMar As Long: Dim lnTEAbr As Long: Dim lnTEMay As Long: Dim lnTEJun As Long: Dim lnTEJul As Long: Dim lnTEAgo As Long: Dim lnTESep As Long: Dim lnTEOct As Long: Dim lnTENov As Long: Dim lnTEDic As Long
Dim lnTDTot As Long: Dim lnTDEne As Long: Dim lnTDFeb As Long: Dim lnTDMar As Long: Dim lnTDAbr As Long: Dim lnTDMay As Long: Dim lnTDJun As Long: Dim lnTDJul As Long: Dim lnTDAgo As Long: Dim lnTDSep As Long: Dim lnTDOct As Long: Dim lnTDNov As Long: Dim lnTDDic As Long
Dim lnTTot As Long: Dim lnTEne As Long: Dim lnTFeb As Long: Dim lnTMar As Long: Dim lnTAbr As Long: Dim lnTMay As Long: Dim lnTJun As Long: Dim lnTJul As Long: Dim lnTAgo As Long: Dim lnTSep As Long: Dim lnTOct As Long: Dim lnTNov As Long: Dim lnTDic As Long

Dim lnCTSTot As Double: Dim lnCTSEne As Long: Dim lnCTSFeb As Long: Dim lnCTSMar As Long: Dim lnCTSAbr As Long: Dim lnCTSMay As Long: Dim lnCTSJun As Long: Dim lnCTSJul As Long: Dim lnCTSAgo As Long: Dim lnCTSSep As Long: Dim lnCTSOct As Long: Dim lnCTSNov As Long: Dim lnCTSDic As Long
Dim lnCTETot As Double: Dim lnCTEEne As Long: Dim lnCTEFeb As Long: Dim lnCTEMar As Long: Dim lnCTEAbr As Long: Dim lnCTEMay As Long: Dim lnCTEJun As Long: Dim lnCTEJul As Long: Dim lnCTEAgo As Long: Dim lnCTESep As Long: Dim lnCTEOct As Long: Dim lnCTENov As Long: Dim lnCTEDic As Long
Dim lnCTDTot As Double: Dim lnCTDEne As Long: Dim lnCTDFeb As Long: Dim lnCTDMar As Long: Dim lnCTDAbr As Long: Dim lnCTDMay As Long: Dim lnCTDJun As Long: Dim lnCTDJul As Long: Dim lnCTDAgo As Long: Dim lnCTDSep As Long: Dim lnCTDOct As Long: Dim lnCTDNov As Long: Dim lnCTDDic As Long
Dim lnCTTot As Double:  Dim lnCTEne As Long: Dim lnCTFeb As Long: Dim lnCTMar As Long: Dim lnCTAbr As Long: Dim lnCTMay As Long: Dim lnCTJun As Long: Dim lnCTJul As Long: Dim lnCTAgo As Long: Dim lnCTSep As Long: Dim lnCTOct As Long: Dim lnCTNov As Long: Dim lnCTDic As Long
Dim lnCCTEEne As Double: Dim lnCCTEFeb As Double: Dim lnCCTEMar As Double: Dim lnCCTEAbr As Double: Dim lnCCTEMay As Double: Dim lnCCTEJun As Double: Dim lnCCTEJul As Double: Dim lnCCTEAgo As Double: Dim lnCCTESep As Double: Dim lnCCTEOct As Double: Dim lnCCTENov As Double: Dim lnCCTEDic As Double
Dim lnCCTDEne As Double: Dim lnCCTDFeb As Double: Dim lnCCTDMar As Double: Dim lnCCTDAbr As Double: Dim lnCCTDMay As Double: Dim lnCCTDJun As Double: Dim lnCCTDJul As Double: Dim lnCCTDAgo As Double: Dim lnCCTDSep As Double: Dim lnCCTDOct As Double: Dim lnCCTDNov As Double: Dim lnCCTDDic As Double
Dim lnCCTEne As Double: Dim lnCCTFeb As Double: Dim lnCCTMar As Double: Dim lnCCTAbr As Double: Dim lnCCTMay As Double: Dim lnCCTJun As Double: Dim lnCCTJul As Double: Dim lnCCTAgo As Double: Dim lnCCTSep As Double: Dim lnCCTOct As Double: Dim lnCCTNov As Double: Dim lnCCTDic As Double

Dim lcDpto As String: Dim lnIdDpto As Long
Dim lcEspecialidad As String: Dim lnIdEspecialidad As Long
Dim lcServicio As String: Dim lnIdServicio As Long
Dim lnMesDato As Integer: Dim lnDiasEstancia As Integer
Dim lcFiltro As String
Dim lcHoraEgreso As String: Dim ldFechaEgreso As Date
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
    
    'Filtra los Datos
    Set rsReporte = AtencionesSeleccionarPorAnio(ml_FechaAltaMedica, ml_lnAnio)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HEgresoHosp.ods"
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
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgresoHosp.xls")
            oWorkBookPlantilla.Worksheets("HEgresoHosp").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 1).setFormula(ml_Titulo)
            Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
        Else
            'Inicio de Impresion
            oWorkSheet.Cells(2, 3).Value = ml_Titulo
            oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
        End If
            iFila = 7
            iCol = 5
            lnTTot = 0: lnTEne = 0: lnTFeb = 0: lnTMar = 0: lnTAbr = 0: lnTMay = 0: lnTJun = 0: lnTJul = 0: lnTAgo = 0: lnTSep = 0: lnTOct = 0: lnTNov = 0: lnTDic = 0
            lnCTTot = 0: lnCTEne = 0: lnCTFeb = 0: lnCTMar = 0: lnCTAbr = 0: lnCTMay = 0: lnCTJun = 0: lnCTJul = 0: lnCTAgo = 0: lnCTSep = 0: lnCTOct = 0: lnCTNov = 0: lnCTDic = 0
            'lncCTTot = 0
            lnCCTEne = 0: lnCCTFeb = 0: lnCCTMar = 0: lnCCTAbr = 0: lnCCTMay = 0: lnCCTJun = 0: lnCCTJul = 0: lnCCTAgo = 0: lnCCTSep = 0: lnCCTOct = 0: lnCCTNov = 0: lnCCTDic = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                lnCTDTot = 0: lnCTDEne = 0: lnCTDFeb = 0: lnCTDMar = 0: lnCTDAbr = 0: lnCTDMay = 0: lnCTDJun = 0: lnCTDJul = 0: lnCTDAgo = 0: lnCTDSep = 0: lnCTDOct = 0: lnCTDNov = 0: lnCTDDic = 0
                lnCCTDEne = 0: lnCCTDFeb = 0: lnCCTDMar = 0: lnCCTDAbr = 0: lnCCTDMay = 0: lnCCTDJun = 0: lnCCTDJul = 0: lnCCTDAgo = 0: lnCCTDSep = 0: lnCCTDOct = 0: lnCCTDNov = 0: lnCCTDDic = 0
                lcDpto = rsReporte.Fields("dDpto").Value
                lnIdDpto = rsReporte.Fields("CodDpto").Value
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("dDpto").Value)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                End If
                iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lnCTETot = 0: lnCTEEne = 0: lnCTEFeb = 0: lnCTEMar = 0: lnCTEAbr = 0: lnCTEMay = 0: lnCTEJun = 0: lnCTEJul = 0: lnCTEAgo = 0: lnCTESep = 0: lnCTEOct = 0: lnCTENov = 0: lnCTEDic = 0
                   lnCCTEEne = 0: lnCCTEFeb = 0: lnCCTEMar = 0: lnCCTEAbr = 0: lnCCTEMay = 0: lnCCTEJun = 0: lnCCTEJul = 0: lnCCTEAgo = 0: lnCCTESep = 0: lnCCTEOct = 0: lnCCTENov = 0: lnCCTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("dEspecialidad").Value)
                    Else
                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                    End If
                   iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
                        lnCTSTot = 0: lnCTSEne = 0: lnCTSFeb = 0: lnCTSMar = 0: lnCTSAbr = 0: lnCTSMay = 0: lnCTSJun = 0: lnCTSJul = 0: lnCTSAgo = 0: lnCTSSep = 0: lnCTSOct = 0: lnCTSNov = 0: lnCTSDic = 0
                        lcServicio = rsReporte.Fields("dServicio").Value
                        lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value)
                        Else
                            oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value
                        End If
                        Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value And lcServicio = rsReporte.Fields("dServicio").Value And lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                            lcHoraEgreso = ""
                            If ml_FechaAltaMedica Then
                               lnMesDato = Month(rsReporte.Fields("FechaEgreso").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgreso").Value
                               If Not IsNull(rsReporte.Fields("HoraEgreso").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgreso").Value
                               End If
                            Else
                               lnMesDato = Month(rsReporte.Fields("FechaEgresoAdministrativo").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgresoAdministrativo").Value
                               If Not IsNull(rsReporte.Fields("HoraEgresoAdministrativo").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgresoAdministrativo").Value
                               End If
                            End If
                            
'                            lnDiasEstancia = ldFechaEgreso - rsReporte.Fields("FechaIngreso").Value
'                            If lnDiasEstancia = 0 Then lnDiasEstancia = 1
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte.Fields("FechaIngreso").Value, rsReporte.Fields("HoraIngreso").Value, ldFechaEgreso, lcHoraEgreso)
                            
                            Select Case lnMesDato
                            Case 1
                                 lnTSEne = lnTSEne + lnDiasEstancia
                                 lnTEEne = lnTEEne + lnDiasEstancia
                                 lnTDEne = lnTDEne + lnDiasEstancia
                                 lnTEne = lnTEne + lnDiasEstancia
                                 
                                 lnCTSEne = lnCTSEne + 1
                                 lnCTEEne = lnCTEEne + 1
                                 lnCTDEne = lnCTDEne + 1
                                 lnCTEne = lnCTEne + 1
                                 
                            Case 2
                                 lnTSFeb = lnTSFeb + lnDiasEstancia
                                 lnTEFeb = lnTEFeb + lnDiasEstancia
                                 lnTDFeb = lnTDFeb + lnDiasEstancia
                                 lnTFeb = lnTFeb + lnDiasEstancia
                            
                            
                                 lnCTSFeb = lnCTSFeb + 1
                                 lnCTEFeb = lnCTEFeb + 1
                                 lnCTDFeb = lnCTDFeb + 1
                                 lnCTFeb = lnCTFeb + 1
                            Case 3
                                 lnTSMar = lnTSMar + lnDiasEstancia
                                 lnTEMar = lnTEMar + lnDiasEstancia
                                 lnTDMar = lnTDMar + lnDiasEstancia
                                 lnTMar = lnTMar + lnDiasEstancia
                            
                            
                                 lnCTSMar = lnCTSMar + 1
                                 lnCTEMar = lnCTEMar + 1
                                 lnCTDMar = lnCTDMar + 1
                                 lnCTMar = lnCTMar + 1
                            Case 4
                                 lnTSAbr = lnTSAbr + lnDiasEstancia
                                 lnTEAbr = lnTEAbr + lnDiasEstancia
                                 lnTDAbr = lnTDAbr + lnDiasEstancia
                                 lnTAbr = lnTAbr + lnDiasEstancia
                            
                            
                                 lnCTSAbr = lnCTSAbr + 1
                                 lnCTEAbr = lnCTEAbr + 1
                                 lnCTDAbr = lnCTDAbr + 1
                                 lnCTAbr = lnCTAbr + 1
                            Case 5
                                 lnTSMay = lnTSMay + lnDiasEstancia
                                 lnTEMay = lnTEMay + lnDiasEstancia
                                 lnTDMay = lnTDMay + lnDiasEstancia
                                 lnTMay = lnTMay + lnDiasEstancia
                            
                            
                                 lnCTSMay = lnCTSMay + 1
                                 lnCTEMay = lnCTEMay + 1
                                 lnCTDMay = lnCTDMay + 1
                                 lnCTMay = lnCTMay + 1
                            Case 6
                                 lnTSJun = lnTSJun + lnDiasEstancia
                                 lnTEJun = lnTEJun + lnDiasEstancia
                                 lnTDJun = lnTDJun + lnDiasEstancia
                                 lnTJun = lnTJun + lnDiasEstancia
                            
                                 lnCTSJun = lnCTSJun + 1
                                 lnCTEJun = lnCTEJun + 1
                                 lnCTDJun = lnCTDJun + 1
                                 lnCTJun = lnCTJun + 1
                            Case 7
                                 lnTSJul = lnTSJul + lnDiasEstancia
                                 lnTEJul = lnTEJul + lnDiasEstancia
                                 lnTDJul = lnTDJul + lnDiasEstancia
                                 lnTJul = lnTJul + lnDiasEstancia
                            
                            
                                 lnCTSJul = lnCTSJul + 1
                                 lnCTEJul = lnCTEJul + 1
                                 lnCTDJul = lnCTDJul + 1
                                 lnCTJul = lnCTJul + 1
                            Case 8
                                 lnTSAgo = lnTSAgo + lnDiasEstancia
                                 lnTEAgo = lnTEAgo + lnDiasEstancia
                                 lnTDAgo = lnTDAgo + lnDiasEstancia
                                 lnTAgo = lnTAgo + lnDiasEstancia
                            
                            
                                 lnCTSAgo = lnCTSAgo + 1
                                 lnCTEAgo = lnCTEAgo + 1
                                 lnCTDAgo = lnCTDAgo + 1
                                 lnCTAgo = lnCTAgo + 1
                            Case 9
                                 lnTSSep = lnTSSep + lnDiasEstancia
                                 lnTESep = lnTESep + lnDiasEstancia
                                 lnTDSep = lnTDSep + lnDiasEstancia
                                 lnTSep = lnTSep + lnDiasEstancia
                            
                            
                                 lnCTSSep = lnCTSSep + 1
                                 lnCTESep = lnCTESep + 1
                                 lnCTDSep = lnCTDSep + 1
                                 lnCTSep = lnCTSep + 1
                            Case 10
                                 lnTSOct = lnTSOct + lnDiasEstancia
                                 lnTEOct = lnTEOct + lnDiasEstancia
                                 lnTDOct = lnTDOct + lnDiasEstancia
                                 lnTOct = lnTOct + lnDiasEstancia
                            
                            
                                 lnCTSOct = lnCTSOct + 1
                                 lnCTEOct = lnCTEOct + 1
                                 lnCTDOct = lnCTDOct + 1
                                 lnCTOct = lnCTOct + 1
                            Case 11
                                 lnTSNov = lnTSNov + lnDiasEstancia
                                 lnTENov = lnTENov + lnDiasEstancia
                                 lnTDNov = lnTDNov + lnDiasEstancia
                                 lnTNov = lnTNov + lnDiasEstancia
                            
                            
                                 lnCTSNov = lnCTSNov + 1
                                 lnCTENov = lnCTENov + 1
                                 lnCTDNov = lnCTDNov + 1
                                 lnCTNov = lnCTNov + 1
                            Case 12
                                 lnTSDic = lnTSDic + lnDiasEstancia
                                 lnTEDic = lnTEDic + lnDiasEstancia
                                 lnTDDic = lnTDDic + lnDiasEstancia
                                 lnTDic = lnTDic + lnDiasEstancia
                            
                            
                                 lnCTSDic = lnCTSDic + 1
                                 lnCTEDic = lnCTEDic + 1
                                 lnCTDDic = lnCTDDic + 1
                                 lnCTDic = lnCTDic + 1
                            End Select
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        
                        
                        If lnCTSEne = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 1).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(Round(lnTSEne / lnCTSEne, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 1).Value = Round(lnTSEne / lnCTSEne, 2)
                            End If
                           lnCTSTot = lnCTSTot + Round(lnTSEne / lnCTSEne, 2)
                           lnCCTEEne = lnCCTEEne + Round(lnTSEne / lnCTSEne, 2)
                           lnCCTDEne = lnCCTDEne + Round(lnTSEne / lnCTSEne, 2)
                           lnCCTEne = lnCCTEne + Round(lnTSEne / lnCTSEne, 2)
                        End If
                        If lnCTSFeb = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 2).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(Round(lnTSFeb / lnCTSFeb, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 2).Value = Round(lnTSFeb / lnCTSFeb, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSFeb / lnCTSFeb, 2)
                            lnCCTEFeb = lnCCTEFeb + Round(lnTSFeb / lnCTSFeb, 2)
                            lnCCTDFeb = lnCCTDFeb + Round(lnTSFeb / lnCTSFeb, 2)
                            lnCCTFeb = lnCCTFeb + Round(lnTSFeb / lnCTSFeb, 2)
                        End If
                        If lnCTSMar = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 3).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(Round(lnTSMar / lnCTSMar, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 3).Value = Round(lnTSMar / lnCTSMar, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSMar / lnCTSMar, 2)
                            lnCCTEMar = lnCCTEMar + Round(lnTSMar / lnCTSMar, 2)
                            lnCCTDMar = lnCCTDMar + Round(lnTSMar / lnCTSMar, 2)
                            lnCCTMar = lnCCTMar + Round(lnTSMar / lnCTSMar, 2)
                        End If
                        If lnCTSAbr = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 4).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(Round(lnTSAbr / lnCTSAbr, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 4).Value = Round(lnTSAbr / lnCTSAbr, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSAbr / lnCTSAbr, 2)
                            lnCCTEAbr = lnCCTEAbr + Round(lnTSAbr / lnCTSAbr, 2)
                            lnCCTDAbr = lnCCTDAbr + Round(lnTSAbr / lnCTSAbr, 2)
                            lnCCTAbr = lnCCTAbr + Round(lnTSAbr / lnCTSAbr, 2)
                        End If
                        If lnCTSMay = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round(lnTSMay / lnCTSMay, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 5).Value = Round(lnTSMay / lnCTSMay, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSMay / lnCTSMay, 2)
                            lnCCTEMay = lnCCTEMay + Round(lnTSMay / lnCTSMay, 2)
                            lnCCTDMay = lnCCTDMay + Round(lnTSMay / lnCTSMay, 2)
                            lnCCTMay = lnCCTMay + Round(lnTSMay / lnCTSMay, 2)
                        End If
                        If lnCTSJun = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(lnTSJun / lnCTSJun, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 6).Value = Round(lnTSJun / lnCTSJun, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSJun / lnCTSJun, 2)
                            lnCCTEJun = lnCCTEJun + Round(lnTSJun / lnCTSJun, 2)
                            lnCCTDJun = lnCCTDJun + Round(lnTSJun / lnCTSJun, 2)
                            lnCCTJun = lnCCTJun + Round(lnTSJun / lnCTSJun, 2)
                           
                        End If
                        If lnCTSJul = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round(lnTSJul / lnCTSJul, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 7).Value = Round(lnTSJul / lnCTSJul, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSJul / lnCTSJul, 2)
                            lnCCTEJul = lnCCTEJul + Round(lnTSJul / lnCTSJul, 2)
                            lnCCTDJul = lnCCTDJul + Round(lnTSJul / lnCTSJul, 2)
                            lnCCTJul = lnCCTJul + Round(lnTSJul / lnCTSJul, 2)
                        End If
                        If lnCTSAgo = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(lnTSAgo / lnCTSAgo, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 8).Value = Round(lnTSAgo / lnCTSAgo, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSAgo / lnCTSAgo, 2)
                            lnCCTEAgo = lnCCTEAgo + Round(lnTSAgo / lnCTSAgo, 2)
                            lnCCTDAgo = lnCCTDAgo + Round(lnTSAgo / lnCTSAgo, 2)
                            lnCCTAgo = lnCCTAgo + Round(lnTSAgo / lnCTSAgo, 2)
                           
                        End If
                        If lnCTSSep = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 9).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(Round(lnTSSep / lnCTSSep, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 9).Value = Round(lnTSSep / lnCTSSep, 2)
                            End If
                           lnCTSTot = lnCTSTot + Round(lnTSSep / lnCTSSep, 2)
                           lnCCTESep = lnCCTESep + Round(lnTSSep / lnCTSSep, 2)
                           lnCCTDSep = lnCCTDSep + Round(lnTSSep / lnCTSSep, 2)
                           lnCCTSep = lnCCTSep + Round(lnTSSep / lnCTSSep, 2)
                        End If
                        If lnCTSOct = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 10).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(Round(lnTSOct / lnCTSOct, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 10).Value = Round(lnTSOct / lnCTSOct, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSOct / lnCTSOct, 2)
                            lnCCTEOct = lnCCTEOct + Round(lnTSOct / lnCTSOct, 2)
                            lnCCTDOct = lnCCTDOct + Round(lnTSOct / lnCTSOct, 2)
                            lnCCTOct = lnCCTOct + Round(lnTSOct / lnCTSOct, 2)
                        
                        End If
                        If lnCTSNov = 0 Then
                           If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(0)
                            Else
                                 oWorkSheet.Cells(iFila, iCol + 11).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(Round(lnTSNov / lnCTSNov, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 11).Value = Round(lnTSNov / lnCTSNov, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSNov / lnCTSNov, 2)
                            lnCCTENov = lnCCTENov + Round(lnTSNov / lnCTSNov, 2)
                            lnCCTDNov = lnCCTDNov + Round(lnTSNov / lnCTSNov, 2)
                            lnCCTNov = lnCCTNov + Round(lnTSNov / lnCTSNov, 2)
                           
                        End If
                        If lnCTSDic = 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 12).Value = 0
                            End If
                        Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(Round(lnTSDic / lnCTSDic, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 12).Value = Round(lnTSDic / lnCTSDic, 2)
                            End If
                            lnCTSTot = lnCTSTot + Round(lnTSDic / lnCTSDic, 2)
                            lnCCTEDic = lnCCTEDic + Round(lnTSDic / lnCTSDic, 2)
                            lnCCTDDic = lnCCTDDic + Round(lnTSDic / lnCTSDic, 2)
                            lnCCTDic = lnCCTDic + Round(lnTSDic / lnCTSDic, 2)
                        
                        End If
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTSTot)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTSTot
                        End If
                        lnCTETot = lnCTETot + lnCTSTot
                        lnCTDTot = lnCTDTot + lnCTSTot
                        lnCTTot = lnCTTot + lnCTSTot
                        
                        iFila = iFila + 1
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop
                    If lbEsOpenOffice = True Then
                        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":Q" & CStr(iFila))
                        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTETot)
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTEEne)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTEFeb)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTEMar)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTEAbr)
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTEMay)
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTEJun)
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTEJul)
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTEAgo)
                        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTESep)
                        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTEOct)
                        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTENov)
                        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTEDic)
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 4, iFila, iCol + 12
                        oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTETot
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTEEne
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTEFeb
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTEMar
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTEAbr
                        oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTEMay
                        oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTEJun
                        oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTEJul
                        oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTEAgo
                        oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTESep
                        oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTEOct
                        oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTENov
                        oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTEDic
                    End If
                   iFila = iFila + 1
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
                iFila = iFila + 1
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":Q" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTDTot)
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTDEne)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTDFeb)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTDMar)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTDAbr)
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTDMay)
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTDJun)
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTDJul)
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTDAgo)
                        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTDSep)
                        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTDOct)
                        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTDNov)
                        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTDDic)
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 12
                        oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTDTot
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTDEne
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTDFeb
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTDMar
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTDAbr
                        oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTDMay
                        oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTDJun
                        oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTDJul
                        oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTDAgo
                        oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTDSep
                        oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTDOct
                        oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTDNov
                        oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTDDic
                    End If
                iFila = iFila + 1
                If rsReporte.EOF Then
                   Exit Do
                End If
             Loop
             iFila = iFila + 1
                If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Q" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTTot)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTEne)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTFeb)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTMar)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTAbr)
                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTMay)
                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTJun)
                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTJul)
                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTAgo)
                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTSep)
                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTOct)
                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTNov)
                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTDic)
                Else
                     mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 12
                     oWorkSheet.Cells(iFila, 2).Value = "Total: "
                     oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTTot
                     oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTEne
                     oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTFeb
                     oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTMar
                     oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTAbr
                     oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTMay
                     oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTJun
                     oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTJul
                     oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTAgo
                     oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTSep
                     oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTOct
                     oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTNov
                     oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTDic
                End If
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 17
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
End Sub

Sub CrearReportePorDepartamentos(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsTmpReporte As New Recordset
Dim iFila As Long: Dim iCol As Integer

Dim lnTSTot As Long: Dim lnTSEne As Long: Dim lnTSFeb As Long: Dim lnTSMar As Long: Dim lnTSAbr As Long: Dim lnTSMay As Long: Dim lnTSJun As Long: Dim lnTSJul As Long: Dim lnTSAgo As Long: Dim lnTSSep As Long: Dim lnTSOct As Long: Dim lnTSNov As Long: Dim lnTSDic As Long
Dim lnTETot As Long: Dim lnTEEne As Long: Dim lnTEFeb As Long: Dim lnTEMar As Long: Dim lnTEAbr As Long: Dim lnTEMay As Long: Dim lnTEJun As Long: Dim lnTEJul As Long: Dim lnTEAgo As Long: Dim lnTESep As Long: Dim lnTEOct As Long: Dim lnTENov As Long: Dim lnTEDic As Long
Dim lnTDTot As Long: Dim lnTDEne As Long: Dim lnTDFeb As Long: Dim lnTDMar As Long: Dim lnTDAbr As Long: Dim lnTDMay As Long: Dim lnTDJun As Long: Dim lnTDJul As Long: Dim lnTDAgo As Long: Dim lnTDSep As Long: Dim lnTDOct As Long: Dim lnTDNov As Long: Dim lnTDDic As Long
Dim lnTTot As Long: Dim lnTEne As Long: Dim lnTFeb As Long: Dim lnTMar As Long: Dim lnTAbr As Long: Dim lnTMay As Long: Dim lnTJun As Long: Dim lnTJul As Long: Dim lnTAgo As Long: Dim lnTSep As Long: Dim lnTOct As Long: Dim lnTNov As Long: Dim lnTDic As Long

Dim lnCTSTot As Double: Dim lnCTSEne As Long: Dim lnCTSFeb As Long: Dim lnCTSMar As Long: Dim lnCTSAbr As Long: Dim lnCTSMay As Long: Dim lnCTSJun As Long: Dim lnCTSJul As Long: Dim lnCTSAgo As Long: Dim lnCTSSep As Long: Dim lnCTSOct As Long: Dim lnCTSNov As Long: Dim lnCTSDic As Long
Dim lnCTETot As Double: Dim lnCTEEne As Long: Dim lnCTEFeb As Long: Dim lnCTEMar As Long: Dim lnCTEAbr As Long: Dim lnCTEMay As Long: Dim lnCTEJun As Long: Dim lnCTEJul As Long: Dim lnCTEAgo As Long: Dim lnCTESep As Long: Dim lnCTEOct As Long: Dim lnCTENov As Long: Dim lnCTEDic As Long
Dim lnCTDTot As Double: Dim lnCTDEne As Long: Dim lnCTDFeb As Long: Dim lnCTDMar As Long: Dim lnCTDAbr As Long: Dim lnCTDMay As Long: Dim lnCTDJun As Long: Dim lnCTDJul As Long: Dim lnCTDAgo As Long: Dim lnCTDSep As Long: Dim lnCTDOct As Long: Dim lnCTDNov As Long: Dim lnCTDDic As Long
Dim lnCTTot As Double:  Dim lnCTEne As Long: Dim lnCTFeb As Long: Dim lnCTMar As Long: Dim lnCTAbr As Long: Dim lnCTMay As Long: Dim lnCTJun As Long: Dim lnCTJul As Long: Dim lnCTAgo As Long: Dim lnCTSep As Long: Dim lnCTOct As Long: Dim lnCTNov As Long: Dim lnCTDic As Long
Dim lnCCTEEne As Double: Dim lnCCTEFeb As Double: Dim lnCCTEMar As Double: Dim lnCCTEAbr As Double: Dim lnCCTEMay As Double: Dim lnCCTEJun As Double: Dim lnCCTEJul As Double: Dim lnCCTEAgo As Double: Dim lnCCTESep As Double: Dim lnCCTEOct As Double: Dim lnCCTENov As Double: Dim lnCCTEDic As Double
Dim lnCCTDEne As Double: Dim lnCCTDFeb As Double: Dim lnCCTDMar As Double: Dim lnCCTDAbr As Double: Dim lnCCTDMay As Double: Dim lnCCTDJun As Double: Dim lnCCTDJul As Double: Dim lnCCTDAgo As Double: Dim lnCCTDSep As Double: Dim lnCCTDOct As Double: Dim lnCCTDNov As Double: Dim lnCCTDDic As Double
Dim lnCCTEne As Double: Dim lnCCTFeb As Double: Dim lnCCTMar As Double: Dim lnCCTAbr As Double: Dim lnCCTMay As Double: Dim lnCCTJun As Double: Dim lnCCTJul As Double: Dim lnCCTAgo As Double: Dim lnCCTSep As Double: Dim lnCCTOct As Double: Dim lnCCTNov As Double: Dim lnCCTDic As Double


Dim lcDpto As String: Dim lnIdDpto As Long
Dim lcEspecialidad As String: Dim lnIdEspecialidad As Long
Dim lcServicio As String: Dim lnIdServicio As Long
Dim lnMesDato As Integer: Dim lnDiasEstancia As Integer
Dim lcFiltro As String
Dim lcHoraEgreso As String: Dim ldFechaEgreso As Date
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError1


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
    
    'Filtra los Datos
    Set rsReporte = AtencionesSeleccionarPorAnio(ml_FechaAltaMedica, ml_lnAnio)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HEgresoHosp.ods"
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
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgresoHosp.xls")
            oWorkBookPlantilla.Worksheets("HEgresoHosp").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 1).setFormula(ml_Titulo)
                Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
            Else
                'Inicio de Impresion
                oWorkSheet.Cells(2, 3).Value = ml_Titulo
                oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
            End If
            iFila = 7
            iCol = 5
            lnTTot = 0: lnTEne = 0: lnTFeb = 0: lnTMar = 0: lnTAbr = 0: lnTMay = 0: lnTJun = 0: lnTJul = 0: lnTAgo = 0: lnTSep = 0: lnTOct = 0: lnTNov = 0: lnTDic = 0
            lnCTTot = 0: lnCTEne = 0: lnCTFeb = 0: lnCTMar = 0: lnCTAbr = 0: lnCTMay = 0: lnCTJun = 0: lnCTJul = 0: lnCTAgo = 0: lnCTSep = 0: lnCTOct = 0: lnCTNov = 0: lnCTDic = 0
            lnCCTEne = 0: lnCCTFeb = 0: lnCCTMar = 0: lnCCTAbr = 0: lnCCTMay = 0: lnCCTJun = 0: lnCCTJul = 0: lnCCTAgo = 0: lnCCTSep = 0: lnCCTOct = 0: lnCCTNov = 0: lnCCTDic = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                lnCTDTot = 0: lnCTDEne = 0: lnCTDFeb = 0: lnCTDMar = 0: lnCTDAbr = 0: lnCTDMay = 0: lnCTDJun = 0: lnCTDJul = 0: lnCTDAgo = 0: lnCTDSep = 0: lnCTDOct = 0: lnCTDNov = 0: lnCTDDic = 0
                lnCCTDEne = 0: lnCCTDFeb = 0: lnCCTDMar = 0: lnCCTDAbr = 0: lnCCTDMay = 0: lnCCTDJun = 0: lnCCTDJul = 0: lnCCTDAgo = 0: lnCCTDSep = 0: lnCCTDOct = 0: lnCCTDNov = 0: lnCCTDDic = 0
                lcDpto = rsReporte.Fields("dDpto").Value
                lnIdDpto = rsReporte.Fields("CodDpto").Value
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("dDpto").Value)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                End If
                'iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lnCTETot = 0: lnCTEEne = 0: lnCTEFeb = 0: lnCTEMar = 0: lnCTEAbr = 0: lnCTEMay = 0: lnCTEJun = 0: lnCTEJul = 0: lnCTEAgo = 0: lnCTESep = 0: lnCTEOct = 0: lnCTENov = 0: lnCTEDic = 0
                   lnCCTEEne = 0: lnCCTEFeb = 0: lnCCTEMar = 0: lnCCTEAbr = 0: lnCCTEMay = 0: lnCCTEJun = 0: lnCCTEJul = 0: lnCCTEAgo = 0: lnCCTESep = 0: lnCCTEOct = 0: lnCCTENov = 0: lnCCTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                 '  oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                 '  iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
                        lnCTSTot = 0: lnCTSEne = 0: lnCTSFeb = 0: lnCTSMar = 0: lnCTSAbr = 0: lnCTSMay = 0: lnCTSJun = 0: lnCTSJul = 0: lnCTSAgo = 0: lnCTSSep = 0: lnCTSOct = 0: lnCTSNov = 0: lnCTSDic = 0
                        lcServicio = rsReporte.Fields("dServicio").Value
                        lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                  '      oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value
                        Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value And lcServicio = rsReporte.Fields("dServicio").Value And lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                            lcHoraEgreso = ""
                            If ml_FechaAltaMedica Then
                               lnMesDato = Month(rsReporte.Fields("FechaEgreso").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgreso").Value
                               If Not IsNull(rsReporte.Fields("HoraEgreso").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgreso").Value
                               End If
                            Else
                               lnMesDato = Month(rsReporte.Fields("FechaEgresoAdministrativo").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgresoAdministrativo").Value
                               If Not IsNull(rsReporte.Fields("HoraEgresoAdministrativo").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgresoAdministrativo").Value
                               End If
                            End If
                            
'                            lnDiasEstancia = ldFechaEgreso - rsReporte.Fields("FechaIngreso").Value
'                            If lnDiasEstancia = 0 Then lnDiasEstancia = 1
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte.Fields("FechaIngreso").Value, rsReporte.Fields("HoraIngreso").Value, ldFechaEgreso, lcHoraEgreso)
                            
                            Select Case lnMesDato
                            Case 1
                                 lnTSEne = lnTSEne + lnDiasEstancia
                                 lnTEEne = lnTEEne + lnDiasEstancia
                                 lnTDEne = lnTDEne + lnDiasEstancia
                                 lnTEne = lnTEne + lnDiasEstancia
                                 
                                 lnCTSEne = lnCTSEne + 1
                                 lnCTEEne = lnCTEEne + 1
                                 lnCTDEne = lnCTDEne + 1
                                 lnCTEne = lnCTEne + 1
                                 
                            Case 2
                                 lnTSFeb = lnTSFeb + lnDiasEstancia
                                 lnTEFeb = lnTEFeb + lnDiasEstancia
                                 lnTDFeb = lnTDFeb + lnDiasEstancia
                                 lnTFeb = lnTFeb + lnDiasEstancia
                            
                            
                                 lnCTSFeb = lnCTSFeb + 1
                                 lnCTEFeb = lnCTEFeb + 1
                                 lnCTDFeb = lnCTDFeb + 1
                                 lnCTFeb = lnCTFeb + 1
                            Case 3
                                 lnTSMar = lnTSMar + lnDiasEstancia
                                 lnTEMar = lnTEMar + lnDiasEstancia
                                 lnTDMar = lnTDMar + lnDiasEstancia
                                 lnTMar = lnTMar + lnDiasEstancia
                            
                            
                                 lnCTSMar = lnCTSMar + 1
                                 lnCTEMar = lnCTEMar + 1
                                 lnCTDMar = lnCTDMar + 1
                                 lnCTMar = lnCTMar + 1
                            Case 4
                                 lnTSAbr = lnTSAbr + lnDiasEstancia
                                 lnTEAbr = lnTEAbr + lnDiasEstancia
                                 lnTDAbr = lnTDAbr + lnDiasEstancia
                                 lnTAbr = lnTAbr + lnDiasEstancia
                            
                            
                                 lnCTSAbr = lnCTSAbr + 1
                                 lnCTEAbr = lnCTEAbr + 1
                                 lnCTDAbr = lnCTDAbr + 1
                                 lnCTAbr = lnCTAbr + 1
                            Case 5
                                 lnTSMay = lnTSMay + lnDiasEstancia
                                 lnTEMay = lnTEMay + lnDiasEstancia
                                 lnTDMay = lnTDMay + lnDiasEstancia
                                 lnTMay = lnTMay + lnDiasEstancia
                            
                            
                                 lnCTSMay = lnCTSMay + 1
                                 lnCTEMay = lnCTEMay + 1
                                 lnCTDMay = lnCTDMay + 1
                                 lnCTMay = lnCTMay + 1
                            Case 6
                                 lnTSJun = lnTSJun + lnDiasEstancia
                                 lnTEJun = lnTEJun + lnDiasEstancia
                                 lnTDJun = lnTDJun + lnDiasEstancia
                                 lnTJun = lnTJun + lnDiasEstancia
                            
                                 lnCTSJun = lnCTSJun + 1
                                 lnCTEJun = lnCTEJun + 1
                                 lnCTDJun = lnCTDJun + 1
                                 lnCTJun = lnCTJun + 1
                            Case 7
                                 lnTSJul = lnTSJul + lnDiasEstancia
                                 lnTEJul = lnTEJul + lnDiasEstancia
                                 lnTDJul = lnTDJul + lnDiasEstancia
                                 lnTJul = lnTJul + lnDiasEstancia
                            
                            
                                 lnCTSJul = lnCTSJul + 1
                                 lnCTEJul = lnCTEJul + 1
                                 lnCTDJul = lnCTDJul + 1
                                 lnCTJul = lnCTJul + 1
                            Case 8
                                 lnTSAgo = lnTSAgo + lnDiasEstancia
                                 lnTEAgo = lnTEAgo + lnDiasEstancia
                                 lnTDAgo = lnTDAgo + lnDiasEstancia
                                 lnTAgo = lnTAgo + lnDiasEstancia
                            
                            
                                 lnCTSAgo = lnCTSAgo + 1
                                 lnCTEAgo = lnCTEAgo + 1
                                 lnCTDAgo = lnCTDAgo + 1
                                 lnCTAgo = lnCTAgo + 1
                            Case 9
                                 lnTSSep = lnTSSep + lnDiasEstancia
                                 lnTESep = lnTESep + lnDiasEstancia
                                 lnTDSep = lnTDSep + lnDiasEstancia
                                 lnTSep = lnTSep + lnDiasEstancia
                            
                            
                                 lnCTSSep = lnCTSSep + 1
                                 lnCTESep = lnCTESep + 1
                                 lnCTDSep = lnCTDSep + 1
                                 lnCTSep = lnCTSep + 1
                            Case 10
                                 lnTSOct = lnTSOct + lnDiasEstancia
                                 lnTEOct = lnTEOct + lnDiasEstancia
                                 lnTDOct = lnTDOct + lnDiasEstancia
                                 lnTOct = lnTOct + lnDiasEstancia
                            
                            
                                 lnCTSOct = lnCTSOct + 1
                                 lnCTEOct = lnCTEOct + 1
                                 lnCTDOct = lnCTDOct + 1
                                 lnCTOct = lnCTOct + 1
                            Case 11
                                 lnTSNov = lnTSNov + lnDiasEstancia
                                 lnTENov = lnTENov + lnDiasEstancia
                                 lnTDNov = lnTDNov + lnDiasEstancia
                                 lnTNov = lnTNov + lnDiasEstancia
                            
                            
                                 lnCTSNov = lnCTSNov + 1
                                 lnCTENov = lnCTENov + 1
                                 lnCTDNov = lnCTDNov + 1
                                 lnCTNov = lnCTNov + 1
                            Case 12
                                 lnTSDic = lnTSDic + lnDiasEstancia
                                 lnTEDic = lnTEDic + lnDiasEstancia
                                 lnTDDic = lnTDDic + lnDiasEstancia
                                 lnTDic = lnTDic + lnDiasEstancia
                            
                            
                                 lnCTSDic = lnCTSDic + 1
                                 lnCTEDic = lnCTEDic + 1
                                 lnCTDDic = lnCTDDic + 1
                                 lnCTDic = lnCTDic + 1
                            End Select
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        If lnCTSTot = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 0).Value = 0
                        Else
                           'oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTSTot / lnCTSTot, 2)
                           lnCTSTot = lnCTSTot + Round(lnTSTot / lnCTSTot, 2)
                        End If
                        If lnCTSEne = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 1).Value = 0
                        Else
                           'oWorkSheet.Cells(iFila, iCol + 1).Value = Round(lnTSEne / lnCTSEne, 2)
                           lnCTSTot = lnCTSTot + Round(lnTSEne / lnCTSEne, 2)
                           lnCCTEEne = lnCCTEEne + Round(lnTSEne / lnCTSEne, 2)
                           lnCCTDEne = lnCCTDEne + Round(lnTSEne / lnCTSEne, 2)
                           lnCCTEne = lnCCTEne + Round(lnTSEne / lnCTSEne, 2)
                        End If
                        If lnCTSFeb = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 2).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 2).Value = Round(lnTSFeb / lnCTSFeb, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSFeb / lnCTSFeb, 2)
                            lnCCTEFeb = lnCCTEFeb + Round(lnTSFeb / lnCTSFeb, 2)
                            lnCCTDFeb = lnCCTDFeb + Round(lnTSFeb / lnCTSFeb, 2)
                            lnCCTFeb = lnCCTFeb + Round(lnTSFeb / lnCTSFeb, 2)
                        End If
                        If lnCTSMar = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 3).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 3).Value = Round(lnTSMar / lnCTSMar, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSMar / lnCTSMar, 2)
                            lnCCTEMar = lnCCTEMar + Round(lnTSMar / lnCTSMar, 2)
                            lnCCTDMar = lnCCTDMar + Round(lnTSMar / lnCTSMar, 2)
                            lnCCTMar = lnCCTMar + Round(lnTSMar / lnCTSMar, 2)
                        End If
                        If lnCTSAbr = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 4).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 4).Value = Round(lnTSAbr / lnCTSAbr, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSAbr / lnCTSAbr, 2)
                            lnCCTEAbr = lnCCTEAbr + Round(lnTSAbr / lnCTSAbr, 2)
                            lnCCTDAbr = lnCCTDAbr + Round(lnTSAbr / lnCTSAbr, 2)
                            lnCCTAbr = lnCCTAbr + Round(lnTSAbr / lnCTSAbr, 2)
                        End If
                        If lnCTSMay = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 5).Value = Round(lnTSMay / lnCTSMay, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSMay / lnCTSMay, 2)
                            lnCCTEMay = lnCCTEMay + Round(lnTSMay / lnCTSMay, 2)
                            lnCCTDMay = lnCCTDMay + Round(lnTSMay / lnCTSMay, 2)
                            lnCCTMay = lnCCTMay + Round(lnTSMay / lnCTSMay, 2)
                        End If
                        If lnCTSJun = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 6).Value = Round(lnTSJun / lnCTSJun, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSJun / lnCTSJun, 2)
                            lnCCTEJun = lnCCTEJun + Round(lnTSJun / lnCTSJun, 2)
                            lnCCTDJun = lnCCTDJun + Round(lnTSJun / lnCTSJun, 2)
                            lnCCTJun = lnCCTJun + Round(lnTSJun / lnCTSJun, 2)
                           
                        End If
                        If lnCTSJul = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 7).Value = Round(lnTSJul / lnCTSJul, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSJul / lnCTSJul, 2)
                            lnCCTEJul = lnCCTEJul + Round(lnTSJul / lnCTSJul, 2)
                            lnCCTDJul = lnCCTDJul + Round(lnTSJul / lnCTSJul, 2)
                            lnCCTJul = lnCCTJul + Round(lnTSJul / lnCTSJul, 2)
                           
                        End If
                        If lnCTSAgo = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                        Else
                           ' oWorkSheet.Cells(iFila, iCol + 8).Value = Round(lnTSAgo / lnCTSAgo, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSAgo / lnCTSAgo, 2)
                            lnCCTEAgo = lnCCTEAgo + Round(lnTSAgo / lnCTSAgo, 2)
                            lnCCTDAgo = lnCCTDAgo + Round(lnTSAgo / lnCTSAgo, 2)
                            lnCCTAgo = lnCCTAgo + Round(lnTSAgo / lnCTSAgo, 2)
                           
                        End If
                        If lnCTSSep = 0 Then
                          ' oWorkSheet.Cells(iFila, iCol + 9).Value = 0
                        Else
                          ' oWorkSheet.Cells(iFila, iCol + 9).Value = Round(lnTSSep / lnCTSSep, 2)
                           lnCTSTot = lnCTSTot + Round(lnTSSep / lnCTSSep, 2)
                           lnCCTESep = lnCCTESep + Round(lnTSSep / lnCTSSep, 2)
                           lnCCTDSep = lnCCTDSep + Round(lnTSSep / lnCTSSep, 2)
                           lnCCTSep = lnCCTSep + Round(lnTSSep / lnCTSSep, 2)
                        End If
                        If lnCTSOct = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 10).Value = 0
                        Else
                            'oWorkSheet.Cells(iFila, iCol + 10).Value = Round(lnTSOct / lnCTSOct, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSOct / lnCTSOct, 2)
                            lnCCTEOct = lnCCTEOct + Round(lnTSOct / lnCTSOct, 2)
                            lnCCTDOct = lnCCTDOct + Round(lnTSOct / lnCTSOct, 2)
                            lnCCTOct = lnCCTOct + Round(lnTSOct / lnCTSOct, 2)
                        
                        End If
                        If lnCTSNov = 0 Then
                           'oWorkSheet.Cells(iFila, iCol + 11).Value = 0
                        Else
                           ' oWorkSheet.Cells(iFila, iCol + 11).Value = Round(lnTSNov / lnCTSNov, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSNov / lnCTSNov, 2)
                            lnCCTENov = lnCCTENov + Round(lnTSNov / lnCTSNov, 2)
                            lnCCTDNov = lnCCTDNov + Round(lnTSNov / lnCTSNov, 2)
                            lnCCTNov = lnCCTNov + Round(lnTSNov / lnCTSNov, 2)
                           
                        End If
                        If lnCTSDic = 0 Then
                          ' oWorkSheet.Cells(iFila, iCol + 12).Value = 0
                        Else
                         '   oWorkSheet.Cells(iFila, iCol + 12).Value = Round(lnTSDic / lnCTSDic, 2)
                            lnCTSTot = lnCTSTot + Round(lnTSDic / lnCTSDic, 2)
                            lnCCTEDic = lnCCTEDic + Round(lnTSDic / lnCTSDic, 2)
                            lnCCTDDic = lnCCTDDic + Round(lnTSDic / lnCTSDic, 2)
                            lnCCTDic = lnCCTDic + Round(lnTSDic / lnCTSDic, 2)
                        
                        End If
                        'oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTSTot
                        lnCTETot = lnCTETot + lnCTSTot
                        lnCTDTot = lnCTDTot + lnCTSTot
                        lnCTTot = lnCTTot + lnCTSTot
                        
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop
'                   oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTETot
'                   oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTEEne
'                   oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTEFeb
'                   oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTEMar
'                   oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTEAbr
'                   oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTEMay
'                   oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTEJun
'                   oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTEJul
'                   oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTEAgo
'                   oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTESep
'                   oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTEOct
'                   oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTENov
'                   oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTEDic
'                   iFila = iFila + 1
                   
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
    '            iFila = iFila + 1
    '            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 12
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTDTot)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTDEne)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTDFeb)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTDMar)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTDAbr)
                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTDMay)
                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTDJun)
                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTDJul)
                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTDAgo)
                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTDSep)
                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTDOct)
                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTDNov)
                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTDDic)
                Else
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTDTot
                    oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTDEne
                    oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTDFeb
                    oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTDMar
                    oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTDAbr
                    oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTDMay
                    oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTDJun
                    oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTDJul
                    oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTDAgo
                    oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTDSep
                    oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTDOct
                    oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTDNov
                    oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTDDic
                End If
                iFila = iFila + 1
                If rsReporte.EOF Then
                   Exit Do
                End If
             Loop
             iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Q" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTTot)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTEne)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTFeb)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTMar)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTAbr)
                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTMay)
                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTJun)
                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTJul)
                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTAgo)
                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTSep)
                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTOct)
                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTNov)
                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTDic)
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 12
                    oWorkSheet.Cells(iFila, 2).Value = "Total: "
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTTot
                    oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTEne
                    oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTFeb
                    oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTMar
                    oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTAbr
                    oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTMay
                    oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTJun
                    oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTJul
                    oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTAgo
                    oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTSep
                    oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTOct
                    oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTNov
                    oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTDic
                 End If
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 17
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
ManejadorError1:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub

Sub CrearReportePorServicios(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsTmpReporte As New Recordset
Dim iFila As Long: Dim iCol As Integer
Dim lnTSTot As Long: Dim lnTSEne As Long: Dim lnTSFeb As Long: Dim lnTSMar As Long: Dim lnTSAbr As Long: Dim lnTSMay As Long: Dim lnTSJun As Long: Dim lnTSJul As Long: Dim lnTSAgo As Long: Dim lnTSSep As Long: Dim lnTSOct As Long: Dim lnTSNov As Long: Dim lnTSDic As Long
Dim lnTETot As Long: Dim lnTEEne As Long: Dim lnTEFeb As Long: Dim lnTEMar As Long: Dim lnTEAbr As Long: Dim lnTEMay As Long: Dim lnTEJun As Long: Dim lnTEJul As Long: Dim lnTEAgo As Long: Dim lnTESep As Long: Dim lnTEOct As Long: Dim lnTENov As Long: Dim lnTEDic As Long
Dim lnTDTot As Long: Dim lnTDEne As Long: Dim lnTDFeb As Long: Dim lnTDMar As Long: Dim lnTDAbr As Long: Dim lnTDMay As Long: Dim lnTDJun As Long: Dim lnTDJul As Long: Dim lnTDAgo As Long: Dim lnTDSep As Long: Dim lnTDOct As Long: Dim lnTDNov As Long: Dim lnTDDic As Long
Dim lnTTot As Long: Dim lnTEne As Long: Dim lnTFeb As Long: Dim lnTMar As Long: Dim lnTAbr As Long: Dim lnTMay As Long: Dim lnTJun As Long: Dim lnTJul As Long: Dim lnTAgo As Long: Dim lnTSep As Long: Dim lnTOct As Long: Dim lnTNov As Long: Dim lnTDic As Long

Dim lnCTSTot As Double: Dim lnCTSEne As Long: Dim lnCTSFeb As Long: Dim lnCTSMar As Long: Dim lnCTSAbr As Long: Dim lnCTSMay As Long: Dim lnCTSJun As Long: Dim lnCTSJul As Long: Dim lnCTSAgo As Long: Dim lnCTSSep As Long: Dim lnCTSOct As Long: Dim lnCTSNov As Long: Dim lnCTSDic As Long
Dim lnCTETot As Double: Dim lnCTEEne As Long: Dim lnCTEFeb As Long: Dim lnCTEMar As Long: Dim lnCTEAbr As Long: Dim lnCTEMay As Long: Dim lnCTEJun As Long: Dim lnCTEJul As Long: Dim lnCTEAgo As Long: Dim lnCTESep As Long: Dim lnCTEOct As Long: Dim lnCTENov As Long: Dim lnCTEDic As Long
Dim lnCTDTot As Double: Dim lnCTDEne As Long: Dim lnCTDFeb As Long: Dim lnCTDMar As Long: Dim lnCTDAbr As Long: Dim lnCTDMay As Long: Dim lnCTDJun As Long: Dim lnCTDJul As Long: Dim lnCTDAgo As Long: Dim lnCTDSep As Long: Dim lnCTDOct As Long: Dim lnCTDNov As Long: Dim lnCTDDic As Long
Dim lnCTTot As Double:  Dim lnCTEne As Long: Dim lnCTFeb As Long: Dim lnCTMar As Long: Dim lnCTAbr As Long: Dim lnCTMay As Long: Dim lnCTJun As Long: Dim lnCTJul As Long: Dim lnCTAgo As Long: Dim lnCTSep As Long: Dim lnCTOct As Long: Dim lnCTNov As Long: Dim lnCTDic As Long
Dim lnCCTEEne As Double: Dim lnCCTEFeb As Double: Dim lnCCTEMar As Double: Dim lnCCTEAbr As Double: Dim lnCCTEMay As Double: Dim lnCCTEJun As Double: Dim lnCCTEJul As Double: Dim lnCCTEAgo As Double: Dim lnCCTESep As Double: Dim lnCCTEOct As Double: Dim lnCCTENov As Double: Dim lnCCTEDic As Double
Dim lnCCTDEne As Double: Dim lnCCTDFeb As Double: Dim lnCCTDMar As Double: Dim lnCCTDAbr As Double: Dim lnCCTDMay As Double: Dim lnCCTDJun As Double: Dim lnCCTDJul As Double: Dim lnCCTDAgo As Double: Dim lnCCTDSep As Double: Dim lnCCTDOct As Double: Dim lnCCTDNov As Double: Dim lnCCTDDic As Double
Dim lnCCTEne As Double: Dim lnCCTFeb As Double: Dim lnCCTMar As Double: Dim lnCCTAbr As Double: Dim lnCCTMay As Double: Dim lnCCTJun As Double: Dim lnCCTJul As Double: Dim lnCCTAgo As Double: Dim lnCCTSep As Double: Dim lnCCTOct As Double: Dim lnCCTNov As Double: Dim lnCCTDic As Double

Dim lcDpto As String: Dim lnIdDpto As Long
Dim lcEspecialidad As String: Dim lnIdEspecialidad As Long
Dim lcServicio As String: Dim lnIdServicio As Long
Dim lnMesDato As Integer: Dim lnDiasEstancia As Integer
Dim lcFiltro As String
Dim lcHoraEgreso As String: Dim ldFechaEgreso As Date
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError2

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
    
    'Filtra los Datos
    Set rsReporte = AtencionesSeleccionarPorAnio(ml_FechaAltaMedica, ml_lnAnio)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HEgresoHosp.ods"
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
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgresoHosp.xls")
            oWorkBookPlantilla.Worksheets("HEgresoHosp").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 1).setFormula(ml_Titulo)
            Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
        Else
            'Inicio de Impresion
            oWorkSheet.Cells(2, 3).Value = ml_Titulo
            oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
        End If
            iFila = 7
            iCol = 5
            lnTTot = 0: lnTEne = 0: lnTFeb = 0: lnTMar = 0: lnTAbr = 0: lnTMay = 0: lnTJun = 0: lnTJul = 0: lnTAgo = 0: lnTSep = 0: lnTOct = 0: lnTNov = 0: lnTDic = 0
            lnCTTot = 0: lnCTEne = 0: lnCTFeb = 0: lnCTMar = 0: lnCTAbr = 0: lnCTMay = 0: lnCTJun = 0: lnCTJul = 0: lnCTAgo = 0: lnCTSep = 0: lnCTOct = 0: lnCTNov = 0: lnCTDic = 0
            lnCCTEne = 0: lnCCTFeb = 0: lnCCTMar = 0: lnCCTAbr = 0: lnCCTMay = 0: lnCCTJun = 0: lnCCTJul = 0: lnCCTAgo = 0: lnCCTSep = 0: lnCCTOct = 0: lnCCTNov = 0: lnCCTDic = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                lnCTDTot = 0: lnCTDEne = 0: lnCTDFeb = 0: lnCTDMar = 0: lnCTDAbr = 0: lnCTDMay = 0: lnCTDJun = 0: lnCTDJul = 0: lnCTDAgo = 0: lnCTDSep = 0: lnCTDOct = 0: lnCTDNov = 0: lnCTDDic = 0
                lnCCTDEne = 0: lnCCTDFeb = 0: lnCCTDMar = 0: lnCCTDAbr = 0: lnCCTDMay = 0: lnCCTDJun = 0: lnCCTDJul = 0: lnCCTDAgo = 0: lnCCTDSep = 0: lnCCTDOct = 0: lnCCTDNov = 0: lnCCTDDic = 0
                lcDpto = rsReporte.Fields("dDpto").Value
                lnIdDpto = rsReporte.Fields("CodDpto").Value
                'oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                'iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lnCTETot = 0: lnCTEEne = 0: lnCTEFeb = 0: lnCTEMar = 0: lnCTEAbr = 0: lnCTEMay = 0: lnCTEJun = 0: lnCTEJul = 0: lnCTEAgo = 0: lnCTESep = 0: lnCTEOct = 0: lnCTENov = 0: lnCTEDic = 0
                   lnCCTEEne = 0: lnCCTEFeb = 0: lnCCTEMar = 0: lnCCTEAbr = 0: lnCCTEMay = 0: lnCCTEJun = 0: lnCCTEJul = 0: lnCCTEAgo = 0: lnCCTESep = 0: lnCCTEOct = 0: lnCCTENov = 0: lnCCTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                   'oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                   'iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
                        lnCTSTot = 0: lnCTSEne = 0: lnCTSFeb = 0: lnCTSMar = 0: lnCTSAbr = 0: lnCTSMay = 0: lnCTSJun = 0: lnCTSJul = 0: lnCTSAgo = 0: lnCTSSep = 0: lnCTSOct = 0: lnCTSNov = 0: lnCTSDic = 0
                        lcServicio = rsReporte.Fields("dServicio").Value
                        lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                        If lnIdEspecialidad = ml_idEspecialidad1 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value)
                            Else
                                oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value
                            End If
                        End If
                        Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value And lcServicio = rsReporte.Fields("dServicio").Value And lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                            lcHoraEgreso = ""
                            If ml_FechaAltaMedica Then
                               lnMesDato = Month(rsReporte.Fields("FechaEgreso").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgreso").Value
                               If Not IsNull(rsReporte.Fields("HoraEgreso").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgreso").Value
                               End If
                            Else
                               lnMesDato = Month(rsReporte.Fields("FechaEgresoAdministrativo").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgresoAdministrativo").Value
                               If Not IsNull(rsReporte.Fields("HoraEgresoAdministrativo").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgresoAdministrativo").Value
                               End If
                            End If
                            
'                            lnDiasEstancia = ldFechaEgreso - rsReporte.Fields("FechaIngreso").Value
'                            If lnDiasEstancia = 0 Then lnDiasEstancia = 1
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte.Fields("FechaIngreso").Value, rsReporte.Fields("HoraIngreso").Value, ldFechaEgreso, lcHoraEgreso)
                            
                            If lnIdEspecialidad = ml_idEspecialidad1 Then
                            Select Case lnMesDato
                            Case 1
                                 lnTSEne = lnTSEne + lnDiasEstancia
                                 lnTEEne = lnTEEne + lnDiasEstancia
                                 lnTDEne = lnTDEne + lnDiasEstancia
                                 lnTEne = lnTEne + lnDiasEstancia
                                 
                                 lnCTSEne = lnCTSEne + 1
                                 lnCTEEne = lnCTEEne + 1
                                 lnCTDEne = lnCTDEne + 1
                                 lnCTEne = lnCTEne + 1
                                 
                            Case 2
                                 lnTSFeb = lnTSFeb + lnDiasEstancia
                                 lnTEFeb = lnTEFeb + lnDiasEstancia
                                 lnTDFeb = lnTDFeb + lnDiasEstancia
                                 lnTFeb = lnTFeb + lnDiasEstancia
                            
                            
                                 lnCTSFeb = lnCTSFeb + 1
                                 lnCTEFeb = lnCTEFeb + 1
                                 lnCTDFeb = lnCTDFeb + 1
                                 lnCTFeb = lnCTFeb + 1
                            Case 3
                                 lnTSMar = lnTSMar + lnDiasEstancia
                                 lnTEMar = lnTEMar + lnDiasEstancia
                                 lnTDMar = lnTDMar + lnDiasEstancia
                                 lnTMar = lnTMar + lnDiasEstancia
                            
                            
                                 lnCTSMar = lnCTSMar + 1
                                 lnCTEMar = lnCTEMar + 1
                                 lnCTDMar = lnCTDMar + 1
                                 lnCTMar = lnCTMar + 1
                            Case 4
                                 lnTSAbr = lnTSAbr + lnDiasEstancia
                                 lnTEAbr = lnTEAbr + lnDiasEstancia
                                 lnTDAbr = lnTDAbr + lnDiasEstancia
                                 lnTAbr = lnTAbr + lnDiasEstancia
                            
                            
                                 lnCTSAbr = lnCTSAbr + 1
                                 lnCTEAbr = lnCTEAbr + 1
                                 lnCTDAbr = lnCTDAbr + 1
                                 lnCTAbr = lnCTAbr + 1
                            Case 5
                                 lnTSMay = lnTSMay + lnDiasEstancia
                                 lnTEMay = lnTEMay + lnDiasEstancia
                                 lnTDMay = lnTDMay + lnDiasEstancia
                                 lnTMay = lnTMay + lnDiasEstancia
                            
                            
                                 lnCTSMay = lnCTSMay + 1
                                 lnCTEMay = lnCTEMay + 1
                                 lnCTDMay = lnCTDMay + 1
                                 lnCTMay = lnCTMay + 1
                            Case 6
                                 lnTSJun = lnTSJun + lnDiasEstancia
                                 lnTEJun = lnTEJun + lnDiasEstancia
                                 lnTDJun = lnTDJun + lnDiasEstancia
                                 lnTJun = lnTJun + lnDiasEstancia
                            
                                 lnCTSJun = lnCTSJun + 1
                                 lnCTEJun = lnCTEJun + 1
                                 lnCTDJun = lnCTDJun + 1
                                 lnCTJun = lnCTJun + 1
                            Case 7
                                 lnTSJul = lnTSJul + lnDiasEstancia
                                 lnTEJul = lnTEJul + lnDiasEstancia
                                 lnTDJul = lnTDJul + lnDiasEstancia
                                 lnTJul = lnTJul + lnDiasEstancia
                            
                            
                                 lnCTSJul = lnCTSJul + 1
                                 lnCTEJul = lnCTEJul + 1
                                 lnCTDJul = lnCTDJul + 1
                                 lnCTJul = lnCTJul + 1
                            Case 8
                                 lnTSAgo = lnTSAgo + lnDiasEstancia
                                 lnTEAgo = lnTEAgo + lnDiasEstancia
                                 lnTDAgo = lnTDAgo + lnDiasEstancia
                                 lnTAgo = lnTAgo + lnDiasEstancia
                            
                            
                                 lnCTSAgo = lnCTSAgo + 1
                                 lnCTEAgo = lnCTEAgo + 1
                                 lnCTDAgo = lnCTDAgo + 1
                                 lnCTAgo = lnCTAgo + 1
                            Case 9
                                 lnTSSep = lnTSSep + lnDiasEstancia
                                 lnTESep = lnTESep + lnDiasEstancia
                                 lnTDSep = lnTDSep + lnDiasEstancia
                                 lnTSep = lnTSep + lnDiasEstancia
                            
                            
                                 lnCTSSep = lnCTSSep + 1
                                 lnCTESep = lnCTESep + 1
                                 lnCTDSep = lnCTDSep + 1
                                 lnCTSep = lnCTSep + 1
                            Case 10
                                 lnTSOct = lnTSOct + lnDiasEstancia
                                 lnTEOct = lnTEOct + lnDiasEstancia
                                 lnTDOct = lnTDOct + lnDiasEstancia
                                 lnTOct = lnTOct + lnDiasEstancia
                            
                            
                                 lnCTSOct = lnCTSOct + 1
                                 lnCTEOct = lnCTEOct + 1
                                 lnCTDOct = lnCTDOct + 1
                                 lnCTOct = lnCTOct + 1
                            Case 11
                                 lnTSNov = lnTSNov + lnDiasEstancia
                                 lnTENov = lnTENov + lnDiasEstancia
                                 lnTDNov = lnTDNov + lnDiasEstancia
                                 lnTNov = lnTNov + lnDiasEstancia
                            
                            
                                 lnCTSNov = lnCTSNov + 1
                                 lnCTENov = lnCTENov + 1
                                 lnCTDNov = lnCTDNov + 1
                                 lnCTNov = lnCTNov + 1
                            Case 12
                                 lnTSDic = lnTSDic + lnDiasEstancia
                                 lnTEDic = lnTEDic + lnDiasEstancia
                                 lnTDDic = lnTDDic + lnDiasEstancia
                                 lnTDic = lnTDic + lnDiasEstancia
                            
                            
                                 lnCTSDic = lnCTSDic + 1
                                 lnCTEDic = lnCTEDic + 1
                                 lnCTDDic = lnCTDDic + 1
                                 lnCTDic = lnCTDic + 1
                            End Select
                            End If
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        If lnIdEspecialidad = ml_idEspecialidad1 Then
                            If lnCTSEne = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 1).Value = 0
                                End If
                            Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(Round(lnTSEne / lnCTSEne, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 1).Value = Round(lnTSEne / lnCTSEne, 2)
                            End If
                               lnCTSTot = lnCTSTot + Round(lnTSEne / lnCTSEne, 2)
                               lnCCTEEne = lnCCTEEne + Round(lnTSEne / lnCTSEne, 2)
                               lnCCTDEne = lnCCTDEne + Round(lnTSEne / lnCTSEne, 2)
                               lnCCTEne = lnCCTEne + Round(lnTSEne / lnCTSEne, 2)
                            End If
                            If lnCTSFeb = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 2).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(Round(lnTSFeb / lnCTSFeb, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 2).Value = Round(lnTSFeb / lnCTSFeb, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSFeb / lnCTSFeb, 2)
                                lnCCTEFeb = lnCCTEFeb + Round(lnTSFeb / lnCTSFeb, 2)
                                lnCCTDFeb = lnCCTDFeb + Round(lnTSFeb / lnCTSFeb, 2)
                                lnCCTFeb = lnCCTFeb + Round(lnTSFeb / lnCTSFeb, 2)
                            End If
                            If lnCTSMar = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 3).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(Round(lnTSMar / lnCTSMar, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 3).Value = Round(lnTSMar / lnCTSMar, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSMar / lnCTSMar, 2)
                                lnCCTEMar = lnCCTEMar + Round(lnTSMar / lnCTSMar, 2)
                                lnCCTDMar = lnCCTDMar + Round(lnTSMar / lnCTSMar, 2)
                                lnCCTMar = lnCCTMar + Round(lnTSMar / lnCTSMar, 2)
                            End If
                            If lnCTSAbr = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 4).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(Round(lnTSAbr / lnCTSAbr, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 4).Value = Round(lnTSAbr / lnCTSAbr, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSAbr / lnCTSAbr, 2)
                                lnCCTEAbr = lnCCTEAbr + Round(lnTSAbr / lnCTSAbr, 2)
                                lnCCTDAbr = lnCCTDAbr + Round(lnTSAbr / lnCTSAbr, 2)
                                lnCCTAbr = lnCCTAbr + Round(lnTSAbr / lnCTSAbr, 2)
                            End If
                            If lnCTSMay = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round(lnTSMay / lnCTSMay, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 5).Value = Round(lnTSMay / lnCTSMay, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSMay / lnCTSMay, 2)
                                lnCCTEMay = lnCCTEMay + Round(lnTSMay / lnCTSMay, 2)
                                lnCCTDMay = lnCCTDMay + Round(lnTSMay / lnCTSMay, 2)
                                lnCCTMay = lnCCTMay + Round(lnTSMay / lnCTSMay, 2)
                            End If
                            If lnCTSJun = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(lnTSJun / lnCTSJun, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 6).Value = Round(lnTSJun / lnCTSJun, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSJun / lnCTSJun, 2)
                                lnCCTEJun = lnCCTEJun + Round(lnTSJun / lnCTSJun, 2)
                                lnCCTDJun = lnCCTDJun + Round(lnTSJun / lnCTSJun, 2)
                                lnCCTJun = lnCCTJun + Round(lnTSJun / lnCTSJun, 2)
                               
                            End If
                            If lnCTSJul = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round(lnTSJul / lnCTSJul, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 7).Value = Round(lnTSJul / lnCTSJul, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSJul / lnCTSJul, 2)
                                lnCCTEJul = lnCCTEJul + Round(lnTSJul / lnCTSJul, 2)
                                lnCCTDJul = lnCCTDJul + Round(lnTSJul / lnCTSJul, 2)
                                lnCCTJul = lnCCTJul + Round(lnTSJul / lnCTSJul, 2)
                               
                            End If
                            If lnCTSAgo = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(lnTSAgo / lnCTSAgo, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 8).Value = Round(lnTSAgo / lnCTSAgo, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSAgo / lnCTSAgo, 2)
                                lnCCTEAgo = lnCCTEAgo + Round(lnTSAgo / lnCTSAgo, 2)
                                lnCCTDAgo = lnCCTDAgo + Round(lnTSAgo / lnCTSAgo, 2)
                                lnCCTAgo = lnCCTAgo + Round(lnTSAgo / lnCTSAgo, 2)
                               
                            End If
                            If lnCTSSep = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 9).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(Round(lnTSSep / lnCTSSep, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 9).Value = Round(lnTSSep / lnCTSSep, 2)
                                End If
                               lnCTSTot = lnCTSTot + Round(lnTSSep / lnCTSSep, 2)
                               lnCCTESep = lnCCTESep + Round(lnTSSep / lnCTSSep, 2)
                               lnCCTDSep = lnCCTDSep + Round(lnTSSep / lnCTSSep, 2)
                               lnCCTSep = lnCCTSep + Round(lnTSSep / lnCTSSep, 2)
                            End If
                            If lnCTSOct = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 10).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(Round(lnTSOct / lnCTSOct, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 10).Value = Round(lnTSOct / lnCTSOct, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSOct / lnCTSOct, 2)
                                lnCCTEOct = lnCCTEOct + Round(lnTSOct / lnCTSOct, 2)
                                lnCCTDOct = lnCCTDOct + Round(lnTSOct / lnCTSOct, 2)
                                lnCCTOct = lnCCTOct + Round(lnTSOct / lnCTSOct, 2)
                            
                            End If
                            If lnCTSNov = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 11).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(Round(lnTSNov / lnCTSNov, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 11).Value = Round(lnTSNov / lnCTSNov, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSNov / lnCTSNov, 2)
                                lnCCTENov = lnCCTENov + Round(lnTSNov / lnCTSNov, 2)
                                lnCCTDNov = lnCCTDNov + Round(lnTSNov / lnCTSNov, 2)
                                lnCCTNov = lnCCTNov + Round(lnTSNov / lnCTSNov, 2)
                            End If
                            If lnCTSDic = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 12).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(Round(lnTSDic / lnCTSDic, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 12).Value = Round(lnTSDic / lnCTSDic, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSDic / lnCTSDic, 2)
                                lnCCTEDic = lnCCTEDic + Round(lnTSDic / lnCTSDic, 2)
                                lnCCTDDic = lnCCTDDic + Round(lnTSDic / lnCTSDic, 2)
                                lnCCTDic = lnCCTDic + Round(lnTSDic / lnCTSDic, 2)
                            End If
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTSTot)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTSTot
                            End If
                            lnCTETot = lnCTETot + lnCTSTot
                            lnCTDTot = lnCTDTot + lnCTSTot
                            lnCTTot = lnCTTot + lnCTSTot
                            
                            iFila = iFila + 1
                        End If
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop

                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
 
                If rsReporte.EOF Then
                   Exit Do
                End If
             Loop
             iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Q" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
                Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTTot)
                Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTEne)
                Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTFeb)
                Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTMar)
                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTAbr)
                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTMay)
                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTJun)
                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTJul)
                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTAgo)
                Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTSep)
                Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTOct)
                Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTNov)
                Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTDic)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 12
                oWorkSheet.Cells(iFila, 2).Value = "Total: "
                oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTTot
                oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTEne
                oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTFeb
                oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTMar
                oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTAbr
                oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTMay
                oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTJun
                oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTJul
                oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTAgo
                oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTSep
                oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTOct
                oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTNov
                oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTDic
            End If
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 17
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
ManejadorError2:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub

Sub CrearReportePorDosEspecialidades(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsTmpReporte As New Recordset
Dim iFila As Long: Dim iCol As Integer
Dim lnTSTot As Long: Dim lnTSEne As Long: Dim lnTSFeb As Long: Dim lnTSMar As Long: Dim lnTSAbr As Long: Dim lnTSMay As Long: Dim lnTSJun As Long: Dim lnTSJul As Long: Dim lnTSAgo As Long: Dim lnTSSep As Long: Dim lnTSOct As Long: Dim lnTSNov As Long: Dim lnTSDic As Long
Dim lnTETot As Long: Dim lnTEEne As Long: Dim lnTEFeb As Long: Dim lnTEMar As Long: Dim lnTEAbr As Long: Dim lnTEMay As Long: Dim lnTEJun As Long: Dim lnTEJul As Long: Dim lnTEAgo As Long: Dim lnTESep As Long: Dim lnTEOct As Long: Dim lnTENov As Long: Dim lnTEDic As Long
Dim lnTDTot As Long: Dim lnTDEne As Long: Dim lnTDFeb As Long: Dim lnTDMar As Long: Dim lnTDAbr As Long: Dim lnTDMay As Long: Dim lnTDJun As Long: Dim lnTDJul As Long: Dim lnTDAgo As Long: Dim lnTDSep As Long: Dim lnTDOct As Long: Dim lnTDNov As Long: Dim lnTDDic As Long
Dim lnTTot As Long: Dim lnTEne As Long: Dim lnTFeb As Long: Dim lnTMar As Long: Dim lnTAbr As Long: Dim lnTMay As Long: Dim lnTJun As Long: Dim lnTJul As Long: Dim lnTAgo As Long: Dim lnTSep As Long: Dim lnTOct As Long: Dim lnTNov As Long: Dim lnTDic As Long

Dim lnCTSTot As Double: Dim lnCTSEne As Long: Dim lnCTSFeb As Long: Dim lnCTSMar As Long: Dim lnCTSAbr As Long: Dim lnCTSMay As Long: Dim lnCTSJun As Long: Dim lnCTSJul As Long: Dim lnCTSAgo As Long: Dim lnCTSSep As Long: Dim lnCTSOct As Long: Dim lnCTSNov As Long: Dim lnCTSDic As Long
Dim lnCTETot As Double: Dim lnCTEEne As Long: Dim lnCTEFeb As Long: Dim lnCTEMar As Long: Dim lnCTEAbr As Long: Dim lnCTEMay As Long: Dim lnCTEJun As Long: Dim lnCTEJul As Long: Dim lnCTEAgo As Long: Dim lnCTESep As Long: Dim lnCTEOct As Long: Dim lnCTENov As Long: Dim lnCTEDic As Long
Dim lnCTDTot As Double: Dim lnCTDEne As Long: Dim lnCTDFeb As Long: Dim lnCTDMar As Long: Dim lnCTDAbr As Long: Dim lnCTDMay As Long: Dim lnCTDJun As Long: Dim lnCTDJul As Long: Dim lnCTDAgo As Long: Dim lnCTDSep As Long: Dim lnCTDOct As Long: Dim lnCTDNov As Long: Dim lnCTDDic As Long
Dim lnCTTot As Double:  Dim lnCTEne As Long: Dim lnCTFeb As Long: Dim lnCTMar As Long: Dim lnCTAbr As Long: Dim lnCTMay As Long: Dim lnCTJun As Long: Dim lnCTJul As Long: Dim lnCTAgo As Long: Dim lnCTSep As Long: Dim lnCTOct As Long: Dim lnCTNov As Long: Dim lnCTDic As Long
Dim lnCCTEEne As Double: Dim lnCCTEFeb As Double: Dim lnCCTEMar As Double: Dim lnCCTEAbr As Double: Dim lnCCTEMay As Double: Dim lnCCTEJun As Double: Dim lnCCTEJul As Double: Dim lnCCTEAgo As Double: Dim lnCCTESep As Double: Dim lnCCTEOct As Double: Dim lnCCTENov As Double: Dim lnCCTEDic As Double
Dim lnCCTDEne As Double: Dim lnCCTDFeb As Double: Dim lnCCTDMar As Double: Dim lnCCTDAbr As Double: Dim lnCCTDMay As Double: Dim lnCCTDJun As Double: Dim lnCCTDJul As Double: Dim lnCCTDAgo As Double: Dim lnCCTDSep As Double: Dim lnCCTDOct As Double: Dim lnCCTDNov As Double: Dim lnCCTDDic As Double
Dim lnCCTEne As Double: Dim lnCCTFeb As Double: Dim lnCCTMar As Double: Dim lnCCTAbr As Double: Dim lnCCTMay As Double: Dim lnCCTJun As Double: Dim lnCCTJul As Double: Dim lnCCTAgo As Double: Dim lnCCTSep As Double: Dim lnCCTOct As Double: Dim lnCCTNov As Double: Dim lnCCTDic As Double

Dim lcDpto As String: Dim lnIdDpto As Long
Dim lcEspecialidad As String: Dim lnIdEspecialidad As Long
Dim lcServicio As String: Dim lnIdServicio As Long
Dim lnMesDato As Integer: Dim lnDiasEstancia As Integer
Dim lcFiltro As String
Dim lcHoraEgreso As String: Dim ldFechaEgreso As Date
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError3

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

    'Filtra los Datos
    Set rsReporte = AtencionesSeleccionarPorAnio(ml_FechaAltaMedica, ml_lnAnio)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HEgresoHosp.ods"
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
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HEgresoHosp.xls")
            oWorkBookPlantilla.Worksheets("HEgresoHosp").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(2, 1).setFormula(ml_Titulo)
            Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
        Else
            'Inicio de Impresion
            oWorkSheet.Cells(2, 3).Value = ml_Titulo
            oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
        End If
            iFila = 7
            iCol = 5
            lnTTot = 0: lnTEne = 0: lnTFeb = 0: lnTMar = 0: lnTAbr = 0: lnTMay = 0: lnTJun = 0: lnTJul = 0: lnTAgo = 0: lnTSep = 0: lnTOct = 0: lnTNov = 0: lnTDic = 0
            lnCTTot = 0: lnCTEne = 0: lnCTFeb = 0: lnCTMar = 0: lnCTAbr = 0: lnCTMay = 0: lnCTJun = 0: lnCTJul = 0: lnCTAgo = 0: lnCTSep = 0: lnCTOct = 0: lnCTNov = 0: lnCTDic = 0
            lnCCTEne = 0: lnCCTFeb = 0: lnCCTMar = 0: lnCCTAbr = 0: lnCCTMay = 0: lnCCTJun = 0: lnCCTJul = 0: lnCCTAgo = 0: lnCCTSep = 0: lnCCTOct = 0: lnCCTNov = 0: lnCCTDic = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                lnCTDTot = 0: lnCTDEne = 0: lnCTDFeb = 0: lnCTDMar = 0: lnCTDAbr = 0: lnCTDMay = 0: lnCTDJun = 0: lnCTDJul = 0: lnCTDAgo = 0: lnCTDSep = 0: lnCTDOct = 0: lnCTDNov = 0: lnCTDDic = 0
                lnCCTDEne = 0: lnCCTDFeb = 0: lnCCTDMar = 0: lnCCTDAbr = 0: lnCCTDMay = 0: lnCCTDJun = 0: lnCCTDJul = 0: lnCCTDAgo = 0: lnCCTDSep = 0: lnCCTDOct = 0: lnCCTDNov = 0: lnCCTDDic = 0
                lcDpto = rsReporte.Fields("dDpto").Value
                lnIdDpto = rsReporte.Fields("CodDpto").Value
                'oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                'iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lnCTETot = 0: lnCTEEne = 0: lnCTEFeb = 0: lnCTEMar = 0: lnCTEAbr = 0: lnCTEMay = 0: lnCTEJun = 0: lnCTEJul = 0: lnCTEAgo = 0: lnCTESep = 0: lnCTEOct = 0: lnCTENov = 0: lnCTEDic = 0
                   lnCCTEEne = 0: lnCCTEFeb = 0: lnCCTEMar = 0: lnCCTEAbr = 0: lnCCTEMay = 0: lnCCTEJun = 0: lnCCTEJul = 0: lnCCTEAgo = 0: lnCCTESep = 0: lnCCTEOct = 0: lnCCTENov = 0: lnCCTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                   'oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                   'iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
                        lnCTSTot = 0: lnCTSEne = 0: lnCTSFeb = 0: lnCTSMar = 0: lnCTSAbr = 0: lnCTSMay = 0: lnCTSJun = 0: lnCTSJul = 0: lnCTSAgo = 0: lnCTSSep = 0: lnCTSOct = 0: lnCTSNov = 0: lnCTSDic = 0
                        lcServicio = rsReporte.Fields("dServicio").Value
                        lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                        If lnIdServicio = ml_idServicio1 Or lnIdServicio = ml_idServicio2 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value)
                            Else
                                oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("CodigoServicio").Value & " - " & rsReporte.Fields("dServicio").Value
                            End If
                        End If
                        Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value And lcServicio = rsReporte.Fields("dServicio").Value And lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                            lcHoraEgreso = ""
                            If ml_FechaAltaMedica Then
                               lnMesDato = Month(rsReporte.Fields("FechaEgreso").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgreso").Value
                               If Not IsNull(rsReporte.Fields("HoraEgreso").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgreso").Value
                               End If
                            Else
                               lnMesDato = Month(rsReporte.Fields("FechaEgresoAdministrativo").Value)
                               ldFechaEgreso = rsReporte.Fields("FechaEgresoAdministrativo").Value
                               If Not IsNull(rsReporte.Fields("HoraEgresoAdministrativo").Value) Then
                                  lcHoraEgreso = rsReporte.Fields("HoraEgresoAdministrativo").Value
                               End If
                            End If
                            
'                            lnDiasEstancia = ldFechaEgreso - rsReporte.Fields("FechaIngreso").Value
'                            If lnDiasEstancia = 0 Then lnDiasEstancia = 1
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte.Fields("FechaIngreso").Value, rsReporte.Fields("HoraIngreso").Value, ldFechaEgreso, lcHoraEgreso)
                            
                            If lnIdServicio = ml_idServicio1 Or lnIdServicio = ml_idServicio2 Then
                            Select Case lnMesDato
                            Case 1
                                 lnTSEne = lnTSEne + lnDiasEstancia
                                 lnTEEne = lnTEEne + lnDiasEstancia
                                 lnTDEne = lnTDEne + lnDiasEstancia
                                 lnTEne = lnTEne + lnDiasEstancia
                                 
                                 lnCTSEne = lnCTSEne + 1
                                 lnCTEEne = lnCTEEne + 1
                                 lnCTDEne = lnCTDEne + 1
                                 lnCTEne = lnCTEne + 1
                                 
                            Case 2
                                 lnTSFeb = lnTSFeb + lnDiasEstancia
                                 lnTEFeb = lnTEFeb + lnDiasEstancia
                                 lnTDFeb = lnTDFeb + lnDiasEstancia
                                 lnTFeb = lnTFeb + lnDiasEstancia
                            
                            
                                 lnCTSFeb = lnCTSFeb + 1
                                 lnCTEFeb = lnCTEFeb + 1
                                 lnCTDFeb = lnCTDFeb + 1
                                 lnCTFeb = lnCTFeb + 1
                            Case 3
                                 lnTSMar = lnTSMar + lnDiasEstancia
                                 lnTEMar = lnTEMar + lnDiasEstancia
                                 lnTDMar = lnTDMar + lnDiasEstancia
                                 lnTMar = lnTMar + lnDiasEstancia
                            
                            
                                 lnCTSMar = lnCTSMar + 1
                                 lnCTEMar = lnCTEMar + 1
                                 lnCTDMar = lnCTDMar + 1
                                 lnCTMar = lnCTMar + 1
                            Case 4
                                 lnTSAbr = lnTSAbr + lnDiasEstancia
                                 lnTEAbr = lnTEAbr + lnDiasEstancia
                                 lnTDAbr = lnTDAbr + lnDiasEstancia
                                 lnTAbr = lnTAbr + lnDiasEstancia
                            
                            
                                 lnCTSAbr = lnCTSAbr + 1
                                 lnCTEAbr = lnCTEAbr + 1
                                 lnCTDAbr = lnCTDAbr + 1
                                 lnCTAbr = lnCTAbr + 1
                            Case 5
                                 lnTSMay = lnTSMay + lnDiasEstancia
                                 lnTEMay = lnTEMay + lnDiasEstancia
                                 lnTDMay = lnTDMay + lnDiasEstancia
                                 lnTMay = lnTMay + lnDiasEstancia
                            
                            
                                 lnCTSMay = lnCTSMay + 1
                                 lnCTEMay = lnCTEMay + 1
                                 lnCTDMay = lnCTDMay + 1
                                 lnCTMay = lnCTMay + 1
                            Case 6
                                 lnTSJun = lnTSJun + lnDiasEstancia
                                 lnTEJun = lnTEJun + lnDiasEstancia
                                 lnTDJun = lnTDJun + lnDiasEstancia
                                 lnTJun = lnTJun + lnDiasEstancia
                            
                                 lnCTSJun = lnCTSJun + 1
                                 lnCTEJun = lnCTEJun + 1
                                 lnCTDJun = lnCTDJun + 1
                                 lnCTJun = lnCTJun + 1
                            Case 7
                                 lnTSJul = lnTSJul + lnDiasEstancia
                                 lnTEJul = lnTEJul + lnDiasEstancia
                                 lnTDJul = lnTDJul + lnDiasEstancia
                                 lnTJul = lnTJul + lnDiasEstancia
                            
                            
                                 lnCTSJul = lnCTSJul + 1
                                 lnCTEJul = lnCTEJul + 1
                                 lnCTDJul = lnCTDJul + 1
                                 lnCTJul = lnCTJul + 1
                            Case 8
                                 lnTSAgo = lnTSAgo + lnDiasEstancia
                                 lnTEAgo = lnTEAgo + lnDiasEstancia
                                 lnTDAgo = lnTDAgo + lnDiasEstancia
                                 lnTAgo = lnTAgo + lnDiasEstancia
                            
                            
                                 lnCTSAgo = lnCTSAgo + 1
                                 lnCTEAgo = lnCTEAgo + 1
                                 lnCTDAgo = lnCTDAgo + 1
                                 lnCTAgo = lnCTAgo + 1
                            Case 9
                                 lnTSSep = lnTSSep + lnDiasEstancia
                                 lnTESep = lnTESep + lnDiasEstancia
                                 lnTDSep = lnTDSep + lnDiasEstancia
                                 lnTSep = lnTSep + lnDiasEstancia
                            
                            
                                 lnCTSSep = lnCTSSep + 1
                                 lnCTESep = lnCTESep + 1
                                 lnCTDSep = lnCTDSep + 1
                                 lnCTSep = lnCTSep + 1
                            Case 10
                                 lnTSOct = lnTSOct + lnDiasEstancia
                                 lnTEOct = lnTEOct + lnDiasEstancia
                                 lnTDOct = lnTDOct + lnDiasEstancia
                                 lnTOct = lnTOct + lnDiasEstancia
                            
                            
                                 lnCTSOct = lnCTSOct + 1
                                 lnCTEOct = lnCTEOct + 1
                                 lnCTDOct = lnCTDOct + 1
                                 lnCTOct = lnCTOct + 1
                            Case 11
                                 lnTSNov = lnTSNov + lnDiasEstancia
                                 lnTENov = lnTENov + lnDiasEstancia
                                 lnTDNov = lnTDNov + lnDiasEstancia
                                 lnTNov = lnTNov + lnDiasEstancia
                            
                            
                                 lnCTSNov = lnCTSNov + 1
                                 lnCTENov = lnCTENov + 1
                                 lnCTDNov = lnCTDNov + 1
                                 lnCTNov = lnCTNov + 1
                            Case 12
                                 lnTSDic = lnTSDic + lnDiasEstancia
                                 lnTEDic = lnTEDic + lnDiasEstancia
                                 lnTDDic = lnTDDic + lnDiasEstancia
                                 lnTDic = lnTDic + lnDiasEstancia
                            
                            
                                 lnCTSDic = lnCTSDic + 1
                                 lnCTEDic = lnCTEDic + 1
                                 lnCTDDic = lnCTDDic + 1
                                 lnCTDic = lnCTDic + 1
                            End Select
                            End If
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        If lnIdServicio = ml_idServicio1 Or lnIdServicio = ml_idServicio2 Then
                            If lnCTSEne = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 1).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(Round(lnTSEne / lnCTSEne, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 1).Value = Round(lnTSEne / lnCTSEne, 2)
                                End If
                               lnCTSTot = lnCTSTot + Round(lnTSEne / lnCTSEne, 2)
                               lnCCTEEne = lnCCTEEne + Round(lnTSEne / lnCTSEne, 2)
                               lnCCTDEne = lnCCTDEne + Round(lnTSEne / lnCTSEne, 2)
                               lnCCTEne = lnCCTEne + Round(lnTSEne / lnCTSEne, 2)
                            End If
                            If lnCTSFeb = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 2).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(Round(lnTSFeb / lnCTSFeb, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 2).Value = Round(lnTSFeb / lnCTSFeb, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSFeb / lnCTSFeb, 2)
                                lnCCTEFeb = lnCCTEFeb + Round(lnTSFeb / lnCTSFeb, 2)
                                lnCCTDFeb = lnCCTDFeb + Round(lnTSFeb / lnCTSFeb, 2)
                                lnCCTFeb = lnCCTFeb + Round(lnTSFeb / lnCTSFeb, 2)
                            End If
                            If lnCTSMar = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 3).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(Round(lnTSMar / lnCTSMar, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 3).Value = Round(lnTSMar / lnCTSMar, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSMar / lnCTSMar, 2)
                                lnCCTEMar = lnCCTEMar + Round(lnTSMar / lnCTSMar, 2)
                                lnCCTDMar = lnCCTDMar + Round(lnTSMar / lnCTSMar, 2)
                                lnCCTMar = lnCCTMar + Round(lnTSMar / lnCTSMar, 2)
                            End If
                            If lnCTSAbr = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 4).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(Round(lnTSAbr / lnCTSAbr, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 4).Value = Round(lnTSAbr / lnCTSAbr, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSAbr / lnCTSAbr, 2)
                                lnCCTEAbr = lnCCTEAbr + Round(lnTSAbr / lnCTSAbr, 2)
                                lnCCTDAbr = lnCCTDAbr + Round(lnTSAbr / lnCTSAbr, 2)
                                lnCCTAbr = lnCCTAbr + Round(lnTSAbr / lnCTSAbr, 2)
                            End If
                            If lnCTSMay = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round(lnTSMay / lnCTSMay, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 5).Value = Round(lnTSMay / lnCTSMay, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSMay / lnCTSMay, 2)
                                lnCCTEMay = lnCCTEMay + Round(lnTSMay / lnCTSMay, 2)
                                lnCCTDMay = lnCCTDMay + Round(lnTSMay / lnCTSMay, 2)
                                lnCCTMay = lnCCTMay + Round(lnTSMay / lnCTSMay, 2)
                            End If
                            If lnCTSJun = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(lnTSJun / lnCTSJun, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 6).Value = Round(lnTSJun / lnCTSJun, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSJun / lnCTSJun, 2)
                                lnCCTEJun = lnCCTEJun + Round(lnTSJun / lnCTSJun, 2)
                                lnCCTDJun = lnCCTDJun + Round(lnTSJun / lnCTSJun, 2)
                                lnCCTJun = lnCCTJun + Round(lnTSJun / lnCTSJun, 2)
                               
                            End If
                            If lnCTSJul = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round(lnTSJul / lnCTSJul, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 7).Value = Round(lnTSJul / lnCTSJul, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSJul / lnCTSJul, 2)
                                lnCCTEJul = lnCCTEJul + Round(lnTSJul / lnCTSJul, 2)
                                lnCCTDJul = lnCCTDJul + Round(lnTSJul / lnCTSJul, 2)
                                lnCCTJul = lnCCTJul + Round(lnTSJul / lnCTSJul, 2)
                               
                            End If
                            If lnCTSAgo = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(lnTSAgo / lnCTSAgo, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 8).Value = Round(lnTSAgo / lnCTSAgo, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSAgo / lnCTSAgo, 2)
                                lnCCTEAgo = lnCCTEAgo + Round(lnTSAgo / lnCTSAgo, 2)
                                lnCCTDAgo = lnCCTDAgo + Round(lnTSAgo / lnCTSAgo, 2)
                                lnCCTAgo = lnCCTAgo + Round(lnTSAgo / lnCTSAgo, 2)
                               
                            End If
                            If lnCTSSep = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 9).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(Round(lnTSSep / lnCTSSep, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 9).Value = Round(lnTSSep / lnCTSSep, 2)
                                End If
                               lnCTSTot = lnCTSTot + Round(lnTSSep / lnCTSSep, 2)
                               lnCCTESep = lnCCTESep + Round(lnTSSep / lnCTSSep, 2)
                               lnCCTDSep = lnCCTDSep + Round(lnTSSep / lnCTSSep, 2)
                               lnCCTSep = lnCCTSep + Round(lnTSSep / lnCTSSep, 2)
                            End If
                            If lnCTSOct = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 10).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(Round(lnTSOct / lnCTSOct, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 10).Value = Round(lnTSOct / lnCTSOct, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSOct / lnCTSOct, 2)
                                lnCCTEOct = lnCCTEOct + Round(lnTSOct / lnCTSOct, 2)
                                lnCCTDOct = lnCCTDOct + Round(lnTSOct / lnCTSOct, 2)
                                lnCCTOct = lnCCTOct + Round(lnTSOct / lnCTSOct, 2)
                            
                            End If
                            If lnCTSNov = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 11).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(Round(lnTSNov / lnCTSNov, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 11).Value = Round(lnTSNov / lnCTSNov, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSNov / lnCTSNov, 2)
                                lnCCTENov = lnCCTENov + Round(lnTSNov / lnCTSNov, 2)
                                lnCCTDNov = lnCCTDNov + Round(lnTSNov / lnCTSNov, 2)
                                lnCCTNov = lnCCTNov + Round(lnTSNov / lnCTSNov, 2)
                               
                            End If
                            If lnCTSDic = 0 Then
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(0)
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 12).Value = 0
                                End If
                            Else
                                If lbEsOpenOffice = True Then
                                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(Round(lnTSDic / lnCTSDic, 2))
                                Else
                                    oWorkSheet.Cells(iFila, iCol + 12).Value = Round(lnTSDic / lnCTSDic, 2)
                                End If
                                lnCTSTot = lnCTSTot + Round(lnTSDic / lnCTSDic, 2)
                                lnCCTEDic = lnCCTEDic + Round(lnTSDic / lnCTSDic, 2)
                                lnCCTDDic = lnCCTDDic + Round(lnTSDic / lnCTSDic, 2)
                                lnCCTDic = lnCCTDic + Round(lnTSDic / lnCTSDic, 2)
                            
                            End If
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTSTot)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTSTot
                            End If
                            lnCTETot = lnCTETot + lnCTSTot
                            lnCTDTot = lnCTDTot + lnCTSTot
                            lnCTTot = lnCTTot + lnCTSTot
                            
                            iFila = iFila + 1
                        End If
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop
  
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
 
                If rsReporte.EOF Then
                   Exit Do
                End If
             Loop
             iFila = iFila + 1
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":Q" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnCTTot)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnCCTEne)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnCCTFeb)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnCCTMar)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnCCTAbr)
                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnCCTMay)
                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnCCTJun)
                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnCCTJul)
                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnCCTAgo)
                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnCCTSep)
                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnCCTOct)
                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnCCTNov)
                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnCCTDic)
                Else
                     mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 12
                     oWorkSheet.Cells(iFila, 2).Value = "Total: "
                     oWorkSheet.Cells(iFila, iCol + 0).Value = lnCTTot
                     oWorkSheet.Cells(iFila, iCol + 1).Value = lnCCTEne
                     oWorkSheet.Cells(iFila, iCol + 2).Value = lnCCTFeb
                     oWorkSheet.Cells(iFila, iCol + 3).Value = lnCCTMar
                     oWorkSheet.Cells(iFila, iCol + 4).Value = lnCCTAbr
                     oWorkSheet.Cells(iFila, iCol + 5).Value = lnCCTMay
                     oWorkSheet.Cells(iFila, iCol + 6).Value = lnCCTJun
                     oWorkSheet.Cells(iFila, iCol + 7).Value = lnCCTJul
                     oWorkSheet.Cells(iFila, iCol + 8).Value = lnCCTAgo
                     oWorkSheet.Cells(iFila, iCol + 9).Value = lnCCTSep
                     oWorkSheet.Cells(iFila, iCol + 10).Value = lnCCTOct
                     oWorkSheet.Cells(iFila, iCol + 11).Value = lnCCTNov
                     oWorkSheet.Cells(iFila, iCol + 12).Value = lnCCTDic
                End If
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 17
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
ManejadorError3:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub

Function AtencionesSeleccionarPorAnio(lbPorFechaAltaMedica As Boolean, lnAnio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set AtencionesSeleccionarPorAnio = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        If lbPorFechaAltaMedica Then
           .CommandText = "AtencionesSeleccionarPorAnioDeEgresoMedico"
        Else
           .CommandText = "AtencionesSeleccionarPorAnioDeEgresoAdministrativo"
        End If
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorAnio = oRecordset
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


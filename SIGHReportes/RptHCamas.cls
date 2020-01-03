VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHCamas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Camas
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim rsReporte As New Recordset
Dim mrs_Tmp As New ADODB.Recordset
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_lnAnio As Integer
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
Dim ml_SoloHospit As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let SoloHospit(lValue As Boolean)
    ml_SoloHospit = lValue
End Property

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
    Dim oFormulario As New HCamas
    oFormulario.Show 1
End Sub
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Programación de Camas
'*****************Consideraciones:  Solo Hospitalizados,
'*****************                  se considera a partir  de la Segunda transferencia (secuencia=2)
'*****************                  no se consideran pacientes que llega y sale en 1 solo servicio
'*****************                  Si el paciente aun no tiene ALTA MEDICA, no se considera ese ultimo Servicio
Sub CrearReporteDetallado(lnHwnd As Long)
Dim iFila As Long: Dim iCol As Integer
Dim lnTSTot As Long: Dim lnTSEne As Long: Dim lnTSFeb As Long: Dim lnTSMar As Long: Dim lnTSAbr As Long: Dim lnTSMay As Long: Dim lnTSJun As Long: Dim lnTSJul As Long: Dim lnTSAgo As Long: Dim lnTSSep As Long: Dim lnTSOct As Long: Dim lnTSNov As Long: Dim lnTSDic As Long
Dim lnTETot As Long: Dim lnTEEne As Long: Dim lnTEFeb As Long: Dim lnTEMar As Long: Dim lnTEAbr As Long: Dim lnTEMay As Long: Dim lnTEJun As Long: Dim lnTEJul As Long: Dim lnTEAgo As Long: Dim lnTESep As Long: Dim lnTEOct As Long: Dim lnTENov As Long: Dim lnTEDic As Long
Dim lnTDTot As Long: Dim lnTDEne As Long: Dim lnTDFeb As Long: Dim lnTDMar As Long: Dim lnTDAbr As Long: Dim lnTDMay As Long: Dim lnTDJun As Long: Dim lnTDJul As Long: Dim lnTDAgo As Long: Dim lnTDSep As Long: Dim lnTDOct As Long: Dim lnTDNov As Long: Dim lnTDDic As Long
Dim lnTTot As Long: Dim lnTEne As Long: Dim lnTFeb As Long: Dim lnTMar As Long: Dim lnTAbr As Long: Dim lnTMay As Long: Dim lnTJun As Long: Dim lnTJul As Long: Dim lnTAgo As Long: Dim lnTSep As Long: Dim lnTOct As Long: Dim lnTNov As Long: Dim lnTDic As Long
Dim lcDpto As String: Dim lnIdDpto As Long
Dim lcEspecialidad As String: Dim lnIdEspecialidad As Long
Dim lcServicio As String: Dim lnIdServicio As Long
Dim lnMesDato As Integer: Dim lnDiasEstancia As Integer
Dim lcFiltro As String
Dim lnIdAtencion As Long: Dim lbContinua As Boolean
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
    Set mrs_Tmp = mo_ReglasHoteleria.CamasProgramadasPorAnio(ml_lnAnio, ml_SoloHospit)
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        GenerarRecordsetTemporal
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
           lbContinua = True
           Select Case ml_TipoReporte
           Case 2   'Por un Servicio
                  lbContinua = IIf(mrs_Tmp.Fields("CodEspecialidad").Value = ml_idEspecialidad1, True, False)
           Case 3   'Por dos Especialidades
                  lbContinua = IIf(mrs_Tmp.Fields("IdServicio").Value = ml_idServicio1 Or mrs_Tmp.Fields("IdServicio").Value = ml_idServicio2, True, False)
           End Select
           If lbContinua Then
                rsReporte.AddNew
                rsReporte.Fields("dDpto").Value = mrs_Tmp.Fields("dDpto").Value
                rsReporte.Fields("CodDpto").Value = mrs_Tmp.Fields("codDpto").Value
                rsReporte.Fields("dEspecialidad").Value = mrs_Tmp.Fields("dEspecialidad").Value
                rsReporte.Fields("CodEspecialidad").Value = mrs_Tmp.Fields("CodEspecialidad").Value
                rsReporte.Fields("dServicio").Value = mrs_Tmp.Fields("dServicio").Value
                rsReporte.Fields("CodServicio").Value = mrs_Tmp.Fields("CodServicio").Value
                rsReporte.Fields("idServicio").Value = mrs_Tmp.Fields("idServicio").Value
                rsReporte.Update
           End If
           lnIdServicio = mrs_Tmp.Fields("idServicio").Value
           Do While Not mrs_Tmp.EOF And lnIdServicio = mrs_Tmp.Fields("idServicio").Value
               lnMesDato = Month(mrs_Tmp.Fields("fechaIngreso").Value)
               iFila = 0
               Do While Not mrs_Tmp.EOF And lnIdServicio = mrs_Tmp.Fields("idServicio").Value And lnMesDato = Month(mrs_Tmp.Fields("fechaIngreso").Value)
                   iFila = iFila + 1
                   mrs_Tmp.MoveNext
                   If mrs_Tmp.EOF Then
                       Exit Do
                   End If
               Loop
               If lbContinua Then
                    Select Case lnMesDato
                    Case 1
                       rsReporte.Fields("Ene").Value = rsReporte.Fields("Ene").Value + iFila
                    Case 2
                       rsReporte.Fields("Feb").Value = rsReporte.Fields("Feb").Value + iFila
                    Case 3
                       rsReporte.Fields("Mar").Value = rsReporte.Fields("Mar").Value + iFila
                    Case 4
                       rsReporte.Fields("Abr").Value = rsReporte.Fields("Abr").Value + iFila
                    Case 5
                       rsReporte.Fields("May").Value = rsReporte.Fields("May").Value + iFila
                    Case 6
                       rsReporte.Fields("Jun").Value = rsReporte.Fields("Jun").Value + iFila
                    Case 7
                       rsReporte.Fields("Jul").Value = rsReporte.Fields("Jul").Value + iFila
                    Case 8
                       rsReporte.Fields("Ago").Value = rsReporte.Fields("Ago").Value + iFila
                    Case 9
                       rsReporte.Fields("Sep").Value = rsReporte.Fields("Sep").Value + iFila
                    Case 10
                       rsReporte.Fields("Oct").Value = rsReporte.Fields("Oct").Value + iFila
                    Case 11
                       rsReporte.Fields("Nov").Value = rsReporte.Fields("Nov").Value + iFila
                    Case 12
                       rsReporte.Fields("Dic").Value = rsReporte.Fields("Dic").Value + iFila
                    End Select
                    rsReporte.Update
               End If
               If mrs_Tmp.EOF Then
                 Exit Do
               End If
           Loop
        Loop
        If rsReporte.RecordCount = 0 Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            rsReporte.Sort = "dDpto,CodDpto,dEspecialidad,CodEspecialidad,dServicio,IdServicio"
            If ml_TipoReporte = 1 Then   'Consolidado Por Departamentos
                If lbEsOpenOffice = True Then
                    'Abre el archivo ExcelOpenOffice
                    lcArchivoExcel = App.Path + "\Plantillas\HEgresoHosp.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
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
                        Call Feuille.getcellbyposition(4, 5).setFormula("Promedio Mensual")
                    Else
                        'Inicio de Impresion
                        oWorkSheet.Cells(2, 3).Value = ml_Titulo
                        oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
                        oWorkSheet.Cells(6, 5).Value = "Promedio Mensual"
                    End If
                iFila = 7
                iCol = 5
                lnTTot = 0: lnTEne = 0: lnTFeb = 0: lnTMar = 0: lnTAbr = 0: lnTMay = 0: lnTJun = 0: lnTJul = 0: lnTAgo = 0: lnTSep = 0: lnTOct = 0: lnTNov = 0: lnTDic = 0
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                    lcDpto = rsReporte.Fields("dDpto").Value
                    lnIdDpto = rsReporte.Fields("CodDpto").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("dDpto").Value)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                    End If
                    Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                       lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                       lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                       lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                       Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                            lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
                            lcServicio = rsReporte.Fields("dServicio").Value
                            lnIdServicio = rsReporte.Fields("IdServicio").Value
                            lnTSEne = lnTSEne + rsReporte.Fields("Ene").Value
                            lnTEEne = lnTEEne + rsReporte.Fields("Ene").Value
                            lnTDEne = lnTDEne + rsReporte.Fields("Ene").Value
                            lnTEne = lnTEne + rsReporte.Fields("Ene").Value
                            lnTSFeb = lnTSFeb + rsReporte.Fields("Feb").Value
                            lnTEFeb = lnTEFeb + rsReporte.Fields("Feb").Value
                            lnTDFeb = lnTDFeb + rsReporte.Fields("Feb").Value
                            lnTFeb = lnTFeb + rsReporte.Fields("Feb").Value
                            lnTSMar = lnTSMar + rsReporte.Fields("Mar").Value
                            lnTEMar = lnTEMar + rsReporte.Fields("Mar").Value
                            lnTDMar = lnTDMar + rsReporte.Fields("Mar").Value
                            lnTMar = lnTMar + rsReporte.Fields("Mar").Value
                            lnTSAbr = lnTSAbr + rsReporte.Fields("Abr").Value
                            lnTEAbr = lnTEAbr + rsReporte.Fields("Abr").Value
                            lnTDAbr = lnTDAbr + rsReporte.Fields("Abr").Value
                            lnTAbr = lnTAbr + rsReporte.Fields("Abr").Value
                            lnTSMay = lnTSMay + rsReporte.Fields("May").Value
                            lnTEMay = lnTEMay + rsReporte.Fields("May").Value
                            lnTDMay = lnTDMay + rsReporte.Fields("May").Value
                            lnTMay = lnTMay + rsReporte.Fields("May").Value
                            lnTSJun = lnTSJun + rsReporte.Fields("Jun").Value
                            lnTEJun = lnTEJun + rsReporte.Fields("Jun").Value
                            lnTDJun = lnTDJun + rsReporte.Fields("Jun").Value
                            lnTJun = lnTJun + rsReporte.Fields("Jun").Value
                            lnTSJul = lnTSJul + rsReporte.Fields("Jul").Value
                            lnTEJul = lnTEJul + rsReporte.Fields("Jul").Value
                            lnTDJul = lnTDJul + rsReporte.Fields("Jul").Value
                            lnTJul = lnTJul + rsReporte.Fields("Jul").Value
                            lnTSAgo = lnTSAgo + rsReporte.Fields("Ago").Value
                            lnTEAgo = lnTEAgo + rsReporte.Fields("Ago").Value
                            lnTDAgo = lnTDAgo + rsReporte.Fields("Ago").Value
                            lnTAgo = lnTAgo + rsReporte.Fields("Ago").Value
                            lnTSSep = lnTSSep + rsReporte.Fields("Sep").Value
                            lnTESep = lnTESep + rsReporte.Fields("Sep").Value
                            lnTDSep = lnTDSep + rsReporte.Fields("Sep").Value
                            lnTSep = lnTSep + rsReporte.Fields("Sep").Value
                            lnTSOct = lnTSOct + rsReporte.Fields("Oct").Value
                            lnTEOct = lnTEOct + rsReporte.Fields("Oct").Value
                            lnTDOct = lnTDOct + rsReporte.Fields("Oct").Value
                            lnTOct = lnTOct + rsReporte.Fields("Oct").Value
                            lnTSNov = lnTSNov + rsReporte.Fields("Nov").Value
                            lnTENov = lnTENov + rsReporte.Fields("Nov").Value
                            lnTDNov = lnTDNov + rsReporte.Fields("Nov").Value
                            lnTNov = lnTNov + rsReporte.Fields("Nov").Value
                            lnTSDic = lnTSDic + rsReporte.Fields("Dic").Value
                            lnTEDic = lnTEDic + rsReporte.Fields("Dic").Value
                            lnTDDic = lnTDDic + rsReporte.Fields("Dic").Value
                            lnTDic = lnTDic + rsReporte.Fields("Dic").Value
                            
                            lnTSTot = lnTSTot + lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                            lnTETot = lnTETot + lnTSTot
                            lnTDTot = lnTDTot + lnTSTot
                            lnTTot = lnTTot + lnTSTot
                            
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                       Loop
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(Round(lnTDTot / 12, 2))
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTDEne)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTDFeb)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTDMar)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTDAbr)
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTDMay)
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTDJun)
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTDJul)
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTDAgo)
                        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTDSep)
                        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTDOct)
                        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTDNov)
                        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTDDic)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTDTot / 12, 2)
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTDEne
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTDFeb
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTDMar
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTDAbr
                        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTDMay
                        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTDJun
                        oWorkSheet.Cells(iFila, iCol + 7).Value = lnTDJul
                        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTDAgo
                        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTDSep
                        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTDOct
                        oWorkSheet.Cells(iFila, iCol + 11).Value = lnTDNov
                        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTDDic
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
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(Round(lnTTot / 12, 2))
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTEne)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTFeb)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTMar)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTAbr)
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTMay)
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTJun)
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTJul)
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTAgo)
                        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTSep)
                        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTOct)
                        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTNov)
                        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTDic)
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 12
                        oWorkSheet.Cells(iFila, 2).Value = "Total: "
                        oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTTot / 12, 2)
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTEne
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTFeb
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTMar
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTAbr
                        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTMay
                        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTJun
                        oWorkSheet.Cells(iFila, iCol + 7).Value = lnTJul
                        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTAgo
                        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTSep
                        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTOct
                        oWorkSheet.Cells(iFila, iCol + 11).Value = lnTNov
                        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTDic
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
                    Else
                        oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                            If oWorkSheet.PageSetup.PrintArea <> "" Then
                               oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                            End If
                        oExcel.Visible = True
                        oWorkSheet.PrintPreview
                    End If
            Else
                If lbEsOpenOffice = True Then
                    'Abre el archivo ExcelOpenOffice
                    lcArchivoExcel = App.Path + "\Plantillas\HEgresoHosp.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
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
                        Call Feuille.getcellbyposition(4, 5).setFormula("Promedio Mensual")
                    Else
                        'Inicio de Impresion
                        oWorkSheet.Cells(2, 3).Value = ml_Titulo
                        oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
                        oWorkSheet.Cells(6, 5).Value = "Promedio Mensual"
                    End If
                iFila = 7
                iCol = 5
                lnTTot = 0: lnTEne = 0: lnTFeb = 0: lnTMar = 0: lnTAbr = 0: lnTMay = 0: lnTJun = 0: lnTJul = 0: lnTAgo = 0: lnTSep = 0: lnTOct = 0: lnTNov = 0: lnTDic = 0
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
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
                            lcServicio = rsReporte.Fields("dServicio").Value
                            lnIdServicio = rsReporte.Fields("IdServicio").Value
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("codServicio").Value & " - " & rsReporte.Fields("dServicio").Value)
                            Else
                                oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("codServicio").Value & " - " & rsReporte.Fields("dServicio").Value
                            End If
                            lnTSEne = lnTSEne + rsReporte.Fields("Ene").Value
                            lnTEEne = lnTEEne + rsReporte.Fields("Ene").Value
                            lnTDEne = lnTDEne + rsReporte.Fields("Ene").Value
                            lnTEne = lnTEne + rsReporte.Fields("Ene").Value
                            lnTSFeb = lnTSFeb + rsReporte.Fields("Feb").Value
                            lnTEFeb = lnTEFeb + rsReporte.Fields("Feb").Value
                            lnTDFeb = lnTDFeb + rsReporte.Fields("Feb").Value
                            lnTFeb = lnTFeb + rsReporte.Fields("Feb").Value
                            lnTSMar = lnTSMar + rsReporte.Fields("Mar").Value
                            lnTEMar = lnTEMar + rsReporte.Fields("Mar").Value
                            lnTDMar = lnTDMar + rsReporte.Fields("Mar").Value
                            lnTMar = lnTMar + rsReporte.Fields("Mar").Value
                            lnTSAbr = lnTSAbr + rsReporte.Fields("Abr").Value
                            lnTEAbr = lnTEAbr + rsReporte.Fields("Abr").Value
                            lnTDAbr = lnTDAbr + rsReporte.Fields("Abr").Value
                            lnTAbr = lnTAbr + rsReporte.Fields("Abr").Value
                            lnTSMay = lnTSMay + rsReporte.Fields("May").Value
                            lnTEMay = lnTEMay + rsReporte.Fields("May").Value
                            lnTDMay = lnTDMay + rsReporte.Fields("May").Value
                            lnTMay = lnTMay + rsReporte.Fields("May").Value
                            lnTSJun = lnTSJun + rsReporte.Fields("Jun").Value
                            lnTEJun = lnTEJun + rsReporte.Fields("Jun").Value
                            lnTDJun = lnTDJun + rsReporte.Fields("Jun").Value
                            lnTJun = lnTJun + rsReporte.Fields("Jun").Value
                            lnTSJul = lnTSJul + rsReporte.Fields("Jul").Value
                            lnTEJul = lnTEJul + rsReporte.Fields("Jul").Value
                            lnTDJul = lnTDJul + rsReporte.Fields("Jul").Value
                            lnTJul = lnTJul + rsReporte.Fields("Jul").Value
                            lnTSAgo = lnTSAgo + rsReporte.Fields("Ago").Value
                            lnTEAgo = lnTEAgo + rsReporte.Fields("Ago").Value
                            lnTDAgo = lnTDAgo + rsReporte.Fields("Ago").Value
                            lnTAgo = lnTAgo + rsReporte.Fields("Ago").Value
                            lnTSSep = lnTSSep + rsReporte.Fields("Sep").Value
                            lnTESep = lnTESep + rsReporte.Fields("Sep").Value
                            lnTDSep = lnTDSep + rsReporte.Fields("Sep").Value
                            lnTSep = lnTSep + rsReporte.Fields("Sep").Value
                            lnTSOct = lnTSOct + rsReporte.Fields("Oct").Value
                            lnTEOct = lnTEOct + rsReporte.Fields("Oct").Value
                            lnTDOct = lnTDOct + rsReporte.Fields("Oct").Value
                            lnTOct = lnTOct + rsReporte.Fields("Oct").Value
                            lnTSNov = lnTSNov + rsReporte.Fields("Nov").Value
                            lnTENov = lnTENov + rsReporte.Fields("Nov").Value
                            lnTDNov = lnTDNov + rsReporte.Fields("Nov").Value
                            lnTNov = lnTNov + rsReporte.Fields("Nov").Value
                            lnTSDic = lnTSDic + rsReporte.Fields("Dic").Value
                            lnTEDic = lnTEDic + rsReporte.Fields("Dic").Value
                            lnTDDic = lnTDDic + rsReporte.Fields("Dic").Value
                            lnTDic = lnTDic + rsReporte.Fields("Dic").Value
                            
                            lnTSTot = lnTSTot + lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                            lnTETot = lnTETot + lnTSTot
                            lnTDTot = lnTDTot + lnTSTot
                            lnTTot = lnTTot + lnTSTot
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(Round(lnTSTot / 12, 2))
                                Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTSEne)
                                Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTSFeb)
                                Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTSMar)
                                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTSAbr)
                                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTSMay)
                                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTSJun)
                                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTSJul)
                                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTSAgo)
                                Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTSSep)
                                Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTSOct)
                                Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTSNov)
                                Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTSDic)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTSTot / 12, 2)
                                oWorkSheet.Cells(iFila, iCol + 1).Value = lnTSEne
                                oWorkSheet.Cells(iFila, iCol + 2).Value = lnTSFeb
                                oWorkSheet.Cells(iFila, iCol + 3).Value = lnTSMar
                                oWorkSheet.Cells(iFila, iCol + 4).Value = lnTSAbr
                                oWorkSheet.Cells(iFila, iCol + 5).Value = lnTSMay
                                oWorkSheet.Cells(iFila, iCol + 6).Value = lnTSJun
                                oWorkSheet.Cells(iFila, iCol + 7).Value = lnTSJul
                                oWorkSheet.Cells(iFila, iCol + 8).Value = lnTSAgo
                                oWorkSheet.Cells(iFila, iCol + 9).Value = lnTSSep
                                oWorkSheet.Cells(iFila, iCol + 10).Value = lnTSOct
                                oWorkSheet.Cells(iFila, iCol + 11).Value = lnTSNov
                                oWorkSheet.Cells(iFila, iCol + 12).Value = lnTSDic
                            End If
                            iFila = iFila + 1
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                       Loop
                        If lbEsOpenOffice = True Then
                            Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":Q" & CStr(iFila))
                            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(Round(lnTETot / 12, 2))
                            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTEEne)
                            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTEFeb)
                            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTEMar)
                            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTEAbr)
                            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTEMay)
                            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTEJun)
                            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTEJul)
                            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTEAgo)
                            Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTESep)
                            Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTEOct)
                            Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTENov)
                            Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTEDic)
                        Else
                            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 4, iFila, iCol + 12
                            oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTETot / 12, 2)
                            oWorkSheet.Cells(iFila, iCol + 1).Value = lnTEEne
                            oWorkSheet.Cells(iFila, iCol + 2).Value = lnTEFeb
                            oWorkSheet.Cells(iFila, iCol + 3).Value = lnTEMar
                            oWorkSheet.Cells(iFila, iCol + 4).Value = lnTEAbr
                            oWorkSheet.Cells(iFila, iCol + 5).Value = lnTEMay
                            oWorkSheet.Cells(iFila, iCol + 6).Value = lnTEJun
                            oWorkSheet.Cells(iFila, iCol + 7).Value = lnTEJul
                            oWorkSheet.Cells(iFila, iCol + 8).Value = lnTEAgo
                            oWorkSheet.Cells(iFila, iCol + 9).Value = lnTESep
                            oWorkSheet.Cells(iFila, iCol + 10).Value = lnTEOct
                            oWorkSheet.Cells(iFila, iCol + 11).Value = lnTENov
                            oWorkSheet.Cells(iFila, iCol + 12).Value = lnTEDic
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
                            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(Round(lnTDTot / 12, 2))
                            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTDEne)
                            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTDFeb)
                            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTDMar)
                            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTDAbr)
                            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTDMay)
                            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTDJun)
                            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTDJul)
                            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTDAgo)
                            Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTDSep)
                            Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTDOct)
                            Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTDNov)
                            Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTDDic)
                        Else
                            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 12
                            oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTDTot / 12, 2)
                            oWorkSheet.Cells(iFila, iCol + 1).Value = lnTDEne
                            oWorkSheet.Cells(iFila, iCol + 2).Value = lnTDFeb
                            oWorkSheet.Cells(iFila, iCol + 3).Value = lnTDMar
                            oWorkSheet.Cells(iFila, iCol + 4).Value = lnTDAbr
                            oWorkSheet.Cells(iFila, iCol + 5).Value = lnTDMay
                            oWorkSheet.Cells(iFila, iCol + 6).Value = lnTDJun
                            oWorkSheet.Cells(iFila, iCol + 7).Value = lnTDJul
                            oWorkSheet.Cells(iFila, iCol + 8).Value = lnTDAgo
                            oWorkSheet.Cells(iFila, iCol + 9).Value = lnTDSep
                            oWorkSheet.Cells(iFila, iCol + 10).Value = lnTDOct
                            oWorkSheet.Cells(iFila, iCol + 11).Value = lnTDNov
                            oWorkSheet.Cells(iFila, iCol + 12).Value = lnTDDic
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
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(Round(lnTTot / 12, 2))
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTEne)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTFeb)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTMar)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTAbr)
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTMay)
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTJun)
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTJul)
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTAgo)
                        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(lnTSep)
                        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(lnTOct)
                        Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(lnTNov)
                        Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(lnTDic)
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 12
                        oWorkSheet.Cells(iFila, 2).Value = "Total: "
                        oWorkSheet.Cells(iFila, iCol + 0).Value = Round(lnTTot / 12, 2)
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTEne
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTFeb
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTMar
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTAbr
                        oWorkSheet.Cells(iFila, iCol + 5).Value = lnTMay
                        oWorkSheet.Cells(iFila, iCol + 6).Value = lnTJun
                        oWorkSheet.Cells(iFila, iCol + 7).Value = lnTJul
                        oWorkSheet.Cells(iFila, iCol + 8).Value = lnTAgo
                        oWorkSheet.Cells(iFila, iCol + 9).Value = lnTSep
                        oWorkSheet.Cells(iFila, iCol + 10).Value = lnTOct
                        oWorkSheet.Cells(iFila, iCol + 11).Value = lnTNov
                        oWorkSheet.Cells(iFila, iCol + 12).Value = lnTDic
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
                    Else
                        oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                            If oWorkSheet.PageSetup.PrintArea <> "" Then
                               oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                            End If
                        oExcel.Visible = True
                        oWorkSheet.PrintPreview
                    End If
            End If
        End If
    End If
        If lbEsOpenOffice = True Then
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
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

End Sub


Sub GenerarRecordsetTemporal()
    With rsReporte
          .Fields.Append "dDpto", adVarChar, 150, adFldIsNullable
          .Fields.Append "CodDpto", adInteger, 4, adFldIsNullable
          .Fields.Append "dEspecialidad", adVarChar, 150, adFldIsNullable
          .Fields.Append "CodEspecialidad", adInteger, 4, adFldIsNullable
          .Fields.Append "dServicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "CodServicio", adVarChar, 15, adFldIsNullable
          .Fields.Append "IdServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "Total", adInteger
          .Fields.Append "Ene", adInteger
          .Fields.Append "Feb", adInteger
          .Fields.Append "Mar", adInteger
          .Fields.Append "Abr", adInteger
          .Fields.Append "May", adInteger
          .Fields.Append "Jun", adInteger
          .Fields.Append "Jul", adInteger
          .Fields.Append "Ago", adInteger
          .Fields.Append "Sep", adInteger
          .Fields.Append "Oct", adInteger
          .Fields.Append "Nov", adInteger
          .Fields.Append "Dic", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
End Sub





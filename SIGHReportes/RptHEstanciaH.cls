VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHEstanciaH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Estancia Hospitalaria
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
    Dim oFormulario As New HEstanciaH
    oFormulario.Show 1
End Sub

'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Estancia Hospitalaria
'*****************Consideraciones: Solo ultimo Servicio, solo Hospitalizados
'*****************                 Si el Paciente esta solo 1 dia entonces ESTANCIA=1
'*****************
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
                            
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, ldFechaEgreso, lcHoraEgreso)

                            
                            Select Case lnMesDato
                            Case 1
                                 lnTSEne = lnTSEne + lnDiasEstancia
                                 lnTEEne = lnTEEne + lnDiasEstancia
                                 lnTDEne = lnTDEne + lnDiasEstancia
                                 lnTEne = lnTEne + lnDiasEstancia
                            Case 2
                                 lnTSFeb = lnTSFeb + lnDiasEstancia
                                 lnTEFeb = lnTEFeb + lnDiasEstancia
                                 lnTDFeb = lnTDFeb + lnDiasEstancia
                                 lnTFeb = lnTFeb + lnDiasEstancia
                            Case 3
                                 lnTSMar = lnTSMar + lnDiasEstancia
                                 lnTEMar = lnTEMar + lnDiasEstancia
                                 lnTDMar = lnTDMar + lnDiasEstancia
                                 lnTMar = lnTMar + lnDiasEstancia
                            Case 4
                                 lnTSAbr = lnTSAbr + lnDiasEstancia
                                 lnTEAbr = lnTEAbr + lnDiasEstancia
                                 lnTDAbr = lnTDAbr + lnDiasEstancia
                                 lnTAbr = lnTAbr + lnDiasEstancia
                            Case 5
                                 lnTSMay = lnTSMay + lnDiasEstancia
                                 lnTEMay = lnTEMay + lnDiasEstancia
                                 lnTDMay = lnTDMay + lnDiasEstancia
                                 lnTMay = lnTMay + lnDiasEstancia
                            Case 6
                                 lnTSJun = lnTSJun + lnDiasEstancia
                                 lnTEJun = lnTEJun + lnDiasEstancia
                                 lnTDJun = lnTDJun + lnDiasEstancia
                                 lnTJun = lnTJun + lnDiasEstancia
                            Case 7
                                 lnTSJul = lnTSJul + lnDiasEstancia
                                 lnTEJul = lnTEJul + lnDiasEstancia
                                 lnTDJul = lnTDJul + lnDiasEstancia
                                 lnTJul = lnTJul + lnDiasEstancia
                            Case 8
                                 lnTSAgo = lnTSAgo + lnDiasEstancia
                                 lnTEAgo = lnTEAgo + lnDiasEstancia
                                 lnTDAgo = lnTDAgo + lnDiasEstancia
                                 lnTAgo = lnTAgo + lnDiasEstancia
                            Case 9
                                 lnTSSep = lnTSSep + lnDiasEstancia
                                 lnTESep = lnTESep + lnDiasEstancia
                                 lnTDSep = lnTDSep + lnDiasEstancia
                                 lnTSep = lnTSep + lnDiasEstancia
                            Case 10
                                 lnTSOct = lnTSOct + lnDiasEstancia
                                 lnTEOct = lnTEOct + lnDiasEstancia
                                 lnTDOct = lnTDOct + lnDiasEstancia
                                 lnTOct = lnTOct + lnDiasEstancia
                            Case 11
                                 lnTSNov = lnTSNov + lnDiasEstancia
                                 lnTENov = lnTENov + lnDiasEstancia
                                 lnTDNov = lnTDNov + lnDiasEstancia
                                 lnTNov = lnTNov + lnDiasEstancia
                            Case 12
                                 lnTSDic = lnTSDic + lnDiasEstancia
                                 lnTEDic = lnTEDic + lnDiasEstancia
                                 lnTDDic = lnTDDic + lnDiasEstancia
                                 lnTDic = lnTDic + lnDiasEstancia
                            End Select
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        lnTSTot = lnTSTot + lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                        lnTETot = lnTETot + lnTSTot
                        lnTDTot = lnTDTot + lnTSTot
                        lnTTot = lnTTot + lnTSTot
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTSTot)
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
                            oWorkSheet.Cells(iFila, iCol + 0).Value = lnTSTot
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
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop
                    If lbEsOpenOffice = True Then
                        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":Q" & CStr(iFila))
                        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTETot)
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
                        oWorkSheet.Cells(iFila, iCol + 0).Value = lnTETot
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
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTDTot)
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
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTDTot
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
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTTot)
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
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTTot
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
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                        If oWorkSheet.PageSetup.PrintArea <> "" Then
                          oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                        End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
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
            Call Feuille.getcellbyposition(2, 2).setFormula(ml_Titulo)
        Else
            'Inicio de Impresion
            oWorkSheet.Cells(2, 3).Value = ml_Titulo
            oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
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
                'iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                 '  oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                 '  iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
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
                            
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, ldFechaEgreso, lcHoraEgreso)
                            
                            Select Case lnMesDato
                            Case 1
                                 lnTSEne = lnTSEne + lnDiasEstancia
                                 lnTEEne = lnTEEne + lnDiasEstancia
                                 lnTDEne = lnTDEne + lnDiasEstancia
                                 lnTEne = lnTEne + lnDiasEstancia
                            Case 2
                                 lnTSFeb = lnTSFeb + lnDiasEstancia
                                 lnTEFeb = lnTEFeb + lnDiasEstancia
                                 lnTDFeb = lnTDFeb + lnDiasEstancia
                                 lnTFeb = lnTFeb + lnDiasEstancia
                            Case 3
                                 lnTSMar = lnTSMar + lnDiasEstancia
                                 lnTEMar = lnTEMar + lnDiasEstancia
                                 lnTDMar = lnTDMar + lnDiasEstancia
                                 lnTMar = lnTMar + lnDiasEstancia
                            Case 4
                                 lnTSAbr = lnTSAbr + lnDiasEstancia
                                 lnTEAbr = lnTEAbr + lnDiasEstancia
                                 lnTDAbr = lnTDAbr + lnDiasEstancia
                                 lnTAbr = lnTAbr + lnDiasEstancia
                            Case 5
                                 lnTSMay = lnTSMay + lnDiasEstancia
                                 lnTEMay = lnTEMay + lnDiasEstancia
                                 lnTDMay = lnTDMay + lnDiasEstancia
                                 lnTMay = lnTMay + lnDiasEstancia
                            Case 6
                                 lnTSJun = lnTSJun + lnDiasEstancia
                                 lnTEJun = lnTEJun + lnDiasEstancia
                                 lnTDJun = lnTDJun + lnDiasEstancia
                                 lnTJun = lnTJun + lnDiasEstancia
                            Case 7
                                 lnTSJul = lnTSJul + lnDiasEstancia
                                 lnTEJul = lnTEJul + lnDiasEstancia
                                 lnTDJul = lnTDJul + lnDiasEstancia
                                 lnTJul = lnTJul + lnDiasEstancia
                            Case 8
                                 lnTSAgo = lnTSAgo + lnDiasEstancia
                                 lnTEAgo = lnTEAgo + lnDiasEstancia
                                 lnTDAgo = lnTDAgo + lnDiasEstancia
                                 lnTAgo = lnTAgo + lnDiasEstancia
                            Case 9
                                 lnTSSep = lnTSSep + lnDiasEstancia
                                 lnTESep = lnTESep + lnDiasEstancia
                                 lnTDSep = lnTDSep + lnDiasEstancia
                                 lnTSep = lnTSep + lnDiasEstancia
                            Case 10
                                 lnTSOct = lnTSOct + lnDiasEstancia
                                 lnTEOct = lnTEOct + lnDiasEstancia
                                 lnTDOct = lnTDOct + lnDiasEstancia
                                 lnTOct = lnTOct + lnDiasEstancia
                            Case 11
                                 lnTSNov = lnTSNov + lnDiasEstancia
                                 lnTENov = lnTENov + lnDiasEstancia
                                 lnTDNov = lnTDNov + lnDiasEstancia
                                 lnTNov = lnTNov + lnDiasEstancia
                            Case 12
                                 lnTSDic = lnTSDic + lnDiasEstancia
                                 lnTEDic = lnTEDic + lnDiasEstancia
                                 lnTDDic = lnTDDic + lnDiasEstancia
                                 lnTDic = lnTDic + lnDiasEstancia
                            End Select
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        lnTSTot = lnTSTot + lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                        lnTETot = lnTETot + lnTSTot
                        lnTDTot = lnTDTot + lnTSTot
                        lnTTot = lnTTot + lnTSTot
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTDTot)
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
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTDTot
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
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTTot)
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
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTTot
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
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                        If oWorkSheet.PageSetup.PrintArea <> "" Then
                          oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                        End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
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
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                lcDpto = rsReporte.Fields("dDpto").Value
                lnIdDpto = rsReporte.Fields("CodDpto").Value
                'oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                'iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                   'oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                   'iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
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
                            
                            
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, ldFechaEgreso, lcHoraEgreso)
                            
                            If lnIdEspecialidad = ml_idEspecialidad1 Then
                                Select Case lnMesDato
                                Case 1
                                     lnTSEne = lnTSEne + lnDiasEstancia
                                     lnTEEne = lnTEEne + lnDiasEstancia
                                     lnTDEne = lnTDEne + lnDiasEstancia
                                     lnTEne = lnTEne + lnDiasEstancia
                                Case 2
                                     lnTSFeb = lnTSFeb + lnDiasEstancia
                                     lnTEFeb = lnTEFeb + lnDiasEstancia
                                     lnTDFeb = lnTDFeb + lnDiasEstancia
                                     lnTFeb = lnTFeb + lnDiasEstancia
                                Case 3
                                     lnTSMar = lnTSMar + lnDiasEstancia
                                     lnTEMar = lnTEMar + lnDiasEstancia
                                     lnTDMar = lnTDMar + lnDiasEstancia
                                     lnTMar = lnTMar + lnDiasEstancia
                                Case 4
                                    lnTSAbr = lnTSAbr + lnDiasEstancia
                                    lnTEAbr = lnTEAbr + lnDiasEstancia
                                    lnTDAbr = lnTDAbr + lnDiasEstancia
                                    lnTAbr = lnTAbr + lnDiasEstancia
                                Case 5
                                     lnTSMay = lnTSMay + lnDiasEstancia
                                     lnTEMay = lnTEMay + lnDiasEstancia
                                     lnTDMay = lnTDMay + lnDiasEstancia
                                     lnTMay = lnTMay + lnDiasEstancia
                                Case 6
                                     lnTSJun = lnTSJun + lnDiasEstancia
                                     lnTEJun = lnTEJun + lnDiasEstancia
                                     lnTDJun = lnTDJun + lnDiasEstancia
                                     lnTJun = lnTJun + lnDiasEstancia
                                Case 7
                                     lnTSJul = lnTSJul + lnDiasEstancia
                                     lnTEJul = lnTEJul + lnDiasEstancia
                                     lnTDJul = lnTDJul + lnDiasEstancia
                                     lnTJul = lnTJul + lnDiasEstancia
                                Case 8
                                     lnTSAgo = lnTSAgo + lnDiasEstancia
                                     lnTEAgo = lnTEAgo + lnDiasEstancia
                                     lnTDAgo = lnTDAgo + lnDiasEstancia
                                     lnTAgo = lnTAgo + lnDiasEstancia
                                Case 9
                                     lnTSSep = lnTSSep + lnDiasEstancia
                                     lnTESep = lnTESep + lnDiasEstancia
                                     lnTDSep = lnTDSep + lnDiasEstancia
                                     lnTSep = lnTSep + lnDiasEstancia
                                Case 10
                                     lnTSOct = lnTSOct + lnDiasEstancia
                                     lnTEOct = lnTEOct + lnDiasEstancia
                                     lnTDOct = lnTDOct + lnDiasEstancia
                                     lnTOct = lnTOct + lnDiasEstancia
                                Case 11
                                     lnTSNov = lnTSNov + lnDiasEstancia
                                     lnTENov = lnTENov + lnDiasEstancia
                                     lnTDNov = lnTDNov + lnDiasEstancia
                                     lnTNov = lnTNov + lnDiasEstancia
                                Case 12
                                     lnTSDic = lnTSDic + lnDiasEstancia
                                     lnTEDic = lnTEDic + lnDiasEstancia
                                     lnTDDic = lnTDDic + lnDiasEstancia
                                     lnTDic = lnTDic + lnDiasEstancia
                                End Select
                            End If
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        If lnIdEspecialidad = ml_idEspecialidad1 Then
                            lnTSTot = lnTSTot + lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                            lnTETot = lnTETot + lnTSTot
                            lnTDTot = lnTDTot + lnTSTot
                            lnTTot = lnTTot + lnTSTot
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTSTot)
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
                                oWorkSheet.Cells(iFila, iCol + 0).Value = lnTSTot
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
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTTot)
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
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTTot
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
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                        If oWorkSheet.PageSetup.PrintArea <> "" Then
                          oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                        End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
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
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTDTot = 0: lnTDEne = 0: lnTDFeb = 0: lnTDMar = 0: lnTDAbr = 0: lnTDMay = 0: lnTDJun = 0: lnTDJul = 0: lnTDAgo = 0: lnTDSep = 0: lnTDOct = 0: lnTDNov = 0: lnTDDic = 0
                lcDpto = rsReporte.Fields("dDpto").Value
                lnIdDpto = rsReporte.Fields("CodDpto").Value
                'oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                'iFila = iFila + 1
                Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                   lnTETot = 0: lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                   lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                   lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                   'oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                   'iFila = iFila + 1
                   Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
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
                            
                            lnDiasEstancia = SIGHEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, ldFechaEgreso, lcHoraEgreso)
                            
                            If lnIdServicio = ml_idServicio1 Or lnIdServicio = ml_idServicio2 Then
                                Select Case lnMesDato
                                Case 1
                                     lnTSEne = lnTSEne + lnDiasEstancia
                                     lnTEEne = lnTEEne + lnDiasEstancia
                                     lnTDEne = lnTDEne + lnDiasEstancia
                                     lnTEne = lnTEne + lnDiasEstancia
                                Case 2
                                     lnTSFeb = lnTSFeb + lnDiasEstancia
                                     lnTEFeb = lnTEFeb + lnDiasEstancia
                                     lnTDFeb = lnTDFeb + lnDiasEstancia
                                     lnTFeb = lnTFeb + lnDiasEstancia
                                Case 3
                                     lnTSMar = lnTSMar + lnDiasEstancia
                                     lnTEMar = lnTEMar + lnDiasEstancia
                                     lnTDMar = lnTDMar + lnDiasEstancia
                                     lnTMar = lnTMar + lnDiasEstancia
                                Case 4
                                    lnTSAbr = lnTSAbr + lnDiasEstancia
                                    lnTEAbr = lnTEAbr + lnDiasEstancia
                                    lnTDAbr = lnTDAbr + lnDiasEstancia
                                    lnTAbr = lnTAbr + lnDiasEstancia
                                Case 5
                                     lnTSMay = lnTSMay + lnDiasEstancia
                                     lnTEMay = lnTEMay + lnDiasEstancia
                                     lnTDMay = lnTDMay + lnDiasEstancia
                                     lnTMay = lnTMay + lnDiasEstancia
                                Case 6
                                     lnTSJun = lnTSJun + lnDiasEstancia
                                     lnTEJun = lnTEJun + lnDiasEstancia
                                     lnTDJun = lnTDJun + lnDiasEstancia
                                     lnTJun = lnTJun + lnDiasEstancia
                                Case 7
                                     lnTSJul = lnTSJul + lnDiasEstancia
                                     lnTEJul = lnTEJul + lnDiasEstancia
                                     lnTDJul = lnTDJul + lnDiasEstancia
                                     lnTJul = lnTJul + lnDiasEstancia
                                Case 8
                                     lnTSAgo = lnTSAgo + lnDiasEstancia
                                     lnTEAgo = lnTEAgo + lnDiasEstancia
                                     lnTDAgo = lnTDAgo + lnDiasEstancia
                                     lnTAgo = lnTAgo + lnDiasEstancia
                                Case 9
                                     lnTSSep = lnTSSep + lnDiasEstancia
                                     lnTESep = lnTESep + lnDiasEstancia
                                     lnTDSep = lnTDSep + lnDiasEstancia
                                     lnTSep = lnTSep + lnDiasEstancia
                                Case 10
                                     lnTSOct = lnTSOct + lnDiasEstancia
                                     lnTEOct = lnTEOct + lnDiasEstancia
                                     lnTDOct = lnTDOct + lnDiasEstancia
                                     lnTOct = lnTOct + lnDiasEstancia
                                Case 11
                                     lnTSNov = lnTSNov + lnDiasEstancia
                                     lnTENov = lnTENov + lnDiasEstancia
                                     lnTDNov = lnTDNov + lnDiasEstancia
                                     lnTNov = lnTNov + lnDiasEstancia
                                Case 12
                                     lnTSDic = lnTSDic + lnDiasEstancia
                                     lnTEDic = lnTEDic + lnDiasEstancia
                                     lnTDDic = lnTDDic + lnDiasEstancia
                                     lnTDic = lnTDic + lnDiasEstancia
                                End Select
                            End If
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                        If lnIdServicio = ml_idServicio1 Or lnIdServicio = ml_idServicio2 Then
                            lnTSTot = lnTSTot + lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                            lnTETot = lnTETot + lnTSTot
                            lnTDTot = lnTDTot + lnTSTot
                            lnTTot = lnTTot + lnTSTot
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTSTot)
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
                                oWorkSheet.Cells(iFila, iCol + 0).Value = lnTSTot
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
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTTot)
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
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTTot
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
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                        If oWorkSheet.PageSetup.PrintArea <> "" Then
                          oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                        End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
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


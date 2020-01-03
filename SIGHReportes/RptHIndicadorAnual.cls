VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHIndicadorAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Indicador Anual
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
Dim ml_RangoMeses As Integer
Dim lcHoraEstanciaMax As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let RangoMeses(lValue As Integer)
    ml_RangoMeses = lValue
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
Property Let Anio(lValue As Integer)
    ml_lnAnio = lValue
End Property


Sub EjecutaFormulario()
    Dim oFormulario As New HIndicadorAnual
    oFormulario.Show 1
End Sub


Sub CrearReporteDetallado(lnHwnd As Long)
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
Dim lcServicio As String: Dim lnIdServicio As Long: Dim lcCodigoServicio As String
Dim lnMesDato As Integer: Dim lnDiasEstancia As Long
Dim lcFiltro As String: Dim lbContinua As Boolean
Dim lnEgresos As Long: Dim lnCamasProgramadas As Integer: Dim lnDiasCama As Integer
Dim lnDiasPaciente As Integer: Dim lnNroMeses As Integer
Dim lcHoraEgreso As String: Dim ldFechaEgreso As Date

Dim lnTEEgresos As Long: Dim lnTECamas As Long: Dim lnTEDiaCamas As Long: Dim lnTEDiaPaciente As Long: Dim lnTEDiaEstancia As Long
Dim lnTSEgresos As Long: Dim lnTSCamas As Long: Dim lnTSDiaCamas As Long: Dim lnTSDiaPaciente As Long: Dim lnTSDiaEstancia As Long
Dim lnTDEgresos As Long: Dim lnTDCamas As Long: Dim lnTDDiaCamas As Long: Dim lnTDDiaPaciente As Long: Dim lnTDDiaEstancia As Long
Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
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
    Set mrs_Tmp = AtencionesSeleccionarPorAnio(ml_FechaAltaMedica, ml_lnAnio)
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        Dim lcErr As String: Dim lnDiasPacientePorCien As Long
        lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
        GenerarRecordsetTemporal
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
           lcDpto = mrs_Tmp.Fields("dDpto").Value
           lnIdDpto = mrs_Tmp.Fields("CodDpto").Value
           Do While Not mrs_Tmp.EOF And lcDpto = mrs_Tmp.Fields("dDpto").Value And lnIdDpto = mrs_Tmp.Fields("CodDpto").Value
                lcEspecialidad = mrs_Tmp.Fields("dEspecialidad").Value
                lnIdEspecialidad = mrs_Tmp.Fields("CodEspecialidad").Value
                Do While Not mrs_Tmp.EOF And lcDpto = mrs_Tmp.Fields("dDpto").Value And lnIdDpto = mrs_Tmp.Fields("CodDpto").Value And lcEspecialidad = mrs_Tmp.Fields("dEspecialidad").Value And lnIdEspecialidad = mrs_Tmp.Fields("CodEspecialidad").Value
                    lnTSTot = 0: lnTSEne = 0: lnTSFeb = 0: lnTSMar = 0: lnTSAbr = 0: lnTSMay = 0: lnTSJun = 0: lnTSJul = 0: lnTSAgo = 0: lnTSSep = 0: lnTSOct = 0: lnTSNov = 0: lnTSDic = 0
                    lnCTSTot = 0: lnCTSEne = 0: lnCTSFeb = 0: lnCTSMar = 0: lnCTSAbr = 0: lnCTSMay = 0: lnCTSJun = 0: lnCTSJul = 0: lnCTSAgo = 0: lnCTSSep = 0: lnCTSOct = 0: lnCTSNov = 0: lnCTSDic = 0
                    lnTEEne = 0: lnTEFeb = 0: lnTEMar = 0: lnTEAbr = 0: lnTEMay = 0: lnTEJun = 0: lnTEJul = 0: lnTEAgo = 0: lnTESep = 0: lnTEOct = 0: lnTENov = 0: lnTEDic = 0
                    lcServicio = mrs_Tmp.Fields("dServicio").Value
                    lnIdServicio = mrs_Tmp.Fields("IdServicioEgreso").Value
                    lcCodigoServicio = mrs_Tmp.Fields("CodigoServicio").Value
                    lbContinua = True
                    Select Case ml_TipoReporte
                    Case 1   'Por departamento
                           lbContinua = IIf(lnIdDpto = ml_idDepartamento1, True, False)
                    Case 2   'Por un Servicio
                           lbContinua = IIf(mrs_Tmp.Fields("CodEspecialidad").Value = ml_idEspecialidad1, True, False)
                    Case 3   'Por dos Especialidades
                           lbContinua = IIf(mrs_Tmp.Fields("IdServicioEgreso").Value = ml_idServicio1 Or mrs_Tmp.Fields("IdServicioEgreso").Value = ml_idServicio2, True, False)
                    End Select
                    Do While Not mrs_Tmp.EOF And lcDpto = mrs_Tmp.Fields("dDpto").Value And lnIdDpto = mrs_Tmp.Fields("CodDpto").Value And lcEspecialidad = mrs_Tmp.Fields("dEspecialidad").Value And lnIdEspecialidad = mrs_Tmp.Fields("CodEspecialidad").Value And lcServicio = mrs_Tmp.Fields("dServicio").Value And lnIdServicio = mrs_Tmp.Fields("IdServicioEgreso").Value
                        lcHoraEgreso = ""
                        If ml_FechaAltaMedica Then
                           lnMesDato = Month(mrs_Tmp.Fields("FechaEgreso").Value)
                           ldFechaEgreso = mrs_Tmp.Fields("FechaEgreso").Value
                           If Not IsNull(mrs_Tmp.Fields("HoraEgreso").Value) Then
                              lcHoraEgreso = mrs_Tmp.Fields("HoraEgreso").Value
                           End If
                        Else
                           lnMesDato = Month(mrs_Tmp.Fields("FechaEgresoAdministrativo").Value)
                           ldFechaEgreso = mrs_Tmp.Fields("FechaEgresoAdministrativo").Value
                           If Not IsNull(mrs_Tmp.Fields("HoraEgresoAdministrativo").Value) Then
                              lcHoraEgreso = mrs_Tmp.Fields("HoraEgresoAdministrativo").Value
                           End If
                        End If
                        
                        'Estancia
                        lnDiasEstancia = sighEntidades.DiasDeEstanciaEnHospitalizacionEmergencia(mrs_Tmp!FechaIngreso, mrs_Tmp!HoraIngreso, ldFechaEgreso, lcHoraEgreso)
                        
                        'Dias Paciente
                        lnDiasPaciente = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(mrs_Tmp!FechaIngreso, mrs_Tmp!HoraIngreso, ldFechaEgreso, lcHoraEgreso, lcHoraEstanciaMax)
                        '
                        Select Case lnMesDato
                        Case 1
                             lnTSEne = lnTSEne + lnDiasEstancia
                             lnCTSEne = lnCTSEne + 1
                             lnTEEne = lnTEEne + lnDiasPaciente
                        Case 2
                             lnTSFeb = lnTSFeb + lnDiasEstancia
                             lnCTSFeb = lnCTSFeb + 1
                             lnTEFeb = lnTEFeb + lnDiasPaciente
                        Case 3
                             lnTSMar = lnTSMar + lnDiasEstancia
                             lnCTSMar = lnCTSMar + 1
                             lnTEMar = lnTEMar + lnDiasPaciente
                             
                        Case 4
                             lnTSAbr = lnTSAbr + lnDiasEstancia
                             lnCTSAbr = lnCTSAbr + 1
                             lnTEAbr = lnTEAbr + lnDiasPaciente
                        Case 5
                             lnTSMay = lnTSMay + lnDiasEstancia
                             lnCTSMay = lnCTSMay + 1
                             lnTEMay = lnTEMay + lnDiasPaciente
                        Case 6
                             lnTSJun = lnTSJun + lnDiasEstancia
                             lnCTSJun = lnCTSJun + 1
                             lnTEJun = lnTEJun + lnDiasPaciente
                        Case 7
                             lnTSJul = lnTSJul + lnDiasEstancia
                             lnCTSJul = lnCTSJul + 1
                             lnTEJul = lnTEJul + lnDiasPaciente
                        Case 8
                             lnTSAgo = lnTSAgo + lnDiasEstancia
                             lnCTSAgo = lnCTSAgo + 1
                             lnTEAgo = lnTEAgo + lnDiasPaciente
                        Case 9
                             lnTSSep = lnTSSep + lnDiasEstancia
                             lnCTSSep = lnCTSSep + 1
                             lnTESep = lnTESep + lnDiasPaciente
                        Case 10
                             lnTSOct = lnTSOct + lnDiasEstancia
                             lnCTSOct = lnCTSOct + 1
                             lnTEOct = lnTEOct + lnDiasPaciente
                        Case 11
                             lnTSNov = lnTSNov + lnDiasEstancia
                             lnCTSNov = lnCTSNov + 1
                             lnTENov = lnTENov + lnDiasPaciente
                        Case 12
                             lnTSDic = lnTSDic + lnDiasEstancia
                             lnCTSDic = lnCTSDic + 1
                             lnTEDic = lnTEDic + lnDiasPaciente
                        End Select
                        mrs_Tmp.MoveNext
                        If mrs_Tmp.EOF Then
                           Exit Do
                        End If
                    Loop
                    If lbContinua Then
                       lnNroMeses = 1
                       Select Case ml_RangoMeses
                       Case 0        'Anual
                            lnEgresos = lnCTSEne + lnCTSFeb + lnCTSMar + lnCTSAbr + lnCTSMay + lnCTSJun + lnCTSJul + lnCTSAgo + lnCTSSep + lnCTSOct + lnCTSNov + lnCTSDic
                            lnDiasEstancia = lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun + lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1) + sighEntidades.diasdelmes(ml_lnAnio, 2) + sighEntidades.diasdelmes(ml_lnAnio, 3) + sighEntidades.diasdelmes(ml_lnAnio, 4) + sighEntidades.diasdelmes(ml_lnAnio, 5) + sighEntidades.diasdelmes(ml_lnAnio, 6) + sighEntidades.diasdelmes(ml_lnAnio, 7) + sighEntidades.diasdelmes(ml_lnAnio, 8) + sighEntidades.diasdelmes(ml_lnAnio, 9) + sighEntidades.diasdelmes(ml_lnAnio, 10) + sighEntidades.diasdelmes(ml_lnAnio, 11) + sighEntidades.diasdelmes(ml_lnAnio, 12)
                            lnDiasPaciente = lnTEEne + lnTEFeb + lnTEMar + lnTEAbr + lnTEMay + lnTEJun + lnTEJul + lnTEAgo + lnTESep + lnTEOct + lnTENov + lnTEDic
                            lnNroMeses = 12
                       Case 1        'Primer Semestre
                            lnEgresos = lnCTSEne + lnCTSFeb + lnCTSMar + lnCTSAbr + lnCTSMay + lnCTSJun
                            lnDiasEstancia = lnTSEne + lnTSFeb + lnTSMar + lnTSAbr + lnTSMay + lnTSJun
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1) + sighEntidades.diasdelmes(ml_lnAnio, 2) + sighEntidades.diasdelmes(ml_lnAnio, 3) + sighEntidades.diasdelmes(ml_lnAnio, 4) + sighEntidades.diasdelmes(ml_lnAnio, 5) + sighEntidades.diasdelmes(ml_lnAnio, 6)
                            lnDiasPaciente = lnTEEne + lnTEFeb + lnTEMar + lnTEAbr + lnTEMay + lnTEJun
                            lnNroMeses = 6
                      Case 2        'Segundo Semestre
                            lnEgresos = lnCTSJul + lnCTSAgo + lnCTSSep + lnCTSOct + lnCTSNov + lnCTSDic
                            lnDiasEstancia = lnTSJul + lnTSAgo + lnTSSep + lnTSOct + lnTSNov + lnTSDic
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 7) + sighEntidades.diasdelmes(ml_lnAnio, 8) + sighEntidades.diasdelmes(ml_lnAnio, 9) + sighEntidades.diasdelmes(ml_lnAnio, 10) + sighEntidades.diasdelmes(ml_lnAnio, 11) + sighEntidades.diasdelmes(ml_lnAnio, 12)
                            lnDiasPaciente = lnTEJul + lnTEAgo + lnTESep + lnTEOct + lnTENov + lnTEDic
                            lnNroMeses = 6
                       Case 3        'Primer Trimestre
                            lnEgresos = lnCTSEne + lnCTSFeb + lnCTSMar
                            lnDiasEstancia = lnTSEne + lnTSFeb + lnTSMar
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1) + sighEntidades.diasdelmes(ml_lnAnio, 2) + sighEntidades.diasdelmes(ml_lnAnio, 3)
                            lnDiasPaciente = lnTEEne + lnTEFeb + lnTEMar
                            lnNroMeses = 3
                       Case 4       'Segundo Trimestre
                            lnEgresos = lnCTSAbr + lnCTSMay + lnCTSJun
                            lnDiasEstancia = lnTSAbr + lnTSMay + lnTSJun
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 4) + sighEntidades.diasdelmes(ml_lnAnio, 5) + sighEntidades.diasdelmes(ml_lnAnio, 6)
                            lnDiasPaciente = lnTEAbr + lnTEMay + lnTEJun
                            lnNroMeses = 3
                       Case 5        'Tercer Trimestre
                            lnEgresos = lnCTSJul + lnCTSAgo + lnCTSSep
                            lnDiasEstancia = lnTSJul + lnTSAgo + lnTSSep
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 7) + sighEntidades.diasdelmes(ml_lnAnio, 8) + sighEntidades.diasdelmes(ml_lnAnio, 9)
                            lnDiasPaciente = lnTEJul + lnTEAgo + lnTESep
                            lnNroMeses = 3
                       Case 6       'Cuarto Trimestre
                            lnEgresos = lnCTSOct + lnCTSNov + lnCTSDic
                            lnDiasEstancia = lnTSOct + lnTSNov + lnTSDic
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 10) + sighEntidades.diasdelmes(ml_lnAnio, 11) + sighEntidades.diasdelmes(ml_lnAnio, 12)
                            lnDiasPaciente = lnTEOct + lnTENov + lnTEDic
                            lnNroMeses = 3
                       Case 7        'Enero
                            lnEgresos = lnCTSEne
                            lnDiasEstancia = lnTSEne
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1)
                            lnDiasPaciente = lnTEEne
                       Case 8       'Febrero
                            lnEgresos = lnCTSFeb
                            lnDiasEstancia = lnTSFeb
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 2)
                            lnDiasPaciente = lnTEFeb
                       Case 9       'Marzo
                            lnEgresos = lnCTSMar
                            lnDiasEstancia = lnTSMar
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 3)
                            lnDiasPaciente = lnTEMar
                            
                       Case 10       'Abril
                            lnEgresos = lnCTSAbr
                            lnDiasEstancia = lnTSAbr
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 4)
                            lnDiasPaciente = lnTEAbr
                            
                       Case 11       'Mayo
                            lnEgresos = lnCTSMay
                            lnDiasEstancia = lnTSMay
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 5)
                            lnDiasPaciente = lnTEMay
                       Case 12       'Junio
                            lnEgresos = lnCTSJun
                            lnDiasEstancia = lnTSJun
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 6)
                            lnDiasPaciente = lnTEJun
                       Case 13       'Julio
                            lnEgresos = lnCTSJul
                            lnDiasEstancia = lnTSJul
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 7)
                            lnDiasPaciente = lnTEJul
                       Case 14       'Agosto
                            lnEgresos = lnCTSAgo
                            lnDiasEstancia = lnTSAgo
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 8)
                            lnDiasPaciente = lnTEAgo
                       Case 15       'Setiembre
                            lnEgresos = lnCTSSep
                            lnDiasEstancia = lnTSSep
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 9)
                            lnDiasPaciente = lnTESep
                       Case 16       'Octubre
                            lnEgresos = lnCTSOct
                            lnDiasEstancia = lnTSOct
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 10)
                            lnDiasPaciente = lnTEOct
                       Case 17       'Noviembre
                            lnEgresos = lnCTSNov
                            lnDiasEstancia = lnTSNov
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 11)
                            lnDiasPaciente = lnTENov
                       Case 18       'Diciembre
                            lnEgresos = lnCTSDic
                            lnDiasEstancia = lnTSDic
                            lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                            lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 12)
                            lnDiasPaciente = lnTEDic
                       End Select
                       If lnEgresos > 0 Or lnCamasProgramadas > 0 Or lnDiasPaciente > 0 Or lnDiasEstancia > 0 Then
                            rsReporte.AddNew
                            rsReporte.Fields("dDpto").Value = lcDpto
                            rsReporte.Fields("CodDpto").Value = lnIdDpto
                            rsReporte.Fields("dEspecialidad").Value = lcEspecialidad
                            rsReporte.Fields("CodEspecialidad").Value = lnIdEspecialidad
                            rsReporte.Fields("dServicio").Value = lcServicio
                            rsReporte.Fields("IdServicioEgreso").Value = lnIdServicio
                            rsReporte.Fields("CodServicio").Value = lcCodigoServicio
                            rsReporte.Fields("Egresos").Value = lnEgresos
                            'rsReporte.Fields("Camas").Value = lnCamasProgramadas
                            'rsReporte.Fields("DiaCamas").Value = lnCamasProgramadas * lnDiasCama
                            rsReporte.Fields("DiaPaciente").Value = lnDiasPaciente
                            rsReporte.Fields("DiaEstancia").Value = lnDiasEstancia
                            rsReporte.Update
                       End If
                    End If
                    If mrs_Tmp.EOF Then
                       Exit Do
                    End If
                Loop
                If mrs_Tmp.EOF Then
                   Exit Do
                End If
           Loop
        Loop
        If rsReporte.RecordCount = 0 Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
           'Busca Camas Programadas
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
                 lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                 lnNroMeses = 1
                 Select Case ml_RangoMeses
                 Case 0        'Anual
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1) + sighEntidades.diasdelmes(ml_lnAnio, 2) + sighEntidades.diasdelmes(ml_lnAnio, 3) + sighEntidades.diasdelmes(ml_lnAnio, 4) + sighEntidades.diasdelmes(ml_lnAnio, 5) + sighEntidades.diasdelmes(ml_lnAnio, 6) + sighEntidades.diasdelmes(ml_lnAnio, 7) + sighEntidades.diasdelmes(ml_lnAnio, 8) + sighEntidades.diasdelmes(ml_lnAnio, 9) + sighEntidades.diasdelmes(ml_lnAnio, 10) + sighEntidades.diasdelmes(ml_lnAnio, 11) + sighEntidades.diasdelmes(ml_lnAnio, 12)
                      lnNroMeses = 12
                 Case 1        'Primer Semestre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1) + sighEntidades.diasdelmes(ml_lnAnio, 2) + sighEntidades.diasdelmes(ml_lnAnio, 3) + sighEntidades.diasdelmes(ml_lnAnio, 4) + sighEntidades.diasdelmes(ml_lnAnio, 5) + sighEntidades.diasdelmes(ml_lnAnio, 6)
                      lnNroMeses = 6
                Case 2        'Segundo Semestre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 7) + sighEntidades.diasdelmes(ml_lnAnio, 8) + sighEntidades.diasdelmes(ml_lnAnio, 9) + sighEntidades.diasdelmes(ml_lnAnio, 10) + sighEntidades.diasdelmes(ml_lnAnio, 11) + sighEntidades.diasdelmes(ml_lnAnio, 12)
                      lnNroMeses = 6
                 Case 3        'Primer Trimestre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1) + sighEntidades.diasdelmes(ml_lnAnio, 2) + sighEntidades.diasdelmes(ml_lnAnio, 3)
                      lnNroMeses = 3
                 Case 4       'Segundo Trimestre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 4) + sighEntidades.diasdelmes(ml_lnAnio, 5) + sighEntidades.diasdelmes(ml_lnAnio, 6)
                      lnNroMeses = 3
                 Case 5        'Tercer Trimestre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 7) + sighEntidades.diasdelmes(ml_lnAnio, 8) + sighEntidades.diasdelmes(ml_lnAnio, 9)
                      lnNroMeses = 3
                 Case 6       'Cuarto Trimestre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio) + CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 10) + sighEntidades.diasdelmes(ml_lnAnio, 11) + sighEntidades.diasdelmes(ml_lnAnio, 12)
                      lnNroMeses = 3
                 Case 7        'Enero
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 1, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 1)
                 Case 8       'Febrero
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 2, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 2)
                 Case 9       'Marzo
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 3, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 3)
                 Case 10       'Abril
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 4, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 4)
                 Case 11       'Mayo
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 5, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 5)
                 Case 12       'Junio
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 6, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 6)
                 Case 13       'Julio
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 7, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 7)
                 Case 14       'Agosto
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 8, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 8)
                 Case 15       'Setiembre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 9, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 9)
                 Case 16       'Octubre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 10, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 10)
                 Case 17       'Noviembre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 11, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 11)
                 Case 18       'Diciembre
                      lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                      lnDiasCama = sighEntidades.diasdelmes(ml_lnAnio, 12)
                 End Select
                'lnCamasProgramadas = CamasProgramadasPorServicioAnioMes(ml_lnAnio, 12, lnIdServicio)
                lnDiasEstancia = rsReporte.Fields("DiaEstancia").Value
                lnEgresos = rsReporte.Fields("Egresos").Value
                lnDiasPaciente = rsReporte.Fields("DiaPaciente").Value
                rsReporte.Fields("Camas").Value = lnCamasProgramadas
                rsReporte.Fields("DiaCamas").Value = lnCamasProgramadas * lnDiasCama
                If lnEgresos > 0 Then
                   rsReporte.Fields("PromPermanencia").Value = Round((lnDiasEstancia / lnEgresos), 2)
                End If
                If lnEgresos > 0 Then
                   rsReporte.Fields("InteSustitucion").Value = Round(((lnCamasProgramadas * lnDiasCama) - lnDiasPaciente) / lnEgresos, 2)
                End If
                If lnCamasProgramadas > 0 And lnDiasCama > 0 Then
                   rsReporte.Fields("PorcOcupacion").Value = Round((lnDiasPaciente / (lnCamasProgramadas * lnDiasCama)) * 100, 2)
                End If
                If lnCamasProgramadas > 0 Then
                   rsReporte.Fields("RendCama").Value = Round((lnEgresos / lnCamasProgramadas) / lnNroMeses, 2)
                End If
                rsReporte.Update
                rsReporte.MoveNext
           Loop

            If ml_TipoReporte = 1 Then                 'Por Departamentos
                If lbEsOpenOffice = True Then
                    'Abre el archivo ExcelOpenOffice
                    lcArchivoExcel = App.Path + "\Plantillas\Hindicadores.ods"
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
                    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Hindicadores.xls")
                    oWorkBookPlantilla.Worksheets("Hindicadores").Copy Before:=oWorkBook.Sheets(1)
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
                lnCTTot = 0: lnCCTEne = 0: lnCCTFeb = 0: lnCCTMar = 0: lnCCTAbr = 0: lnCCTMay = 0: lnCCTJun = 0: lnCCTJul = 0: lnCCTAgo = 0
                lnTDEgresos = 0: lnTDCamas = 0: lnTDDiaCamas = 0: lnTDDiaPaciente = 0: lnTDDiaEstancia = 0
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnCTDTot = 0: lnCCTDEne = 0: lnCCTDFeb = 0: lnCCTDMar = 0: lnCCTDAbr = 0: lnCCTDMay = 0: lnCCTDJun = 0: lnCCTDJul = 0: lnCCTDAgo = 0
                    lnTSEgresos = 0: lnTSCamas = 0: lnTSDiaCamas = 0: lnTSDiaPaciente = 0: lnTSDiaEstancia = 0
                    lcDpto = rsReporte.Fields("dDpto").Value
                    lnIdDpto = rsReporte.Fields("CodDpto").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("dDpto").Value)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                    End If
 '                   iFila = iFila + 1
                    Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                       lnCTETot = 0: lnCCTEEne = 0: lnCCTEFeb = 0: lnCCTEMar = 0: lnCCTEAbr = 0: lnCCTEMay = 0: lnCCTEJun = 0: lnCCTEJul = 0: lnCCTEAgo = 0
                       lnTEEgresos = 0: lnTECamas = 0: lnTEDiaCamas = 0: lnTEDiaPaciente = 0: lnTEDiaEstancia = 0
                       lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                       lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
'                       oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
'                       iFila = iFila + 1
                       Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                            lcServicio = rsReporte.Fields("dServicio").Value
                            lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                            lnTEEgresos = lnTEEgresos + rsReporte.Fields("Egresos").Value
                            lnTSEgresos = lnTSEgresos + rsReporte.Fields("Egresos").Value
                            lnTDEgresos = lnTDEgresos + rsReporte.Fields("Egresos").Value
                            
                            lnTECamas = lnTECamas + rsReporte.Fields("Camas").Value
                            lnTSCamas = lnTSCamas + rsReporte.Fields("Camas").Value
                            lnTDCamas = lnTDCamas + rsReporte.Fields("Camas").Value
                            
                            lnTEDiaCamas = lnTEDiaCamas + rsReporte.Fields("DiaCamas").Value
                            lnTSDiaCamas = lnTSDiaCamas + rsReporte.Fields("DiaCamas").Value
                            lnTDDiaCamas = lnTDDiaCamas + rsReporte.Fields("DiaCamas").Value
                            
                            lnTEDiaPaciente = lnTEDiaPaciente + rsReporte.Fields("Diapaciente").Value
                            lnTSDiaPaciente = lnTSDiaPaciente + rsReporte.Fields("Diapaciente").Value
                            lnTDDiaPaciente = lnTDDiaPaciente + rsReporte.Fields("Diapaciente").Value
                            
                            lnTEDiaEstancia = lnTEDiaEstancia + rsReporte.Fields("DiaEstancia").Value
                            lnTSDiaEstancia = lnTSDiaEstancia + rsReporte.Fields("DiaEstancia").Value
                            lnTDDiaEstancia = lnTDDiaEstancia + rsReporte.Fields("DiaEstancia").Value
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                       Loop

                       
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
'                    iFila = iFila + 1
'                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 8
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTSEgresos)
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTSCamas)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTSDiaCamas)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTSDiaPaciente)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTSDiaEstancia)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 0).Value = lnTSEgresos
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTSCamas
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTSDiaCamas
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTSDiaPaciente
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTSDiaEstancia
                    End If
                    lnTDCamas = lnTSCamas
                    If lnTSEgresos > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round((lnTSDiaEstancia / lnTSEgresos), 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 5).Value = Round((lnTSDiaEstancia / lnTSEgresos), 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                        End If
                    End If
                    If lnTSEgresos > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(((lnTSDiaCamas - lnTSDiaPaciente) / lnTSEgresos), 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 6).Value = Round(((lnTSDiaCamas - lnTSDiaPaciente) / lnTSEgresos), 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                        End If
                    End If
                    If lnTSDiaCamas > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round((lnTSDiaPaciente / lnTSDiaCamas) * 100, 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 7).Value = Round((lnTSDiaPaciente / lnTSDiaCamas) * 100, 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                        End If
                    End If
                    If lnTSCamas > 0 And lnNroMeses > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(((lnTSEgresos / lnTSCamas) / lnNroMeses), 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 8).Value = Round(((lnTSEgresos / lnTSCamas) / lnNroMeses), 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                        End If
                    End If
                    iFila = iFila + 1
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                 Loop
                 iFila = iFila + 1
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":M" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTDEgresos)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTDCamas)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTDDiaCamas)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTDDiaPaciente)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTDDiaEstancia)
                 Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 8
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTDEgresos
                    oWorkSheet.Cells(iFila, iCol + 1).Value = lnTDCamas
                    oWorkSheet.Cells(iFila, iCol + 2).Value = lnTDDiaCamas
                    oWorkSheet.Cells(iFila, iCol + 3).Value = lnTDDiaPaciente
                    oWorkSheet.Cells(iFila, iCol + 4).Value = lnTDDiaEstancia
                End If
                 If lnTDEgresos > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round((lnTDDiaEstancia / lnTDEgresos), 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 5).Value = Round((lnTDDiaEstancia / lnTDEgresos), 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                    End If
                 End If
                 If lnTDEgresos > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(((lnTDDiaCamas - lnTDDiaPaciente) / lnTDEgresos), 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 6).Value = Round(((lnTDDiaCamas - lnTDDiaPaciente) / lnTDEgresos), 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                    End If
                 End If
                 If lnTDDiaCamas > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round((lnTDDiaPaciente / lnTDDiaCamas) * 100, 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 7).Value = Round((lnTDDiaPaciente / lnTDDiaCamas) * 100, 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                    End If
                 End If
                 If lnTDCamas > 0 And lnNroMeses > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(((lnTDEgresos / lnTDCamas) / lnNroMeses), 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 8).Value = Round(((lnTDEgresos / lnTDCamas) / lnNroMeses), 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                    End If
                 End If
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 13
                    PrintArea(0).EndRow = iFila
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    'Falta que salga el nombre del responsable
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                    End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
                    'oWorkSheet.PrintOut
                End If
            Else
                If lbEsOpenOffice = True Then
                    'Abre el archivo ExcelOpenOffice
                    lcArchivoExcel = App.Path + "\Plantillas\Hindicadores.ods"
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
                    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Hindicadores.xls")
                    oWorkBookPlantilla.Worksheets("Hindicadores").Copy Before:=oWorkBook.Sheets(1)
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
                lnCTTot = 0: lnCCTEne = 0: lnCCTFeb = 0: lnCCTMar = 0: lnCCTAbr = 0: lnCCTMay = 0: lnCCTJun = 0: lnCCTJul = 0: lnCCTAgo = 0
                lnTDEgresos = 0: lnTDCamas = 0: lnTDDiaCamas = 0: lnTDDiaPaciente = 0: lnTDDiaEstancia = 0
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnCTDTot = 0: lnCCTDEne = 0: lnCCTDFeb = 0: lnCCTDMar = 0: lnCCTDAbr = 0: lnCCTDMay = 0: lnCCTDJun = 0: lnCCTDJul = 0: lnCCTDAgo = 0
                    lnTSEgresos = 0: lnTSCamas = 0: lnTSDiaCamas = 0: lnTSDiaPaciente = 0: lnTSDiaEstancia = 0
                    lcDpto = rsReporte.Fields("dDpto").Value
                    lnIdDpto = rsReporte.Fields("CodDpto").Value
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("dDpto").Value)
                    Else
                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("dDpto").Value
                    End If
                    iFila = iFila + 1
                    Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value
                       lnCTETot = 0: lnCCTEEne = 0: lnCCTEFeb = 0: lnCCTEMar = 0: lnCCTEAbr = 0: lnCCTEMay = 0: lnCCTEJun = 0: lnCCTEJul = 0: lnCCTEAgo = 0
                       lnTEEgresos = 0: lnTECamas = 0: lnTEDiaCamas = 0: lnTEDiaPaciente = 0: lnTEDiaEstancia = 0
                       lcEspecialidad = rsReporte.Fields("dEspecialidad").Value
                       lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("dEspecialidad").Value)
                        Else
                            oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("dEspecialidad").Value
                        End If
                       iFila = iFila + 1
                       Do While Not rsReporte.EOF And lcDpto = rsReporte.Fields("dDpto").Value And lnIdDpto = rsReporte.Fields("CodDpto").Value And lcEspecialidad = rsReporte.Fields("dEspecialidad").Value And lnIdEspecialidad = rsReporte.Fields("CodEspecialidad").Value
                            lcServicio = rsReporte.Fields("dServicio").Value
                            lnIdServicio = rsReporte.Fields("IdServicioEgreso").Value
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("CodServicio").Value & " - " & rsReporte.Fields("dServicio").Value)
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("Egresos").Value)
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(rsReporte.Fields("Camas").Value)
                                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte.Fields("DiaCamas").Value)
                                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(rsReporte.Fields("Diapaciente").Value)
                                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(rsReporte.Fields("DiaEstancia").Value)
                                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(rsReporte.Fields("PromPermanencia").Value)
                                Call Feuille.getcellbyposition(10, iFila - 1).setFormula(rsReporte.Fields("InteSustitucion").Value)
                                Call Feuille.getcellbyposition(11, iFila - 1).setFormula(rsReporte.Fields("PorcOcupacion").Value)
                                Call Feuille.getcellbyposition(12, iFila - 1).setFormula(rsReporte.Fields("RendCama").Value)
                            Else
                                oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("CodServicio").Value & " - " & rsReporte.Fields("dServicio").Value
                                oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("Egresos").Value
                                oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("Camas").Value
                                oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields("DiaCamas").Value
                                oWorkSheet.Cells(iFila, 8).Value = rsReporte.Fields("Diapaciente").Value
                                oWorkSheet.Cells(iFila, 9).Value = rsReporte.Fields("DiaEstancia").Value
                                oWorkSheet.Cells(iFila, 10).Value = rsReporte.Fields("PromPermanencia").Value
                                oWorkSheet.Cells(iFila, 11).Value = rsReporte.Fields("InteSustitucion").Value
                                oWorkSheet.Cells(iFila, 12).Value = rsReporte.Fields("PorcOcupacion").Value
                                oWorkSheet.Cells(iFila, 13).Value = rsReporte.Fields("RendCama").Value
                            End If
                            lnTEEgresos = lnTEEgresos + rsReporte.Fields("Egresos").Value
                            lnTSEgresos = lnTSEgresos + rsReporte.Fields("Egresos").Value
                            lnTDEgresos = lnTDEgresos + rsReporte.Fields("Egresos").Value
                            
                            lnTECamas = lnTECamas + rsReporte.Fields("Camas").Value
                            lnTSCamas = lnTSCamas + rsReporte.Fields("Camas").Value
                            lnTDCamas = lnTDCamas + rsReporte.Fields("Camas").Value
                            
                            lnTEDiaCamas = lnTEDiaCamas + rsReporte.Fields("DiaCamas").Value
                            lnTSDiaCamas = lnTSDiaCamas + rsReporte.Fields("DiaCamas").Value
                            lnTDDiaCamas = lnTDDiaCamas + rsReporte.Fields("DiaCamas").Value
                            
                            lnTEDiaPaciente = lnTEDiaPaciente + rsReporte.Fields("Diapaciente").Value
                            lnTSDiaPaciente = lnTSDiaPaciente + rsReporte.Fields("Diapaciente").Value
                            lnTDDiaPaciente = lnTDDiaPaciente + rsReporte.Fields("Diapaciente").Value
                            
                            lnTEDiaEstancia = lnTEDiaEstancia + rsReporte.Fields("DiaEstancia").Value
                            lnTSDiaEstancia = lnTSDiaEstancia + rsReporte.Fields("DiaEstancia").Value
                            lnTDDiaEstancia = lnTDDiaEstancia + rsReporte.Fields("DiaEstancia").Value
                            
                            iFila = iFila + 1
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                       Loop
                        If lbEsOpenOffice = True Then
                            Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":M" & CStr(iFila))
                            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                            Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnTEEgresos)
                            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnTECamas)
                            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnTEDiaCamas)
                            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(lnTEDiaPaciente)
                            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(lnTEDiaEstancia)
                        Else
                           mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 4, iFila, iCol + 8
                           oWorkSheet.Cells(iFila, 5).Value = lnTEEgresos
                           oWorkSheet.Cells(iFila, 6).Value = lnTECamas
                           oWorkSheet.Cells(iFila, 7).Value = lnTEDiaCamas
                           oWorkSheet.Cells(iFila, 8).Value = lnTEDiaPaciente
                           oWorkSheet.Cells(iFila, 9).Value = lnTEDiaEstancia
                       End If
                       If lnTEEgresos > 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round((lnTEDiaEstancia / lnTEEgresos), 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 5).Value = Round((lnTEDiaEstancia / lnTEEgresos), 2)
                            End If
                       Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                            End If
                       End If
                       If lnTEEgresos > 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(((lnTEDiaCamas - lnTEDiaPaciente) / lnTEEgresos), 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 6).Value = Round(((lnTEDiaCamas - lnTEDiaPaciente) / lnTEEgresos), 2)
                            End If
                       Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                            End If
                       End If
                       If lnTEDiaCamas > 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round((lnTEDiaPaciente / lnTEDiaCamas) * 100, 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 7).Value = Round((lnTEDiaPaciente / lnTEDiaCamas) * 100, 2)
                            End If
                       Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                            End If
                       End If
                       If lnTECamas > 0 And lnNroMeses > 0 Then
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(((lnTEEgresos / lnTECamas) / lnNroMeses), 2))
                            Else
                                oWorkSheet.Cells(iFila, iCol + 8).Value = Round(((lnTEEgresos / lnTECamas) / lnNroMeses), 2)
                            End If
                       Else
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                            Else
                                oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                            End If
                       End If
                       iFila = iFila + 1
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    iFila = iFila + 1
                    If lbEsOpenOffice = True Then
                        Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":M" & CStr(iFila))
                        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTSEgresos)
                        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTSCamas)
                        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTSDiaCamas)
                        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTSDiaPaciente)
                        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTSDiaEstancia)
                    Else
                        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, iCol + 8
                        oWorkSheet.Cells(iFila, iCol + 0).Value = lnTSEgresos
                        oWorkSheet.Cells(iFila, iCol + 1).Value = lnTSCamas
                        oWorkSheet.Cells(iFila, iCol + 2).Value = lnTSDiaCamas
                        oWorkSheet.Cells(iFila, iCol + 3).Value = lnTSDiaPaciente
                        oWorkSheet.Cells(iFila, iCol + 4).Value = lnTSDiaEstancia
                    End If
                    lnTDCamas = lnTSCamas
                    If lnTSEgresos > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round((lnTSDiaEstancia / lnTSEgresos), 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 5).Value = Round((lnTSDiaEstancia / lnTSEgresos), 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                        End If
                    End If
                    If lnTSEgresos > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(((lnTSDiaCamas - lnTSDiaPaciente) / lnTSEgresos), 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 6).Value = Round(((lnTSDiaCamas - lnTSDiaPaciente) / lnTSEgresos), 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                        End If
                    End If
                    If lnTSDiaCamas > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round((lnTSDiaPaciente / lnTSDiaCamas) * 100, 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 7).Value = Round((lnTSDiaPaciente / lnTSDiaCamas) * 100, 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                        End If
                    End If
                    If lnTSCamas > 0 And lnNroMeses > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(((lnTSEgresos / lnTSCamas) / lnNroMeses), 2))
                        Else
                            oWorkSheet.Cells(iFila, iCol + 8).Value = Round(((lnTSEgresos / lnTSCamas) / lnNroMeses), 2)
                        End If
                    Else
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                        Else
                            oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                        End If
                    End If
                    iFila = iFila + 1
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                 Loop
                iFila = iFila + 1
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":M" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(lnTDEgresos)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTDCamas)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTDDiaCamas)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTDDiaPaciente)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTDDiaEstancia)
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 8
                    oWorkSheet.Cells(iFila, 2).Value = "Total: "
                    oWorkSheet.Cells(iFila, iCol + 0).Value = lnTDEgresos
                    oWorkSheet.Cells(iFila, iCol + 1).Value = lnTDCamas
                    oWorkSheet.Cells(iFila, iCol + 2).Value = lnTDDiaCamas
                    oWorkSheet.Cells(iFila, iCol + 3).Value = lnTDDiaPaciente
                    oWorkSheet.Cells(iFila, iCol + 4).Value = lnTDDiaEstancia
                 End If
                 If lnTDEgresos > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(Round((lnTDDiaEstancia / lnTDEgresos), 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 5).Value = Round((lnTDDiaEstancia / lnTDEgresos), 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 5).Value = 0
                    End If
                 End If
                 If lnTDEgresos > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(Round(((lnTDDiaCamas - lnTDDiaPaciente) / lnTDEgresos), 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 6).Value = Round(((lnTDDiaCamas - lnTDDiaPaciente) / lnTDEgresos), 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 6).Value = 0
                    End If
                 End If
                 If lnTDDiaCamas > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(Round((lnTDDiaPaciente / lnTDDiaCamas) * 100, 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 7).Value = Round((lnTDDiaPaciente / lnTDDiaCamas) * 100, 2)
                    End If
                 Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 7).Value = 0
                    End If
                 End If
                 If lnTDCamas > 0 And lnNroMeses > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(Round(((lnTDEgresos / lnTDCamas) / lnNroMeses), 2))
                    Else
                        oWorkSheet.Cells(iFila, iCol + 8).Value = Round(((lnTDEgresos / lnTDCamas) / lnNroMeses), 2)
                    End If
                 Else
                 If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(0)
                    Else
                        oWorkSheet.Cells(iFila, iCol + 8).Value = 0
                    End If
                 End If

                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 13
                    PrintArea(0).EndRow = iFila
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    'Falta que salga el nombre del responsable
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                    End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
                    'oWorkSheet.PrintOut
                End If
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
    Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    MsgBox lcErr

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
          .Fields.Append "IdServicioEgreso", adInteger, 4, adFldIsNullable
          .Fields.Append "Egresos", adUnsignedBigInt
          .Fields.Append "Camas", adUnsignedBigInt
          .Fields.Append "DiaCamas", adUnsignedBigInt
          .Fields.Append "DiaPaciente", adUnsignedBigInt
          .Fields.Append "DiaEstancia", adUnsignedBigInt
          .Fields.Append "PromPermanencia", adDouble
          .Fields.Append "InteSustitucion", adDouble
          .Fields.Append "PorcOcupacion", adDouble
          .Fields.Append "RendCama", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Function AtencionesSeleccionarPorAnio(lbPorFechaAltaMedica As Boolean, lnAnio As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set AtencionesSeleccionarPorAnio = Nothing
    ms_MensajeError = ""
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        If lbPorFechaAltaMedica Then
           .CommandText = "AtencionesSeleccionarPorAnioDeEgresoMedico"
        Else
           .CommandText = "AtencionesSeleccionarPorAnioDeEgresoAdministrativo"
        End If
        .CommandTimeout = 150
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

Function CamasProgramadasPorServicioAnioMes(lnAnio As Integer, lnMes As Integer, lnServicio As Long) As Integer
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    CamasProgramadasPorServicioAnioMes = 0
    ms_MensajeError = ""
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "CamasMovimientosSeleccionarPorServicioAnioMes"
        .CommandTimeout = 150
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, lnServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   CamasProgramadasPorServicioAnioMes = oRecordset.RecordCount
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



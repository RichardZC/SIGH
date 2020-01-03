VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptConstanciaAMAlternativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Constancias
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_NroConstancia As String
Dim ms_TipoConstancia As String
Dim ms_NombrePaciente As String
Dim ms_EdadPaciente As String
Dim ms_DescripcionServicio As String
Dim ms_FechasEstancia As String
Dim ms_NroHC As String

Dim ms_NombreMedico As String
Dim ms_CodigoMedico As String
Dim ms_Diagnostico As String
Dim ms_FechaDescriptivaActual As String
Dim mb_Tabla As ADODB.Recordset

Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
Dim mo_Conexion As ADODB.Connection

Property Let NroConstancia(lValue As String)
  ms_NroConstancia = lValue
End Property

Property Let TipoConstancia(lValue As String)
  ms_TipoConstancia = lValue
End Property

Property Let NombrePaciente(lValue As String)
  ms_NombrePaciente = lValue
End Property

Property Let EdadPaciente(lValue As String)
  ms_EdadPaciente = lValue
End Property

Property Let DescripcionServicio(lValue As String)
  ms_DescripcionServicio = lValue
End Property

Property Let FechasEstancia(lValue As String)
  ms_FechasEstancia = lValue
End Property

Property Let NroHC(lValue As String)
  ms_NroHC = lValue
End Property

Property Let NombreMedico(lValue As String)
  ms_NombreMedico = lValue
End Property

Property Let CodigoMedico(lValue As String)
  ms_CodigoMedico = lValue
End Property

Property Let Diagnostico(lValue As String)
  ms_Diagnostico = lValue
End Property

Property Let Datos(lValue As ADODB.Recordset)
  Set mb_Tabla = lValue
End Property

Property Let FechaDescriptivaActual(lValue As String)
  ms_FechaDescriptivaActual = lValue
End Property

Sub EjecutaFormulario()
  Dim oFormulario As New RpConstanciaAMAlternativa
  oFormulario.Show vbModal
End Sub

Sub CrearReporte()
  On Error GoTo ManejadorError
  Dim oRsTmp1 As New Recordset
  Set oRsTmp1 = mo_ReglasServGeograf.DepartamentosSeleccionarXidentificador(Val(Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)))
  If oRsTmp1.RecordCount > 0 Then
     ms_FechaDescriptivaActual = Trim(oRsTmp1.Fields!Nombre) & ms_FechaDescriptivaActual
  End If
  oRsTmp1.Close
  Set oRsTmp1 = Nothing
  Set RpConstanciaAMAlternativa.DataSource = mb_Tabla
  
  RpConstanciaAMAlternativa.RightMargin = 10
  RpConstanciaAMAlternativa.TopMargin = 10
  RpConstanciaAMAlternativa.LeftMargin = 10
  RpConstanciaAMAlternativa.BottomMargin = 10
  
  'CABECERA
  RpConstanciaAMAlternativa.Sections("Cabecera").Controls("lblNroConstancia").Caption = ms_NroConstancia
  RpConstanciaAMAlternativa.Sections("Cabecera").Controls("lblTipoConstancia").Caption = ms_TipoConstancia
  'Set RpConstanciaAMAlternativa.Sections("Cabecera").Controls("LogoEstablecimiento").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
  
  'DETALLE
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblNombrePaciente").Caption = UCase(ms_NombrePaciente)
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblEdadPaciente").Caption = ms_EdadPaciente
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblDescripcionServicio").Caption = ms_DescripcionServicio
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("FechasEstancia").Caption = ms_FechasEstancia
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblNroHistoriaClinica").Caption = ms_NroHC

    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblMedicoTurno1").Caption = ms_NombreMedico
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblMedicoTurno2").Caption = ms_NombreMedico
    
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblCodigoMedico").Caption = ms_CodigoMedico
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblDiagnostico").Caption = ms_Diagnostico
    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblFechaActual").Caption = ms_FechaDescriptivaActual
    
    '"Del " & lcBuscaParametro.SeleccionaFilaParametro(205) & " en la fecha: " & ml_FechaAt
  
'  If ml_Hospitaliza = True Then
'    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblHospitalizado").Visible = ml_Hospitaliza
'    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblDescripcionServicio").Left = RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblConsultorio").Left - 700
'    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblCama").Visible = ml_Hospitaliza
'    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblCama").Caption = RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblCama").Caption & ml_Cama
'  Else
'    RpConstanciaAMAlternativa.Sections("Detalle").Controls("lblAtendido").Visible = Not (ml_Hospitaliza)
'  End If
  
  RpConstanciaAMAlternativa.Orientation = rptOrientPortrait
  RpConstanciaAMAlternativa.Show 1
  Exit Sub

ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "SIGH"
    Case Else
      MsgBox Err.Description
  End Select
  Resume
  Exit Sub
End Sub




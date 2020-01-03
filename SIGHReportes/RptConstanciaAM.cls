VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptConstanciaAM"
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
Dim ml_Paciente As String
Dim ml_HC As String
Dim ml_Consultorio As String
Dim ml_Fecha As String
Dim ml_FechaAt As String
Dim ml_Diagnostico As String
Dim ml_Medico As String
Dim ml_idMedico As String
Dim ml_Observacion As String
Dim ml_Recibo As String
Dim ml_PC As String
Dim ml_NConstancia As String
Dim mb_Tabla As ADODB.Recordset
Dim ml_IdResponsable As Long
Dim ml_idPaciente As Long
Dim ml_idAtencion As Long
Dim ml_Hospitaliza As Boolean
Dim ml_Cama As String
Dim lcBuscaParametro As New SIGHDatos.Parametros

Dim mo_Conexion As ADODB.Connection

Property Let Cama(lValue As String)
  ml_Cama = lValue
End Property

Property Let Hospitaliza(lValue As Boolean)
  ml_Hospitaliza = lValue
End Property

Property Let NConstancia(lValue As String)
  ml_NConstancia = lValue
End Property

Property Let Paciente(lValue As String)
  ml_Paciente = lValue
End Property

Property Let HC(lValue As String)
  ml_HC = lValue
End Property

Property Let Consultorio(lValue As String)
  ml_Consultorio = lValue
End Property

Property Let Fecha(lValue As String)
  ml_Fecha = lValue
End Property

Property Let FechaAt(lValue As String)
  ml_FechaAt = lValue
End Property

Property Let Diagnostico(lValue As String)
  ml_Diagnostico = lValue
End Property

Property Let Medico(lValue As String)
  ml_Medico = lValue
End Property

Property Let idMedico(lValue As String)
  ml_idMedico = lValue
End Property

Property Let Recibo(lValue As String)
  ml_Recibo = lValue
End Property

Property Let Observacion(lValue As String)
  ml_Observacion = lValue
End Property

Property Let PC(lValue As String)
  ml_PC = lValue
End Property

Property Let Tabla(lValue As ADODB.Recordset)
  Set mb_Tabla = lValue
End Property

Property Let IdPaciente(lValue As Long)
  ml_idPaciente = lValue
End Property

Property Let IdResponsable(lValue As Long)
  ml_IdResponsable = lValue
End Property

Property Let idAtencion(lValue As Long)
  ml_idAtencion = lValue
End Property

Sub EjecutaFormulario()
  Dim oFormulario As New RpConstanciaAM
  oFormulario.Show vbModal
End Sub

Sub CrearReporte()
  On Error GoTo ManejadorError
  Dim oRsTmp1 As New Recordset
  Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
  Set oRsTmp1 = mo_ReglasServGeograf.DepartamentosSeleccionarXidentificador(Val(Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)))
  If oRsTmp1.RecordCount > 0 Then
     ml_Fecha = Trim(oRsTmp1.Fields!Nombre) & ml_Fecha
  End If
  oRsTmp1.Close
  Set oRsTmp1 = Nothing
  Set RpConstanciaAM.DataSource = mb_Tabla
  RpConstanciaAM.RightMargin = 10
  RpConstanciaAM.TopMargin = 10
  RpConstanciaAM.LeftMargin = 10
  RpConstanciaAM.BottomMargin = 10
  RpConstanciaAM.Sections("Cabecera").Controls("lblTConstancia").Caption = ml_NConstancia
  Set RpConstanciaAM.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
  RpConstanciaAM.Sections("Detalle").Controls("lblPaciente").Caption = UCase(ml_Paciente)
  RpConstanciaAM.Sections("Detalle").Controls("lblidpaciente").Caption = ml_HC
  RpConstanciaAM.Sections("Detalle").Controls("lblConsultorio").Caption = UCase(ml_Consultorio)
 'RpConstanciaAM.Sections("Detalle").Controls("lblFechas").Caption = ml_FechaAt
  RpConstanciaAM.Sections("Detalle").Controls("lblDiagnostico").Caption = ml_Diagnostico
  RpConstanciaAM.Sections("Detalle").Controls("lblMedico").Caption = UCase(ml_Medico)
  RpConstanciaAM.Sections("Detalle").Controls("lblidMedico").Caption = ml_idMedico
  RpConstanciaAM.Sections("Detalle").Controls("lblFecha").Caption = ml_Fecha
  RpConstanciaAM.Sections("Detalle").Controls("lblRecibo").Caption = ml_Recibo
  RpConstanciaAM.Sections("Detalle").Controls("lblObservaciones").Caption = ml_Observacion
  RpConstanciaAM.Sections("Detalle").Controls("etiqueta7").Caption = "Del " & lcBuscaParametro.SeleccionaFilaParametro(205) & " en la fecha: " & ml_FechaAt
  If ml_Hospitaliza = True Then
    RpConstanciaAM.Sections("Detalle").Controls("lblHospitalizado").Visible = ml_Hospitaliza
    RpConstanciaAM.Sections("Detalle").Controls("lblConsultorio").Left = RpConstanciaAM.Sections("Detalle").Controls("lblConsultorio").Left - 700
    RpConstanciaAM.Sections("Detalle").Controls("lblCama").Visible = ml_Hospitaliza
    RpConstanciaAM.Sections("Detalle").Controls("lblCama").Caption = RpConstanciaAM.Sections("Detalle").Controls("lblCama").Caption & ml_Cama
  Else
    RpConstanciaAM.Sections("Detalle").Controls("lblAtendido").Visible = Not (ml_Hospitaliza)
  End If
  RpConstanciaAM.Orientation = rptOrientPortrait
  RpConstanciaAM.Show 1
  Set mo_ReglasServGeograf = Nothing
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



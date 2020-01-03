VERSION 5.00
Begin VB.Form LoginActualizaClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualiza Clave"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LoginActualizaClave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCita 
      Caption         =   "Imprime Ticket Cita"
      Height          =   3375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4665
      Begin VB.CommandButton btnImprimePreCta 
         Caption         =   "Cuenta"
         Height          =   615
         Left            =   1800
         Picture         =   "LoginActualizaClave.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2700
         Width           =   1245
      End
      Begin VB.TextBox txtPaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   20
         Top             =   705
         Width           =   4380
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   19
         Top             =   1080
         Width           =   4380
      End
      Begin VB.TextBox txtNroOrdenPago 
         Height          =   330
         Left            =   210
         TabIndex        =   18
         Top             =   1500
         Width           =   4380
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Height          =   330
         Left            =   2400
         Picture         =   "LoginActualizaClave.frx":11A3
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox txtNcuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1140
         TabIndex        =   14
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
         Height          =   210
         Index           =   4
         Left            =   210
         TabIndex        =   16
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.Frame Frame 
      Height          =   645
      Left            =   15
      TabIndex        =   10
      Top             =   1620
      Width           =   4665
      Begin VB.TextBox txtLicenciaClinica 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1815
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   210
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Clinica"
         Height          =   210
         Index           =   3
         Left            =   270
         TabIndex        =   12
         Top             =   255
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   15
      TabIndex        =   6
      Top             =   0
      Width           =   4665
      Begin VB.TextBox txtClaveNew 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1815
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1110
         Width           =   2325
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1815
         TabIndex        =   0
         Top             =   330
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1815
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Contraseña Nueva"
         Height          =   210
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   1170
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   390
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Contraseña Actual"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   780
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   15
      TabIndex        =   5
      Top             =   2310
      Width           =   4665
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LoginActualizaClave.frx":172D
         DownPicture     =   "LoginActualizaClave.frx":1BF1
         Height          =   700
         Left            =   2445
         Picture         =   "LoginActualizaClave.frx":20DD
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LoginActualizaClave.frx":25C9
         DownPicture     =   "LoginActualizaClave.frx":2A29
         Height          =   700
         Left            =   900
         Picture         =   "LoginActualizaClave.frx":2E9E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "LoginActualizaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Actualiza clave del Usuario activo
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New sighEntidades.ReporteUtil
Dim ml_idUsuario  As Long
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Procesos As New SIGHProxies.Procesos
Dim oRsUsuario As New Recordset
Dim mo_Pacientes As New doPaciente
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lnCuentaDesdeOtroFormulario As Long
Property Let CuentaDesdeOtroFormulario(lValue As Long)
    lnCuentaDesdeOtroFormulario = lValue
End Property

Property Let ImprimeCuenta(lValue As Boolean)
    If lValue = True Then
       Me.fraCita.Visible = True
       Frame1.Visible = False
       Frame.Visible = False
       Frame3.Visible = False
    End If
End Property


Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Private Sub btnAceptar_Click()
     If txtClaveNew.Text = "" Then
        MsgBox "Debe registrar la nueva Contraseña", vbInformation, Me.Caption
        Exit Sub
     End If
     If txtPassword.Text = "" Then
        MsgBox "Debe registrar la Contraseña actual", vbInformation, Me.Caption
        Exit Sub
     End If
     
     Dim oCrypKey As New CrypKey.Util
     If UCase(txtPassword.Text) <> UCase(oCrypKey.DecryptString(oRsUsuario.Fields!Clave)) Then
        MsgBox "La Contraseña actual no corresponde a ese usuario", vbInformation, Me.Caption
        Exit Sub
     End If
     
     Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsCitasWeb As New Recordset
     lcMensaje = ""
     Set oRsCitasWeb = Nothing
     If lcMensaje <> "" Then
        MsgBox lcMensaje, vbInformation, Me.Caption
        Exit Sub
     End If
     
     Dim oDOEmpleado As New dOEmpleado, oEmpleados As New Empleados
     Dim oConexion As New Connection
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighEntidades.CadenaConexion
     Set oEmpleados.Conexion = oConexion
     oDOEmpleado.IdEmpleado = ml_idUsuario
     If oEmpleados.SeleccionarPorId(oDOEmpleado) = True Then
        oDOEmpleado.Clave = txtClaveNew.Text
        If oEmpleados.Modificar(oDOEmpleado) = True Then
            
        End If
     End If
     oConexion.Close
     Set oDOEmpleado = Nothing
     Set oEmpleados = Nothing
     Set oConexion = Nothing
     Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    On Error GoTo errActiv
    If lnCuentaDesdeOtroFormulario > 0 Then
       txtNcuenta.Text = lnCuentaDesdeOtroFormulario
       txtNcuenta_LostFocus
       btnImprimePreCta_Click
       Unload Me
    Else
        If Frame1.Visible = False Then
           Me.fraCita.Visible = True
           Me.Caption = "Imprime Ticket de CITA"
        Else
           Me.fraCita.Visible = False
        End If
    End If
    Exit Sub
errActiv:
    If lnCuentaDesdeOtroFormulario > 0 Then
       Unload Me
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub

Private Sub txtNcuenta_LostFocus()
    txtPaciente.Text = ""
    txtDatosDeCuenta.Text = ""
    txtNroOrdenPago.Text = ""
    If Val(txtNcuenta.Text) > 0 Then
       Dim oConexion As New Connection
       Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
       sighEntidades.AbreConexionSIGH oConexion
       Set oRsUsuario = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       If oRsUsuario.RecordCount > 0 Then
          Set mo_Pacientes = mo_ReglasAdmision.PacientesSeleccionarPorId(oRsUsuario!idPaciente, oConexion)
          txtPaciente.Text = mo_Pacientes.ApellidoPaterno & " " & mo_Pacientes.ApellidoMaterno & " " & mo_Pacientes.PrimerNombre & _
                            "  (N° Historia: " & mo_Pacientes.NroHistoriaClinica
          txtDatosDeCuenta.Text = "F. Cita: " & Format(oRsUsuario!FechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & _
                                  oRsUsuario!HoraIngreso & "    Consultorio: " & oRsUsuario!Consultorio
          VerSiTieneServicioAutomaticoPorEstancia oConexion
       End If
       Set oConexion = Nothing
       Set mo_ReglasFarmacia = Nothing
    End If
End Sub

Sub VerSiTieneServicioAutomaticoPorEstancia(oConexion As Connection)
    Dim oRsTmp As New ADODB.Recordset
    txtNroOrdenPago.Text = ""
    Set oRsTmp = mo_AdminFacturacion.FactOrdenServicioPagosPorIdAtencion(oRsUsuario!idAtencion, oConexion)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.Filter = "idPuntoCarga=6"
    End If
    If oRsTmp.RecordCount > 0 Then
       txtNroOrdenPago.Text = oRsTmp.Fields!IdOrdenPago
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub



Private Sub btnImprimePreCta_Click()
    On Error GoTo errbtnimp
    If txtPaciente.Text = "" Then
       MsgBox "Tiene que ingresar una CUENTA correcta", vbInformation, ""
       Exit Sub
    End If
    
    Dim lcPaciente As String
    Dim lcMedico As String
    Dim lcCola As String
    Dim ms_NroCola As String, lcEPS As String
    Dim oRsTmp3 As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim oReporte As New RptCaja
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    wxParametro216 = lcBuscaParametro.SeleccionaFilaParametro(216)
    wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
    lcPaciente = Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre)
    If mo_Pacientes.SegundoNombre <> "" Then
       lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.SegundoNombre)
    End If
    If mo_Pacientes.TercerNombre <> "" Then
      lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.TercerNombre)
    End If
    
    lcMedico = ""
    Set oRsTmp3 = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(oRsUsuario!IdEmpleado)
    If oRsTmp3.RecordCount > 0 Then
       lcMedico = oRsTmp3!ApellidoPaterno & " " & oRsTmp3!ApellidoMaterno & " " & oRsTmp3!Nombres
    End If
    oRsTmp3.Close
    
    ms_NroCola = mo_ReglasComunes.DevuelveNumeroColaEnCita(Val(Me.txtNcuenta.Text), oRsUsuario!HoraIngreso)
    
    
    If Val(txtNroOrdenPago.Text) > 0 Then
       Set oRsTmp3 = mo_AdminFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & oRsUsuario!IdServicioIngreso, sghPorCodigo)
       If oRsTmp3!CostoCeroCE = "S" Then
          'solo Para PLANIFICACION FAMILIAR
          lcCola = ms_NroCola
       Else
          lcCola = ms_NroCola & Space(5) & "N°Ord.Pago: " & txtNroOrdenPago.Text
       End If
       oRsTmp3.Close
    Else
       lcCola = ms_NroCola
    End If
    
    lcEPS = ""
    If Not IsNull(oRsUsuario!EpsPorcentaje) Then
       lcEPS = IIf(oRsUsuario!EpsPorcentaje > 0, mo_ReporteUtil.DevuelveEPScubre(oRsUsuario!EpsPorcentaje), "")
    End If
     
    oReporte.ImpresionPreCuenta oRsUsuario!FechaIngreso, oRsUsuario!HoraIngreso, lcPaciente, mo_Pacientes.NroHistoriaClinica, _
                                oRsUsuario!Consultorio, lcMedico, "CONSULTORIO EXTERNO", oRsUsuario!idAtencion, _
                                txtNroOrdenPago.Text, oRsUsuario!idCuentaAtencion, _
                                oRsUsuario!dFuenteFinanciamiento & _
                                lcEPS, _
                                lcCola, sighEntidades.Usuario, "CONSULTA MEDICA EN CE", mo_Pacientes.FichaFamiliar, _
                                mo_Pacientes.idTipoNumeracion, wxParametro216, wxParametro306, _
                                IIf(lnCuentaDesdeOtroFormulario > 0, True, False), oRsUsuario!IdMedicoIngreso
    Set oReporte = Nothing
    Set oRsTmp3 = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_ReglasDeProgMedica = Nothing
errbtnimp:
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oConexion As New Connection
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    sighEntidades.AbreConexionSIGH oConexion
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set mo_Pacientes = mo_ReglasAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not mo_Pacientes Is Nothing Then
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(mo_Pacientes.idPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set mo_ReglasFarmacia = Nothing

End Sub

Private Sub Form_Load()
'    If Frame1.Visible = False Then
'       Me.fraCita.Visible = True
'       Me.Caption = "Imprime Ticket de CITA"
'    Else
'       Me.fraCita.Visible = False
'    End If
    mo_Formulario.HabilitarDeshabilitar Me.txtUsuario, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroOrdenPago, False
    Set oRsUsuario = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    txtUsuario.Text = oRsUsuario.Fields!Usuario
End Sub





Private Sub Form_Unload(Cancel As Integer)
    oRsUsuario.Close
    Set oRsUsuario = Nothing
End Sub





Private Sub txtClaveNew_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtClaveNew

End Sub


Private Sub txtLicenciaClinica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       
    End If
End Sub





Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPassword

End Sub






VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CitasDetalle 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "CitasDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Datos del paciente"
      Height          =   1665
      Left            =   60
      TabIndex        =   21
      Top             =   1950
      Width           =   7695
      Begin VB.CheckBox chkSolicitarHistoria 
         Appearance      =   0  'Flat
         Caption         =   "Solicitar automatica historia clinica al archivo clínico"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1500
         TabIndex        =   35
         Top             =   1350
         Width           =   4725
      End
      Begin VB.CheckBox chkPacienteNuevo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Paciente nuevo"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5820
         TabIndex        =   11
         Top             =   360
         Width           =   1635
      End
      Begin VB.CommandButton btnBuscarPacientes 
         Caption         =   "..."
         Height          =   315
         Left            =   2490
         TabIndex        =   10
         Top             =   270
         Width           =   345
      End
      Begin VB.TextBox txtNroHistoria 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1500
         TabIndex        =   9
         Top             =   270
         Width           =   945
      End
      Begin VB.TextBox txtPrimerNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         MaxLength       =   35
         TabIndex        =   14
         Top             =   630
         Width           =   1935
      End
      Begin VB.TextBox txtApellidoPaterno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1500
         MaxLength       =   35
         TabIndex        =   12
         Top             =   630
         Width           =   1755
      End
      Begin VB.TextBox txtApellidoMaterno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1500
         MaxLength       =   35
         TabIndex        =   13
         Top             =   1005
         Width           =   1755
      End
      Begin VB.TextBox txtSegundoNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         MaxLength       =   35
         TabIndex        =   15
         Top             =   1005
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo cmbIdTipoGeneracionHistoria 
         Height          =   315
         Left            =   2880
         TabIndex        =   36
         Top             =   270
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         Text            =   ""
      End
      Begin VB.Label Label9 
         Caption         =   "N° Historia:"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Paterno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4170
         TabIndex        =   24
         Top             =   735
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Materno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   1035
         Width           =   1605
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4020
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Cita"
      Height          =   1875
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   7695
      Begin VB.TextBox txtFecha 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1515
         TabIndex        =   6
         Top             =   1380
         Width           =   1065
      End
      Begin VB.TextBox txtMedico 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   300
         Width           =   3975
      End
      Begin VB.TextBox txtIdCita 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6420
         TabIndex        =   0
         Top             =   300
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo cmbIdTipoConsulta 
         Height          =   315
         Left            =   5490
         TabIndex        =   5
         Top             =   1020
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIdTipoCita 
         Height          =   315
         Left            =   5490
         TabIndex        =   4
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIdEspecialidad 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   660
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIdServicio 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   1020
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSMask.MaskEdBox txtHoraInicio 
         Height          =   315
         Left            =   4410
         TabIndex        =   7
         Top             =   1380
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraFin 
         Height          =   315
         Left            =   6390
         TabIndex        =   8
         Top             =   1380
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         Height          =   315
         Left            =   90
         TabIndex        =   34
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label lblHoraInicio 
         Caption         =   "HoraInicio"
         Height          =   315
         Left            =   3390
         TabIndex        =   33
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label lblHoraFin 
         Caption         =   "HoraFin"
         Height          =   315
         Left            =   5625
         TabIndex        =   32
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "Servicio"
         Height          =   255
         Left            =   90
         TabIndex        =   31
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Especialidad"
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   690
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Medico"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   330
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo consulta"
         Height          =   255
         Left            =   4470
         TabIndex        =   27
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "IdCita"
         Height          =   315
         Left            =   5610
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo cita"
         Height          =   225
         Left            =   4500
         TabIndex        =   20
         Top             =   690
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   18
      Top             =   3630
      Width           =   7695
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   2370
         Picture         =   "CitasDetalle.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   3930
         Picture         =   "CitasDetalle.frx":0D3F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CitasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POCitas
'        Autor: William Castro Grijalva
'        Fecha: 08/08/2004 04:19:12 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim ml_IdCita As Long
Dim ml_IdEstadoCita As Long
Dim ml_IdAtencion As Long
Dim mo_Diario As PVDayView.PVDayView
Dim mo_Calendario As PVCalendar
Dim mi_Opcion As sghOpciones
Dim ml_IdPaciente As Long
Dim mo_Cita As New doCita
Dim mo_Paciente As New doPaciente
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim mda_UltimoSlotSeleccionado As Date
Dim mo_AdminAdmision As New SIGHReglasNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHReglasNegocios.ReglasComunes
Dim mo_AdminServiciosHosp As New SIGHReglasNegocios.ReglasServiciosHosp
Dim mo_AdminArchivoClinico As New SIGHReglasNegocios.ReglasArchivoClinico
Dim mo_Especialidad As New doEspecialidad
Dim mb_TieneHistoriaClinicaDefinitva As sghTipoGeneracionDeNroHistoria

Property Set Diario(oValue As PVDayView.PVDayView)
   Set mo_Diario = oValue
End Property
Property Get Diario() As PVDayView.PVDayView
   Set Diario = mo_Diario
End Property
Property Set Calendario(oValue As PVCalendar)
   Set mo_Calendario = oValue
End Property
Property Get Calendario() As PVCalendar
   Set mo_Calendario = mo_Diario
End Property
Property Let IdMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get IdMedico() As Long
   IdMedico = ml_IdMedico
End Property
Property Let NombreMedico(sValue As String)
   ms_NombreMedico = sValue
End Property
Property Get NombreMedico() As String
   NombreMedico = ms_NombreMedico
End Property
Property Let UltimoSlotSeleccionado(lValue As Date)
   mda_UltimoSlotSeleccionado = lValue
End Property
Property Get UltimoSlotSeleccionado() As Date
   UltimoSlotSeleccionado = mda_UltimoSlotSeleccionado
End Property

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdCita(lValue As Long)
   ml_IdCita = lValue
End Property
Property Get IdCita() As Long
   IdCita = ml_IdCita
End Property
Property Let IdEstadoCita(lValue As Long)
   ml_IdEstadoCita = lValue
End Property
Property Get IdEstadoCita() As Long
   IdEstadoCita = ml_IdEstadoCita
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property

Private Sub btnBuscarPacientes_Click()
Dim oBusqueda As New PacientesBusqueda

    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set mo_Paciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not mo_Paciente Is Nothing Then
            Me.txtNroHistoria.Tag = mo_Paciente.IdPaciente
            Me.txtNroHistoria = mo_Paciente.NroHistoriaClinica
            Me.txtApellidoPaterno = mo_Paciente.ApellidoPaterno
            Me.txtApellidoMaterno = mo_Paciente.ApellidoMaterno
            Me.txtPrimerNombre = mo_Paciente.PrimerNombre
            Me.txtSegundoNombre = mo_Paciente.SegundoNombre
            Me.cmbIdTipoGeneracionHistoria.BoundText = mo_Paciente.IdtipoGeneracion
        End If
    End If
    
    
End Sub

Private Sub chkPacienteNuevo_Click()
    
    Me.txtNroHistoria.Tag = ""
    Me.txtNroHistoria = ""
    Me.cmbIdTipoGeneracionHistoria.BoundText = ""
    
    Me.txtApellidoMaterno.Text = ""
    Me.txtApellidoPaterno.Text = ""
    Me.txtPrimerNombre.Text = ""
    Me.txtSegundoNombre.Text = ""
    
    If chkPacienteNuevo.Value = 1 Then
    
        Me.btnBuscarPacientes.Enabled = False
        mo_Formulario.HabilitarDeshabilitar Me.txtNroHistoria, False
    
        mo_Formulario.HabilitarDeshabilitar Me.txtApellidoMaterno, True
        mo_Formulario.HabilitarDeshabilitar Me.txtApellidoPaterno, True
        mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombre, True
        mo_Formulario.HabilitarDeshabilitar Me.txtSegundoNombre, True
        
        If Me.txtFecha = Date Then
            Me.chkSolicitarHistoria.Value = 1
        Else
            Me.chkSolicitarHistoria.Value = 0
        End If
        
    Else
        Me.btnBuscarPacientes.Enabled = True
        mo_Formulario.HabilitarDeshabilitar Me.txtNroHistoria, True
         mo_Formulario.HabilitarDeshabilitar Me.txtApellidoMaterno, False
        mo_Formulario.HabilitarDeshabilitar Me.txtApellidoPaterno, False
        mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombre, False
        mo_Formulario.HabilitarDeshabilitar Me.txtSegundoNombre, False
        Me.chkSolicitarHistoria.Value = 1
    End If
    
End Sub

Private Sub cmbIdEspecialidad_Change()
Dim rsServicio As New Recordset

    Set mo_Especialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarporId(Val(cmbIdEspecialidad.BoundText))
    CalculaLaHoraFinal
    
    cmbIdServicio.BoundColumn = "IdServicio"
    cmbIdServicio.ListField = "DescripcionLarga"
    Set rsServicio = mo_AdminServiciosHosp.ServiciosSeleccionarConsultoriosPorEspecialidad(Val(cmbIdEspecialidad.BoundText))
    Set cmbIdServicio.RowSource = rsServicio
    
    If rsServicio.RecordCount = 1 Then
            rsServicio.MoveFirst
            cmbIdServicio.BoundText = rsServicio!IdServicio
            cmbIdServicio.Enabled = False
    End If

    
    If mo_AdminServiciosComunes.MensajeError <> "" Then
        MsgBox mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption
    End If
    
End Sub

Sub CalculaLaHoraFinal()
Dim daHoraFin  As Date

    On Error Resume Next
    If Me.txtHoraInicio <> "__:__" Then
        daHoraFin = DateAdd("n", mo_Especialidad.TiempoPromedioConsulta, CDate(Me.txtHoraInicio))
        Me.txtHoraFin = Format(daHoraFin, "hh:mm")
    Else
        Me.txtHoraFin = "__:__"
    End If
    
End Sub
Private Sub cmbIdEspecialidad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEspecialidad
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdEspecialidad_LostFocus()
   If cmbIdEspecialidad.Text <> "" Then
       cmbIdEspecialidad.BoundText = Val(Split(cmbIdEspecialidad.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdEspecialidad
End Sub

Private Sub cmbIdEspecialidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicio
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdServicio_LostFocus()
   If cmbIdServicio.Text <> "" Then
       cmbIdServicio.BoundText = Val(Split(cmbIdServicio.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdServicio
End Sub

Private Sub cmbIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoConsulta
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoConsulta_LostFocus()
   If cmbIdTipoConsulta.Text <> "" Then
       cmbIdTipoConsulta.BoundText = Val(Split(cmbIdTipoConsulta.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoConsulta
End Sub

Private Sub cmbIdTipoConsulta_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoCita_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoCita
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoCita_LostFocus()
   If cmbIdTipoCita.Text <> "" Then
       cmbIdTipoCita.BoundText = Val(Split(cmbIdTipoCita.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoCita
End Sub

Private Sub cmbIdTipoCita_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoGeneracionHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoGeneracionHistoria
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoGeneracionHistoria_LostFocus()
   If cmbIdTipoGeneracionHistoria.Text <> "" Then
       cmbIdTipoGeneracionHistoria.BoundText = Val(Split(cmbIdTipoGeneracionHistoria.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdTipoGeneracionHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraFin
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraFin_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraFin
End Sub

Private Sub txtHoraFin_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraInicio
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraInicio_LostFocus()
    CalculaLaHoraFinal
    mo_Formulario.MarcarComoVacio txtHoraInicio
End Sub

Private Sub txtHoraInicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFecha
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFecha_LostFocus()
   mo_Formulario.MarcarComoVacio txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    Me.txtMedico = Me.NombreMedico
    Me.txtFecha = Format(Diario.CurrentDate, "dd/mm/yyyy")

    Select Case mi_Opcion
        Case sghAgregar
            Dim rsEspecialidad As Recordset
            Set rsEspecialidad = Me.cmbIdEspecialidad.RowSource
            If rsEspecialidad.RecordCount = 1 Then
                rsEspecialidad.MoveFirst
                Me.cmbIdEspecialidad.BoundText = rsEspecialidad!IdEspecialidad
            End If
            Me.cmbIdTipoCita.BoundText = 1
            Me.cmbIdTipoConsulta.BoundText = 1
            Me.txtHoraInicio = Format(Me.UltimoSlotSeleccionado, "hh:mm")
            CalculaLaHoraFinal
        Case sghModificar
            CargarDatosALosControles
        Case sghConsultar
            CargarDatosALosControles
        Case sghEliminar
            CargarDatosALosControles
    End Select

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Citas"
       Case sghModificar
           Me.Caption = "Modificar Citas"
       Case sghConsultar
           Me.Caption = "Consultar Citas"
           btnAceptar.Enabled = False
           Me.Frame1.Enabled = False
           Me.Frame3.Enabled = False
           
       Case sghEliminar
           Me.Frame1.Enabled = False
           Me.Frame3.Enabled = False
           Me.Caption = "Eliminar Citas"
           
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me

        Me.chkSolicitarHistoria.Value = SIGHComun.SolicitarHistoriaEnFormaAutomatica
        mo_Formulario.HabilitarDeshabilitar Me.txtApellidoMaterno, False
        mo_Formulario.HabilitarDeshabilitar Me.txtApellidoPaterno, False
        mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombre, False
        mo_Formulario.HabilitarDeshabilitar Me.txtSegundoNombre, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoGeneracionHistoria, False


End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
               
                    Dim sTexto As String
                    Dim dHoraIni As Double
                    Dim dHoraFin As Double
                    Dim iHoras() As Integer
                    Dim sHoras() As String
                    Dim programacion As PVAppointment
                    
                    sHoras = Split(Me.txtHoraInicio.Text, ":")
                    dHoraIni = CDbl(Val(sHoras(0)) + Val(sHoras(1)) / 60)
                    
                    sHoras = Split(Me.txtHoraFin, ":")
                    dHoraFin = CDbl(Val(sHoras(0)) + Val(sHoras(1)) / 60)
               
                    'sTexto = mo_Cita.IdCita & Chr(13)
                    'sTexto = sTexto + mo_Cita.HoraInicio + " - " + mo_Cita.HoraFin
                    sTexto = mo_Paciente.ApellidoPaterno + " " + mo_Paciente.ApellidoMaterno + " " + mo_Paciente.PrimerNombre
                    
                    Set programacion = mo_Diario.AppointmentSet.Add(sTexto, mo_Diario.CurrentDate + dHoraIni / 24, mo_Diario.CurrentDate + dHoraFin / 24)
                    programacion.DataVariant = mo_Cita
                    Me.Visible = False
                Else
                    If mo_AdminAdmision.MensajeError <> "" Then
                        MsgBox mo_AdminAdmision.MensajeError, vbCritical, Me.Caption
                        Exit Sub
                    End If
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               'If ModificarDatos() Then
               
               'End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron con éxito", vbInformation, Me.Caption
                    Me.Visible = False
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String

   ValidarDatosObligatorios = False
   If Me.txtFecha.Text = "__/__/____" Then
       sMensaje = sMensaje + "Ingrese el valor de Fecha" + Chr(13)
   End If
   If Me.txtHoraInicio.Text = "__:__" Then
       sMensaje = sMensaje + "Ingrese el valor de HoraInicio" + Chr(13)
   End If
   If Me.txtHoraFin.Text = "__:__" Then
       sMensaje = sMensaje + "Ingrese el valor de HoraFin" + Chr(13)
   End If
   If Me.cmbIdTipoCita.BoundText = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdTipoCita" + Chr(13)
   End If
   If Me.cmbIdTipoConsulta.BoundText = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdTipoConsulta" + Chr(13)
   End If
   If Me.cmbIdEspecialidad.BoundText = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdEspecialidad" + Chr(13)
   End If
   
    If Me.cmbIdServicio.BoundText = "" Then
        sMensaje = sMensaje + "Ingrese el valor del servicio" + Chr(13)
    End If

   If Me.chkPacienteNuevo.Value = 1 Then
        If Me.txtApellidoPaterno = "" Then
            sMensaje = sMensaje + "Ingrese el apellido paterno" + Chr(13)
        End If
        If Me.txtApellidoMaterno = "" Then
            sMensaje = sMensaje + "Ingrese el apellido materno" + Chr(13)
        End If
        If Me.txtPrimerNombre = "" Then
            sMensaje = sMensaje + "Ingrese el primer nombre" + Chr(13)
        End If
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Cita
           .IdCita = Me.IdCita
           .Fecha = Me.txtFecha.Text
           .HoraInicio = Me.txtHoraInicio.Text
           .HoraFin = Me.txtHoraFin.Text
           .IdTipoCita = Me.cmbIdTipoCita.BoundText
           .IdTipoConsulta = Me.cmbIdTipoConsulta.BoundText
           .IdMedico = Me.IdMedico
           .IdEspecialidad = Me.cmbIdEspecialidad.BoundText
           .IdPaciente = Val(Me.txtNroHistoria.Tag)
           .IdServicio = Me.cmbIdServicio.BoundText
           .IdEstadoCita = 1
           .IdAtencion = 0
   End With
   

   With mo_Paciente
            .IdPaciente = Val(Me.txtNroHistoria.Tag)
            .ApellidoMaterno = Me.txtApellidoMaterno
            .ApellidoPaterno = Me.txtApellidoPaterno
            .PrimerNombre = Me.txtPrimerNombre
            .SegundoNombre = Me.txtSegundoNombre
            .IdtipoGeneracion = Val(Me.cmbIdTipoGeneracionHistoria.BoundText)
   End With
   
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    AgregarDatos = mo_AdminAdmision.CitasAgregar(mo_Cita, mo_Paciente, Me.chkPacienteNuevo.Value)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
'Function ModificarDatos() As Boolean
'    CargaDatosAlObjetosDeDatos
'    AgregarDatos = mo_AdminAdmision.CitasAgregar(mo_Cita)
'End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_AdminAdmision.CitasEliminar(mo_Cita)
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
Dim doCita As doCita

    Dim programacion As PVAppointment
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    
    If programacion Is Nothing Then
       Exit Sub
    End If
    
    Set doCita = programacion.DataVariant

    mb_ExistenDatos = mo_AdminAdmision.CitasSeleccionarPorId(doCita.IdCita, mo_Cita, mo_Paciente)
    
    If mo_AdminAdmision.MensajeError <> "" Then
         MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbCritical, Me.Caption
         mb_ExistenDatos = False
         Exit Sub
    End If
       
    If mb_ExistenDatos Then
         With mo_Cita
             Me.IdCita = .IdCita
             Me.IdEstadoCita = .IdEstadoCita
             Me.IdAtencion = .IdAtencion
             Me.IdPaciente = .IdPaciente
             
             Me.txtIdCita = .IdCita
             Me.txtFecha.Text = .Fecha
             Me.txtHoraInicio.Text = .HoraInicio
             Me.txtHoraFin.Text = .HoraFin
             Me.cmbIdTipoCita.BoundText = .IdTipoCita
             Me.cmbIdTipoConsulta.BoundText = .IdTipoConsulta
             Me.cmbIdEspecialidad.BoundText = .IdEspecialidad
             Me.cmbIdServicio.BoundText = .IdServicio
             mb_ExistenDatos = True
         End With
         With mo_Paciente
            Me.txtNroHistoria = .NroHistoriaClinica
            Me.txtNroHistoria.Tag = .IdPaciente
            Me.txtPrimerNombre = .PrimerNombre
            Me.txtSegundoNombre = .SegundoNombre
            Me.txtApellidoMaterno.Text = .ApellidoMaterno
            Me.txtApellidoPaterno.Text = .ApellidoPaterno
            Me.cmbIdTipoGeneracionHistoria.BoundText = .IdtipoGeneracion
         End With
         mb_ExistenDatos = True
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
       
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Citas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdCita = 0
           Me.IdEstadoCita = 0
           Me.IdAtencion = 0
           Me.IdPaciente = 0
           Me.txtFecha.Text = ""
           Me.txtHoraInicio.Text = ""
           Me.txtHoraFin.Text = ""
           Me.cmbIdTipoCita.BoundText = ""
           Me.cmbIdTipoConsulta.BoundText = ""
           Me.txtMedico.Text = ""
           Me.cmbIdEspecialidad.BoundText = ""
   
End Sub


Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
Dim rsEspecialidad As New Recordset
       
       cmbIdTipoCita.BoundColumn = "IdTipoCita"
       cmbIdTipoCita.ListField = "DescripcionLarga"
       Set cmbIdTipoCita.RowSource = mo_AdminServiciosComunes.TiposCitaSeleccionarTodos()
       sMensaje = sMensaje + IIf(mo_AdminServiciosComunes.MensajeError <> "", mo_AdminServiciosComunes.MensajeError + Chr(13), "")
       
       cmbIdTipoConsulta.BoundColumn = "IdTipoConsulta"
       cmbIdTipoConsulta.ListField = "DescripcionLarga"
       Set cmbIdTipoConsulta.RowSource = mo_AdminServiciosComunes.TiposConsultaSeleccionarTodos()
       sMensaje = sMensaje + IIf(mo_AdminServiciosComunes.MensajeError <> "", mo_AdminServiciosComunes.MensajeError + Chr(13), "")
       
       cmbIdTipoGeneracionHistoria.BoundColumn = "IdTipoGeneracionNroHistoria"
       cmbIdTipoGeneracionHistoria.ListField = "DescripcionLarga"
       Set cmbIdTipoGeneracionHistoria.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
       sMensaje = sMensaje + IIf(mo_AdminServiciosComunes.MensajeError <> "", mo_AdminServiciosComunes.MensajeError + Chr(13), "")
       
       cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
       cmbIdEspecialidad.ListField = "DescripcionLarga"
       Set rsEspecialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarporMedico(ml_IdMedico)
       Set cmbIdEspecialidad.RowSource = rsEspecialidad
       
       If rsEspecialidad.RecordCount = 1 Then
            rsEspecialidad.MoveFirst
            cmbIdEspecialidad.BoundText = rsEspecialidad!IdEspecialidad
            cmbIdEspecialidad.Enabled = False
       End If
       
       sMensaje = sMensaje + IIf(mo_AdminServiciosComunes.MensajeError <> "", mo_AdminServiciosComunes.MensajeError + Chr(13), "")
       If sMensaje <> "" Then
           MsgBox sMensaje, vbCritical, Me.Caption
       End If

End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroHistoria_LostFocus()
    
    If txtNroHistoria.Text <> "" Then
        Dim oDOPAciente As doPaciente
        Set oDOPAciente = mo_AdminAdmision.PacientesSeleccionarPorHistoriaClinicaDefinitiva(Val(txtNroHistoria.Text))
        If Not oDOPAciente Is Nothing Then
            Me.txtNroHistoria.Tag = oDOPAciente.IdPaciente
            Me.txtNroHistoria = oDOPAciente.NroHistoriaClinica
            Me.txtApellidoPaterno = oDOPAciente.ApellidoPaterno
            Me.txtApellidoMaterno = oDOPAciente.ApellidoMaterno
            Me.txtPrimerNombre = oDOPAciente.PrimerNombre
            Me.txtSegundoNombre = oDOPAciente.SegundoNombre
            Me.cmbIdTipoGeneracionHistoria.BoundText = oDOPAciente.IdtipoGeneracion
            Set mo_Paciente = oDOPAciente
        Else
            MsgBox "El Nº de historia clínica ingresado no existe", vbExclamation, Me.Caption
            
            Me.txtNroHistoria.Tag = ""
            Me.txtApellidoPaterno = ""
            Me.txtApellidoMaterno = ""
            Me.txtPrimerNombre = ""
            Me.txtSegundoNombre = ""
            Me.cmbIdTipoGeneracionHistoria.BoundText = ""
        End If
    End If
    
    mo_Formulario.MarcarComoVacio txtNroHistoria
    
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtSegundoNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtSegundoNombre_LostFocus()
txtSegundoNombre.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombre.Text)
   mo_Formulario.MarcarComoVacio txtSegundoNombre
End Sub

Private Sub txtSegundoNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombre_LostFocus()
txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombre.Text)
   mo_Formulario.MarcarComoVacio txtPrimerNombre
End Sub

Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaterno_LostFocus()
txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
   mo_Formulario.MarcarComoVacio txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_LostFocus()
txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
   mo_Formulario.MarcarComoVacio txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ProgramacionDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "ProgramacionDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnColor 
      Caption         =   "..."
      Height          =   315
      Left            =   3735
      TabIndex        =   12
      Top             =   3900
      Width           =   315
   End
   Begin VB.Frame Frame2 
      Height          =   1080
      Left            =   75
      TabIndex        =   20
      Top             =   4365
      Width           =   5085
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ProgramacionDetalle.frx":08CA
         DownPicture     =   "ProgramacionDetalle.frx":0D8E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2580
         Picture         =   "ProgramacionDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ProgramacionDetalle.frx":1766
         DownPicture     =   "ProgramacionDetalle.frx":1BC6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   1035
         Picture         =   "ProgramacionDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraProg 
      Height          =   4350
      Left            =   75
      TabIndex        =   15
      Top             =   15
      Width           =   5100
      Begin VB.ComboBox cmbIdServicio 
         DataField       =   "IdEspecialidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ProgramacionDetalle.frx":24B0
         Left            =   1440
         List            =   "ProgramacionDetalle.frx":24B2
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1335
         Width           =   3500
      End
      Begin VB.TextBox lblColor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1455
         TabIndex        =   11
         Top             =   3885
         Width           =   2130
      End
      Begin VB.ComboBox cmbIdTipoServicio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   3500
      End
      Begin VB.ComboBox cmbIdTurno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2505
         Width           =   3495
      End
      Begin VB.ComboBox cmbIdTipoProgramacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2115
         Width           =   3500
      End
      Begin VB.ComboBox cmbIdEspecialidadMedico 
         DataField       =   "IdEspecialidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "ProgramacionDetalle.frx":24B4
         Left            =   1440
         List            =   "ProgramacionDetalle.frx":24B6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   3500
      End
      Begin MSComDlg.CommonDialog dlgColor 
         Left            =   810
         Top             =   3420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtFechaFin 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   5
         Top             =   1740
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtMedico 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   3500
      End
      Begin VB.TextBox txtFechaIni 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1740
         Width           =   1200
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3270
         Width           =   3500
      End
      Begin MSMask.MaskEdBox txtHoraInicio 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   2895
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraFin 
         Height          =   315
         Left            =   4170
         TabIndex        =   9
         Top             =   2880
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblServicio 
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   1365
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   26
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label lblHoraFin 
         Caption         =   "Hora final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3225
         TabIndex        =   25
         Top             =   2925
         Width           =   1005
      End
      Begin VB.Label lblHoraInicio 
         Caption         =   "Hora inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   2940
         Width           =   1005
      End
      Begin VB.Label Label43 
         Caption         =   "Especialidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label Label44 
         Caption         =   "Médico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   2550
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   19
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo prog."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   18
         Top             =   2175
         Width           =   1665
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   17
         Top             =   1770
         Width           =   1005
      End
      Begin VB.Label lblObservacion 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   16
         Top             =   3255
         Width           =   1005
      End
   End
End
Attribute VB_Name = "ProgramacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programar a Médicos
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProgramacion As Long
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim ml_IdDepartamento As Long
Dim ml_IdEspecialidad As Long
Dim ml_IdTipoServicio As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_NroCuposCE As Long
Dim mo_Diario As PVDayView.PVDayView
Dim mo_Calendario As PVCalendar
Dim mb_SeHaModificadoProgramacion As Boolean
Dim mo_FechaHora As New sighEntidades.FechaHora

Dim mo_cmbIdTipoProgramacion As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidadMedico As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTurno As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
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
   Set Calendario = mo_Calendario
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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdProgramacion(lValue As Long)
   ml_IdProgramacion = lValue
End Property
Property Get IdProgramacion() As Long
   IdProgramacion = ml_IdProgramacion
End Property
Property Let idMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_IdMedico
End Property
Property Let NombreMedico(sValue As String)
   ms_NombreMedico = sValue
End Property
Property Get NombreMedico() As String
   NombreMedico = ms_NombreMedico
End Property

Property Let IdDepartamento(lValue As Long)
   ml_IdDepartamento = lValue
End Property
Property Get IdDepartamento() As Long
   IdDepartamento = ml_IdDepartamento
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let idTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get idTipoServicio() As Long
   idTipoServicio = ml_IdTipoServicio
End Property
Property Let NroCuposCE(lValue As Long)
   ml_NroCuposCE = lValue
End Property
Property Get NroCuposCE() As Long
   NroCuposCE = ml_NroCuposCE
End Property
Property Get SeHaModificadoProgramacion() As Long
   SeHaModificadoProgramacion = mb_SeHaModificadoProgramacion
End Property

Private Sub btnColor_Click()
    
    Me.dlgColor.CancelError = False
    Me.dlgColor.ShowColor
    
    lblColor.BackColor = &H80000005
    On Error Resume Next
    lblColor.BackColor = Me.dlgColor.Color
    
End Sub

Private Sub cmbIdEspecialidadMedico_Click()
Dim rsServicio As New Recordset
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    If Val(mo_cmbIdTipoServicio.BoundText) > 0 And Val(mo_cmbIdEspecialidadMedico.BoundText) > 0 Then
        If Val(mo_cmbIdTipoServicio.BoundText) = 1 Then
            'Selecciona los consultorios en caso de consulta externa
            mo_cmbIdServicio.BoundColumn = "IdServicio"
            mo_cmbIdServicio.ListField = "DescripcionLarga"
            Set rsServicio = mo_AdminServiciosHosp.ServiciosSeleccionarConsultoriosPorEspecialidaddebb(Val(mo_cmbIdEspecialidadMedico.BoundText), sghFiltraSoloActivos, oConexion)
        Else
            'Selecciona Servicios Hosp/Emerg/otros
            mo_cmbIdServicio.BoundColumn = "IdServicio"
            mo_cmbIdServicio.ListField = "DescripcionLarga"
            Set rsServicio = mo_AdminServiciosComunes.ServiciosXEspecialidadYtipoServicio(Val(mo_cmbIdEspecialidadMedico.BoundText), Val(mo_cmbIdTipoServicio.BoundText), oConexion)
        End If
        Set mo_cmbIdServicio.RowSource = rsServicio
        mo_Formulario.HabilitarDeshabilitar cmbIdServicio, True
        If rsServicio.RecordCount = 1 Then
                rsServicio.MoveFirst
                mo_cmbIdServicio.BoundText = rsServicio!IdServicio
                mo_Formulario.HabilitarDeshabilitar cmbIdServicio, False
        End If
        If mo_AdminServiciosComunes.MensajeError <> "" Then
            'MsgBox mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
            Dim oMensaje As New SIGHNegocios.clMensaje
            oMensaje.MostrarFormulario mo_AdminServiciosComunes.MensajeError, Me.Caption
            Set oMensaje = Nothing

        End If
    End If
    Set rsServicio = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub cmbIdEspecialidadMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdEspecialidadMedico
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicio
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbIdTipoProgramacion_Click()
    mo_cmbIdTipoProgramacion.BoundText = Val(Split(cmbIdTipoProgramacion.Text, " = ")(0))
End Sub

Private Sub cmbIdTipoProgramacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoProgramacion
AdministrarKeyPreview KeyCode
End Sub
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
Dim oConexion       As New Connection
       oConexion.Open sighEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       
       mo_cmbIdTipoProgramacion.BoundColumn = "IdTipoProgramacion"
       mo_cmbIdTipoProgramacion.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoProgramacion.RowSource = mo_AdminServiciosComunes.TiposProgramacionSeleccionarTodos()
       
       mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
       mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
       
        mo_cmbIdEspecialidadMedico.BoundColumn = "IdEspecialidad"
        mo_cmbIdEspecialidadMedico.ListField = "DescripcionLarga"
        Dim rsEspecialidad As New Recordset
        Set rsEspecialidad = mo_AdminServiciosHosp.EspecialidadesSeleccionarporMedico(ml_IdMedico, oConexion)
        Set mo_cmbIdEspecialidadMedico.RowSource = rsEspecialidad
        
        If rsEspecialidad.RecordCount = 1 Then
             rsEspecialidad.MoveFirst
             mo_cmbIdEspecialidadMedico.BoundText = rsEspecialidad!IdEspecialidad
             cmbIdEspecialidadMedico.Enabled = False
        End If
       
       If mo_AdminServiciosComunes.MensajeError <> "" Then
            'MsgBox mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
            Dim oMensaje As New SIGHNegocios.clMensaje
            oMensaje.MostrarFormulario mo_AdminServiciosComunes.MensajeError, Me.Caption
            Set oMensaje = Nothing
           
       End If
       oConexion.Close
       Set oConexion = Nothing

End Sub

Private Sub cmbIdTipoProgramacion_LostFocus()
   
   If cmbIdTipoProgramacion.Text <> "" Then
        mo_cmbIdTipoProgramacion.BoundText = Val(Split(cmbIdTipoProgramacion.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoProgramacion

End Sub

Private Sub cmbIdTipoProgramacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoServicio_Click()
    mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
    
    Select Case mo_cmbIdTipoServicio.BoundText
    Case sghConsultaExterna
        lblServicio.Caption = "Consultorio"
        Me.cmbIdServicio.Visible = True
        cmbIdEspecialidadMedico_Click
    Case sghHospitalizacion, sghEmergenciaConsultorios, sghEmergenciaObservacion
        lblServicio.Caption = "Servicio"
        cmbIdEspecialidadMedico_Click
       ' Me.cmbIdServicio.Clear
       ' Me.cmbIdServicio.Visible = False
    Case Else
        Me.cmbIdServicio.Clear
        Me.cmbIdServicio.Visible = False
    End Select
    'Selecciona turnos
    If Val(mo_cmbIdTipoServicio.BoundText) > 0 Then
        mo_cmbIdTurno.BoundColumn = "IdTurno"
        mo_cmbIdTurno.ListField = "DescripcionLarga"
        Set mo_cmbIdTurno.RowSource = mo_AdminProgMedica.TurnosSeleccionarPorIdTipoServicio(Val(mo_cmbIdTipoServicio.BoundText))
    End If
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
        mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
End Sub


Private Sub cmbIdTurno_Click()
Dim oDOTurno As New doTurno
Dim IdTurno As Long

    IdTurno = mo_cmbIdTurno.ObtenerItemDataDeComboxBox(cmbIdTurno)
    mo_cmbIdTurno.BoundText = IdTurno

    Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorId(Val(mo_cmbIdTurno.BoundText))
    If Not oDOTurno Is Nothing Then
        Me.txtHoraInicio.Text = oDOTurno.HoraInicio
        Me.txtHoraFin.Text = oDOTurno.HoraFin
    End If
End Sub

Private Sub cmbIdTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTurno
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTurno_LostFocus()
Dim oDOTurno As New doTurno

    If cmbIdTurno.Text = "" Then
        Exit Sub
    End If
    
    'Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorCodigo(Trim(Split(cmbIdTurno.Text, "=")(0)))
    Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorId(mo_cmbIdTurno.BoundText)
    If oDOTurno.IdTurno <> 0 Then
        mo_cmbIdTurno.BoundText = oDOTurno.IdTurno
   End If
   
End Sub

Private Sub Form_Initialize()
    
    Set mo_cmbIdTipoProgramacion.MiComboBox = Me.cmbIdTipoProgramacion
    Set mo_cmbIdTipoServicio.MiComboBox = Me.cmbIdTipoServicio
    Set mo_cmbIdServicio.MiComboBox = Me.cmbIdServicio
    Set mo_cmbIdTurno.MiComboBox = Me.cmbIdTurno
    Set mo_cmbIdEspecialidadMedico.MiComboBox = Me.cmbIdEspecialidadMedico
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub lblColor_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblColor
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcion_LostFocus()
   mo_Formulario.MarcarComoVacio txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraFin
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraFin_LostFocus()
  If Not sighEntidades.ValidaHora(txtHoraFin.Text) Then
            'MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            Dim oMensaje As New SIGHNegocios.clMensaje
            oMensaje.MostrarFormulario "La hora ingresada no es correcta", Me.Caption
            Set oMensaje = Nothing
            
            txtHoraFin.Text = sighEntidades.HORA_VACIA_HM
        End If
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
    
       If Not sighEntidades.ValidaHora(txtHoraInicio.Text) Then
            'MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            Dim oMensaje As New SIGHNegocios.clMensaje
            oMensaje.MostrarFormulario "La hora ingresada no es correcta", Me.Caption
            Set oMensaje = Nothing
            
             txtHoraInicio.Text = sighEntidades.HORA_VACIA_HM
        End If
    
    
End Sub

Private Sub txtHoraInicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
        ValoresPorDefecto
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
 End Select
 
        Select Case mi_Opcion
       Case sghAgregar
       Case sghModificar
       Case sghConsultar
            fraProg.Enabled = False
            Me.btnAceptar.Enabled = False
       Case sghEliminar
            fraProg.Enabled = False
       End Select
 
End Sub
Sub ValoresPorDefecto()

    Me.txtMedico = Me.NombreMedico
    mo_cmbIdTipoProgramacion.BoundText = 1
    mo_cmbIdTipoServicio.BoundText = 1
    
    If mo_Calendario.SelectedDateCount = 1 Then
        Me.txtFechaIni = Format(mo_Calendario.Value, sighEntidades.DevuelveFechaSoloFormato_DMY)
        Me.txtFechaFin.Visible = False
    ElseIf mo_Calendario.SelectedDateCount > 1 Then
        Dim DiaSeleccionado As Date
        DiaSeleccionado = mo_Calendario.Value
        Me.txtFechaIni = Format(mo_Calendario.Value, sighEntidades.DevuelveFechaSoloFormato_DMY)
        Do While DiaSeleccionado <> 0
            Me.txtFechaFin.Text = Format(DiaSeleccionado, sighEntidades.DevuelveFechaSoloFormato_DMY)
            DiaSeleccionado = mo_Calendario.NextSelectedDate(DiaSeleccionado)
        Loop
        Me.txtFechaFin.Visible = True
    End If
    
    Dim rsEspecialidad As Recordset

End Sub


'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar programación médica"
       Case sghModificar
           Me.Caption = "Modificar programación médica"
       Case sghConsultar
           Me.Caption = "Consultar programación médica"
       Case sghEliminar
           Me.Caption = "Eliminar programación médica"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
       End If
   Else
        If CDate(txtFechaIni.Text) < lcBuscaParametro.RetornaFechaServidorSQL Then
          ' MsgBox "Sólo puede programar Fechas mayores a " & lcBuscaParametro.RetornaFechaServidorSQL, vbInformation, Me.Caption
           Dim oMensaje As New SIGHNegocios.clMensaje
           oMensaje.MostrarFormulario "Sólo puede programar Fechas mayores a " & lcBuscaParametro.RetornaFechaServidorSQL, Me.Caption
           Set oMensaje = Nothing
           
           
           Me.Visible = False
           LimpiarVariablesDeMemoria
        End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
               End If
           End If
       End If
   Case sghEliminar
            Dim oMensaje2 As New SIGHNegocios.clMensaje
            oMensaje2.MostrarFormulario Chr(13) & "Esta seguro ?", Me.Caption, 20, , , True
            If oMensaje2.BotonPresionado = sghAceptar Then
                If EliminarDatos() Then
                End If
            End If
            Set oMensaje2 = Nothing
   End Select
End Sub

Private Sub btnCancelar_Click()
    mb_SeHaModificadoProgramacion = False
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String

   ValidarDatosObligatorios = False
   
    If mo_cmbIdTipoServicio.BoundText = "" Then
        sMensaje = sMensaje + "Ingrese el tipo de servicio" + Chr(13)
    Else
        If Val(mo_cmbIdTipoServicio.BoundText) = 1 Then
            If Val(mo_cmbIdServicio.BoundText) = 0 Then
                sMensaje = sMensaje + "Ingrese el consultorio de consulta externa" + Chr(13)
            End If
        End If
    End If
    
    If Me.txtHoraInicio.Text = sighEntidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Ingrese la hora inicio" + Chr(13)
    Else
       'A.Yañez 06-11-2014***************************************
        If Not sighEntidades.EsHora(txtHoraInicio) Then
'        If Not mo_FechaHora.ValidaHora(txtHoraInicio) Then
            sMensaje = sMensaje + "La hora de inicio no es correcta" + Chr(13)
        End If
    End If
    If Me.txtHoraFin.Text = sighEntidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Ingrese la hora fin" + Chr(13)
    Else
      'A.Yañez 06-11-2014****************************************
        If Not sighEntidades.EsHora(txtHoraFin) Then
'        If Not mo_FechaHora.ValidaHora(txtHoraFin) Then
            sMensaje = sMensaje + "La hora final no es correcta" + Chr(13)
        End If
   End If
   If mo_cmbIdTipoProgramacion.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el tipo de programacion" + Chr(13)
   End If
   If mo_cmbIdTurno.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el turno" + Chr(13)
   End If
   If Val(mo_cmbIdTipoProgramacion.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el tipoProgramacion" + Chr(13)
   End If
   If sMensaje <> "" Then
       'MsgBox sMensaje, vbInformation, Me.Caption
       Dim oMensaje As New SIGHNegocios.clMensaje
       oMensaje.MostrarFormulario sMensaje, Me.Caption
       Set oMensaje = Nothing
       
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   Dim sMensaje As String
   ValidarReglas = False
   If Not sighEntidades.ValidaHora(Me.txtHoraInicio) Then
        'MsgBox "La hora inicial ingresada no es válida", vbExclamation, Me.Caption
        Dim oMensaje As New SIGHNegocios.clMensaje
        oMensaje.MostrarFormulario "La hora inicial ingresada no es válida", Me.Caption
        Set oMensaje = Nothing
        
        Exit Function
   End If
   
   If Not sighEntidades.ValidaHora(Me.txtHoraFin) Then
        'MsgBox "La hora final ingresada no es válida", vbExclamation, Me.Caption
        Dim oMensaje0 As New SIGHNegocios.clMensaje
        oMensaje0.MostrarFormulario "La hora final ingresada no es válida", Me.Caption
        Set oMensaje0 = Nothing
        
        Exit Function
   End If
   
   If mi_Opcion <> sghAgregar Then
        If -DateDiff("h", CDate(Date & " " & Me.txtHoraInicio), CDate(Date & " " & txtHoraFin)) > 0 Then
            Dim oMensaje1 As New SIGHNegocios.clMensaje
            oMensaje1.MostrarFormulario "La hora de salida no puede ser menor a la hora de ingreso", Me.Caption
            Set oMensaje1 = Nothing
            
            Exit Function
        End If
    End If
    'Actualizado 30102014 yamill palomino
    If validarProgramacionMedicoEnOtrosServicios(Val(mo_cmbIdServicio.BoundText), Me.idMedico, _
                    txtFechaIni.Text, txtHoraInicio.Text, txtHoraFin.Text) = False Then
        Exit Function
    End If
    '
    Dim oRsTmp As New Recordset
    Dim oMensaje2 As New SIGHNegocios.clMensaje
    
    Dim dFechaInicio As Date, dFechaFin As Date
    Dim dFechaInicioProgramacion As Date, dFechaFinProgracion As Date
    
    dFechaInicioProgramacion = getFechaInicio(CDate(txtFechaIni.Text), Me.txtHoraInicio.Text)
    dFechaFinProgracion = getFechaFin(CDate(txtFechaIni.Text), Me.txtHoraInicio.Text, Me.txtHoraFin.Text)
    
    Set oRsTmp = mo_AdminProgMedica.ProgramacionMedicaSeleccionarPorIdServicio(Val(mo_cmbIdServicio.BoundText))
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!fecha = CDate(txtFechaIni.Text) Then
                'Actualizado 27102014 yamill palomino
                dFechaInicio = getFechaInicio(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio)
                dFechaFin = getFechaFin(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio, oRsTmp.Fields!HoraFin)
                
                'If (oRsTmp.Fields!HoraInicio <= Me.txtHoraInicio.Text And oRsTmp.Fields!HoraFin >= Me.txtHoraInicio.Text) Or (oRsTmp.Fields!HoraInicio <= Me.txtHoraFin.Text And oRsTmp.Fields!HoraFin >= Me.txtHoraFin.Text) Then
'                    If (oRsTmp.Fields!HoraInicio <= Me.txtHoraInicio.Text And oRsTmp.Fields!HoraFin >= Me.txtHoraFin.Text) _
'                    Or (Me.txtHoraInicio.Text <= oRsTmp.Fields!HoraInicio And Me.txtHoraFin.Text >= oRsTmp.Fields!HoraFin) _
'                    Or (Me.txtHoraInicio.Text > oRsTmp.Fields!HoraInicio And Me.txtHoraInicio.Text < oRsTmp.Fields!HoraFin) _
'                    Or (Me.txtHoraFin.Text > oRsTmp.Fields!HoraInicio And Me.txtHoraFin.Text < oRsTmp.Fields!HoraFin) Then ' Actualizado 16/10/2014 Yamill palomino

                If (dFechaInicio <= dFechaInicioProgramacion And dFechaFin >= dFechaFinProgracion) _
                    Or (dFechaInicioProgramacion <= dFechaInicio And dFechaFinProgracion >= dFechaFin) _
                    Or (dFechaInicioProgramacion > dFechaInicio And dFechaInicioProgramacion < dFechaFin) _
                    Or (dFechaFinProgracion > dFechaInicio And dFechaFinProgracion < dFechaFin) Then ' Actualizado 16/10/2014 Yamill palomino
                    If mi_Opcion = sghAgregar Then
                        oMensaje2.MostrarFormulario "El Servicio elegido ya fué programado para esa FECHA/HORA, al Médico: " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!Nombres), Me.Caption
                        Set oMensaje2 = Nothing
                        Exit Function
                    ElseIf mi_Opcion = sghModificar And oRsTmp!IdProgramacion <> ml_IdProgramacion Then
                        oMensaje2.MostrarFormulario "El Servicio elegido ya fué programado para esa FECHA/HORA, al Médico: " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!Nombres), Me.Caption
                        Set oMensaje2 = Nothing
                        Exit Function
                    End If
                End If
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oMensaje2 = Nothing
    '
    sMensaje = HayCruceDeProgramacionesDeConsultaExterna()
    If sMensaje <> "" Then
         'MsgBox "Hay un cruces con la(s) siguiente(s) programacion(es) de consulta externa " & Chr(13) & sMensaje, vbExclamation, Me.Caption
         Dim oMensaje22 As New SIGHNegocios.clMensaje
         oMensaje22.MostrarFormulario "Hay un cruces con la(s) siguiente(s) programacion(es) de consulta externa " & Chr(13) & sMensaje, Me.Caption
         Set oMensaje22 = Nothing
         
         Exit Function
    End If
   
   ValidarReglas = True
End Function
Function HayCruceDeProgramacionesDeConsultaExterna() As String
Dim daDiaSeleccionado As Date
Dim programacion As PVAppointment
Dim sTitulo As String
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim bTurnoProgramado As Boolean
Dim oDOProgramacion As New DOProgramacionMedica
Dim sMensaje  As String

    sMensaje = ""
    daDiaSeleccionado = Me.Calendario.Value
    Do While daDiaSeleccionado <> 0
        Set programacion = Me.Diario.AppointmentSet.Get(daDiaSeleccionado)
        bTurnoProgramado = False
        Do While Not programacion Is Nothing
            'Verifica que la programacion sea del mismo dia
             Set oDOProgramacion = programacion.DataVariant
            'Verifica si la nueva porg es de CE
             If Val(Me.cmbIdTipoServicio) = 1 Then
                If Me.IdProgramacion <> oDOProgramacion.IdProgramacion Then
                    'Verifica si la prog existente tambien es de CE
                    If oDOProgramacion.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText) Then
                        If oDOProgramacion.fecha = daDiaSeleccionado Then
                            'Las validaciones se hacen solo si la hora inicio nueva es menor que la hora de fin existente
                            If ConvertirAMinutos(Me.txtHoraInicio) < ConvertirAMinutos(oDOProgramacion.HoraFin) Then
                                If (ConvertirAMinutos(Me.txtHoraInicio) <= ConvertirAMinutos(oDOProgramacion.HoraInicio) And ConvertirAMinutos(Me.txtHoraFin) >= ConvertirAMinutos(oDOProgramacion.HoraInicio)) Or _
                                    (ConvertirAMinutos(Me.txtHoraInicio) >= ConvertirAMinutos(oDOProgramacion.HoraInicio) And ConvertirAMinutos(Me.txtHoraFin) <= ConvertirAMinutos(oDOProgramacion.HoraFin)) Or _
                                    (ConvertirAMinutos(Me.txtHoraInicio) <= ConvertirAMinutos(oDOProgramacion.HoraFin) And ConvertirAMinutos(Me.txtHoraFin) >= ConvertirAMinutos(oDOProgramacion.HoraFin)) Then
                                    sMensaje = sMensaje & programacion.Description & " " & oDOProgramacion.fecha & " Hora Inicio " & oDOProgramacion.HoraInicio & " Hora Fin " & oDOProgramacion.HoraFin & Chr(13)
                                End If
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                End If
            End If
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
        daDiaSeleccionado = Me.Calendario.NextSelectedDate(daDiaSeleccionado)
    Loop

    HayCruceDeProgramacionesDeConsultaExterna = sMensaje

End Function
Function ConvertirAMinutos(sHora As String) As Integer
Dim sHoras() As String

    sHoras = Split(sHora, ":")
    ConvertirAMinutos = Val(sHoras(0)) * 60 + Val(sHoras(1))

End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
Dim daDiaSeleccionado As Date
Dim programacion As PVAppointment
Dim sTitulo As String
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim iHoraIni1 As Double
Dim iHoraFin1 As Double
Dim bTurnoProgramado As Boolean
Dim doProgramacion As DOProgramacionMedica
Dim programacionAdicional As PVAppointment
Dim daDiaSeleccionadoFantasma As Date
Dim bAgregarProgFicticio As Boolean
Dim sHoraFin As String
Dim oCollProgramaciones As New Collection
Dim ldFechaHoraServidor As Date
Dim oRsTmp As New Recordset
Dim lnTiempoPromedioAtencion As Long
    ldFechaHoraServidor = lcBuscaParametro.RetornaFechaHoraServidorSQL
    mb_SeHaModificadoProgramacion = False
    daDiaSeleccionado = mo_Calendario.Value
    '
    lnTiempoPromedioAtencion = 0
    Set oRsTmp = mo_AdminServiciosHosp.EspecialidadCESeleccionarPorIdServicio(Val(mo_cmbIdServicio.BoundText))
    If oRsTmp.RecordCount > 0 Then
        If Not IsNull(oRsTmp.Fields!TiempoPromedioAtencion) Then
            lnTiempoPromedioAtencion = oRsTmp.Fields!TiempoPromedioAtencion
        End If
    End If
    oRsTmp.Close
    '
    Do While daDiaSeleccionado <> 0
        sTitulo = ""
        Set programacion = mo_Diario.AppointmentSet.Get(daDiaSeleccionado)
        bTurnoProgramado = False
        
        'Agrega programacion
        sHoras = Split(Me.txtHoraInicio, ":")
        iHoraIni = Val(sHoras(0)) + Val(sHoras(1)) / 60
        iHoraIni1 = Val(sHoras(0)) + Val(sHoras(1)) / 60
        
        sHoras = Split(Me.txtHoraFin, ":")
        iHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
        iHoraFin1 = Val(sHoras(0)) + Val(sHoras(1)) / 60
        
        If iHoraIni1 < iHoraFin1 Then
            Set programacion = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, daDiaSeleccionado + iHoraIni / 24, daDiaSeleccionado + iHoraFin / 24)
            On Error Resume Next
            programacion.BackColor = Me.lblColor.BackColor
            sHoraFin = Me.txtHoraFin
            bAgregarProgFicticio = False
        Else
            iHoraFin = 23 + 59 / 60
            sHoraFin = "23:59"
            Set programacion = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, daDiaSeleccionado + iHoraIni / 24, daDiaSeleccionado + iHoraFin / 24)
            programacion.BackColor = Me.lblColor.BackColor
            bAgregarProgFicticio = True
        End If
        mo_Calendario.DATEText(daDiaSeleccionado) = IIf(mo_Calendario.DATEText(daDiaSeleccionado) = "", sighEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="), mo_Calendario.DATEText(daDiaSeleccionado) + "," + sighEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="))
        Me.IdDepartamento = mo_ReglasFacturacion.ServiciosDevuelveIdDepartamento(Val(mo_cmbIdServicio.BoundText))  'Me.IdDepartamento
        Set doProgramacion = New DOProgramacionMedica
        'Datos de la programacion
        doProgramacion.fecha = Format(daDiaSeleccionado, sighEntidades.DevuelveFechaSoloFormato_DMY)
        doProgramacion.HoraInicio = Me.txtHoraInicio
        doProgramacion.HoraFin = sHoraFin
        doProgramacion.IdTipoProgramacion = Val(mo_cmbIdTipoProgramacion.BoundText)
        doProgramacion.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
        doProgramacion.IdTurno = Val(mo_cmbIdTurno.BoundText)
        doProgramacion.IdUsuarioAuditoria = ml_idUsuario
        doProgramacion.Descripcion = Me.txtDescripcion
        doProgramacion.IdEspecialidad = Val(mo_cmbIdEspecialidadMedico.BoundText)
        doProgramacion.Color = Me.lblColor.BackColor
        doProgramacion.IdDepartamento = Me.IdDepartamento
        doProgramacion.idMedico = Me.idMedico
        doProgramacion.IdProgramacion = 0
        doProgramacion.IdServicio = Val(mo_cmbIdServicio.BoundText)
        doProgramacion.FechaReg = ldFechaHoraServidor
        doProgramacion.TiempoPromedioAtencion = lnTiempoPromedioAtencion
        
        oCollProgramaciones.Add doProgramacion
        
        programacion.DataVariant = doProgramacion
    
        'Para los turnos que continuan mañana
        If bAgregarProgFicticio Then
            iHoraIni = 0 + 0 / 60
            iHoraFin = Val(sHoras(0)) + Val(sHoras(1)) / 60
            daDiaSeleccionadoFantasma = DateAdd("d", 1, daDiaSeleccionado)
            
            Set doProgramacion = New DOProgramacionMedica
            doProgramacion.fecha = Format(daDiaSeleccionadoFantasma, sighEntidades.DevuelveFechaSoloFormato_DMY)
            doProgramacion.HoraInicio = "00:00"
            doProgramacion.HoraFin = Me.txtHoraFin
            doProgramacion.IdTipoProgramacion = mo_cmbIdTipoProgramacion.BoundText
            doProgramacion.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
            doProgramacion.IdTurno = Val(mo_cmbIdTurno.BoundText)
            doProgramacion.IdUsuarioAuditoria = ml_idUsuario
            doProgramacion.Descripcion = Me.txtDescripcion
            doProgramacion.IdEspecialidad = Val(mo_cmbIdEspecialidadMedico.BoundText)
            doProgramacion.Color = Me.lblColor.BackColor
            doProgramacion.IdDepartamento = Me.IdDepartamento
            doProgramacion.idMedico = Me.idMedico
            doProgramacion.IdProgramacion = 0
            doProgramacion.FechaReg = ldFechaHoraServidor
            doProgramacion.IdServicio = Val(mo_cmbIdServicio.BoundText)
            
            oCollProgramaciones.Add doProgramacion
            
            Set programacionAdicional = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, daDiaSeleccionadoFantasma + iHoraIni / 24, daDiaSeleccionadoFantasma + iHoraFin / 24)
            
            programacionAdicional.DataVariant = doProgramacion
            programacionAdicional.BackColor = Me.lblColor.BackColor
            mo_Calendario.DATEText(daDiaSeleccionadoFantasma) = IIf(mo_Calendario.DATEText(daDiaSeleccionadoFantasma) = "", sighEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="), mo_Calendario.DATEText(daDiaSeleccionadoFantasma) + "," + sighEntidades.ExtraerCadena(Me.cmbIdTurno.Text, 1, "="))
        End If
        
        daDiaSeleccionado = mo_Calendario.NextSelectedDate(daDiaSeleccionado)
        mb_SeHaModificadoProgramacion = True
    Loop

    
    AgregarDatos = mo_AdminProgMedica.ProgramacionMedicaAgregar(oCollProgramaciones, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtMedico.Text) & " " & Trim(cmbIdServicio.Text) & txtFechaIni.Text)

    Me.Visible = False
    LimpiarVariablesDeMemoria
    Set oRsTmp = Nothing
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim iHoras() As Integer
Dim sHoras() As String
Dim programacion As PVAppointment
Dim lnTiempoPromedioAtencion As Long
Dim oRsTmp As New Recordset
    '
    lnTiempoPromedioAtencion = 0
    Set oRsTmp = mo_AdminServiciosHosp.EspecialidadCESeleccionarPorIdServicio(Val(mo_cmbIdServicio.BoundText))
    If oRsTmp.RecordCount > 0 Then
       lnTiempoPromedioAtencion = oRsTmp.Fields!TiempoPromedioAtencion
    End If
    oRsTmp.Close
    '

    mb_SeHaModificadoProgramacion = True
    
    sHoras = Split(Me.txtHoraInicio.Text, ":")
    iHoraIni = CInt(Val(sHoras(0)) + Val(sHoras(1)) / 60)
    
    sHoras = Split(Me.txtHoraFin, ":")
    iHoraFin = CInt(Val(sHoras(0)) + Val(sHoras(1)) / 60)
    
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    On Error Resume Next
    mo_Diario.AppointmentSet.Remove programacion.Key
    
    Set programacion = mo_Diario.AppointmentSet.Add(Me.cmbIdTurno.Text, mo_Diario.CurrentDate + iHoraIni / 24, mo_Diario.CurrentDate + iHoraFin / 24)
    programacion.BackColor = Me.lblColor.BackColor
    
    Dim doProgramacion As New DOProgramacionMedica
    
    'Datos de la programacion
    doProgramacion.IdProgramacion = Me.IdProgramacion
    doProgramacion.fecha = Format(programacion.StartDateTime, sighEntidades.DevuelveFechaSoloFormato_DMY)
    doProgramacion.HoraInicio = Me.txtHoraInicio
    doProgramacion.HoraFin = Me.txtHoraFin
    doProgramacion.IdTipoProgramacion = mo_cmbIdTipoProgramacion.BoundText
    doProgramacion.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    doProgramacion.IdTurno = mo_cmbIdTurno.BoundText
    doProgramacion.IdUsuarioAuditoria = ml_idUsuario
    doProgramacion.Descripcion = Me.txtDescripcion
    doProgramacion.IdEspecialidad = Val(mo_cmbIdEspecialidadMedico.BoundText)
    doProgramacion.Color = Me.lblColor.BackColor
    doProgramacion.IdDepartamento = mo_ReglasFacturacion.ServiciosDevuelveIdDepartamento(Val(mo_cmbIdServicio.BoundText))    'Me.IdDepartamento
    doProgramacion.idMedico = Me.idMedico
    doProgramacion.IdServicio = Val(mo_cmbIdServicio.BoundText)
    doProgramacion.TiempoPromedioAtencion = lnTiempoPromedioAtencion

        
    programacion.DataVariant = doProgramacion

    mo_Calendario.DATEText(mo_Diario.CurrentDate) = ObtieneCodigoDeProgramacionPorDia(mo_Diario.CurrentDate)
    Me.Visible = False
    LimpiarVariablesDeMemoria

    ModificarDatos = mo_AdminProgMedica.ProgramacionMedicaModificar(doProgramacion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtMedico.Text) & " " & Trim(cmbIdServicio.Text) & txtFechaIni.Text)
    Set oRsTmp = Nothing
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
Dim programacion As PVAppointment
Dim oDOProgramacion As New DOProgramacionMedica
Dim oProgMedicas As New Collection
    
    mb_SeHaModificadoProgramacion = True
    
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    mo_Diario.AppointmentSet.Remove programacion.Key

    mo_Calendario.DATEText(mo_Diario.CurrentDate) = ObtieneCodigoDeProgramacionPorDia(mo_Diario.CurrentDate)

    'Agrega la unica programacion medica a eliminar
    
    EliminarDatos = mo_AdminProgMedica.ProgramacionMedicaEliminar(programacion.DataVariant, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtMedico.Text) & " " & Trim(cmbIdServicio.Text) & txtFechaIni.Text)
    
    Me.Visible = False
    LimpiarVariablesDeMemoria

End Function


'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
    Dim doProgramacion As DOProgramacionMedica
    Dim oRsTmp1 As New Recordset
    Dim programacion As PVAppointment
    Set programacion = mo_Diario.AppointmentSet.GetSelectedAppointment
    
    If programacion Is Nothing Then
        'MsgBox "Por favor seleccione la programación que desea modificar o eliminar", vbInformation, Me.Caption
        Dim oMensaje2 As New SIGHNegocios.clMensaje
        oMensaje2.MostrarFormulario "Por favor seleccione la programación que desea modificar o eliminar", Me.Caption
        Set oMensaje2 = Nothing
        
        Exit Sub
    End If
    
    Set doProgramacion = programacion.DataVariant
    Me.IdProgramacion = doProgramacion.IdProgramacion


    If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
        Set oRsTmp1 = mo_AdminProgMedica.CitasSeleccionarPorServicioYfecha(doProgramacion.IdServicio, doProgramacion.fecha)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If oRsTmp1.Fields!IdProgramacion = doProgramacion.IdProgramacion Then
                 Dim oMensaje As New SIGHNegocios.clMensaje
                 oMensaje.MostrarFormulario "No se puede modificar o eliminar porque ya existe un " & _
                                            "Paciente con CITA", Me.Caption
                 Set oMensaje = Nothing
                 '
                 Set oRsTmp1 = Nothing
                 btnCancelar_Click
                 Exit Sub
              End If
              oRsTmp1.MoveNext
           Loop
        End If
    End If
    
    Me.txtMedico = Me.NombreMedico

    Me.txtFechaIni = Format(doProgramacion.fecha, sighEntidades.DevuelveFechaSoloFormato_DMY)
    Me.txtFechaFin.Visible = False
    Me.txtDescripcion = doProgramacion.Descripcion
    
    mo_cmbIdTipoProgramacion.BoundText = doProgramacion.IdTipoProgramacion
    mo_cmbIdTipoServicio.BoundText = doProgramacion.idTipoServicio
    mo_cmbIdEspecialidadMedico.BoundText = doProgramacion.IdEspecialidad
    mo_cmbIdServicio.BoundText = doProgramacion.IdServicio
    
    mo_cmbIdTurno.BoundText = doProgramacion.IdTurno
    Me.txtHoraInicio = doProgramacion.HoraInicio
    Me.txtHoraFin = doProgramacion.HoraFin
    
    Me.lblColor.BackColor = doProgramacion.Color
    mb_ExistenDatos = True
   
    Set oRsTmp1 = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla ProgramacionMedica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdProgramacion = 0
           Me.idMedico = 0
           Me.IdDepartamento = 0
           Me.IdEspecialidad = 0
           Me.idTipoServicio = 0
           Me.NroCuposCE = 0
           Me.txtDescripcion.Text = ""
           mo_cmbIdTipoProgramacion.BoundText = ""
   
End Sub

Function ObtieneCodigoDeProgramacionPorDia(daDiaSeleccionado As Date)
Dim sTitulo As String
Dim programacion As PVAppointment
Dim oDOProgramacion As New DOProgramacionMedica
Dim oDOTurno As New doTurno


        sTitulo = ""
        Set programacion = Diario.AppointmentSet.Get(daDiaSeleccionado)
        Do While Not programacion Is Nothing
            If Format(programacion.StartDateTime, sighEntidades.DevuelveFechaSoloFormato_DMY) <> daDiaSeleccionado Then
                Exit Do
            End If
            
            Set oDOProgramacion = programacion.DataVariant
            Set oDOTurno = mo_AdminProgMedica.TurnosSeleccionarPorId(oDOProgramacion.IdTurno)
            
            sTitulo = sTitulo + oDOTurno.Codigo + Chr(13)
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
        If sTitulo <> "" Then
            ObtieneCodigoDeProgramacionPorDia = Left(sTitulo, Len(sTitulo) - 1)
        End If


End Function



Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminServiciosHosp = Nothing
    Set mo_AdminProgMedica = Nothing
    Set mo_FechaHora = Nothing
    
    Set mo_cmbIdTipoProgramacion = Nothing
    Set mo_cmbIdEspecialidadMedico = Nothing
    Set mo_cmbIdTurno = Nothing
    Set mo_cmbIdTipoServicio = Nothing
    Set mo_cmbIdServicio = Nothing
    Set lcBuscaParametro = Nothing
End Sub


'actualizado 26102014 yamill
Private Function getFechaInicio(sFecha As Date, sHoraInicio As String) As Date
    getFechaInicio = Format(sFecha, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraInicio
End Function


Private Function getFechaFin(sFecha As Date, sHoraInicio As String, sHoraFin As String) As Date
    sHoraInicio = Format(sHoraInicio, "hh:mm")
    sHoraFin = Format(sHoraFin, "hh:mm")
    If sHoraFin > sHoraInicio Then
        getFechaFin = Format(sFecha, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraFin
    Else
        sFecha = DateAdd("d", 1, sFecha)
        getFechaFin = Format(sFecha, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & sHoraFin
    End If
End Function

'Actualizado 30102014 yamill palomino
Private Function validarProgramacionMedicoEnOtrosServicios(lIdServicio As Long, lIDMedico As Long, _
                    sFecha As String, sHoraInicio As String, sHoraFin As String) As Boolean
                    
    Dim oRsTmp As ADODB.Recordset
    Dim dFechaInicio As Date, dFechaFin As Date
    Dim dFechaInicioProgramacion As Date, dFechaFinProgracion As Date
    Dim oDOProgramacionMedica As New DOProgramacionMedica
    Dim returnValue As Boolean
   
    
    oDOProgramacionMedica.IdServicio = lIdServicio
    oDOProgramacionMedica.idMedico = lIDMedico
    oDOProgramacionMedica.fecha = sFecha
    Dim oMensaje2 As New SIGHNegocios.clMensaje
    
    validarProgramacionMedicoEnOtrosServicios = True
    
    dFechaInicioProgramacion = getFechaInicio(CDate(sFecha), sHoraInicio)
    dFechaFinProgracion = getFechaFin(CDate(sFecha), sHoraInicio, sHoraFin)
    
    Set oRsTmp = mo_AdminProgMedica.ProgramacionMedicaSeleccionarPorMedicoEnOtrosServicio(oDOProgramacionMedica)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If oRsTmp.Fields!fecha = CDate(sFecha) Then
                
                dFechaInicio = getFechaInicio(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio)
                dFechaFin = getFechaFin(oRsTmp.Fields!fecha, oRsTmp.Fields!HoraInicio, oRsTmp.Fields!HoraFin)

                If (dFechaInicio <= dFechaInicioProgramacion And dFechaFin >= dFechaFinProgracion) _
                    Or (dFechaInicioProgramacion <= dFechaInicio And dFechaFinProgracion >= dFechaFin) _
                    Or (dFechaInicioProgramacion > dFechaInicio And dFechaInicioProgramacion < dFechaFin) _
                    Or (dFechaFinProgracion > dFechaInicio And dFechaFinProgracion < dFechaFin) Then
                    If mi_Opcion = sghAgregar Then
                        oMensaje2.MostrarFormulario "Médico fué programado en otro servicio para esa FECHA/HORA", Me.Caption
                        Set oMensaje2 = Nothing
                        validarProgramacionMedicoEnOtrosServicios = False
                        Exit Function
                    ElseIf mi_Opcion = sghModificar And oRsTmp!IdProgramacion <> ml_IdProgramacion Then
                        oMensaje2.MostrarFormulario "Médico fué programado en otro servicio para esa FECHA/HORA", Me.Caption
                        Set oMensaje2 = Nothing
                        validarProgramacionMedicoEnOtrosServicios = False
                        Exit Function
                    End If
                End If
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function



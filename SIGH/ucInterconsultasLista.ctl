VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucInterconsultasLista 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   10005
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   75
      TabIndex        =   5
      Top             =   540
      Width           =   9915
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8955
         Picture         =   "ucInterconsultasLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7590
         Picture         =   "ucInterconsultasLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   3765
         MaxLength       =   20
         TabIndex        =   2
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   1845
         MaxLength       =   20
         TabIndex        =   1
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   135
         MaxLength       =   9
         TabIndex        =   0
         Top             =   435
         Width           =   1635
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   5685
         MaxLength       =   20
         TabIndex        =   3
         Top             =   450
         Width           =   1845
      End
      Begin VB.Label Label2 
         Caption         =   "  Nº Historia clínica        Apellido paterno         Apellido materno          Primer nombre                    "
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
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   7095
      End
   End
   Begin UltraGrid.SSUltraGrid grdInterconsultas 
      Height          =   4245
      Left            =   75
      TabIndex        =   4
      Top             =   1500
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   7488
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de interconsultas"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Interconsultas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   15
      Width           =   10005
   End
End
Attribute VB_Name = "ucInterconsultasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_Teclado As New SIGHComun.Teclado
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoServicio As sghTipoServicio

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdInterconsultas.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdInterconsultas.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoServicio(lValue As sghTipoServicio)
    ml_TipoServicio = lValue
End Property
Property Get TipoServicio() As sghTipoServicio
    TipoServicio = ml_TipoServicio
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim oDOPaciente As New doPaciente
        
        oDOPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oDOPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oDOPaciente.PrimerNombre = UserControl.txtPrimerNombre
        oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
        
        Select Case ml_TipoServicio
        Case sghConsultaExterna
            Set grdInterconsultas.DataSource = mo_AdminAdmision.AtencionesInterconsultasFiltrarConsultaExterna(oDOPaciente)
        Case sghEmergenciaConsultorios
            Set grdInterconsultas.DataSource = mo_AdminAdmision.AtencionesInterconsultasFiltrarConsultorioEmergencia(oDOPaciente)
        Case sghEmergenciaObservacion
            Set grdInterconsultas.DataSource = mo_AdminAdmision.AtencionesInterconsultasFiltrarObservacionEmergencia(oDOPaciente)
        Case sghHospitalizacion
            Set grdInterconsultas.DataSource = mo_AdminAdmision.AtencionesInterconsultasFiltrarHospitalizacion(oDOPaciente)
        End Select
        
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdInterconsultas, SIGHComun.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtPrimerNombre = ""
        UserControl.txtNroHistoria = ""
End Sub

Private Sub grdInterconsultas_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdInterconsultas.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdInterconsulta")
End Sub

Private Sub grdInterconsultas_Click()
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdInterconsultas.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdInterconsulta")
    
End Sub


Private Sub grdInterconsultas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdInterconsultas.Bands(0).Columns("IdInterconsulta").Hidden = True
    grdInterconsultas.Bands(0).Columns("IdAtencion").Hidden = True
    grdInterconsultas.Bands(0).Columns("IdCuentaAtencion").Hidden = True
    
    grdInterconsultas.Bands(0).Columns("FechaRealizacion").Header.Caption = "Fecha Realización"
    grdInterconsultas.Bands(0).Columns("FechaRealizacion").Width = 2000
    
    grdInterconsultas.Bands(0).Columns("HoraRealizacion").Header.Caption = "Hora Realizacion"
    grdInterconsultas.Bands(0).Columns("HoraRealizacion").Width = 1000
    
    grdInterconsultas.Bands(0).Columns("FechaSolicitud").Header.Caption = "Fecha Solicitud"
    grdInterconsultas.Bands(0).Columns("FechaSolicitud").Width = 2000
    
    grdInterconsultas.Bands(0).Columns("HoraSolicitud").Header.Caption = "Hora Solicitud"
    grdInterconsultas.Bands(0).Columns("HoraSolicitud").Width = 1000
    
    grdInterconsultas.Bands(0).Columns("NombrePaciente").Header.Caption = "Paciente"
    grdInterconsultas.Bands(0).Columns("NombrePaciente").Width = 3000
    
    grdInterconsultas.Bands(0).Columns("NombreMedico").Header.Caption = "Medico Realiza"
    grdInterconsultas.Bands(0).Columns("NombreMedico").Width = 3000
    

End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
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
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdInterconsultas.Width = fraBusqueda.Width
   grdInterconsultas.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub



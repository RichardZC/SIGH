VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucMedicosLista 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   10140
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
      Top             =   510
      Width           =   10050
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8295
         Picture         =   "ucMensajesParpadeando.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   6930
         Picture         =   "ucMensajesParpadeando.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtCodigoPlanilla 
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   0
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtNombres 
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
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   3
         Top             =   450
         Width           =   1845
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
         Left            =   3090
         MaxLength       =   30
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
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   1
         Top             =   450
         Width           =   1845
      End
      Begin VB.Label Label2 
         Caption         =   "Cod. planilla      Apellido paterno         Apellido materno               Nombres                           "
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
         TabIndex        =   6
         Top             =   240
         Width           =   7395
      End
   End
   Begin UltraGrid.SSUltraGrid grdMedicos 
      Height          =   4290
      Left            =   75
      TabIndex        =   4
      Top             =   1455
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7567
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
      Caption         =   "Relación de médicos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Profesional de la Salúd"
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
      Left            =   15
      TabIndex        =   7
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "ucMedicosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_Apariencia As New sighcomun.GridInfragistic
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New sighcomun.Teclado
Dim ml_IdEspecialidad As Long
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)

Property Let IdEspecialidad(lValue As Long)
    ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
    IdEspecialidad = ml_IdEspecialidad
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdMedicos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdMedicos.DataSource
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
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado

        oDOEmpleado.ApellidoPaterno = UserControl.txtApellidoPaterno
        oDOEmpleado.ApellidoMaterno = UserControl.txtApellidoMaterno
        oDOEmpleado.Nombres = UserControl.txtNombres
        oDOEmpleado.CodigoPlanilla = UserControl.txtCodigoPlanilla
        
        Set grdMedicos.DataSource = mo_AdminProgramacionMedica.MedicosFiltrar(oDoMedico, oDOEmpleado, ml_IdEspecialidad)
        
        If mo_AdminProgramacionMedica.MensajeError <> "" Then
            MsgBox "Error leyendo datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbCritical, "Profesional de la Salud"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdMedicos, sighcomun.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNombres = ""
        UserControl.txtCodigoPlanilla = ""
End Sub

Private Sub grdMedicos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdMedicos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdMedico")
End Sub




Private Sub grdMedicos_DblClick()
    Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdMedicos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdMedico")
    RaiseEvent SeleccionaRegistro(ml_IdRegistroSeleccionado)
End Sub

Private Sub grdMedicos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdMedicos.Bands(0).Columns("IdMedico").Hidden = True
    
    grdMedicos.Bands(0).Columns("CodigoPlanilla").Header.Caption = "Cod. Planilla"
    grdMedicos.Bands(0).Columns("CodigoPlanilla").Width = 800
    
    grdMedicos.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Apellido Paterno"
    grdMedicos.Bands(0).Columns("ApellidoPaterno").Width = 2000
    
    grdMedicos.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Apellido Materno"
    grdMedicos.Bands(0).Columns("ApellidoMaterno").Width = 2000
    
    grdMedicos.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdMedicos.Bands(0).Columns("Nombres").Width = 3000
    
    grdMedicos.Bands(0).Columns("Especialidad").Header.Caption = "Especialidad"
    grdMedicos.Bands(0).Columns("Especialidad").Width = 2000
    
End Sub

Private Sub grdMedicos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdMedicos_DblClick
    End If
End Sub

Private Sub txtCodigoPlanilla_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoPlanilla
End Sub

Private Sub txtCodigoPlanilla_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombres
End Sub


Private Sub txtNombres_KeyPress(KeyAscii As Integer)
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
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = UserControl.Width
   grdMedicos.Width = fraBusqueda.Width
   grdMedicos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub




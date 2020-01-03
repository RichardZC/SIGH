VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucHISEstablecimientos 
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   ScaleHeight     =   6030
   ScaleWidth      =   10200
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
      Height          =   705
      Left            =   75
      TabIndex        =   0
      Top             =   510
      Width           =   10080
      Begin VB.TextBox txtNombre 
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
         Left            =   1350
         TabIndex        =   3
         Top             =   240
         Width           =   2715
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4140
         Picture         =   "ucHISEstablecimientos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   5505
         Picture         =   "ucHISEstablecimientos.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   4
         Top             =   270
         Width           =   675
      End
   End
   Begin UltraGrid.SSUltraGrid grdTurnos 
      Height          =   4665
      Left            =   90
      TabIndex        =   5
      Top             =   1275
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   8229
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
      Caption         =   "Lista de Turnos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Turnos"
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
      TabIndex        =   6
      Top             =   0
      Width           =   10155
   End
End
Attribute VB_Name = "ucHISEstablecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Turnos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim ml_idRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdTurnos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdTurnos.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
        Set grdTurnos.DataSource = mo_AdminProgramacionMedica.TurnosSeleccionarTodos()
        mo_Apariencia.ConfigurarFilasBiColores grdTurnos, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNombre = ""
End Sub

Private Sub grdTurnos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTurnos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdTurno")
End Sub

Private Sub grdTurnos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
Cancel = True
End Sub

Private Sub grdTurnos_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTurnos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdTurno")
    
End Sub


Private Sub grdTurnos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdTurnos.Bands(0).Columns("IdTurno").Hidden = True
    
    grdTurnos.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdTurnos.Bands(0).Columns("Codigo").Width = 750
    
    grdTurnos.Bands(0).Columns("Descripcion").Header.Caption = "Descripcion"
    grdTurnos.Bands(0).Columns("Descripcion").Width = 5000
    
    grdTurnos.Bands(0).Columns("HoraInicio").Header.Caption = "Hora Inicio"
    grdTurnos.Bands(0).Columns("HoraInicio").Width = 1500
    
    grdTurnos.Bands(0).Columns("HoraFin").Header.Caption = "Hora Fin"
    grdTurnos.Bands(0).Columns("HoraFin").Width = 1500
    
    grdTurnos.Bands(0).Columns("TipoServicio").Hidden = True
    grdTurnos.Bands(0).Columns("TipoServicio").Header.Caption = "Tipo Servicios"
    grdTurnos.Bands(0).Columns("TipoServicio").Width = 2000
    
    grdTurnos.Bands(0).Columns("Especialidad").Hidden = True
    grdTurnos.Bands(0).Columns("Especialidad").Header.Caption = "Especialidad"
    grdTurnos.Bands(0).Columns("Especialidad").Width = 2000
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = UserControl.Width
   grdTurnos.Width = UserControl.Width - 150
   grdTurnos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 100)
   
End Sub



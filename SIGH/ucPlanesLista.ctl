VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucPlanesLista 
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   LockControls    =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   10095
   Begin UltraGrid.SSUltraGrid grdPlanes 
      Height          =   4605
      Left            =   75
      TabIndex        =   5
      Top             =   1335
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   8123
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
      Caption         =   "Lista de planes"
   End
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
      Left            =   60
      TabIndex        =   2
      Top             =   555
      Width           =   10020
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1350
         TabIndex        =   0
         Top             =   240
         Width           =   2715
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4140
         Picture         =   "ucPlanesLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "[F6]"
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
         Left            =   4710
         TabIndex        =   6
         Top             =   270
         Width           =   330
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
         TabIndex        =   3
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Planes de Atención"
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
      TabIndex        =   4
      Top             =   15
      Width           =   10095
   End
End
Attribute VB_Name = "ucPlanesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdPlanes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdPlanes.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
        
        Set grdPlanes.DataSource = mo_AdminFacturacion.PlanesSeleccionarTodos()
        mo_Apariencia.ConfigurarFilasBiColores grdPlanes, SIGHComun.GrillaConFilasBicolor

End Sub

Private Sub grdPlanes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPlanes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdPlan")
End Sub

Private Sub grdPlanes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPlanes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdPlan")
    
End Sub

Private Sub grdPlanes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdPlanes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IDPlan")
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdPlanes.Width = fraBusqueda.Width
   grdPlanes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub



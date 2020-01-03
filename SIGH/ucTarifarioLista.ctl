VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucTarifarioLista 
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   LockControls    =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   10125
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
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
      TabIndex        =   3
      Top             =   570
      Width           =   9975
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
         TabIndex        =   0
         Top             =   240
         Width           =   2715
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4140
         Picture         =   "ucTarifarioLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
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
         Height          =   270
         Left            =   4710
         TabIndex        =   7
         Top             =   255
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
         TabIndex        =   4
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Frame fraResultado 
      Height          =   4575
      Left            =   60
      TabIndex        =   2
      Top             =   1290
      Width           =   9975
      Begin UltraGrid.SSUltraGrid grdTarifario 
         Height          =   4215
         Left            =   135
         TabIndex        =   6
         Top             =   225
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   7435
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
         Caption         =   "Lista de precios"
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Tarifario"
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
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   10035
   End
End
Attribute VB_Name = "ucTarifarioLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdTarifario.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdTarifario.DataSource
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
    Set grdTarifario.DataSource = mo_AdminFacturacion.ProductosSeleccionarTodos()
    mo_Apariencia.ConfigurarFilasBiColores grdTarifario, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub grdTarifario_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTarifario.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
    
End Sub

Private Sub grdTarifario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdTarifario.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = fraBusqueda.Width
   
   fraResultado.Width = UserControl.Width - 150
   grdTarifario.Width = fraResultado.Width - 250
   
   fraResultado.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   grdTarifario.Height = fraResultado.Height - 320
   
End Sub




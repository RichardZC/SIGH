VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.UserControl ucPreLiqRecetasLista 
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   ScaleHeight     =   5970
   ScaleWidth      =   10125
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      Height          =   705
      Left            =   60
      TabIndex        =   2
      Top             =   570
      Width           =   9975
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   240
         Width           =   2715
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4140
         Picture         =   "ucPreLiqRecetas.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   345
         Left            =   300
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Frame fraResultado 
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   1290
      Width           =   9975
      Begin MSDataGridLib.DataGrid grdRecetas 
         Height          =   4185
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   7382
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Pre-Liquidacion de Recetas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   10035
   End
End
Attribute VB_Name = "ucPreLiqRecetasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminProgramacionMedica As New SIGHReglasNegocios.ReglasDeProgMedica
Dim ml_IdRegistroSeleccionado As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdRecetas.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdRecetas.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property

Private Sub btnBuscar_Click()
        Set grdRecetas.DataSource = mo_AdminProgramacionMedica.TurnosSeleccionarTodos()
End Sub

Private Sub grdRecetas_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdRecetas.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdTurno")
    
End Sub

Private Sub grdRecetas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdRecetas.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdTurno")
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 100
   lblNombre.Width = fraBusqueda.Width
   
   fraResultado.Width = UserControl.Width - 100
   grdRecetas.Width = fraResultado.Width - 260
   
   fraResultado.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 100)
   grdRecetas.Height = fraResultado.Height - 320
   
End Sub




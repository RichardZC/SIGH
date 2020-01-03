VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucInterconsulta 
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11040
   ScaleHeight     =   5475
   ScaleWidth      =   11040
   Begin VB.CommandButton btnAgregarDx 
      DisabledPicture =   "ucInterconsulta.ctx":0000
      DownPicture     =   "ucInterconsulta.ctx":03DF
      Height          =   315
      Left            =   180
      Picture         =   "ucInterconsulta.ctx":07CD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   390
      Width           =   1005
   End
   Begin VB.CommandButton btnQuitarDx 
      DisabledPicture =   "ucInterconsulta.ctx":0BD9
      DownPicture     =   "ucInterconsulta.ctx":0F64
      Height          =   315
      Left            =   1260
      Picture         =   "ucInterconsulta.ctx":12F7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   390
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Interconsultas"
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   10845
   End
   Begin UltraGrid.SSUltraGrid grdInterconsultas 
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7752
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "Lista de interconsultas"
   End
End
Attribute VB_Name = "ucInterconsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

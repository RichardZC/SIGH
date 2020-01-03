VERSION 5.00
Begin VB.UserControl UcTipoImpresion 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   LockControls    =   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   1410
   Begin VB.Frame fraTreporte 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1410
      Begin VB.OptionButton optImpresora 
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
         Left            =   45
         Picture         =   "UcTipoImpresion.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Impresión normal"
         Top             =   150
         Value           =   -1  'True
         Width           =   420
      End
      Begin VB.OptionButton optPDF 
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
         Left            =   930
         Picture         =   "UcTipoImpresion.ctx":03DE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Genera archivo PDF"
         Top             =   150
         Width           =   420
      End
      Begin VB.OptionButton optExcel 
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
         Left            =   495
         Picture         =   "UcTipoImpresion.ctx":0802
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Impresión en EXCEL"
         Top             =   150
         Width           =   420
      End
   End
End
Attribute VB_Name = "UcTipoImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_OpcionImpresionElegida As sghTipoImpresion

Property Get OpcionImpresionElejida() As sghTipoImpresion
    OpcionImpresionElejida = m_OpcionImpresionElegida
End Property
Private Sub optExcel_Click()
    m_OpcionImpresionElegida = sghTIexcel
End Sub

Private Sub optImpresora_Click()
    m_OpcionImpresionElegida = sghTIimpresora
End Sub

Private Sub optPDF_Click()
    m_OpcionImpresionElegida = sghTIpdf
End Sub

Private Sub UserControl_Initialize()
    m_OpcionImpresionElegida = sghTIimpresora
End Sub

VERSION 5.00
Begin VB.UserControl ucMensajeParpadeando 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ScaleHeight     =   405
   ScaleWidth      =   3735
   Begin VB.Label labelFrente 
      BackStyle       =   0  'Transparent
      Caption         =   "R E P O R T E S    1 9 8 . 7 9"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label labelFondo 
      BackStyle       =   0  'Transparent
      Caption         =   "R E P O R T E S    1 9 8 . 7 9    "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   20
      TabIndex        =   1
      Top             =   20
      Width           =   3735
   End
End
Attribute VB_Name = "ucMensajeParpadeando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Mensajes en color ROJO
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Private mcharaccept As String




Private Sub UserControl_Resize()
   On Error Resume Next
   labelFondo.Top = 20
   labelFondo.Left = 20
   labelFondo.Width = UserControl.Width
   labelFondo.Height = UserControl.Height
   '
   labelFrente.Top = 0
   labelFrente.Left = 0
   labelFrente.Width = UserControl.Width
   labelFrente.Height = UserControl.Height
End Sub



Public Property Get MensajeDeTexto() As String
   MensajeDeTexto = mcharaccept
End Property

Public Property Let MensajeDeTexto(ByVal lcNuevoValorTexto As String)
    mcharaccept = lcNuevoValorTexto
    PropertyChanged "MensajeDeTexto"
    labelFondo.Caption = lcNuevoValorTexto
    labelFrente.Caption = lcNuevoValorTexto
End Property

Function DevuelveTextoConBlancoEntreLetras(lcTexto As String) As String
   Dim lcNewTexto As String
   Dim lnLen As Integer
   lcNewTexto = ""
   For lnLen = 1 To Len(lcTexto)
       lcNewTexto = lcNewTexto & Mid(lcTexto, lnLen, 1) & " "
   Next
   DevuelveTextoConBlancoEntreLetras = lcNewTexto
End Function

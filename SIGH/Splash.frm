VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   6465
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splash.frx":000C
   ScaleHeight     =   6465
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SISGalenPLUS v3.07062021u75hra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   5760
      Width           =   5685
   End
   Begin VB.Label lblBuild 
      BackStyle       =   0  'Transparent
      Caption         =   "Revisión:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B9553C&
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   4110
      Visible         =   0   'False
      Width           =   3450
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Splash
'        Programado por: William C
'        Fecha: Enero 2006
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mb_FormLoad  As Boolean
Dim mb_MostrarCreditos  As Boolean
Property Let MostrarCreditos(bValue As Boolean)
    mb_MostrarCreditos = bValue
End Property

Private Sub Form_Activate()
    WxLcVersionSisGalenPlus = Label1.Caption
    If mb_FormLoad Then
        mb_FormLoad = False
        If mb_MostrarCreditos Then
            'Creditos.Show 1
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblBuild.Caption = "Revisión: " & App.Major & "." & App.Minor & "." & App.Revision
    mb_FormLoad = True
End Sub


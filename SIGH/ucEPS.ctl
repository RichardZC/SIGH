VERSION 5.00
Begin VB.UserControl ucEPS 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   345
   ScaleWidth      =   5385
   Begin VB.TextBox txtPorcentaje 
      Alignment       =   1  'Right Justify
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
      Left            =   1740
      TabIndex        =   2
      Text            =   "100.00"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Porcentaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   885
   End
   Begin VB.Label lblPrestacion 
      AutoSize        =   -1  'True
      Caption         =   "% que la ASEGURADORA cubre"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2610
      TabIndex        =   0
      Top             =   30
      Width           =   2190
   End
End
Attribute VB_Name = "ucEPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Historia Clinica
'        Programado por: Barrantes D
'        Fecha: Octubre 2018
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ms_MensajeError As String
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Property Get MensajeError() As String
  MensajeError = ms_MensajeError
End Property


Property Let Porcentaje(lValue As Double)
    txtPorcentaje.Text = Trim(Str(lValue))
End Property
Property Get Porcentaje() As Double
    Porcentaje = CCur(txtPorcentaje.Text)
End Property


Public Sub inicializar()
    txtPorcentaje.Text = "0.00"
End Sub





Public Sub HabilitaPorcentaje(lbHabilitar As Boolean)
       If lbHabilitar = False Then
            txtPorcentaje.Locked = True
            txtPorcentaje.BackColor = &HF9EADF
            txtPorcentaje.ForeColor = &H808080
       Else
            txtPorcentaje.Locked = False
            txtPorcentaje.BackColor = &HFFFFFF
            txtPorcentaje.ForeColor = &H0&
       End If
End Sub

Public Sub FocusEnCodigoPrestacion()
    On Error Resume Next
    txtPorcentaje.SetFocus
End Sub


Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPorcentaje
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtPorcentaje_LostFocus()
    If Val(txtPorcentaje.Text) >= 1 And Val(txtPorcentaje.Text) <= 100 Then
    Else
       MsgBox "El porcentaje debe estar entre 1 y 100", vbInformation, ""
       txtPorcentaje.Text = "0.00"
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    If Val(txtPorcentaje.Text) >= 1 And Val(txtPorcentaje.Text) <= 100 Then
       ValidaDatosObligatorios = True
    Else
       ms_MensajeError = "El PORCENTAJE debe estar entre 1 y 100"
       ValidaDatosObligatorios = False
    End If
End Function

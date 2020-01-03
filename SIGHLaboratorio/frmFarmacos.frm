VERSION 5.00
Begin VB.Form frmFarmacos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FÁRMACOS"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   ForeColor       =   &H00000000&
   Icon            =   "frmFarmacos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "frmFarmacos.frx":000C
      DownPicture     =   "frmFarmacos.frx":04D0
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Picture         =   "frmFarmacos.frx":09BC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1365
   End
   Begin VB.ListBox lstF 
      Height          =   2790
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox cboGF 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre Genérico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grupo Farmacológico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "frmFarmacos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ml_idFormulario As Long
Dim I As Integer

Property Let idFormulario(lValue As Long)
   ml_idFormulario = lValue
End Property

Property Get idFormulario() As Long
   idFormulario = ml_idFormulario
End Property

Public Sub LlenaComboFarmacos(GF As ComboBox)
  GF.AddItem "Todos los grupos"
  GF.AddItem "PENICILINAS"
  GF.AddItem "CEFALOSPIRINAS"
  GF.AddItem "MACROLIDOS"
  GF.AddItem "AMINOGLUCOSIDOS"
  GF.AddItem "TETRACICLINAS"
  GF.AddItem "QUINOLONAS"
  GF.AddItem "SULFONAMIDAS"
  GF.AddItem "NITROFURANOS"
  GF.AddItem "DER. ÁCIDO TRICLORACET"
  GF.AddItem "LINCOSAMIDA"
  GF.AddItem "GLICOPEPTIDO"
  GF.AddItem "MONOBACTANS"
  GF.AddItem "DERIVADOS ÁCIDO FOSFÓRICO"
  GF.AddItem "NITROIMIDAZOL"
End Sub

Public Sub LlenaListFarmacos(Indice As Integer, F As ListBox)
  If Indice = 1 Then
    F.AddItem "Penicilina"
    F.AddItem "Ampicilina"
    F.AddItem "Amox. + Ac. Clavulánico"
    F.AddItem "Oxacilina"
    F.AddItem "Dicloxacilina"
  ElseIf Indice = 2 Then
    F.AddItem "Cefazolina"
    F.AddItem "Ceftazidina"
    F.AddItem "Ceftriaxona"
    F.AddItem "Cefradina"
    F.AddItem "Cefpirome"
  ElseIf Indice = 3 Then
    F.AddItem "Eritromicina"
    F.AddItem "Claritromicina"
    F.AddItem "Azitromicina"
  ElseIf Indice = 4 Then
    F.AddItem "Amikacina"
    F.AddItem "Gentamicina"
  ElseIf Indice = 5 Then
    F.AddItem "Tetraciclina"
    F.AddItem "Doxiciclina"
  ElseIf Indice = 6 Then
    F.AddItem "Ciprofloxacina"
    F.AddItem "Norfloxacina"
    F.AddItem "Ofloxacina"
    F.AddItem "Ácido Pipemídico"
    F.AddItem "Ácido Nalidíxico"
  ElseIf Indice = 7 Then
    F.AddItem "Cotrimoxazol"
  ElseIf Indice = 8 Then
    F.AddItem "Nitrofurantoína"
    F.AddItem "Furazolidona"
  ElseIf Indice = 9 Then
    F.AddItem "Cloranfenicol"
  ElseIf Indice = 10 Then
    F.AddItem "Clindamicina"
  ElseIf Indice = 11 Then
    F.AddItem "Vancomicina"
  ElseIf Indice = 12 Then
    F.AddItem "Aztreonam"
    F.AddItem "Inipenem"
  ElseIf Indice = 13 Then
    F.AddItem "Fosfomicina Trometanol"
  Else
    F.AddItem "Metronidazol"
  End If
End Sub

Private Sub cboGF_Click()
  lstF.Clear
  If cboGF.ListIndex = 0 Then
    For I = 1 To cboGF.ListCount - 1
      Call LlenaListFarmacos(I, lstF)
    Next I
  Else
    Call LlenaListFarmacos(cboGF.ListIndex, lstF)
  End If
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
  If idFormulario = 0 Then
    'frmMicrobiologia.MIC001_05.SetFocus
  Else
    'frmMicrobiologia.MIC001_06.SetFocus
  End If
End Sub

Private Sub Form_Load()
  'Me.Top = frmMicrobiologia.Top + 1700 '+ 2000
  'Me.Left = frmMicrobiologia.Left + frmMicrobiologia.Width - 600
  Call LlenaComboFarmacos(cboGF)
  cboGF.ListIndex = 0
End Sub

Private Sub lstF_Click()
  Dim Temp1 As String, Temp2 As String
'  MsgBox 1
  If idFormulario = 0 Then
    Temp1 = Trim(frmMicrobiologia.MIC001_05.Text)
  Else
    Temp1 = Trim(frmMicrobiologia.MIC001_06.Text)
  End If
  
  Temp2 = lstF.List(lstF.ListIndex)
  If InStr(1, Temp1, Temp2, vbTextCompare) <> 0 Then
    MsgBox "El fármaco " & Chr(34) & UCase(Temp2) & Chr(34) & " ya ha sido agregado a la lista."
    Exit Sub
  End If
  If Len(Temp1) = 0 Then
    If idFormulario = 0 Then
      frmMicrobiologia.MIC001_05.Text = Temp2
    Else
      frmMicrobiologia.MIC001_06.Text = Temp2
    End If
  Else
    If idFormulario = 0 Then
      frmMicrobiologia.MIC001_05.Text = Temp1 & ", " & Temp2
    Else
      frmMicrobiologia.MIC001_06.Text = Temp1 & ", " & Temp2
    End If
  End If
End Sub

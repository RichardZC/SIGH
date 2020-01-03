VERSION 5.00
Begin VB.Form Camara1 
   Caption         =   "WebCam  de la PC"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   Icon            =   "Camara1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   3795
      Index           =   1
      Left            =   4965
      TabIndex        =   4
      Top             =   0
      Width           =   1545
      Begin VB.CommandButton Command4 
         Caption         =   "Configurar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   90
         TabIndex        =   8
         Top             =   2985
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   90
         TabIndex        =   7
         Top             =   2050
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Detener"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   90
         TabIndex        =   6
         Top             =   1115
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3780
      Left            =   0
      ScaleHeight     =   3720
      ScaleWidth      =   4860
      TabIndex        =   3
      Top             =   0
      Width           =   4920
   End
   Begin VB.Frame Frame 
      Height          =   1005
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   3885
      Width           =   6465
      Begin VB.CommandButton Command5 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         DisabledPicture =   "Camara1.frx":000C
         DownPicture     =   "Camara1.frx":04D0
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3308
         Picture         =   "Camara1.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         DisabledPicture =   "Camara1.frx":0EA8
         DownPicture     =   "Camara1.frx":1308
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   1868
         Picture         =   "Camara1.frx":177D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
      Begin VB.PictureBox Picture2 
         Height          =   540
         Left            =   5070
         Picture         =   "Camara1.frx":1BF2
         ScaleHeight     =   480
         ScaleWidth      =   615
         TabIndex        =   1
         Top             =   255
         Visible         =   0   'False
         Width           =   675
      End
   End
End
Attribute VB_Name = "Camara1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temp As Long


  
Private Sub cmdGrabar_Click()
On Error GoTo Fallo
'**************captura imagen
temp = SendMessage(hwdc, WM_CAP_EDIT_COPY, 1, 0)
Picture2.Picture = Clipboard.GetData
'*************graba imagen
Dim Ruta As String
Dim x As Variant
Ruta = "c:\imagen.jpg"

x = GetAttr(Ruta)
'v = MsgBox("Ya se ha tomado una foto para este alumno, ¿Desea remplazarla?.", vbYesNo, "Precaución")
'If MsgBox("Ya se ha tomado una foto para este alumno, ¿Desea remplazarla?.", vbYesNo, "Precaución") = vbYes Then
SavePicture Picture2, Ruta
'MsgBox "La foto a sido Guardada en " & Ruta & " exitosamente."
Command2_Click
Me.Visible = False

'Else
'End If
Exit Sub
Fallo:

MsgBox "no se grabó la foto" & Chr(13) & Err.Description



End Sub


' botón que inicia la captura
'''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
Dim temp As Long
  
  hwdc = capCreateCaptureWindow("CapWindow", ws_child Or ws_visible, _
                                    0, 0, 320, 240, Picture1.hwnd, 0)
  If (hwdc <> 0) Then
    temp = SendMessage(hwdc, wm_cap_driver_connect, 0, 0)
    temp = SendMessage(hwdc, wm_cap_set_preview, 1, 0)
    temp = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    temp = SendMessage(hwdc, WM_CAP_SET_SCALE, True, 0)
    'esto hace que la imagen recibida por el dispositivo se ajuste
    'al tamaño de la ventana de captura (justo lo que yo buscaba)
    DoEvents
    startcap = True
    Else
    MsgBox "No hay Camara Web", 48, "Error"
  End If
  
End Sub
  
' botón para detener la captura
'''''''''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
      
    temp = DestroyWindow(hwdc)
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
        DoEvents
        startcap = False
    End If
  
End Sub
  
' Botón que abre el dialogo de formato
''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command3_Click()
        If startcap = True Then
              
            temp = SendMessage(hwdc, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
            DoEvents
        End If
End Sub
' Mostrar dialogo de Configuracion de la WebCam
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command4_Click()
 Dim temp As Long
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DLG_VIDEOCONFIG, 0&, 0&)
        DoEvents
    End If
End Sub
  




Private Sub Command5_Click()
    Command2_Click
    Me.Visible = False
End Sub

Private Sub Form_Load()
    Command1.Caption = "Iniciar"
    Command2.Caption = "Detener"
    Command3.Caption = "Formato"
    Command4.Caption = "Configurar"
    Me.Caption = "Capturador de Web Cam"
    Command1_Click

End Sub
  
Private Sub Form_Resize()
    On Error Resume Next
    Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29
End Sub
  
Private Sub Form_Unload(Cancel As Integer)
  
    temp = DestroyWindow(hwdc)
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
        DoEvents
        startcap = False
    End If
End Sub







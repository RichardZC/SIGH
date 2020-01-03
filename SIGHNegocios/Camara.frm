VERSION 5.00
Begin VB.Form Camara 
   Caption         =   "WebCam  de la PC"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   Icon            =   "Camara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   75
      Top             =   7725
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   900
      ScaleHeight     =   735
      ScaleWidth      =   1020
      TabIndex        =   9
      Top             =   7605
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      Height          =   810
      Left            =   2205
      ScaleHeight     =   750
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   7605
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtCamarita 
      Height          =   285
      Left            =   9105
      TabIndex        =   7
      Text            =   "TxtCamarita"
      Top             =   6090
      Width           =   3600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   15
      TabIndex        =   6
      Text            =   "imagen"
      Top             =   6060
      Width           =   8850
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   12705
      Begin VB.CommandButton Command3 
         Height          =   330
         Left            =   7995
         TabIndex        =   3
         Top             =   390
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
         DisabledPicture =   "Camara.frx":000C
         DownPicture     =   "Camara.frx":046C
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
         Left            =   6397
         Picture         =   "Camara.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Imagen"
         DisabledPicture =   "Camara.frx":0D56
         DownPicture     =   "Camara.frx":121A
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
         Left            =   5002
         Picture         =   "Camara.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "...."
         Height          =   330
         Left            =   2880
         TabIndex        =   5
         Top             =   405
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "...."
         Height          =   225
         Left            =   2355
         TabIndex        =   4
         Top             =   495
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Image Image1 
      Height          =   5985
      Left            =   0
      Top             =   0
      Width           =   12705
   End
End
Attribute VB_Name = "Camara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
'Image1.Visible = True
'On Error Resume Next
''captura la imagen desde la webcan
'Imagen Camara.hwnd, GET_FRAME, 0, 0
''luego la copia al portapapeles
'Imagen Camara.hwnd, COPY, 0, 0
''posteriormente la pega en el picturebox utilizado
'Image1.Picture = Clipboard.GetData
'Picture2.Picture = Clipboard.GetData
''por ultimo limpia el picturebox
''Clipboard.Clear
''Label2.Visible = True
'Command2.Visible = True
''Command3.Visible = True



Image1.Visible = True
On Error Resume Next
'captura la imagen desde la webcan
Imagen Me.hwnd, GET_FRAME, 0, 0
'luego la copia al portapapeles
Imagen Me.hwnd, COPY, 0, 0
'posteriormente la pega en el picturebox utilizado
Image1.Picture = Clipboard.GetData
Picture2.Picture = Clipboard.GetData
'por ultimo limpia el picturebox
'Clipboard.Clear
'Label2.Visible = True
Command2.Visible = True
'Command3.Visible = True

End Sub

Private Sub Command2_Click()
Dim Ruta As String
Dim x As Variant
Ruta = Text1.Text
'Ruta = "c:\imagen.jpg"
On Error GoTo Fallo
DoEvents: Imagen Auxiliar, DISCONNECT, 0, 0
'x = GetAttr(Ruta)
'v = MsgBox("Ya se ha tomado una foto para este alumno, ¿Desea remplazarla?.", vbYesNo, "Precaución")
'If MsgBox("Ya se ha tomado una foto para este alumno, ¿Desea remplazarla?.", vbYesNo, "Precaución") = vbYes Then
SavePicture Picture2, Ruta
'MsgBox "La foto a sido Guardada en " & Ruta & " exitosamente."
Command3.value = True
'Else
'End If

Imagen Me.hwnd, DISCONNECT, 0, 0
sighentidades.RutaImagenConPermiso = Text1.Text
sighentidades.NombreCamarita = Trim(TxtCamarita.Text)
Timer1.Interval = 0

Me.Visible = False

Exit Sub
Fallo:

 Imagen Auxiliar, DISCONNECT, 0, 0
Timer1.Interval = 0
Me.Visible = False
'SavePicture Picture2, Ruta
'MsgBox "La foto a sido Guardada en " & Ruta & " exitosamente."
'Command3.value = True
End Sub



Private Sub Command3_Click()
Image1.Visible = False
Label2.Visible = False
Command2.Visible = False
Command3.Visible = False
End Sub

Private Sub Command4_Click()
DoEvents: Imagen Auxiliar, DISCONNECT, 0, 0
End Sub

Private Sub Form_Load()
 
 If sighentidades.NombreCamarita = "" Then
    TxtCamarita.Text = "WebcamCapture"
 Else
    TxtCamarita.Text = sighentidades.NombreCamarita
 End If
 
 If sighentidades.RutaImagenConPermiso = "" Then
   Text1.Text = App.Path & "\imagen.jpg"
 Else
   Text1.Text = sighentidades.RutaImagenConPermiso
 End If
 
 
'Auxiliar = video("WebcamCapture", 0, 0, 0, 160, 120, Me.hwnd, 0)
Auxiliar = video(TxtCamarita.Text, 0, 0, 0, 160, 120, Me.hwnd, 0)

DoEvents: Imagen Auxiliar, CONNECT, 0, 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Imagen Me.hwnd, DISCONNECT, 0, 0
Timer1.Interval = 0
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
'captura la imagen desde la webcan
Imagen Auxiliar, GET_FRAME, 0, 0
'luego la copia al portapapeles
Imagen Auxiliar, COPY, 0, 0
'posteriormente la pega en el picturebox utilizado
Picture1.Picture = Clipboard.GetData
'por ultimo limpia el picturebox
'Clipboard.Clear
End Sub



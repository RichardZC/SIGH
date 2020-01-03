VERSION 5.00
Begin VB.Form ArchivoBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de Archivo"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "BusquedaArchivo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4965
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   7035
      Begin VB.DriveListBox Drive1 
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3195
      End
      Begin VB.DirListBox Dir1 
         ForeColor       =   &H8000000D&
         Height          =   4140
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3225
      End
      Begin VB.FileListBox File1 
         ForeColor       =   &H8000000D&
         Height          =   1650
         Left            =   3540
         Pattern         =   "*.img;*.bmp;*.jpg;*.gif;*.dib"
         TabIndex        =   4
         Top             =   210
         Width           =   3375
      End
      Begin VB.Image pi_actual 
         BorderStyle     =   1  'Fixed Single
         Height          =   2805
         Left            =   3540
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   45
      TabIndex        =   0
      Top             =   5040
      Width           =   7050
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BusquedaArchivo.frx":0CCA
         DownPicture     =   "BusquedaArchivo.frx":118E
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
         Left            =   3660
         Picture         =   "BusquedaArchivo.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BusquedaArchivo.frx":1B66
         DownPicture     =   "BusquedaArchivo.frx":1FC6
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
         Left            =   2115
         Picture         =   "BusquedaArchivo.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ArchivoBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Búsqueda de archivo en Windows
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lb_MuestraImagen As Boolean
Dim lc_ArchivoElegido As String
Dim lc_PathDefault As String
Dim lc_TipoArchivo As String            'DEBB2014a

'DEBB2014a
Property Let TipoArchivo(lValue As String)
    lc_TipoArchivo = lValue
End Property

Property Get ArchivoElegido() As String
    ArchivoElegido = lc_ArchivoElegido
End Property
Property Let MuestraImagen(lValue As Boolean)
    lb_MuestraImagen = lValue
End Property
Property Let PathDefault(lValue As String)
    lc_PathDefault = lValue
End Property

Private Sub btnAceptar_Click()
   If Right(Dir1.Path, 1) = "\" Then
      lc_ArchivoElegido = Dir1.Path & File1.FileName
   Else
      lc_ArchivoElegido = Dir1.Path & "\" & File1.FileName
   End If
   Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
   lc_ArchivoElegido = ""
   Me.Visible = False
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrorDrive1
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    Exit Sub
ErrorDrive1:
    MsgBox "Error (" & Err.Number & ")  " & Err.Description, vbInformation, "Error"
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub File1_Click()
    If lb_MuestraImagen Then
       pi_actual.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
    End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
   lc_ArchivoElegido = ""
   Dir1.Path = lc_PathDefault
   'DEBB2014a
   If lc_TipoArchivo <> "" Then
      File1.Pattern = lc_TipoArchivo
   End If
End Sub



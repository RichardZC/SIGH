VERSION 5.00
Begin VB.UserControl ucPacientesPDF 
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   ScaleHeight     =   1035
   ScaleWidth      =   11610
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8130
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "ucPacientesPDF.ctx":0000
      Top             =   0
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "ucPacientesPDF.ctx":001D
      Top             =   0
      Width           =   1560
   End
   Begin VB.ListBox grdEpicrisis 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   9630
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.ListBox grdPDF 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   1590
      TabIndex        =   0
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "ucPacientesPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ml_IdPaciente As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_meHwnd As Long
Const lnAltoInicial As Long = 1035
Const grdPDFaltoInicial = 1020
Const grdEpicrisisAltoInicial = 1020

Property Let meHwnd(lValue As Long)
    ml_meHwnd = lValue
End Property

Public Sub Inicializar(lnIdPaciente As Long, lnNroHistoriaClinica As Long)
    ml_IdPaciente = lnIdPaciente
    CargaEpicrisisEscaneadas lnNroHistoriaClinica
    CargaPDFgenerados
End Sub

Sub CargaEpicrisisEscaneadas(lnNroHistoriaClinica As Long)
     Dim lcNombreJpg As String, lcNombre As String
     Dim lnFor As Integer, lcRuta As String
     lcRuta = lcBuscaParametro.SeleccionaFilaParametro(237)
     grdEpicrisis.Clear
     For lnFor = 1 To 30
         lcNombre = Trim(Str(lnNroHistoriaClinica)) & "-" & Trim(Str(lnFor)) & ".jpg"
         lcNombreJpg = lcRuta & "\" & lcNombre
         If sighentidades.ArchivoExiste(lcNombreJpg) Then
            grdEpicrisis.AddItem lcNombre
         End If
     Next
End Sub

Sub CargaPDFgenerados()
    Dim sArchivo As String
    grdPDF.Clear
    If ml_IdPaciente > 0 Then
        sArchivo = Dir(lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & Trim(Str(ml_IdPaciente)) & "*.pdf")
        Do While sArchivo <> ""
            If InStr(sArchivo, "-GRABA-") = 0 Then
               grdPDF.AddItem Mid(sArchivo, InStr(sArchivo, "-") + 1)
            End If
            sArchivo = Dir
        Loop
    End If
End Sub

Private Sub grdEpicrisis_DblClick()
     If Len(grdEpicrisis.Text) > 0 Then
        
        FileCopy lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & grdEpicrisis.Text, "c:\dibujo1.jpg"
        Dim oCargaImg As Long
        oCargaImg = Shell("rundll32.exe url.dll,FileProtocolHandler " & "c:\dibujo1.jpg", vbMaximizedFocus)
     End If
End Sub

Private Sub grdPDF_DblClick()
On Error GoTo ErrPDF
     ShellExecute ml_meHwnd, vbNullString, lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & _
                  Trim(Str(ml_IdPaciente)) & "-" & grdPDF.Text, _
                  vbNullString, "C:\", 1
ErrPDF:
End Sub

Private Sub UserControl_Resize()
     If UserControl.Height > lnAltoInicial Then
        grdPDF.Height = grdPDF.Height + (UserControl.Height - lnAltoInicial) - 100
        grdEpicrisis.Height = grdEpicrisis.Height + (UserControl.Height - lnAltoInicial) - 100
     Else
        grdPDF.Height = grdPDFaltoInicial
        grdEpicrisis.Height = grdEpicrisisAltoInicial
     
     End If
End Sub

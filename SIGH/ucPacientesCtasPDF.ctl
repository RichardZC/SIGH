VERSION 5.00
Begin VB.UserControl ucPacientesCtasPDF 
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   ScaleHeight     =   1275
   ScaleWidth      =   3570
   Begin VB.ListBox grdPDF 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3540
   End
End
Attribute VB_Name = "ucPacientesCtasPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ml_IdPaciente As Long, ml_idCuentaAtencion As Long, lcPDFizquierda As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_meHwnd As Long
Const lnAltoInicial As Long = 1525
Const grdPDFaltoInicial = 1520


Property Let meHwnd(lValue As Long)
    ml_meHwnd = lValue
End Property

Public Sub Inicializar(lnIdPaciente As Long, lnIdCuenta As Long)
    ml_IdPaciente = lnIdPaciente
    ml_idCuentaAtencion = lnIdCuenta
    CargaPDFgenerados
End Sub



Sub CargaPDFgenerados()
    Dim sArchivo As String, lcCuenta As String
    grdPDF.Clear
    If ml_IdPaciente > 0 Then
        sArchivo = Dir(lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & Trim(Str(ml_IdPaciente)) & "*.pdf")
        Do While sArchivo <> ""
            lcCuenta = Mid(sArchivo, InStr(sArchivo, "CTA") + 3, 100)
            lcCuenta = Left(lcCuenta, InStr(lcCuenta, "-") - 1)
            If InStr(sArchivo, "-GRABA-") > 0 And Val(lcCuenta) = ml_idCuentaAtencion Then
               lcPDFizquierda = Left(sArchivo, InStr(sArchivo, "-GRABA-") + 6)
               grdPDF.AddItem Mid(sArchivo, InStr(sArchivo, "-GRABA-") + 7)
            End If
            sArchivo = Dir
        Loop
    End If
End Sub


Private Sub grdPDF_DblClick()
On Error GoTo ErrPDF
'     ShellExecute ml_meHwnd, vbNullString, lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & _
'                  Trim(Str(ml_IdPaciente)) & "-" & grdPDF.Text, _
'                  vbNullString, "C:\", 1

     ShellExecute ml_meHwnd, vbNullString, lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & _
                  lcPDFizquierda & grdPDF.Text, _
                  vbNullString, "C:\", 1
ErrPDF:
End Sub

Private Sub UserControl_Resize()
     'If UserControl.Height > lnAltoInicial Then
        grdPDF.Height = UserControl.Height - 15
        
        
     'Else
     '   grdPDF.Height = grdPDFaltoInicial
     'End If
End Sub



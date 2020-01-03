VERSION 5.00
Begin VB.Form DiagnosticosBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de Diagnósticos"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "DiagnosticosBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHNegocios.ucDiagnosticosLista ucDiagnosticosLista1 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Carga Cie10 detallado"
      Top             =   60
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9446
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   3
      Top             =   5490
      Width           =   9780
      Begin VB.CommandButton btnBuscaPDF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Si no existe Dx busque aquí"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   7005
         MaskColor       =   &H80000004&
         Picture         =   "DiagnosticosBusqueda.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   2655
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "DiagnosticosBusqueda.frx":11CE
         DownPicture     =   "DiagnosticosBusqueda.frx":1692
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
         Left            =   4935
         Picture         =   "DiagnosticosBusqueda.frx":1B7E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "DiagnosticosBusqueda.frx":206A
         DownPicture     =   "DiagnosticosBusqueda.frx":24CA
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
         Left            =   3390
         Picture         =   "DiagnosticosBusqueda.frx":293F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "DiagnosticosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Diagnóstico
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long




Property Set DataSource(oValue As ADODB.Recordset)
    Set ucDiagnosticosLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucDiagnosticosLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucDiagnosticosLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucDiagnosticosLista1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnBuscaPDF_Click()
    ShellExecute Me.hwnd, vbNullString, App.Path & "\archivos\cie10.pdf", vbNullString, "C:\", 1
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub


Private Sub Form_Activate()
    Me.ucDiagnosticosLista1.FocusEnDescripcion
    'mgaray20141022
'    Me.ucDiagnosticosLista1.MostrarSoloActivos = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucDiagnosticosLista1.Titulo = "Búsqueda de diagnósticos"
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucDiagnosticosLista1.RealizarBusqueda True
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Private Sub ucDiagnosticosLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
    If lnIdRegistroSeleccionado > 0 Then
       btnAceptar_Click
    End If
End Sub

Property Let CodigoDx(lValue As String)
    
    ucDiagnosticosLista1.CodigoDx = lValue
End Property


Property Let SoloMuestraDxGalenHos(lValue As Boolean)
    ucDiagnosticosLista1.SoloMuestraDxGalenHos = lValue
End Property

Property Let USAcodigoCIEsinPto(lValue As Boolean)
    ucDiagnosticosLista1.USAcodigoCIEsinPto = lValue
End Property


VERSION 5.00
Begin VB.Form rConsumoPromAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consumo Promedio Anual"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rConsumoPromAnual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9195
      Begin VB.CheckBox chkSinTipoConceptos 
         Caption         =   "Sin considerar salidas por: Distribución/Devoluciones/Inventario"
         Height          =   300
         Left            =   225
         TabIndex        =   7
         Top             =   810
         Value           =   1  'Checked
         Width           =   8790
      End
      Begin VB.CheckBox chkExcel 
         Caption         =   "En Excel"
         Height          =   315
         Left            =   7950
         Picture         =   "rConsumoPromAnual.frx":0CCA
         TabIndex        =   6
         Top             =   210
         Width           =   1035
      End
      Begin VB.ComboBox cmbAnios 
         Height          =   330
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label lblNcuenta 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Width           =   330
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      TabIndex        =   3
      Top             =   1590
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rConsumoPromAnual.frx":0FDC
         DownPicture     =   "rConsumoPromAnual.frx":143C
         Height          =   700
         Left            =   3210
         Picture         =   "rConsumoPromAnual.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rConsumoPromAnual.frx":1D26
         DownPicture     =   "rConsumoPromAnual.frx":21EA
         Height          =   700
         Left            =   4740
         Picture         =   "rConsumoPromAnual.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rConsumoPromAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte Consumo promedio Anual
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ms_MensajeError As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim ml_idUsuario As Long
Dim mo_Formulario As New sighentidades.Formulario

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
         Dim oRptClase As New rCrystal
         oRptClase.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
         oRptClase.IdAnio = Val(cmbAnios.Text)
         oRptClase.TextoDelFiltro = ml_TextoDelFiltro
         oRptClase.TipoReporte = Me.Name
         oRptClase.NOconsiderarTipoConceptos = IIf(Me.chkSinTipoConceptos.Value = 1, True, False)
         oRptClase.Show vbModal
         Set oRptClase = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    ml_TextoDelFiltro = "FILTROS:   Año: " & Trim(cmbAnios.Text) & IIf(Me.chkSinTipoConceptos.Value = 1, "   (" & chkSinTipoConceptos.Caption & ")", "")
    If cmbAnios.Text = "" Then
        ms_MensajeError = ms_MensajeError + "Por favor elija el Año" + Chr(13)
        cmbAnios.SetFocus
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
'           ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub








Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, cmbAnios

End Sub

Private Sub cmbAnios_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, cmbAnios
End Sub


Private Sub Form_Load()
       mo_Formulario.LlenaComboConAnios cmbAnios

End Sub
Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

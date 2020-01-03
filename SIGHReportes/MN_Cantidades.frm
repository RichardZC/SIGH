VERSION 5.00
Begin VB.Form MN_Cantidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cantidades de Mortalidad y Nacimientos"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   90
      TabIndex        =   2
      Top             =   1155
      Width           =   8625
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MN_Cantidades.frx":0000
         DownPicture     =   "MN_Cantidades.frx":04C4
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
         Left            =   4440
         Picture         =   "MN_Cantidades.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MN_Cantidades.frx":0E9C
         DownPicture     =   "MN_Cantidades.frx":12FC
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
         Left            =   2910
         Picture         =   "MN_Cantidades.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   8625
      Begin VB.ComboBox cmbIdDepartamento1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4545
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cmbConsiderar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "MN_Cantidades.frx":1BE6
         Left            =   5660
         List            =   "MN_Cantidades.frx":1BED
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   2865
      End
      Begin VB.ComboBox cmbAnio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "MN_Cantidades.frx":1C02
         Left            =   585
         List            =   "MN_Cantidades.frx":1C04
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label lblDpto1 
         Caption         =   "Considerar para Muertes Neonatales el Departamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   660
         Width           =   4380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Alta"
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
         Left            =   4515
         TabIndex        =   5
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Departamento 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   90
         TabIndex        =   1
         Top             =   285
         Width           =   330
      End
   End
End
Attribute VB_Name = "MN_Cantidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Cantidad de Mortalidad y Nacimientos de NIÑOS
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_cmbIdDepartamento1 As New SIGHEntidades.ListaDespleglable
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_TextoDelFiltro As String

Private Sub btnCancelar_Click()
  Me.Visible = False
  Unload Me
End Sub

Sub btnAceptar_Click()
  If ValidaDatosObligatorios Then
    Me.MousePointer = 11
    Dim oRptMN_Cantidades As New RptMN_Cantidades
    oRptMN_Cantidades.Anio = Val(cmbAnio.Text)
    oRptMN_Cantidades.FechaAltaMedica = IIf(cmbConsiderar.ListIndex = 0, True, False)
    oRptMN_Cantidades.TextoDelFiltro = ml_TextoDelFiltro
    oRptMN_Cantidades.idDepartamento1 = IIf(mo_cmbIdDepartamento1.BoundText = "", 0, mo_cmbIdDepartamento1.BoundText)
    oRptMN_Cantidades.CrearReporteDetallado Me.hwnd
    Me.MousePointer = 1
  End If
End Sub

Function ValidaDatosObligatorios() As Boolean
  Dim sMensaje As String
  sMensaje = "": ml_TextoDelFiltro = ""
  If cmbAnio.Text = "" Then sMensaje = "- Falta Año."
  If cmbConsiderar.Text = "" Then sMensaje = sMensaje & Chr(13) & "- Falta Considerar."
  If cmbIdDepartamento1.Text = "" Then sMensaje = sMensaje & Chr(13) & "- Falta Departamento Hospitalario."
  If sMensaje <> "" Then
    MsgBox sMensaje, vbInformation, Me.Caption
    ValidaDatosObligatorios = False
  Else
    ValidaDatosObligatorios = True
    ml_TextoDelFiltro = "Año: " & cmbAnio.Text & ". Paciente con: " & cmbConsiderar.Text & ". Departamento: " & cmbIdDepartamento1.Text
  End If
End Function

Private Sub Form_Initialize()
  Set mo_cmbIdDepartamento1.MiComboBox = cmbIdDepartamento1
End Sub

Private Sub Form_Load()
  mo_Formulario.LlenaComboConAnios cmbAnio
  cmbConsiderar.ListIndex = 0
  mo_cmbIdDepartamento1.BoundColumn = "IdDepartamento"
  mo_cmbIdDepartamento1.ListField = "DescripcionLarga"
  Set mo_cmbIdDepartamento1.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
  cmbIdDepartamento1.ListIndex = 1
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      btnCancelar_Click
    Case vbKeyF2
      btnAceptar_Click
  End Select
End Sub



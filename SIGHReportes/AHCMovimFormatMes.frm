VERSION 5.00
Begin VB.Form AHCMovimFormatMes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimiento de Formatos de  Historias Mensual"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "AHCMovimFormatMes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   3
      Top             =   1890
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCMovimFormatMes.frx":0CCA
         DownPicture     =   "AHCMovimFormatMes.frx":112A
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
         Left            =   1320
         Picture         =   "AHCMovimFormatMes.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCMovimFormatMes.frx":1A14
         DownPicture     =   "AHCMovimFormatMes.frx":1ED8
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
         Left            =   2850
         Picture         =   "AHCMovimFormatMes.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5370
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
         ItemData        =   "AHCMovimFormatMes.frx":28B0
         Left            =   1680
         List            =   "AHCMovimFormatMes.frx":28B2
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   990
         Width           =   1215
      End
      Begin VB.ComboBox cmbMes 
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
         ItemData        =   "AHCMovimFormatMes.frx":28B4
         Left            =   1680
         List            =   "AHCMovimFormatMes.frx":28B6
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Left            =   3870
         Picture         =   "AHCMovimFormatMes.frx":28B8
         TabIndex        =   6
         Top             =   960
         Width           =   1035
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
         ItemData        =   "AHCMovimFormatMes.frx":2BCA
         Left            =   1680
         List            =   "AHCMovimFormatMes.frx":2BD1
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   105
         TabIndex        =   10
         Top             =   645
         Width           =   720
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
         Left            =   105
         TabIndex        =   9
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Especialidad"
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
         Left            =   105
         TabIndex        =   2
         Top             =   285
         Width           =   1380
      End
   End
End
Attribute VB_Name = "AHCMovimFormatMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Momiento de Formato de Historias Mensual
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Dim mo_Formulario As New SIGHEntidades.Formulario
Private Sub btnAceptar_Click()
        Me.MousePointer = 11
        Dim oRpt As New RptAHCMovimFormatMes
        Dim lcFec1 As String, lcFec2 As String
        Dim lcMES As String, lcANIO As String
        Dim lnUltimoDiaMes As Integer
        Dim ml_TextoDelFiltro As String
        ml_TextoDelFiltro = Trim(cmbConsiderar.Text) & "     Año: " & Trim(cmbAnio.Text) & "     Mes: " & Trim(cmbMes.Text)
        '
        lcMES = Right("0" & Trim(Str(cmbMes.ListIndex + 1)), 2)
        lcANIO = cmbAnio.Text
        lcFec1 = ("01/" & lcMES & "/" & lcANIO)
        lnUltimoDiaMes = DevuelveUltimoDiaDelMes(Val(lcMES), Val(cmbAnio.Text))
        lcFec2 = Right("0" & Trim(Str(lnUltimoDiaMes)), 2) & "/" & lcMES & "/" & lcANIO
        '
        oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), "Movimiento de Formatos de Historia Clinica Mensual", ml_TextoDelFiltro, cmbConsiderar.ListIndex, CDate(Format(lcFec1 & " 00:00:01", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(lcFec2 & " 23:59:59", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), Me.hwnd
        Set oRpt = Nothing
        Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Private Sub Form_Load()
       cmbConsiderar.ListIndex = 0
       mo_Formulario.LlenaComboConAnios Me.cmbAnio
       mo_Formulario.LlenaComboConMeses Me.cmbMes

End Sub




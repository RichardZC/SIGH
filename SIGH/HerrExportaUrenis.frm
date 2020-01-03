VERSION 5.00
Begin VB.Form HerrExportaUrenis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporta datos al Sistema URENIS"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   5
      Top             =   2520
      Width           =   7845
      Begin VB.CheckBox chkSoloConDx 
         Caption         =   "Solo procesa los pacientes atendidos (los que tienen al menos un Diagnósticos"
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
         Left            =   210
         TabIndex        =   11
         Top             =   1830
         Value           =   1  'Checked
         Width           =   7395
      End
      Begin VB.CheckBox chkExportaCPT 
         Caption         =   "Agrega a los CIE los procedimientos (CPT) realizados en el mismo servicio"
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
         Left            =   210
         TabIndex        =   10
         Top             =   1470
         Value           =   1  'Checked
         Width           =   6345
      End
      Begin VB.TextBox txtTarde 
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
         Left            =   1320
         TabIndex        =   8
         Text            =   "13:01"
         Top             =   1050
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3255
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   6956042
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarde"
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
         Left            =   180
         TabIndex        =   9
         Top             =   1110
         Width           =   480
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
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   330
      End
      Begin VB.Label Label4 
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
         Left            =   180
         TabIndex        =   6
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   7860
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   2160
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   7665
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   60
      TabIndex        =   3
      Top             =   5160
      Width           =   7845
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrExportaUrenis.frx":0000
         DownPicture     =   "HerrExportaUrenis.frx":04C4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4080
         Picture         =   "HerrExportaUrenis.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdExportaURENIS 
         Caption         =   "Exporta al URENIS"
         DisabledPicture =   "HerrExportaUrenis.frx":0E9C
         DownPicture     =   "HerrExportaUrenis.frx":12FC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2520
         Picture         =   "HerrExportaUrenis.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "HerrExportaUrenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: MINSA - Proyecto SIGES
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exporta información al SIstema del MINSA URENIS
'        Programado por: Franklin Cachay
'        Fecha: Enero 2015
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_idUsuario As Long
Dim mo_lcNombrePc  As String
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim oConexion As New Connection
Dim oConexionFox As New Connection
Dim lcSql As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property


Private Sub btnAceptar_Click()
       Dim oProcesos As New Procesos
       oProcesos.ExportaDatosAlHis
       Set oProcesos = Nothing
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub cmdExportaURENIS_Click()
    If cmbAnio.Text = "" Then
       MsgBox "Por favor elija el AÑO", vbCritical, "Mensaje"
       Exit Sub
    End If
    If cmbMes.Text = "" Then
       MsgBox "Por favor elija el MES", vbCritical, "Mensaje"
       Exit Sub
    End If
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       Dim oProcesos As New Procesos
       Set oProcesos.progressRpt1 = Me.progressRpt
       oProcesos.idUsuario = ml_idUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       oProcesos.ExportaDAtosAlURENIS_V1 txtTarde.Text, Me.chkExportaCPT.Value, (cmbMes.ListIndex + 1), cmbAnio.Text
'       '
'       Dim oProcesos2 As New Procesos
'       Set oProcesos2.progressRpt8 = Me.progressRpt1
'       oProcesos2.idUsuario = ml_idUsuario
'       oProcesos2.lcNombrePc = mo_lcNombrePc
'       oProcesos2.ExportaDAtosAlHISv4_2 txtTarde.Text, (cmbMes.ListIndex + 1), cmbAnio.Text
'
''       ExportaDAtosAlHISv4_2 yamill
       '
       Me.MousePointer = 1
       If oProcesos.MensajeError = "" Then
          Me.Visible = False
'       Else
'            If oProcesos2.MensajeError = "" Then
'                Me.Visible = False
'            End If
       End If
       Set oProcesos = Nothing
'       Set oProcesos2 = Nothing
    End If
End Sub

Private Sub Form_Load()
  mo_reglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrExportaUrenis"
  '
  mo_Formulario.LlenaComboConAnios cmbAnio
  mo_Formulario.LlenaComboConMeses cmbMes
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub





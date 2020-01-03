VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CEpadronNominal 
   Caption         =   "Padrón Nominal"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   Icon            =   "CEpadronNominal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   6195
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En excel"
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
         Left            =   5085
         TabIndex        =   16
         Top             =   1215
         Width           =   1005
      End
      Begin VB.TextBox TxtEdad 
         Height          =   285
         Left            =   1350
         TabIndex        =   15
         Text            =   "5"
         Top             =   990
         Width           =   405
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Caption         =   "..."
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   675
         Width           =   315
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         TabIndex        =   13
         Top             =   675
         Width           =   3255
      End
      Begin VB.TextBox txtNhistoria 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   12
         ToolTipText     =   "Ingrese el Nro Historia Clínica"
         Top             =   675
         Width           =   1125
      End
      Begin VB.ComboBox cmbSexo 
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
         ItemData        =   "CEpadronNominal.frx":000C
         Left            =   1350
         List            =   "CEpadronNominal.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1260
         Width           =   1515
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   315
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   4695
         TabIndex        =   5
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "años"
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
         Left            =   1800
         TabIndex        =   17
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
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
         TabIndex        =   10
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   4215
         TabIndex        =   9
         Top             =   375
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         TabIndex        =   8
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Paciente"
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
         Top             =   670
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Menor o igual"
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
         TabIndex        =   6
         Top             =   1005
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   45
      TabIndex        =   0
      Top             =   1770
      Width           =   6180
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CEpadronNominal.frx":0039
         DownPicture     =   "CEpadronNominal.frx":04FD
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
         Left            =   3225
         Picture         =   "CEpadronNominal.frx":09E9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CEpadronNominal.frx":0ED5
         DownPicture     =   "CEpadronNominal.frx":1335
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
         Left            =   1695
         Picture         =   "CEpadronNominal.frx":17AA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CEpadronNominal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Imprime Padrón Nominal
'        Programado por: Barrantes D
'        Fecha: Marzo 2015
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idPaciente As Long

Private Sub btnAceptar_Click()
    Dim oRptCEpadronNominal As New RptCEpadronNominal
    Dim lcSubTitulo As String
    lcSubTitulo = "Desde: " & Me.txtFechaInicio.Text & " hasta " & Me.txtFechaFin.Text & _
                   "  (Menor o igual a: " & Me.TxtEdad.Text & " años)" & _
                    IIf(Me.cmbSexo.ListIndex > 0, " (Sexo: " & Me.cmbSexo.Text & ")", "") & _
                    IIf(Me.txtNombrePaciente.Text <> "", " (Paciente: " & Me.txtNombrePaciente.Text & ")", "")
    oRptCEpadronNominal.CrearReporte CDate(Me.txtFechaInicio.Text), CDate(Me.txtFechaFin.Text), _
                                     cmbSexo.ListIndex, Val(Me.TxtEdad.Text), ml_idPaciente, _
                                     IIf(Me.chkExcel.Value = 1, True, False), lcSubTitulo, Me.hwnd
    
End Sub

Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    ml_idPaciente = 0
    txtNhistoria.Text = ""
    txtNombrePaciente.Text = ""
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_idPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
        End If
    End If
    Set oBusqueda = Nothing
    Set oDOPaciente = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
       Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
       cmbSexo.ListIndex = 0
End Sub
Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria
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

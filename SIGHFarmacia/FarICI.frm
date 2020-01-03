VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FarICI 
   Caption         =   "Imprime Formato ICI"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   2
      Top             =   1620
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarICI.frx":0000
         DownPicture     =   "FarICI.frx":0460
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
         Picture         =   "FarICI.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarICI.frx":0D4A
         DownPicture     =   "FarICI.frx":120E
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
         Picture         =   "FarICI.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   5370
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   135
         ScaleHeight     =   240
         ScaleWidth      =   5010
         TabIndex        =   1
         Top             =   2280
         Width           =   5070
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
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
      Begin MSDataListLib.DataCombo cmbFarmacia 
         Height          =   360
         Left            =   1260
         TabIndex        =   7
         Top             =   210
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   714
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   3000
         TabIndex        =   9
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.ADespacho"
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
         Left            =   150
         TabIndex        =   10
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Farmacia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   2790
         TabIndex        =   6
         Top             =   690
         Width           =   120
      End
   End
End
Attribute VB_Name = "FarICI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************daniel barrantes**************
'***************Registro de datos de filtro para el Reporte
'***************Historias Clinicas solicitadas por MEDICO
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim sMensaje As String
Dim mo_Teclado As New sighcomun.Teclado
Dim oRsFarmacias As New ADODB.Recordset



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptHistorias As New RptFarICI
        oRptHistorias.IdResponsable = Val(cmbFarmacia.BoundText)
        oRptHistorias.FechaInicio = txtFechaInicio.Text
        oRptHistorias.FechaFin = txtFechaFin.Text
        oRptHistorias.TextoDelFiltro = "Farmacia: " & Trim(cmbFarmacia.Text) & "       F.Despacho:" & txtFechaInicio.Text & " al " & txtFechaFin.Text
        oRptHistorias.CrearReporte_excel
        Me.MousePointer = 1
    End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If cmbFarmacia.Text = "" Then
        sMensaje = sMensaje + "Por favor elija la Farmacia"
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub






Private Sub Form_Load()
      CargaCombos
       
      Me.txtFechaInicio.Text = sighcomun.PrimerFechaDDMMYYDelMesActual()
      Me.txtFechaFin.Text = Format(Date, "dd/mm/yyyy")
       
End Sub

Sub CargaCombos()
         Dim lcSql As String
         oRsFarmacias.Open "select * from permisos where modulo='Farmacia' order by idPermiso", sighcomun.CadenaConexion, adOpenKeyset, adLockOptimistic
         Set cmbFarmacia.RowSource = oRsFarmacias
         cmbFarmacia.ListField = "descripcion"
         cmbFarmacia.BoundColumn = "idPermiso"
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



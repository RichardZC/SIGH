VERSION 5.00
Begin VB.Form HerrExportaHIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporta datos al Sistema HIS"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "HerrExportaHIS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
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
      Height          =   2940
      Left            =   45
      TabIndex        =   7
      Top             =   4395
      Width           =   7845
      Begin VB.CheckBox chkExportaHISminsa 
         Alignment       =   1  'Right Justify
         Caption         =   "Exporta al HIS MINSA"
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
         Left            =   5520
         TabIndex        =   19
         Top             =   255
         Width           =   2175
      End
      Begin VB.CheckBox chkAgregaOrdenesMedicas 
         Caption         =   "Agrega a las Ordenes Médicas (Laboratorio/Imágenes)"
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
         Left            =   180
         TabIndex        =   18
         Top             =   1425
         Width           =   6345
      End
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
         TabIndex        =   14
         Top             =   1830
         Value           =   1  'Checked
         Width           =   6915
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
         Left            =   180
         TabIndex        =   13
         Top             =   1065
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
         Left            =   3345
         TabIndex        =   10
         Text            =   "13:01"
         Top             =   630
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
         ItemData        =   "HerrExportaHIS.frx":0CCA
         Left            =   840
         List            =   "HerrExportaHIS.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3735
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
         ItemData        =   "HerrExportaHIS.frx":0CCE
         Left            =   840
         List            =   "HerrExportaHIS.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   630
         Width           =   1215
      End
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   180
         TabIndex        =   15
         Top             =   2190
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
      Begin SISGalenPlus.XP_ProgressBar progressRpt1 
         Height          =   300
         Left            =   180
         TabIndex        =   16
         Top             =   2520
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
         Left            =   2820
         TabIndex        =   11
         Top             =   690
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
         TabIndex        =   9
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
         TabIndex        =   8
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
      Height          =   4395
      Left            =   45
      TabIndex        =   6
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
         Height          =   4050
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   7665
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   45
      TabIndex        =   5
      Top             =   7305
      Width           =   7845
      Begin VB.CommandButton btnPruebatrama 
         Caption         =   "Prueba Urenis"
         Height          =   615
         Left            =   6720
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdExportaHISV4 
         Caption         =   "Exporta al His (v.4)"
         DisabledPicture =   "HerrExportaHIS.frx":0CD2
         DownPicture     =   "HerrExportaHIS.frx":1132
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
         Left            =   3292
         Picture         =   "HerrExportaHIS.frx":15A7
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Exporta al NOVAFIS"
         DisabledPicture =   "HerrExportaHIS.frx":1A1C
         DownPicture     =   "HerrExportaHIS.frx":1E7C
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
         Left            =   1815
         Picture         =   "HerrExportaHIS.frx":22F1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrExportaHIS.frx":2766
         DownPicture     =   "HerrExportaHIS.frx":2C2A
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
         Left            =   4770
         Picture         =   "HerrExportaHIS.frx":3116
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "HerrExportaHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exporta información al SIstema del MINSA HIS
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
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
       Dim oProcesos As New SIGHProxies.Procesos
       'Set oProcesos.progressRpt1 = Me.progressRpt
       'Set oProcesos.progressRpt2 = progressRpt1
       oProcesos.ValorActual = Me.txtTarde.Text
       oProcesos.idUsuario = ml_idUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       oProcesos.ExportaDAtosAlHISv4 txtTarde.Text, Me.chkExportaCPT.Value, (cmbMes.ListIndex + 1), cmbAnio.Text, True, False, _
                                     False
       '
       Me.MousePointer = 1
       If oProcesos.MensajeError = "" Then
          Me.Visible = False
       End If
       Set oProcesos = Nothing
    End If

End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub btnPruebatrama_Click()
'       Dim oHerrHIS As New HerrExportaUrenis
'       oHerrHIS.idUsuario = ml_idUsuario
'       oHerrHIS.lcNombrePc = mo_lcNombrePc
'       oHerrHIS.Show 1
'       Set oHerrHIS = Nothing
'       Exit Sub
           On Error GoTo ErrPru
           Dim oRsTmpFox1 As New Recordset
           Dim oRsTmp1 As New Recordset
           Dim oConexion As New Connection
           Dim oConexionFox As New Connection
           Dim lcSql As String
           
           oConexion.CommandTimeout = 300
           oConexion.CursorLocation = adUseClient
           oConexion.Open sighentidades.CadenaConexion
       
           oConexionFox.CommandTimeout = 300
           oConexionFox.Open "DSN=his"
           
           lcSql = "select * from clinica"
           oRsTmpFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
           oRsTmpFox1.MoveFirst
           Do While Not oRsTmpFox1.EOF
              lcSql = "select * from pacientes where nroHistoriaClinica=" & Trim(Str(Val(oRsTmpFox1!nclinica)))
              oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
              If oRsTmp1.RecordCount > 0 Then
                 If UCase(Left(oRsTmpFox1!Sexo, 1)) = "M" Then
                    oRsTmp1.Fields!idTipoSexo = 1
                 Else
                    oRsTmp1.Fields!idTipoSexo = 2
                 End If
                 oRsTmp1.Update
              End If
              oRsTmp1.Close
              oRsTmpFox1.MoveNext
           Loop
           oRsTmpFox1.Close
           oConexionFox.Close
           oConexion.Close
           Unload Me
           Exit Sub
ErrPru:
     MsgBox Err.Description
     Exit Sub
     Resume
     
End Sub

Private Sub cmdExportaHISV4_Click()
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
       Dim oProcesos As New SIGHProxies.Procesos
       'Set oProcesos.progressRpt1 = Me.progressRpt
       oProcesos.idUsuario = ml_idUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       oProcesos.ExportaDAtosAlHISv4 txtTarde.Text, Me.chkExportaCPT.Value, (cmbMes.ListIndex + 1), cmbAnio.Text, _
                                     False, IIf(Me.chkAgregaOrdenesMedicas.Value = 1, True, False), _
                                     IIf(Me.chkExportaHISminsa.Value = 1, True, False)
       '
       Dim oProcesos2 As New Procesos
       Set oProcesos2.progressRpt8 = Me.progressRpt1
       oProcesos2.idUsuario = ml_idUsuario
       oProcesos2.lcNombrePc = mo_lcNombrePc
       oProcesos2.ExportaDAtosAlHISv4_2 txtTarde.Text, (cmbMes.ListIndex + 1), cmbAnio.Text
       
'       ExportaDAtosAlHISv4_2 yamill
       '
       Me.MousePointer = 1
       If oProcesos.MensajeError = "" Then
          Me.Visible = False
       Else
            If oProcesos2.MensajeError = "" Then
                Me.Visible = False
            End If
       End If
       Set oProcesos = Nothing
       Set oProcesos2 = Nothing
    End If
End Sub

Private Sub Form_Load()
  mo_reglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrExportaHIS"
  LlenaConsideracionesParaExportarNOVAHIS
  LlenaConsideracionesHIS_MINSA
  
  '
  mo_Formulario.LlenaComboConAnios cmbAnio
  mo_Formulario.LlenaComboConMeses cmbMes
  
End Sub

Sub LlenaConsideracionesHIS_MINSA()
    cmbConsideraciones.AddItem "Exportar datos hacia HIS_MINSA                                               "
    cmbConsideraciones.AddItem "------------------------------                                               "
    cmbConsideraciones.AddItem "- Los MEDICOS q atienden en CONSULTORIOS deben registrarle F.NACIMIENTO y SEXO"
    cmbConsideraciones.AddItem "- En las ATENCIONES DEL MEDICO debe elejir el LAB y no ingresar cualquier valor"
    cmbConsideraciones.AddItem "- La persona que PROCESA HIS se le debe registrar F.NACIMIENTO y SEXO"

End Sub
Sub LlenaConsideracionesParaExportarNOVAHIS()
    cmbConsideraciones.AddItem "Exportar datos hacia NOVAFIS                                                 "
    cmbConsideraciones.AddItem "----------------------------                                                 "
    cmbConsideraciones.AddItem "- Parametro 270 debe contener el DNI del DIGITADOR                           "
    cmbConsideraciones.AddItem "- Configurar Consultorios (ups), Médicos (dni,lote)                          "
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





VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DiagnosticoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "DiagnosticoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   30
      Top             =   4290
      Width           =   10035
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "DiagnosticoDetalle.frx":0CCA
         DownPicture     =   "DiagnosticoDetalle.frx":118E
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
         Left            =   4980
         Picture         =   "DiagnosticoDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "DiagnosticoDetalle.frx":1B66
         DownPicture     =   "DiagnosticoDetalle.frx":1FC6
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
         Left            =   3450
         Picture         =   "DiagnosticoDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Restriciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   60
      TabIndex        =   26
      Top             =   3375
      Width           =   10050
      Begin VB.ComboBox cmbIdTipoSexo 
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
         Left            =   1095
         TabIndex        =   13
         Top             =   195
         Width           =   1725
      End
      Begin VB.CheckBox chkGestacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Gestación"
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
         Left            =   8460
         TabIndex        =   18
         Top             =   540
         Width           =   1410
      End
      Begin VB.CheckBox chkMorbilidad 
         Caption         =   "Morbilidad"
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
         Left            =   1095
         TabIndex        =   16
         Top             =   570
         Width           =   1365
      End
      Begin VB.CheckBox chkIntrahospitalario 
         Alignment       =   1  'Right Justify
         Caption         =   "Intrahospitalario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4650
         TabIndex        =   17
         Top             =   570
         Width           =   1710
      End
      Begin VB.TextBox txtEdadMaxDias 
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
         Left            =   8985
         MaxLength       =   10
         TabIndex        =   15
         Top             =   195
         Width           =   900
      End
      Begin VB.TextBox txtEdadMinDias 
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
         Left            =   5535
         MaxLength       =   10
         TabIndex        =   14
         Top             =   195
         Width           =   840
      End
      Begin VB.Label lblIdTipoSexo 
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
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblEdadMaxDias 
         Caption         =   "Edad máxima"
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
         Left            =   7830
         TabIndex        =   28
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblEdadMinDias 
         Caption         =   "Edad mínima"
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
         Left            =   4425
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3360
      Left            =   60
      TabIndex        =   20
      Top             =   15
      Width           =   10050
      Begin VB.CheckBox chkEsActivo 
         Caption         =   "Habilitado"
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
         Left            =   4320
         TabIndex        =   10
         Top             =   2640
         Width           =   1905
      End
      Begin VB.TextBox txtDescripcionMINSA 
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
         Left            =   1125
         MaxLength       =   250
         TabIndex        =   8
         Top             =   2250
         Width           =   8805
      End
      Begin VB.ComboBox cmbIdCategoria 
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
         Left            =   1125
         TabIndex        =   2
         Top             =   1068
         Width           =   8790
      End
      Begin VB.ComboBox cmbIdGrupo 
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
         Left            =   1125
         TabIndex        =   1
         Top             =   669
         Width           =   8790
      End
      Begin VB.ComboBox cmbIdCapitulo 
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
         Left            =   1125
         TabIndex        =   0
         Top             =   270
         Width           =   8760
      End
      Begin VB.TextBox txtCodigoCIE9 
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
         Left            =   3090
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1470
         Width           =   1000
      End
      Begin VB.TextBox txtCodigoCIE10 
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
         Left            =   8880
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1470
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox txtCodigoExportacion 
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
         Left            =   5475
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1470
         Width           =   1000
      End
      Begin VB.TextBox txtCodigoCIE2004 
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
         Left            =   1125
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1467
         Width           =   1000
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1125
         MaxLength       =   250
         TabIndex        =   7
         Top             =   1851
         Width           =   8805
      End
      Begin VB.CheckBox chkRestriccion 
         Caption         =   "Tiene restricción"
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
         Left            =   1110
         TabIndex        =   12
         Top             =   3045
         Width           =   1905
      End
      Begin MSMask.MaskEdBox mskFechaInicioVigencia 
         Height          =   330
         Left            =   1125
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin VB.Label Label6 
         Caption         =   "Fecha Vigencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   150
         TabIndex        =   37
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
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
         Left            =   8820
         TabIndex        =   36
         Top             =   2610
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Descripción   (MINSA)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   150
         TabIndex        =   35
         Top             =   2250
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción  (Corta)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   150
         TabIndex        =   34
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Categoría"
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
         Left            =   150
         TabIndex        =   33
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
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
         Left            =   150
         TabIndex        =   32
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Capítulo"
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
         Left            =   150
         TabIndex        =   31
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lblCodigoCIE9 
         Caption         =   "CIE-9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   25
         Top             =   1500
         Width           =   555
      End
      Begin VB.Label lblCodigoCIE10 
         Caption         =   "Código CIE10"
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
         Left            =   7710
         TabIndex        =   24
         Top             =   1515
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblCodigoExportacion 
         Caption         =   "Exportación"
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
         Left            =   4470
         TabIndex        =   23
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblCodigoCIE2004 
         Caption         =   "CIE-10"
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
         Left            =   180
         TabIndex        =   22
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción"
         Height          =   270
         Left            =   2235
         TabIndex        =   21
         Top             =   1860
         Width           =   960
      End
   End
End
Attribute VB_Name = "DiagnosticoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Diagnósticos
'        Programado por: Castro W
'        Fecha: Agosto 2004
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Diagnosticos As New DODiagnostico
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdDiagnostico As Long
Dim mo_AdminServiciosComunes As New ReglasComunes
Dim mo_CmbIdTipoSexo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdCapitulo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdGrupo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdCategoria As New SIGHEntidades.ListaDespleglable
Dim mo_AdminComun As New ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim md_FechaServidor As Date

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

    mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
    mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
    Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
    sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
    
    mo_cmbIdCapitulo.BoundColumn = "IdCapitulo"
    mo_cmbIdCapitulo.ListField = "DescripcionLarga"
    Set mo_cmbIdCapitulo.RowSource = mo_AdminServiciosComunes.DiagnosticosCapitulosSeleccionarTodos()
    sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
    
    If sMensaje <> "" Then
        MsgBox mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
    End If
    md_FechaServidor = lcBuscaParametro.RetornaFechaServidorSQL
End Sub
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdDiagnostico(lValue As Long)
   ml_IdDiagnostico = lValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property


Private Sub chkEsActivo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkEsActivo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdCapitulo_Click()
    
    mo_cmbIdGrupo.BoundText = ""
    mo_cmbIdCategoria.BoundText = ""
    
    mo_cmbIdGrupo.BoundColumn = "IdGrupo"
    mo_cmbIdGrupo.ListField = "DescripcionLarga"
    Set mo_cmbIdGrupo.RowSource = mo_AdminServiciosComunes.DiagnosticosGruposSeleccionarPorCapitulo(Val(mo_cmbIdCapitulo.BoundText))

End Sub

Private Sub cmbIdCapitulo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCapitulo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdCapitulo_LostFocus()
    If cmbIdCapitulo.Text <> "" Then
        mo_cmbIdCapitulo.BoundText = Val(Split(cmbIdCapitulo.Text, " = ")(0))
    End If
    mo_Formulario.MarcarComoVacio cmbIdCapitulo
End Sub

Private Sub cmbIdCapitulo_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdGrupo_Click()
    
    mo_cmbIdCategoria.BoundText = ""
    
    mo_cmbIdCategoria.BoundColumn = "IdCategoria"
    mo_cmbIdCategoria.ListField = "DescripcionLarga"
    Set mo_cmbIdCategoria.RowSource = mo_AdminServiciosComunes.DiagnosticosCategoriaSeleccionarPorGrupo(Val(mo_cmbIdGrupo.BoundText))

End Sub

Private Sub cmbIdGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdGrupo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdGrupo_LostFocus()
    If cmbIdGrupo.Text <> "" Then
        mo_cmbIdGrupo.BoundText = Val(Split(cmbIdGrupo.Text, " = ")(0))
    End If
    mo_Formulario.MarcarComoVacio cmbIdGrupo
End Sub

Private Sub cmbIdGrupo_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCategoria
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdCategoria_LostFocus()
    If cmbIdCategoria.Text <> "" Then
        mo_cmbIdCategoria.BoundText = Val(Split(cmbIdCategoria.Text, " = ")(0))
    End If
    mo_Formulario.MarcarComoVacio cmbIdCategoria
End Sub

Private Sub cmbIdCategoria_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub Form_Initialize()
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
    Set mo_cmbIdCapitulo.MiComboBox = cmbIdCapitulo
    Set mo_cmbIdGrupo.MiComboBox = cmbIdGrupo
    Set mo_cmbIdCategoria.MiComboBox = cmbIdCategoria
End Sub

Private Sub mskFechaInicioVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, mskFechaInicioVigencia
    AdministrarKeyPreview KeyCode
End Sub

Private Sub mskFechaInicioVigencia_LostFocus()
    If mskFechaInicioVigencia.Text <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not esfecha(mskFechaInicioVigencia.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de Diagnosticos"
             mskFechaInicioVigencia.Text = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtCodigoCIE2004_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoCIE2004
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigoCIE2004_LostFocus()
    txtCodigoCIE2004 = UCase(txtCodigoCIE2004)
   mo_Formulario.MarcarComoVacio txtCodigoCIE2004
End Sub

Private Sub txtCodigoCIE2004_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub txtDescripcionMINSA_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcionMINSA
AdministrarKeyPreview KeyCode

End Sub

Private Sub txtDescripcionMINSA_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txtEdadMinDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadMinDias
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtEdadMinDias_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtEdadMaxDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadMaxDias
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtEdadMaxDias_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub chkRestriccion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkRestriccion
AdministrarKeyPreview KeyCode
End Sub

Private Sub chkRestriccion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub chkIntrahospitalario_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkIntrahospitalario
AdministrarKeyPreview KeyCode
End Sub

Private Sub chkIntrahospitalario_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub chkMorbilidad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkMorbilidad
AdministrarKeyPreview KeyCode
End Sub

Private Sub chkMorbilidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub chkGestacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkGestacion
AdministrarKeyPreview KeyCode
End Sub

Private Sub chkGestacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCodigoExportacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoExportacion
AdministrarKeyPreview KeyCode
End Sub
Private Sub txtCodigoExportacion_LostFocus()
    txtCodigoExportacion = UCase(txtCodigoExportacion)
   mo_Formulario.MarcarComoVacio txtCodigoExportacion
End Sub

Private Sub txtCodigoExportacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCodigoCIE10_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoCIE10
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigoCIE10_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCodigoCIE9_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoCIE9
AdministrarKeyPreview KeyCode
End Sub
Private Sub txtCodigoCIE9_LostFocus()
    txtCodigoCIE9 = UCase(txtCodigoCIE9)
   mo_Formulario.MarcarComoVacio txtCodigoCIE2004
End Sub

Private Sub txtCodigoCIE9_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
AdministrarKeyPreview KeyCode
End Sub




Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoSexo_LostFocus()
    If cmbIdTipoSexo.Text <> "" Then
       mo_CmbIdTipoSexo.BoundText = Val(Split(cmbIdTipoSexo.Text, " = ")(0))
    End If
End Sub

Private Sub cmbIdTipoSexo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Diagnosticos"
       Case sghModificar
           Me.Caption = "Modificar Diagnosticos"
       Case sghConsultar
           Me.Caption = "Consultar Diagnosticos"
       Case sghEliminar
           Me.Caption = "Eliminar Diagnosticos"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   lblId.Caption = mo_Diagnosticos.IdDiagnostico
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   Me.cmbIdCapitulo.SetFocus
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   If Me.txtCodigoCIE2004.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código CIE10" + Chr(13)
   End If
   If Me.txtDescripcion = "" Then
       sMensaje = sMensaje + "Ingrese la descripción" + Chr(13)
   End If
   
   If mskFechaInicioVigencia.Text = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = sMensaje + "Ingrese la Fecha de Inicio de Vigencia" + Chr(13)
    End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim sMensaje As String
   sMensaje = ""
   'Valida codigos Repetidos
   Dim oRsBuscaCodigo As New Recordset
   Set oRsBuscaCodigo = mo_AdminComun.DiagnosticosSeleccionarXCodigo(Trim(txtCodigoCIE2004.Text))
   Select Case mi_Opcion
   Case sghAgregar
        If validarDuplicadoDiagnostico(oRsBuscaCodigo, 0) = False Then
            sMensaje = sMensaje + "Ese código y Descripción Corta ya estan Registrado para: " _
                    + oRsBuscaCodigo.Fields!Descripcion + Chr(13)
        End If
        


        If CDate(mskFechaInicioVigencia.Text) > CDate(md_FechaServidor) Then
            sMensaje = sMensaje + "Fecha de Vigencia no puede ser mayor que la fecha actual" + Chr(13)
        End If
   Case sghModificar
        If validarDuplicadoDiagnostico(oRsBuscaCodigo, lblId.Caption) = False Then
            sMensaje = sMensaje + "Ese código y Descripción Corta ya esta Registrado para: " _
                    + oRsBuscaCodigo.Fields!Descripcion + Chr(13)
        End If
        If CDate(mskFechaInicioVigencia.Text) > CDate(md_FechaServidor) Then
            sMensaje = sMensaje + "Fecha de Vigencia no puede ser mayor que la fecha actual" + Chr(13)
        End If
   Case sghEliminar
       Dim oDiagnostico As New DODiagnostico
       
       oDiagnostico.IdDiagnostico = Me.IdDiagnostico
       If mo_AdminComun.diagnosticoValidarEliminar(oDiagnostico) = False Then
            sMensaje = mo_AdminComun.MensajeError
       End If
   End Select
   
   Set oRsBuscaCodigo = Nothing
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Diagnosticos
           .IdDiagnostico = Me.IdDiagnostico
            .IdCapitulo = Val(mo_cmbIdCapitulo.BoundText)
            .IdGrupo = Val(mo_cmbIdGrupo.BoundText)
            .IdCategoria = Val(mo_cmbIdCategoria.BoundText)
           .CodigoCIE2004 = Me.txtCodigoCIE2004.Text
           .EdadMinDias = Val(Me.txtEdadMinDias.Text)
           .EdadMaxDias = Val(Me.txtEdadMaxDias.Text)
           .Restriccion = Me.chkRestriccion.Value
           .Intrahospitalario = Me.chkIntrahospitalario.Value
           .Morbilidad = Me.chkMorbilidad.Value
           .Gestacion = Me.chkGestacion.Value
           .CodigoExportacion = Me.txtCodigoExportacion.Text
           .CodigoCIE10 = Me.txtCodigoCIE10.Text
           .CodigoCIE9 = Me.txtCodigoCIE9.Text
           .Descripcion = Me.txtDescripcion.Text
           .idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
           .IdUsuarioAuditoria = Me.idUsuario
           .DescripcionMINSA = Me.txtDescripcionMINSA.Text
           .codigoCIEsinPto = SIGHEntidades.DevuelveCodigoDxSinPUNTO(Me.txtCodigoCIE2004.Text)
           .EsActivo = Me.chkEsActivo.Value
           If mskFechaInicioVigencia.Text = SIGHEntidades.FECHA_VACIA_DMY Then
                .FechaInicioVigencia = 0
           Else
                .FechaInicioVigencia = mskFechaInicioVigencia.Text
           End If
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminServiciosComunes.DiagnosticosAgregar(mo_Diagnosticos)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminServiciosComunes.DiagnosticosModificar(mo_Diagnosticos)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminServiciosComunes.DiagnosticosEliminar(mo_Diagnosticos)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

       Set mo_Diagnosticos = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(Me.IdDiagnostico)
        If mo_AdminServiciosComunes.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
       
       If Not mo_Diagnosticos Is Nothing Then
           With mo_Diagnosticos
           Me.IdDiagnostico = .IdDiagnostico
           mo_cmbIdCapitulo.BoundText = .IdCapitulo
           mo_cmbIdGrupo.BoundText = .IdGrupo
           mo_cmbIdCategoria.BoundText = .IdCategoria
           Me.txtCodigoCIE2004.Text = .CodigoCIE2004
           Me.txtEdadMinDias.Text = .EdadMinDias
           Me.txtEdadMaxDias.Text = .EdadMaxDias
           Me.chkRestriccion.Value = IIf(.Restriccion, 1, 0)
           Me.chkIntrahospitalario.Value = IIf(.Intrahospitalario, 1, 0)
           Me.chkMorbilidad.Value = IIf(.Morbilidad, 1, 0)
           Me.chkGestacion.Value = IIf(.Gestacion, 1, 0)
           Me.txtCodigoExportacion.Text = .CodigoExportacion
           Me.txtCodigoCIE10.Text = .CodigoCIE10
           Me.txtCodigoCIE9.Text = .CodigoCIE9
           Me.txtDescripcion.Text = .Descripcion
           mo_CmbIdTipoSexo.BoundText = .idTipoSexo
           Me.txtDescripcionMINSA.Text = .DescripcionMINSA
           Me.chkEsActivo.Value = IIf(.EsActivo, 1, 0)
           Me.mskFechaInicioVigencia.Text = IIf(.FechaInicioVigencia = 0, _
                                            SIGHEntidades.FECHA_VACIA_DMY, Format(.FechaInicioVigencia, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
           lblId.Caption = Me.IdDiagnostico
               mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdDiagnostico = 0
           Me.txtCodigoCIE2004.Text = ""
           Me.txtEdadMinDias.Text = ""
           Me.txtEdadMaxDias.Text = ""
           Me.chkRestriccion.Value = 0
           Me.chkIntrahospitalario.Value = 0
           Me.chkMorbilidad.Value = 0
           Me.chkGestacion.Value = 0
           Me.txtCodigoExportacion.Text = ""
           Me.txtCodigoCIE10.Text = ""
           Me.txtCodigoCIE9.Text = ""
           Me.txtDescripcion.Text = ""
           mo_CmbIdTipoSexo.BoundText = ""
           mo_cmbIdCapitulo.BoundText = ""
           mo_cmbIdGrupo.BoundText = ""
           mo_cmbIdCategoria.BoundText = ""
           Me.txtDescripcionMINSA.Text = ""
           Me.mskFechaInicioVigencia.Text = SIGHEntidades.FECHA_VACIA_DMY
           Me.chkEsActivo.Value = 0
           lblId.Caption = ""
End Sub
'mgaray
Private Function validarDuplicadoDiagnostico(oRsBuscaCodigo As ADODB.Recordset, _
            lIdDiagnostico As Long) As Boolean
    Dim bReturnValue As Boolean
    bReturnValue = True
    If oRsBuscaCodigo.RecordCount > 0 Then
       oRsBuscaCodigo.MoveFirst
       Do While Not oRsBuscaCodigo.EOF
          If (UCase(Trim(oRsBuscaCodigo.Fields!CodigoCIE2004)) = UCase(Trim(Me.txtCodigoCIE2004.Text)) _
                    And UCase(Trim(oRsBuscaCodigo.Fields!Descripcion)) = UCase(Trim(Me.txtDescripcion.Text))) _
                    And oRsBuscaCodigo.Fields!IdDiagnostico <> lIdDiagnostico Then
             bReturnValue = False
             Exit Do
          End If
          oRsBuscaCodigo.MoveNext
       Loop
    End If
    validarDuplicadoDiagnostico = bReturnValue
End Function


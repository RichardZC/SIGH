VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form labRepAuditoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorio: Auditoría"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "labAuditoria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11280
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
      Height          =   1035
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   11205
      Begin VB.ComboBox cmbResponsable 
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   3000
      End
      Begin VB.ComboBox cmbIdPuntoDeCarga 
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
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   2595
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   5130
         TabIndex        =   3
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   3060
         TabIndex        =   2
         Top             =   630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   6510
         TabIndex        =   4
         Top             =   630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
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
         Left            =   6150
         Picture         =   "labAuditoria.frx":0CCA
         TabIndex        =   13
         Top             =   570
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pto. Carga"
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
         Left            =   7500
         TabIndex        =   14
         Top             =   330
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   4560
         TabIndex        =   12
         Top             =   690
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F. de Movimiento"
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
         TabIndex        =   11
         Top             =   690
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empleado"
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
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   8
      Top             =   4230
      Width           =   11220
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "labAuditoria.frx":0FDC
         DownPicture     =   "labAuditoria.frx":143C
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
         Left            =   4170
         Picture         =   "labAuditoria.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "labAuditoria.frx":1D26
         DownPicture     =   "labAuditoria.frx":21EA
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
         Left            =   5700
         Picture         =   "labAuditoria.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdAuditoria 
      Height          =   3105
      Left            =   30
      TabIndex        =   7
      Top             =   1110
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   5477
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reporte de Auditoría de Movimientos Entradas y Salidas por Usuario y Punto de Carga"
   End
End
Attribute VB_Name = "labRepAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de auditoría
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim rsReporte As New ADODB.Recordset
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long
Dim FI As Date, FF As Date, HI As String, HF As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsResponsables As New Recordset

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Private Sub ConfiguraFecha()
  If Len(Trim(txtFdesde.Text)) < 10 Or Len(Trim(txtFhasta.Text)) < 10 Then Exit Sub
  If txtHrInicio.Text <> "" Then
    HI = " " & txtHrInicio.Text & ":00"
  Else
    HI = " 00:00:00"
  End If
  If txtHrFin.Text <> "" Then
    HF = " " & txtHrFin.Text & ":59"
  Else
    HF = " 23:59:59"
  End If
  FI = CDate(txtFdesde.Text & HI)
  FF = CDate(txtFhasta.Text & HF)
  If FI > FF Then
     MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
     Exit Sub
  End If
End Sub

Private Function Verifica() As Boolean
  Verifica = False
  If cmbResponsable.Text = "" Or Not (IsDate(txtFdesde.Text)) Or Not (IsDate(txtFhasta.Text)) Then
    Verifica = False
  Else
    Verifica = True
  End If
End Function

Private Sub btnAceptar_Click()
  Dim FI As Date, FF As Date, HI As String, HF As String
  
End Sub

Function ValidaDatosObligatorios() As Boolean
  sMensaje = ""
  ml_TextoDelFiltro = "FILTROS:   Pto.Carga: (" & Trim(cmbIdPuntoDeCarga.Text) & ")      F.Movimiento: (" & txtFdesde.Text & "   al " & txtFhasta.Text & ")     Usuario: " & cmbResponsable.Text
  If cmbResponsable.Text = "" Then
    sMensaje = sMensaje + "- Elija algún empleado de Laboratorio" + Chr(13)
    cmbResponsable.SetFocus
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
  LimpiarVariablesDeMemoria
  Unload Me
End Sub

Private Sub cmbIdPuntoDeCarga_Click()
  If Verifica Then btnAceptar_Click
End Sub

Private Sub cmbIdPuntoDeCarga_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, cmbIdPuntoDeCarga
End Sub

Private Sub cmbResponsable_Click()
  ConfiguraFecha
  If ValidaDatosObligatorios And Verifica Then
    Me.MousePointer = 11
    Set rsReporte = mo_ReglasLaboratorio.LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga(Val(mo_cmbResponsable.BoundText), FI, FF, 0)
    Set grdAuditoria.DataSource = rsReporte
    mo_Apariencia.ConfigurarFilasBiColores grdAuditoria, sighentidades.GrillaConFilasBicolor
    Me.MousePointer = 1
  End If
End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, cmbResponsable
End Sub

Private Sub Form_Initialize()
  'Set mo_cmbResponsable.MiComboBox = cmbIdPuntoDeCarga
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Sub InicializaFechaHora()
  txtFdesde.Text = Date
  txtFhasta.Text = Date
  txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267)
  txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268)

End Sub
Private Sub Form_Load()
  InicializaFechaHora
  CargaComboBox
End Sub

Sub CargaComboBox()
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  Set oRsResponsables = mo_ReglasLaboratorio.TodosEmpleadosDeLab()
  Set mo_cmbResponsable.RowSource = oRsResponsables
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      btnCancelar_Click
    Case vbKeyF2
      btnAceptar_Click
  End Select
End Sub

Private Sub grdAuditoria_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  'grdAuditoria.Bands(0).Columns("CantidadFallada").Hidden = True
  grdAuditoria.Bands(0).Columns("MovTipo").Hidden = True
  grdAuditoria.Bands(0).Columns("idTipoConcepto").Hidden = True
  grdAuditoria.Bands(0).Columns("idProducto").Hidden = True
  grdAuditoria.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  grdAuditoria.Bands(0).Columns("cantidad").Width = 800
  grdAuditoria.Bands(0).Columns("Codigo").Header.Caption = "Código"
  grdAuditoria.Bands(0).Columns("Codigo").Width = 800
  grdAuditoria.Bands(0).Columns("nombre").Header.Caption = "Nombre Insumo"
  grdAuditoria.Bands(0).Columns("nombre").Width = 2500
  grdAuditoria.Bands(0).Columns("idmovimiento").Header.Caption = "Id Movimiento"
  grdAuditoria.Bands(0).Columns("idmovimiento").Width = 800
  grdAuditoria.Bands(0).Columns("concepto").Header.Caption = "Concepto"
  grdAuditoria.Bands(0).Columns("Concepto").Width = 2000
  grdAuditoria.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
  grdAuditoria.Bands(0).Columns("Fecha").Width = 2500
End Sub

Private Sub txtFdesde_Change()
  If Verifica = False Then Exit Sub
  cmbResponsable_Click
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtFdesde
End Sub

Private Sub txtFdesde_LostFocus()
  If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
      MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
      InicializaFechaHora
    End If
  End If
End Sub

Private Sub txtFhasta_Change()
  If Verifica = False Then Exit Sub
  cmbResponsable_Click
End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtFhasta
End Sub

Private Sub txtFhasta_LostFocus()
  If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
      MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
      InicializaFechaHora
    End If
  End If
End Sub

Sub LimpiarVariablesDeMemoria()
  On Error Resume Next
  Set mo_ReglasFarmacia = Nothing
  Set mo_Teclado = Nothing
  Set mo_cmbResponsable = Nothing
  Set mo_ReglasFacturacion = Nothing
  Set mo_reglasComunes = Nothing
  Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_Change()
  If Verifica = False Then Exit Sub
  cmbResponsable_Click
End Sub

Private Sub txtHrFin_LostFocus()
If Not sighentidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_Change()
  If Verifica = False Then Exit Sub
  cmbResponsable_Click
End Sub

Private Sub txtHrInicio_LostFocus()
If Not sighentidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

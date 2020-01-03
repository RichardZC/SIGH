VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCajaNotaCredito 
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12825
   ScaleHeight     =   6255
   ScaleWidth      =   12825
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   75
      TabIndex        =   3
      Top             =   540
      Width           =   12690
      Begin VB.CommandButton bntReporte 
         Enabled         =   0   'False
         Height          =   705
         Left            =   10440
         Picture         =   "ucCajaNotaCredito.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   885
      End
      Begin VB.TextBox TxtRsocial 
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
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtNroSerie 
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
         Left            =   120
         MaxLength       =   3
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtNroDocumento 
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
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbEstadoNota 
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   11355
         Picture         =   "ucCajaNotaCredito.ctx":04D9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   11340
         Picture         =   "ucCajaNotaCredito.ctx":30B5
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   7080
         TabIndex        =   12
         ToolTipText     =   "Fecha desde"
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   8520
         TabIndex        =   13
         ToolTipText     =   "Fecha Hasta"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
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
         Caption         =   "Razón Social o Nombres"
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
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha: Desde - Hasta"
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
         Left            =   7320
         TabIndex        =   14
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label8 
         Caption         =   "Nº Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Documento"
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
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label10 
         Caption         =   "Estado"
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
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdNotasCreditoDebito 
      Height          =   4590
      Left            =   90
      TabIndex        =   2
      Top             =   1560
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   8096
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de Notas de Crédito"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Notas de Credito"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   12735
   End
End
Attribute VB_Name = "ucCajaNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar Cajas
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbEstadoNota  As New sighentidades.ListaDespleglable
Dim ml_lnHwnd As Long
Property Let lnHWnd(lValue As Long)
    ml_lnHwnd = lValue
End Property
Property Get lnHWnd() As Long
    lnHWnd = ml_lnHwnd
End Property
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdNotasCreditoDebito.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdNotasCreditoDebito.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoFiltro(lValue As sghTipoFiltroPacientes)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroPacientes
    TipoFiltro = ml_TipoFiltro
End Property

Private Sub bntReporte_Click()
   On Error GoTo errRp
   Dim oRsTmp1 As New Recordset
   Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
   Set oRsTmp1 = grdNotasCreditoDebito.DataSource
   If oRsTmp1.RecordCount > 0 Then
      mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp1, lblNombre.Caption, txtFdesde.Text & " " & txtFhasta.Text, "", ml_lnHwnd, False, True
   End If
errRp:
   Set oRsTmp1 = Nothing
   Set mo_ReglasReportes = Nothing
End Sub

Private Sub btnBuscar_Click()
    'Valida filtros
    If UserControl.txtFdesde.Text = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "La fecha desde no debe estar vacia", vbInformation, "Nota de Crédito"
        Exit Sub
    End If
    If Not IsDate(UserControl.txtFdesde.Text) Then
        MsgBox "La fecha desde no tiene el formato correcto", vbInformation, "Nota de Crédito"
        Exit Sub
    End If
    If UserControl.txtFhasta.Text = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "La fecha hasta no debe estar vacia", vbInformation, "Nota de Crédito"
        Exit Sub
    End If
    If Not IsDate(UserControl.txtFhasta.Text) Then
        MsgBox "La fecha no tiene el formato correcto", vbInformation, "Nota de Crédito"
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
    Set grdNotasCreditoDebito.DataSource = mo_AdminCaja.NotaCreditoRegistrosTotalesPorNumYFecha(Trim(txtNroSerie.Text), Trim(txtNroDocumento.Text), IIf(mo_cmbEstadoNota.BoundText = "", 0, mo_cmbEstadoNota.BoundText), _
                                                                                                Trim(TxtRsocial.Text), txtFdesde.Text, txtFhasta.Text)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox mo_AdminCaja.MensajeError, vbInformation, "Nota de Crédito"
    End If
    InitializeLayout
   ' mo_Apariencia.ConfigurarFilasBiColores grdNotasCreditoDebito, sighentidades.GrillaConFilasBicolor
    bntReporte.Enabled = True
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtNroSerie.Text = ""
    UserControl.txtNroDocumento.Text = ""
    UserControl.TxtRsocial.Text = ""
    UserControl.txtFdesde.Text = Date
    UserControl.txtFhasta.Text = Date
    mo_cmbEstadoNota.BoundText = ""
'    UserControl.txtCodigo = ""
'    UserControl.txtDescripcion = ""
End Sub
Private Sub cmbCajero_Click()
    'mo_cmbCajero.BoundColumn = "IdCajero"
    'mo_cmbCajero.ListField = "NombreCompleto"
    'Set mo_cmbCajero.RowSource = mo_AdminCaja.ServiciosSeleccionarPorTipoV2(Val(mo_cmbCajero.BoundText))
End Sub

Private Sub grdNotasCreditoDebito_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdNotasCreditoDebito.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdNota")
    
End Sub

Private Sub grdNotasCreditoDebito_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdNotasCreditoDebito.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdNota")
    
End Sub

Sub InitializeLayout()
    grdNotasCreditoDebito.Bands(0).Columns("IdNota").Hidden = True
    grdNotasCreditoDebito.Bands(0).Columns("IdEstadoNota").Hidden = True
    grdNotasCreditoDebito.Bands(0).Columns("Documento").Header.Caption = "Nota de Crédito"
    grdNotasCreditoDebito.Bands(0).Columns("Documento").Width = 3000
    grdNotasCreditoDebito.Bands(0).Columns("FechaAprueba").Header.Caption = "Aprobado"
    grdNotasCreditoDebito.Bands(0).Columns("FechaAprueba").Width = 1200
    grdNotasCreditoDebito.Bands(0).Columns("Comprobante Afectado").Width = 4000
    grdNotasCreditoDebito.Bands(0).Columns("Observaciones").Width = 3000
    grdNotasCreditoDebito.Bands(0).Columns("Total").Width = 1000
    grdNotasCreditoDebito.Bands(0).Columns("EstadoNota").Width = 1500
    grdNotasCreditoDebito.Bands(0).Columns("Cajero").Width = 2500
    grdNotasCreditoDebito.Bands(0).Columns("FechaPagado").Width = 1200
End Sub


Private Sub grdNotasCreditoDebito_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
'    grdNotasCreditoDebito.Bands(0).Columns("IdCaja").Hidden = True
'
'    grdNotasCreditoDebito.Bands(0).Columns("Codigo").Header.Caption = "Código"
'    grdNotasCreditoDebito.Bands(0).Columns("Codigo").Width = 700
'
'    grdNotasCreditoDebito.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
'    grdNotasCreditoDebito.Bands(0).Columns("Descripcion").Width = 4000
    
    'grdNotasCreditoDebito.Bands(0).Columns("NroSerie").Header.Caption = "Nº Serie"
    'grdNotasCreditoDebito.Bands(0).Columns("NroSerie").Width = 1000
    
    'grdNotasCreditoDebito.Bands(0).Columns("NroComprobante").Header.Caption = "Ult.Comprobante Emitido"
    'grdNotasCreditoDebito.Bands(0).Columns("NroComprobante").Width = 3000

End Sub
'Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, UserControl.txtDescripcion
'End Sub
'Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, UserControl.btnBuscar
'End Sub
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    AdministrarKeyPreview KeyCode
'End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
     Case vbKeyF8
    End Select
End Sub

'kike 2017
Private Sub grdNotasCreditoDebito_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        Select Case Val(Row.Cells("IdEstadoNota").GetText())
        Case 2   'anulado
            Row.Appearance.ForeColor = vbRed
        End Select

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdNotasCreditoDebito.Width = fraBusqueda.Width
   grdNotasCreditoDebito.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdNotasCreditoDebito, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdNotasCreditoDebito, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub


Public Sub inicializar()
     SkinConfigura
'    cmbEstadoNota.Clear
    'Buscamos el comprobante
    Set mo_cmbEstadoNota.MiComboBox = cmbEstadoNota
    mo_cmbEstadoNota.BoundColumn = "IdEstado"
    mo_cmbEstadoNota.ListField = "EstadoNota"
    Set mo_cmbEstadoNota.RowSource = mo_AdminCaja.NotaCreditoDebitoCargarEstadoNotaCredito
    txtFdesde.Text = Date
    txtFhasta.Text = Date
'    cmbEstadoNota.AddItem Date
'    cmbEstadoNota.AddItem "Todas"
'    cmbFecha.ListIndex = 0
End Sub







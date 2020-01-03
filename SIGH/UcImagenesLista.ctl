VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcImagenesLista 
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10920
   ScaleHeight     =   8160
   ScaleWidth      =   10920
   Begin VB.Frame fraDetalle 
      Caption         =   "DETALLE DE LA ORDEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2265
      Left            =   0
      TabIndex        =   19
      Top             =   5805
      Width           =   10890
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   195
         Width           =   3570
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   210
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   195
         Width           =   1305
      End
      Begin UltraGrid.SSUltraGrid grdListaOrdenesDetalle 
         Height          =   1320
         Left            =   45
         TabIndex        =   23
         Top             =   585
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   2328
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
         Caption         =   "grdListaOrdenesDetalle"
      End
      Begin VB.Label lblEPSpago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   10605
         TabIndex        =   27
         Top             =   225
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Ayuda: <Doble Click sobre el nombre de la prueba> = Ingreso / Modificación de resultados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   26
         Top             =   2010
         Width           =   8625
      End
      Begin VB.Label Label6 
         Caption         =   "Paciente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N° Historia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4725
         TabIndex        =   24
         Top             =   240
         Width           =   900
      End
   End
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
      Height          =   2235
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6990
      Begin VB.ComboBox cmbResponsable 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4545
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1065
         Width           =   2385
      End
      Begin VB.CheckBox chkPorFcpt 
         Caption         =   "Filtrar por Fechas realizar CPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   1575
         Width           =   2835
      End
      Begin VB.TextBox txtNombres 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4935
         MaxLength       =   40
         TabIndex        =   14
         Top             =   480
         Width           =   1950
      End
      Begin VB.TextBox txtNroCuenta 
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
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   7
         Top             =   480
         Width           =   1425
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   3060
         MaxLength       =   9
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtNroMovimiento 
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
         Left            =   150
         MaxLength       =   9
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   5610
         Picture         =   "UcImagenesLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   1305
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4230
         Picture         =   "UcImagenesLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1815
         Width           =   1305
      End
      Begin VB.ComboBox cmbIdPtoCarga 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   1665
      End
      Begin MSMask.MaskEdBox txtFinicio 
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   1080
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFfinal 
         Height          =   315
         Left            =   1530
         TabIndex        =   9
         Top             =   1080
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFcpt1 
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   1785
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtFCpt2 
         Height          =   315
         Left            =   1530
         TabIndex        =   16
         Top             =   1785
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         BackStyle       =   0  'Transparent
         Caption         =   "N° Movimiento     Nro Cuenta        N° Historia                 Paciente"
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
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   9105
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "            Fechas Movimiento            Punto de Carga           Responsable"
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
         Left            =   135
         TabIndex        =   10
         Top             =   870
         Width           =   6810
      End
   End
   Begin UltraGrid.SSUltraGrid grdListaOrdenes 
      Height          =   3030
      Left            =   0
      TabIndex        =   12
      Top             =   2745
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   5345
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
      Caption         =   "Lista de Movimientos"
   End
   Begin UltraGrid.SSUltraGrid grdBoletas 
      Height          =   2160
      Left            =   6990
      TabIndex        =   13
      Top             =   570
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   3810
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Boletas pendientes  (marcar y  F2->Agregar)"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Imageneología"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
   End
End
Attribute VB_Name = "UcImagenesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Rayos X,Tomografías, Ecografías registradas
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim ml_idRegistroSeleccionado As Long
Dim ml_PuntoCarga As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_cmbIdPuntoCargaB As New sighentidades.ListaDespleglable
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim ml_IdTipoFinanciamiento As Long
Dim oRsFarmacias As New ADODB.Recordset
Dim oRsBoletas As New ADODB.Recordset
Dim oRsLista As New Recordset
Dim rs As New Recordset
Dim rsTmp As New Recordset
Dim rsResultados As New Recordset
Dim ml_SeEligioGridBoleta As Boolean
Dim ml_idOrdenLab As Long, ml_IdPaciente As Long

Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdListaOrdenes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdListaOrdenes.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
    mo_cmbIdPuntoCargaB.BoundText = ml_PuntoCarga
    Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & mo_ReglasFarmacia.EmpleadosDevuelveIdCargoSegunPuntoCarga(ml_PuntoCarga))
End Property
Property Get PuntoCarga() As Long
    PuntoCarga = ml_PuntoCarga
End Property
Property Let HabilitarPuntoCarga(lValue As Long)
    cmbIdPtoCarga.Enabled = lValue
End Property
Property Get HabilitarPuntoCarga() As Long
    HabilitarPuntoCarga = cmbIdPtoCarga.Enabled
End Property
Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property
Property Get idTipoFinanciamiento() As Long
    idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Private Sub btnBuscar_Click()
    If CDate(UserControl.txtFinicio.Text) > CDate(UserControl.txtFfinal.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set grdListaOrdenesDetalle.DataSource = Nothing
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        Dim ldFechaIni As Date
        Dim ldFechaFin As Date
        Dim lcFiltro As String
        Dim oRsTmp1 As New Recordset
        Dim lcBoleta As String, lnTotal As Double, ldFecha As Date, lnIdProducto As Long
        Dim lnIdComprobantePago As Long, lbEsDelPuntoCarga As Boolean, lcPacienteFiltro As String
        lcPacienteFiltro = ""
        ml_SeEligioGridBoleta = False
        If (UserControl.txtNroHistoria = "" And _
            UserControl.txtNroMovimiento = "" And _
            UserControl.txtNroCuenta = "" And mo_cmbIdPuntoCarga.BoundText = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Nro Historia, Nro Cuenta, Nro Movimiento)", vbInformation, "Filtro de ordenes de Búsqueda"
            Exit Sub
        End If
        If txtFinicio.Enabled = False Then
            ldFechaIni = Format(txtFcpt1.Text & " 00:00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
            ldFechaFin = Format(txtFCpt2.Text & " 23:59:59", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
        Else
            ldFechaIni = Format(txtFinicio.Text & " 00:00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
            ldFechaFin = Format(txtFfinal.Text & " 23:59:59", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
            If ldFechaIni > ldFechaFin Then
               MsgBox "La FECHA FINAL debe ser mayor a la FECHA INICIAL", vbInformation, ""
               Exit Sub
            End If
            
        End If
        lcFiltro = ""
        If txtNroMovimiento.Text <> "" Then
           lcFiltro = lcFiltro & "idMovimiento=" & Trim(Str(Val(txtNroMovimiento.Text)))
           txtNroHistoria.Text = ""
           txtNroCuenta.Text = ""
        End If
        If mo_Teclado.TextoEsSoloNumeros(txtNroHistoria.Text) Then
           lcFiltro = lcFiltro & "NroHistoriaClinica=" & txtNroHistoria.Text
           txtNroCuenta.Text = ""
           txtNroMovimiento.Text = ""
        End If
        If mo_Teclado.TextoEsSoloNumeros(txtNroCuenta.Text) Then
           lcFiltro = lcFiltro & "idCuentaAtencion=" & Trim(Str(Val(txtNroCuenta.Text)))
           txtNroMovimiento.Text = ""
           txtNroHistoria.Text = ""
        End If
        If Trim(txtNombres.Text) <> "" Then
           lcPacienteFiltro = Trim(txtNombres.Text)
        End If
        If Val(mo_cmbResponsable.BoundText) > 0 And lcFiltro = "" Then
           lcFiltro = "idPersonaTomaImagen=" & mo_cmbResponsable.BoundText
        End If
        Set oRsLista = mo_reglasImagen.ImagMovimientoSeleccionarPorFechasPuntoCarga(Val(mo_cmbIdPuntoCarga.BoundText), _
                                          ldFechaIni, ldFechaFin, _
                                          lcPacienteFiltro, IIf(txtFinicio.Enabled = False, 0, 1))
        If lcFiltro <> "" Then
           oRsLista.Filter = lcFiltro
        End If
        Set grdListaOrdenes.DataSource = oRsLista
        If mo_reglasImagen.MensajeError <> "" Then
            MsgBox mo_reglasImagen.MensajeError, vbInformation, lblNombre.Caption
        Else
            Set grdBoletas.DataSource = mo_ReglasCaja.BoletasServicioPorPuntoCarga(ldFechaIni, ldFechaFin, Val(mo_cmbIdPuntoCarga.BoundText))
        End If
        Set oRsTmp1 = Nothing
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNroHistoria = ""
        UserControl.txtNroMovimiento = ""
        UserControl.txtNroCuenta = ""
        UserControl.txtNombres.Text = ""
        mo_cmbResponsable.BoundText = ""
        ml_SeEligioGridBoleta = False
        UserControl.txtNroMovimiento.SetFocus
End Sub




Private Sub chkPorFcpt_Click()
    If chkPorFcpt.Value = 1 Then
        txtFcpt1.Enabled = True
        txtFCpt2.Enabled = True
        txtFinicio.Enabled = False
        txtFfinal.Enabled = False
    Else
        txtFcpt1.Enabled = False
        txtFCpt2.Enabled = False
        txtFinicio.Enabled = True
        txtFfinal.Enabled = True
    End If
End Sub

Private Sub cmbIdPtoCarga_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, cmbIdPtoCarga
    AdministrarKeyPreview KeyCode
End Sub





Private Sub cmbResponsable_LostFocus()
    If cmbResponsable.Text <> "" Then
       'cmbResponsable.Text = ""
       txtNroCuenta.Text = ""
       txtNroHistoria.Text = ""
       txtNombres.Text = ""
       UserControl.txtNroMovimiento.Text = ""
    End If
End Sub



Private Sub grdBoletas_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdBoletas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idComprobantePago")
    ml_SeEligioGridBoleta = True
End Sub

Private Sub grdBoletas_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdBoletas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdComprobantePago")
    ml_SeEligioGridBoleta = True
End Sub

Private Sub grdBoletas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdBoletas.Bands(0).Columns("IdComprobantePago").Hidden = True
    grdBoletas.Bands(0).Columns("Boleta").Width = 1300
    grdBoletas.Bands(0).Columns("total").Width = 500
    grdBoletas.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
    grdBoletas.Bands(0).Columns("Fecha").Width = 1500
End Sub

Private Sub grdListaOrdenes_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdListaOrdenes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
    ml_SeEligioGridBoleta = False
End Sub

Private Sub grdListaOrdenes_Click()
Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdListaOrdenes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
    ml_SeEligioGridBoleta = False
End Sub


Private Sub grdListaOrdenes_DblClick()
  ml_SeEligioGridBoleta = False
  Dim rsRecordset As ADODB.Recordset

  
  Text1.Text = ""
  Text2.Text = ""
  
  Text4.Text = ""
  ml_idRegistroSeleccionado = 0



  
  Set rsRecordset = grdListaOrdenes.DataSource
  If rsRecordset.State = adStateClosed Then Exit Sub
  If Not (rsRecordset.EOF = True And rsRecordset.BOF = True) Then

    If rsRecordset("IdImagEstado") = "0" Then
      grdListaOrdenesDetalle.Enabled = False
      MsgBox "Esta Orden fue anulada.", vbInformation, "SIGH "
    Else
        grdListaOrdenesDetalle.Enabled = True
    End If
    
    
    
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
    idRegistroSeleccionado = rsRecordset("IdMovimiento")
    ml_idOrdenLab = rsRecordset("IdOrden")
   ' ml_idPaciente = IIf(IsNull(rsRecordset("idpaciente")), 0, rsRecordset("idpaciente"))
   
   mo_ReglasImagenes.ResultadosAutomaticosActualizaImgHaciaGalenhos ml_idRegistroSeleccionado


    'lcServicioActualPaciente = mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(rsRecordset("IdOrden"))
    '
    Set rs = mo_reglasImagen.ImagMovimientoCPTSeleccionarPorIdMovimiento(ml_idRegistroSeleccionado)
    If rs.RecordCount > 0 Then
         
      Set rsResultados = mo_reglasImagen.ImagMovimientoResultadosSeleccionarPorId(ml_idRegistroSeleccionado)
      
      If rsTmp.State = adStateOpen Then Set rsTmp = Nothing
      With rsTmp
        .Fields.Append "Imprime", adBoolean
        .Fields.Append "NroOrden", adDouble
        .Fields.Append "idOrden", adDouble
        .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
        .Fields.Append "Nombre", adVarChar, 250, adFldIsNullable
        .Fields.Append "idProducto", adDouble
        .Fields.Append "Cantidad", adInteger
        .Fields.Append "Precio", adDouble
        .Fields.Append "Total", adDouble
        .Fields.Append "Resultado", adVarChar, 2, adFldIsNullable
        .Fields.Append "ResultadoAutomatico", adBoolean
        .Fields.Append "ObsReceta", adVarChar, 300, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
      End With
  
      Dim Tot As Double
      Dim TotP As Integer
      Dim TotRes As Integer
      Dim T As Integer
      
      Tot = 0: TotP = 0: TotRes = 0: T = 0
      If rs.RecordCount > 0 Then
      rs.MoveFirst
      Do While Not rs.EOF
      
        Tot = Tot + rs!Cantidad * rs!Precio
        TotP = TotP + 1
        rsTmp.AddNew
        T = T + 1
        rsTmp!NroOrden = T
        rsTmp!IdOrden = ml_idOrdenLab
        rsTmp!idProducto = rs!idProductoCpt
        rsTmp!Cantidad = rs!Cantidad
        rsTmp!Precio = rs!Precio
        rsTmp!Total = Round(rs!Cantidad * rs!Precio, 2)
        rsTmp!Codigo = rs!Codigo
        rsTmp!nombre = Left(rs!nombre, 250)
        '
        rsResultados.Filter = "idProductoCpt=" & rs!idProductoCpt
        If rsResultados.RecordCount > 0 Then
          rsTmp!resultado = "SI"
          rsTmp!Imprime = True
          TotRes = TotRes + 1
        Else
          rsTmp!resultado = "NO"
        End If
        '
        If Not IsNull(rs!LabResultadoAutomatico) Then
           If rs!LabResultadoAutomatico = 1 Then
              rsTmp!ResultadoAutomatico = True
           End If
        End If
        rsTmp.Update
        rs.MoveNext
      Loop
      End If
      rsResultados.Filter = ""
      lblEPSpago.Caption = ""
      '*********Proviene de una Receta (inicio)
      If rsRecordset!idCuentaAtencion > 0 Then
      
            '
            Dim oRsTmp1 As New Recordset
            Dim lcBoletaEPS  As String, lcOrdenPago
            Dim oConexion As New Connection
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighentidades.CadenaConexion
            '
            Set oRsTmp1 = mo_reglasComunes.AtencionesSeleccionarMedicoPorCuenta(rsRecordset!idCuentaAtencion)
            If oRsTmp1.RecordCount > 0 Then
                lcOrdenPago = mo_ReglasFacturacion.DevuelveOrdenPago(oRsTmp1!idAtencion, sghPtoCargaCaja, _
                                                                     rsRecordset!fecha, oConexion, lcBoletaEPS)
                If lcBoletaEPS <> "" Then
                    lblEPSpago.Caption = "Pagó EPS - " & lcBoletaEPS
                    lblEPSpago.ForeColor = vbBlack
                ElseIf lcOrdenPago <> "" Then
                    lblEPSpago.Caption = "NO PAGO EPS - N°OrdenPago: " & lcOrdenPago
                    lblEPSpago.ForeColor = vbRed
                End If
            End If
            '
            
            Dim lnidReceta As Long
            lnidReceta = 0
            Set oRsTmp1 = mo_reglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(Trim(Str(rsRecordset!IdMovimiento)), rsRecordset!idCuentaAtencion)
            If oRsTmp1.RecordCount > 0 Then
               lnidReceta = oRsTmp1.Fields!idReceta
            End If
            oRsTmp1.Close
            If lnidReceta > 0 Then
                Set oRsTmp1 = mo_reglasComunes.RecetaDetalleSeleccioarPorIdReceta(lnidReceta, oConexion)
                oRsTmp1.Filter = "observaciones<>''"
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.MoveFirst
                   Do While Not oRsTmp1.EOF
                      If Not IsNull(oRsTmp1!Observaciones) Then
                         rsTmp.MoveFirst
                         rsTmp.Find "idProducto=" & oRsTmp1!idItem
                         If Not rsTmp.EOF Then
                            rsTmp!obsReceta = oRsTmp1!Observaciones
                            rsTmp.Update
                         End If
                      End If
                      oRsTmp1.MoveNext
                   Loop
                End If
                oRsTmp1.Close
             End If
             oConexion.Close
             Set oConexion = Nothing
             Set oRsTmp1 = Nothing
       End If
       '*********Proviene de una Receta (fin)
    End If

    Set rs = Nothing
    Set grdListaOrdenesDetalle.DataSource = rsTmp
    
   ' mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenesDetalle, sighentidades.GrillaConFilasBicolor
    
    Text1.Text = UCase(rsRecordset("Paciente"))

    '

    Text4.Text = IIf(IsNull(rsRecordset("NroHistoriaClinica")), "", rsRecordset("NroHistoriaClinica"))
    If rsTmp.RecordCount > 0 Then
       rsTmp.MoveFirst
    End If
  End If

End Sub

Private Sub grdListaOrdenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    grdListaOrdenes.Bands(0).Columns("IdImagEstado").Hidden = True
    grdListaOrdenes.Bands(0).Columns("IdOrden").Hidden = True
    grdListaOrdenes.Bands(0).Columns("ApellidoPaterno").Hidden = True
    grdListaOrdenes.Bands(0).Columns("ApellidoMaterno").Hidden = True
    grdListaOrdenes.Bands(0).Columns("PrimerNombre").Hidden = True
    grdListaOrdenes.Bands(0).Columns("SegundoNombre").Hidden = True
    
    grdListaOrdenes.Bands(0).Columns("IdMovimiento").Header.Caption = "N° Movimiento"
    grdListaOrdenes.Bands(0).Columns("IdMovimiento").Width = 1200
    
    grdListaOrdenes.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
    grdListaOrdenes.Bands(0).Columns("Fecha").Width = 1700
        
        
    grdListaOrdenes.Bands(0).Columns("Estado").Header.Caption = "Estado"
    grdListaOrdenes.Bands(0).Columns("Estado").Width = 1000
    
    grdListaOrdenes.Bands(0).Columns("idCuentaAtencion").Header.Caption = "N°Cuenta"
    grdListaOrdenes.Bands(0).Columns("idCuentaAtencion").Width = 1500
    
    grdListaOrdenes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdListaOrdenes.Bands(0).Columns("NroHistoriaClinica").Width = 1500

    grdListaOrdenes.Bands(0).Columns("Paciente").Header.Caption = "Paciente"
    grdListaOrdenes.Bands(0).Columns("Paciente").Width = 3900



End Sub

Private Sub grdListaOrdenes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdImagEstado").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
            'Row.Appearance.Font.Strikethrough = True
        End If

End Sub












Private Sub grdListaOrdenesDetalle_DblClick()
    On Error GoTo ErrGrdListOrd
    
    Dim oResultadosImg As New SIGHImagen.ResultadosImg
    oResultadosImg.Producto = rsTmp!Codigo & " " & rsTmp!nombre
    oResultadosImg.EsResultadoAutomatico = rsTmp!ResultadoAutomatico
    oResultadosImg.idProductoCpt = rsTmp!idProducto
    oResultadosImg.IdMovimiento = ml_idRegistroSeleccionado
    Set oResultadosImg.rsResultados = rsResultados
    oResultadosImg.Paciente = Text1.Text
    oResultadosImg.PuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
    If rsTmp!ResultadoAutomatico = True Then
       oResultadosImg.SoloEsConsulta = True
    Else
       oResultadosImg.SoloEsConsulta = False
    End If
    oResultadosImg.MostrarFormulario
    grdListaOrdenes_DblClick
ErrGrdListOrd:
    Set oResultadosImg = Nothing
End Sub

Private Sub grdListaOrdenesDetalle_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
  
  grdListaOrdenesDetalle.Bands(0).Columns("idOrden").Hidden = True
  grdListaOrdenesDetalle.Bands(0).Columns("ResultadoAutomatico").Activation = ssActivationActivateNoEdit
  
  grdListaOrdenesDetalle.Bands(0).Columns("NroOrden").Header.Caption = "Nº"
  grdListaOrdenesDetalle.Bands(0).Columns("NroOrden").Width = 700
  grdListaOrdenesDetalle.Bands(0).Columns("NroOrden").Activation = ssActivationActivateNoEdit ' = ssActivationAllowEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Hidden = True
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Header.Caption = "C.Producto"
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Width = 900
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  grdListaOrdenesDetalle.Bands(0).Columns("cantidad").Width = 1000
  grdListaOrdenesDetalle.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Width = 1000
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Header.Caption = "Precio"
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Format = "#0.000"
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Hidden = False
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("total").Header.Caption = "Total"
  grdListaOrdenesDetalle.Bands(0).Columns("total").Format = "#0.000"
  grdListaOrdenesDetalle.Bands(0).Columns("total").Width = 1000
  grdListaOrdenesDetalle.Bands(0).Columns("total").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Header.Caption = "C.Prueba"
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Width = "1000"
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Hidden = False
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("nombre").Header.Caption = "Nombre de Prueba"
  grdListaOrdenesDetalle.Bands(0).Columns("nombre").Activation = ssActivationActivateNoEdit
  grdListaOrdenesDetalle.Bands(0).Columns("nombre").Width = 6000
  
  grdListaOrdenesDetalle.Bands(0).Columns("resultado").Width = 1000
    
  grdListaOrdenesDetalle.Bands(0).Columns("obsReceta").Width = 4200
  
  grdListaOrdenesDetalle.Bands(0).Columns("Imprime").Width = 800
  grdListaOrdenesDetalle.Bands(0).Columns("Imprime").Style = ssStyleCheckBox

End Sub

Private Sub grdListaOrdenesDetalle_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdListaOrdenesDetalle_DblClick
    End If
End Sub

Private Sub txtFfinal_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFfinal
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFfinal_LostFocus()
    If Not EsFecha(txtFfinal.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFfinal.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFinicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFinicio
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFinicio_LostFocus()
    If Not EsFecha(txtFinicio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFinicio.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If

End Sub

Private Sub txtNroCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroCuenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroCuenta_LostFocus()
    If Len(txtNroCuenta.Text) > 0 Then
       btnBuscar_Click
    End If

End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtNroHistoria_LostFocus()
    If Len(txtNroHistoria.Text) > 0 Then
       btnBuscar_Click
    End If
    
End Sub

Private Sub txtNroMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroMovimiento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroMovimiento_LostFocus()
    If Len(txtNroMovimiento.Text) > 0 Then
       btnBuscar_Click
    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   'fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdListaOrdenes.Width = UserControl.Width - 110
   grdListaOrdenes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 3100)    '150
   
  FraDetalle.Top = grdListaOrdenes.Top + grdListaOrdenes.Height + 100
  FraDetalle.Height = (UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)) / 2 - 150
  FraDetalle.Width = grdListaOrdenes.Width
  grdListaOrdenesDetalle.Width = FraDetalle.Width - 120
  grdListaOrdenesDetalle.Height = FraDetalle.Height - grdListaOrdenesDetalle.Top - 100
   
   
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, "99"
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenesDetalle, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenesDetalle, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Sub Inicializar()
    SkinConfigura
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    txtFcpt1.Text = Date
    txtFCpt2.Text = Date
    If lcBuscaParametro.SeleccionaFilaParametro(531) = "S" Then
       chkPorFcpt.Value = 1
       chkPorFcpt_Click
    End If
    Set lcBuscaParametro = Nothing
    
    
    ConfigurarPuntosDeCarga
    
    txtFinicio.Text = Date
    txtFfinal.Text = Date
    

   ' mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
   ' mo_Apariencia.ConfigurarFilasBiColores grdBoletas, sighentidades.GrillaConFilasBicolor
    mo_Formulario.HabilitarDeshabilitar Text1, False
    mo_Formulario.HabilitarDeshabilitar Text4, False
End Sub



Sub ConfigurarPuntosDeCarga()
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idUPS=1")
    '
    Set mo_cmbResponsable.MiComboBox = cmbResponsable
    mo_cmbResponsable.BoundColumn = "idEmpleado"
    mo_cmbResponsable.ListField = "ApNom"
    
    
End Sub

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
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub

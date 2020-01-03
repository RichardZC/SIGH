VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCaja 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestión de Caja"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13035
   Icon            =   "frmCaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCajero 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   12900
      Begin VB.TextBox TxtRsocial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2130
         MaxLength       =   30
         TabIndex        =   5
         Top             =   510
         Width           =   5415
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   11535
         Picture         =   "frmCaja.frx":0CCA
         TabIndex        =   4
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "Limpiar"
         Height          =   315
         Left            =   11535
         Picture         =   "frmCaja.frx":3913
         TabIndex        =   3
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox txtNroDocumentoBusqueda 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         MaxLength       =   30
         TabIndex        =   2
         Top             =   510
         Width           =   1155
      End
      Begin VB.TextBox txtNroSerieBusqueda 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   30
         TabIndex        =   1
         Top             =   510
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   7710
         TabIndex        =   6
         Top             =   510
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   9630
         TabIndex        =   7
         Top             =   510
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCaja.frx":64EF
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
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   10815
      End
   End
   Begin UltraGrid.SSUltraGrid grdGestionCaja 
      Height          =   6675
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   11774
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
      Caption         =   "Gestion de Caja"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Gestion de Caja"
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
      Height          =   465
      Left            =   60
      TabIndex        =   10
      Top             =   120
      Width           =   12945
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: busca Documentos de Caja
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim ml_puntoCarga As Long
Dim ml_idOrden As Long
Dim mi_Opcion As sghOpcionesPago
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim mo_doCajaGestion As DOCajaGestion
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_DOComprobantePago As New DOCajaComprobantesPago
Dim mo_DOComprobantePagoDevolucion As New DOCajaComprobantesPago
Dim mo_oComprobantepago As New CajaComprobantesPago
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_cmbIdEstado As New sighentidades.ListaDespleglable
Dim mo_cmbFechaIngreso As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New sighentidades.ListaDespleglable
Dim mo_DOFactOrdenServicio As New DOFactOrdenServicio
Dim mo_DOFactOrdenBienInsumo As New DoFactOrdenesBienes
Dim mo_DOAtencion As New DOAtencion
Dim mo_DoFactOrdenServPagos  As New DoFactOrdenServPagos
Dim ml_IdOrdenDespacho As Long
Dim ml_idPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim md_Total As Double
Dim md_Ingresado As Double
Dim md_PendientePago As Double
Dim md_PagoACuenta As Double
Dim md_Exonerado As Double
Dim ml_TipoProducto As Long
Dim mo_DOCuentaAtencion As DOCuentaAtencion
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim doCajero As New SIGHComun.DOCajaCajero
Dim mo_DOPaciente As New doPaciente
Dim ml_idCuentaAtencion As Long
Dim mo_Teclado As New sighentidades.Teclado
Const ID_TIPO_COMPROBANTE_FACTURA = 2
Dim ml_IdGestionCaja As Long
Dim lbEsDevolucion As Boolean
Dim ml_NombreCajero As String
Dim lnParametrosImprimeBoleta As sghImpresion
Dim ml_IdFormaPago As Long
Dim ml_IdFarmacia As Long
Dim ml_idPreVenta As Long
Dim lbCargaEstadoDeCuentaFarmacia As Boolean    'True=Carga CUENTA DE FARMACIA en CAJA Servicios, false=carga CUENTA DE SERVICIO en CAJA Servicios
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idConfiguracionParaPreventa As Long
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oRsBusquedaRecibos As New ADODB.Recordset
Dim lbBoletaDeServicios As Boolean
Dim lnTotalGrid As Double
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lbEsUnaFactura As Boolean
Dim lbTienePermisoSoloParaBoletaFarmacia As Boolean
Dim lbTienePermisoReimprimeBoleta As Boolean
Dim lbTienePermisoExonerarPacExterno As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let NombreCajero(lValue As String)
   ml_NombreCajero = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

'********************************************************************************************
'                                   COMPROBANTE DE PAGO
'********************************************************************************************
Property Let puntoCarga(lValue As Long)
    ml_puntoCarga = lValue
End Property

Property Get puntoCarga() As Long
    puntoCarga = ml_puntoCarga
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Get IdTipoFinanciamiento() As Long
    IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Property Let idOrden(lValue As Long)
    ml_idOrden = lValue
End Property

Property Get idOrden() As Long
    idOrden = ml_idOrden
End Property

Property Let IdGestionCaja(lValue As Long)
    ml_IdGestionCaja = lValue
End Property

Property Get IdGestionCaja() As Long
    IdGestionCaja = ml_IdGestionCaja
End Property



'********************************************************************************************
'                                   GESTION DE CAJA
'********************************************************************************************
Private Sub btnAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
     AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_KeyPress(KeyAscii As Integer)
     AdministrarKeyPreview KeyAscii
End Sub

Private Sub btnBuscar_Click()
    MousePointer = 11
    RealizarBusqueda
    MousePointer = 1
End Sub

Sub LimpiarOpciones()
    
    Set mo_DOFactOrdenServicio = Nothing
    Set mo_DOAtencion = Nothing
    Set mo_DOComprobantePago = Nothing
      
    mo_DOFactOrdenServicio.idOrden = 0
    mo_DOFactOrdenBienInsumo.idOrden = 0
    
    mo_cmbIdTipoGenHistoriaClinica.BoundText = 0
   
    ml_idPaciente = 0
    ml_IdFormaPago = 1          'Contado
    ml_IdFarmacia = 0           '1=Farmacia Principal,2=Farmacia Emergencia,0-otros
    ml_idPreVenta = 0
    ml_idCuentaAtencion = 0
    ml_IdOrdenDespacho = 0
End Sub

Private Sub btnLimpiar_Click()
     txtNroSerieBusqueda = ""
     txtNroDocumentoBusqueda = ""
     txtFhasta.Text = Date & " 23:59"
     txtFdesde.Text = Date & " 00:01"
     TxtRsocial.Text = ""
End Sub

Private Sub grdGestionCaja_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
End Sub

Private Sub grdGestionCaja_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstadoComprobante").GetText()) = 9 Then
            Row.Appearance.ForeColor = vbRed
            'Row.Appearance.Font.Strikethrough = True
        End If
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFdesde
End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFhasta
End Sub

Private Sub txtNroDocumentoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumentoBusqueda
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroSerieBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroSerieBusqueda
   AdministrarKeyPreview KeyCode
End Sub

Private Sub TxtRsocial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, TxtRsocial
    AdministrarKeyPreview KeyCode
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


Sub ConfigurarGrilla()
    
    grdGestionCaja.Bands(0).Columns("Turno").Width = 800      '1200
    grdGestionCaja.Bands(0).Columns("Fecha").Width = 1600
End Sub

Public Sub RealizarBusqueda()

    Dim lcFechaIni As Date: Dim lcFechaFin As Date
    Dim lnTotalRecaudado As Double
    lcFechaIni = CDate(txtFdesde.Text)
    lcFechaFin = CDate(txtFhasta.Text)
    Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento("", "", lcFechaIni, lcFechaFin)
    ms_MensajeError = ""
    If txtNroSerieBusqueda.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "NroSerie='" & Trim(txtNroSerieBusqueda.Text) & "' and NroDocumento='" & txtNroDocumentoBusqueda.Text & "' and "
    ElseIf TxtRsocial.Text <> "" Then
       ms_MensajeError = ms_MensajeError & "RazonSocial like '%" & Trim(TxtRsocial.Text) & "%' and "
       txtNroSerieBusqueda.Text = ""
       txtNroDocumentoBusqueda.Text = ""
    End If
    If ms_MensajeError <> "" Then
       ms_MensajeError = Left(ms_MensajeError, Len(ms_MensajeError) - 5)
       oRsBusquedaRecibos.Filter = ms_MensajeError
    End If
    If oRsBusquedaRecibos.RecordCount > 0 Then
        lnTotalRecaudado = 0
        oRsBusquedaRecibos.MoveFirst
        Do While Not oRsBusquedaRecibos.EOF
           If oRsBusquedaRecibos.Fields!idEstadoComprobante = 4 Then
              lnTotalRecaudado = lnTotalRecaudado + oRsBusquedaRecibos.Fields!Total
           End If
           oRsBusquedaRecibos.MoveNext
        Loop
    End If
    Set grdGestionCaja.DataSource = oRsBusquedaRecibos
    ConfigurarGrilla
    mo_Apariencia.ConfigurarFilasBiColores grdGestionCaja, sighentidades.GrillaConFilasBicolor
End Sub

'********************************************************************************************
'********************************************************************************************
'********************************************************************************************
'                                   COMPROBANTE DE PAGO
'********************************************************************************************
'********************************************************************************************
'********************************************************************************************

Private Sub Form_Load()
  txtFdesde.Text = Date & " 00:01"
  txtFhasta.Text = Date & " 23:59"
End Sub


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FarmPaquetes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "farmPaquetes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SighFarmacia.ucPaquetes grdProductos 
      Height          =   5025
      Left            =   75
      TabIndex        =   24
      Top             =   2670
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   8864
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   60
      TabIndex        =   18
      Top             =   7770
      Width           =   15120
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "farmPaquetes.frx":0CCA
         DownPicture     =   "farmPaquetes.frx":112A
         Height          =   700
         Left            =   6173
         Picture         =   "farmPaquetes.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "farmPaquetes.frx":1A14
         DownPicture     =   "farmPaquetes.frx":1ED8
         Height          =   700
         Left            =   7703
         Picture         =   "farmPaquetes.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   120
         Picture         =   "farmPaquetes.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.Frame fraCabecera 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   15105
      Begin VB.ComboBox cmbAlmOrigen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1500
         TabIndex        =   22
         Top             =   690
         Width           =   5340
      End
      Begin VB.TextBox txtNotaSalida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   8
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   13470
         MaxLength       =   30
         TabIndex        =   7
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox txtNdocum 
         Height          =   315
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1560
         Width           =   1635
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   315
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1980
         Width           =   5325
      End
      Begin VB.ComboBox cmbConcepto 
         Height          =   330
         Left            =   1500
         TabIndex        =   0
         Top             =   1140
         Width           =   5340
      End
      Begin VB.ComboBox cmbTipoDocum 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1500
         TabIndex        =   6
         Top             =   1560
         Width           =   5340
      End
      Begin VB.ComboBox cmbAlmDestino 
         Height          =   330
         Left            =   8910
         TabIndex        =   1
         Top             =   1140
         Width           =   5970
      End
      Begin VB.TextBox txtHoraRegistro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10380
         MaxLength       =   30
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   8910
         TabIndex        =   9
         Top             =   300
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén origen"
         Height          =   210
         Left            =   150
         TabIndex        =   23
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   150
         TabIndex        =   17
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         Height          =   210
         Left            =   8040
         TabIndex        =   16
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Nota Ingreso"
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   210
         Left            =   12840
         TabIndex        =   14
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   210
         Left            =   8235
         TabIndex        =   13
         Top             =   1185
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Docum"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "N° Docum"
         Height          =   210
         Left            =   8010
         TabIndex        =   11
         Top             =   1590
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   2010
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FarmPaquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Notas de Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim mo_cmbConceptos As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenDestino As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoDocum As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oRsConceptos As New ADODB.Recordset
Dim oRsAlmacenOrigen As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mRs_Productos As New ADODB.Recordset
Dim mRs_ProductosLotes As New ADODB.Recordset

Dim mo_farmMovimiento As New sighComun.DoFarmMovimiento
Const lcConstanteMovimientoSalida As String = "S"
Const lcConstanteMovimientoEntrada As String = "E"
Dim lnTotalDocumento As Double
Dim ms_MensajeError As String
Dim mo_farmMovimientoNotaIngreso As New sighComun.DOfarmMovimientoNotaIngreso
Dim oDoProveedores As New DoProveedores
Dim lcTipoLocalesAlmOrigen As String
Dim lbDocumentoEsAutomatico As Boolean
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcTipoLocalesAlmDestino As String
Dim mo_lbElEstablecimentoEsCS As Boolean
Dim ml_idUsuarioCreo As Long
Dim lcMovNumeroSalida As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property







Private Sub ImprimeDocumento()
    Dim oRptClase As New rCrystal
    Dim oDOfarmAlmacen As New DoFarmAlmacen
    Set oDOfarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(Val(mo_cmbAlmacenOrigen.BoundText))
    oRptClase.MovTipo = "E"
    oRptClase.Documento = Me.txtNotaSalida.Text
    oRptClase.TextoDelFiltro = "NOTA DE INGRESO"
    oRptClase.Almacen = cmbAlmDestino.Text
    oRptClase.AlmacenO = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmOrigen.Text
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = Trim(cmbTipoDocum.Text) & " - " & txtNdocum.Text
    oRptClase.Importe = lnTotalDocumento
    oRptClase.TipoReporte = "NiNs"
    oRptClase.Observaciones = Trim(Me.txtObservaciones.Text) & "  (" & Label1.Caption & ":  " & cmbConcepto.Text & ")"   'debb-07/10/2016
    oRptClase.EsUnaDonacion = IIf(mo_cmbConceptos.BoundText = "3", True, False)
    'If Trim(cmbTipoDocum.Text) <> "" Then
    '    oRptClase.Proveedor = Trim(cmbTipoDocum.Text) & "/" & Trim(txtNdocum.Text)
    'End If
    oRptClase.idUsuario = ml_idUsuarioCreo
    oRptClase.Show vbModal
    Set oRptClase = Nothing
    Set oDOfarmAlmacen = Nothing
End Sub

Private Sub btnImprimir_Click()
   ImprimeDocumento
End Sub



Private Sub cmbAlmDestino_Click()
    '** solo en caso de donaciones
    If mo_cmbConceptos.BoundText = "3" Then
        oRsConceptos.MoveFirst
        oRsConceptos.Find "idTipoConcepto=" & mo_cmbConceptos.BoundText
        mo_cmbTipoDocum.BoundText = oRsConceptos.Fields!DocumentoId
        lcTipoLocalesAlmDestino = ""
        Dim oRsTmp As New Recordset
        Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idAlmacen=" & mo_cmbAlmacenDestino.BoundText)
        If oRsTmp.RecordCount > 0 Then
           lcTipoLocalesAlmDestino = oRsTmp.Fields!idTipoLocales
           If oRsTmp.Fields!idTipoLocales = "F" Then
              mo_cmbTipoDocum.BoundText = "15" 'ppa
              Me.txtNdocum.Text = ""
           End If
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
    End If
End Sub

Private Sub cmbAlmDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmDestino

End Sub

Private Sub cmbAlmOrigen_Click()
    oRsAlmacenOrigen.MoveFirst
    oRsAlmacenOrigen.Find "idAlmacen=" & mo_cmbAlmacenOrigen.BoundText
    Set oRsConceptos = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenOrigen.Fields!idTipoLocales, lcConstanteMovimientoSalida, oRsAlmacenOrigen.Fields!idTipoSuministro)
    mo_cmbConceptos.BoundColumn = "IdTipoConcepto"
    mo_cmbConceptos.ListField = "Concepto"
    Set mo_cmbConceptos.RowSource = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenOrigen.Fields!idTipoLocales, lcConstanteMovimientoSalida, oRsAlmacenOrigen.Fields!idTipoSuministro)
    grdProductos.IdAlmacen = oRsAlmacenOrigen.Fields!IdAlmacen
    lcTipoLocalesAlmOrigen = oRsAlmacenOrigen.Fields!idTipoLocales
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    mo_cmbConceptos.BoundText = "20"
    mo_Formulario.HabilitarDeshabilitar Me.cmbConcepto, False
    
End Sub


Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen

End Sub

Private Sub cmbConcepto_Click()
    If Val(mo_cmbConceptos.BoundText) = 0 Then
       Exit Sub
    End If
    oRsConceptos.MoveFirst
    oRsConceptos.Find "idTipoConcepto=" & mo_cmbConceptos.BoundText
    mo_cmbTipoDocum.BoundText = oRsConceptos.Fields!DocumentoId
    '
    mo_cmbAlmacenDestino.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenDestino.ListField = "Descripcion"
    If mo_lbElEstablecimentoEsCS = True Then
       Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro(oRsConceptos.Fields!NsFiltroAlmacenDestinoCS & " and idEstado=1")
    Else
       Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro(oRsConceptos.Fields!NsFiltroAlmacenDestino & " and idEstado=1")
    End If
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If cmbAlmDestino.ListCount = 1 Then
       cmbAlmDestino.ListIndex = 0
    End If
    '
   ' grdProductos.MuestraLoteParaDespachoNS = IIf(oRsConceptos.Fields!MuestraLoteParaDespachoNS = "S", True, False)
    grdProductos.TipoPrecioParaNiNs = 3   'oRsConceptos.Fields!TipoPrecioParaNiNs
    '
    lbDocumentoEsAutomatico = IIf(oRsConceptos.Fields!DocumentoEsAutomatico = "S", True, False)
    If lbDocumentoEsAutomatico = True Then
      'SCCQ 09/10/2020 Cambio28 Inicio
       'txtNdocum.Text = Val(oRsConceptos.Fields!DocumentoUltimoNumero) + 1
       'SCCQ 09/10/2020 Cambio28 Fin
    Else
       txtNdocum.Text = ""
    End If
    
End Sub

Private Sub cmbConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConcepto

End Sub

Private Sub Form_Activate()
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(mo_farmMovimiento.IdAlmacenOrigen) = True Then
        btnCancelar_Click
        Exit Sub
   End If

End Sub

Private Sub Form_Initialize()
    Set mo_cmbConceptos.MiComboBox = cmbConcepto
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbAlmacenDestino.MiComboBox = cmbAlmDestino
    Set mo_cmbTipoDocum.MiComboBox = cmbTipoDocum

End Sub

Private Sub Form_Load()
    mo_lbElEstablecimentoEsCS = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
    ConfigurarGrdProductos
    CargarComboBoxes
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Paquete"
    Case sghModificar
        Me.Caption = "Modificar Paquete"
    Case sghConsultar
        Me.Caption = "Consultar Paquete"
        btnImprimir.Visible = True
    Case sghEliminar
        Me.Caption = "Anular Paquete"
    End Select
    CargarDatosAlFormulario
End Sub
Sub ConfigurarGrdProductos()
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.FechaMinimaDespacho = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + Val(lcBuscaParametro.SeleccionaFilaParametro(220))
    grdProductos.inicializar
End Sub


Sub CargarComboBoxes()
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    '
    'Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    If mo_lnIdTablaLISTBARITEMS <> 1305 Then
       Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
       Label4.Caption = "Farmacia origen"
    Else
       Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1")
       Label4.Caption = "Almacén origen"
    End If
    '
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    'Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    If mo_lnIdTablaLISTBARITEMS <> 1305 Then
       Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
    Else
       Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1")
    End If
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
'    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacenOrigen.BoundText = "999"
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
       oRsAlmacenOrigen.MoveFirst
       oRsAlmacenOrigen.Find "idAlmacen=999"
       lcTipoLocalesAlmOrigen = oRsAlmacenOrigen.Fields!idTipoLocales
 '   End If
   '
    mo_cmbTipoDocum.BoundColumn = "idTipoDocumento"
    mo_cmbTipoDocum.ListField = "Nombre"
    Set mo_cmbTipoDocum.RowSource = mo_ReglasFarmacia.FarmTipoDocumentosDevuelveTodos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    mo_cmbTipoDocum.BoundText = "10"
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError
    End If
End Sub
Sub CargarDatosAlFormulario()
'SCCQ 14/10/2020 Cambio28 Inicio
mo_Formulario.HabilitarDeshabilitar txtNdocum, False
'SCCQ 14/10/2020 Cambio28 Fin
    mo_Formulario.HabilitarDeshabilitar Me.txtNotaSalida, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraRegistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocum, False
  

     Select Case mi_Opcion
     Case sghAgregar
        txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL      'Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
        txtHoraRegistro.Text = lcBuscaParametro.RetornaHoraServidorSQL
        grdProductos.movNumero = ""
        grdProductos.LimpiarGrilla
        grdProductos.CargaProductosPorMovNumero
        grdProductos.AgregaRegistro
        
     Case sghModificar
        DeshabilitaCabecera
        CargarDatosALosControles
     Case sghConsultar
        DeshabilitaCabecera
        CargarDatosALosControles
        btnAceptar.Enabled = False
     Case sghEliminar
        DeshabilitaCabecera
        CargarDatosALosControles
 End Select
End Sub

Sub CargarDatosALosControles()
 '**************Datos de la tabla FarmMovimiento *****************
   
   Dim oRsTmp1 As New Recordset
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 900
   oConexion.CursorLocation = adUseClient
   oConexion.Open SIGHEntidades.CadenaConexion
   mo_farmMovimiento.movNumero = ml_movNumero
   mo_farmMovimiento.MovTipo = lcConstanteMovimientoEntrada
   If Not mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_farmMovimiento) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   txtNotaSalida.Text = ml_movNumero
   'mo_cmbAlmacenOrigen.BoundText = mo_farmMovimiento.IdAlmacenOrigen
   mo_cmbConceptos.BoundText = mo_farmMovimiento.idTipoConcepto
   mo_cmbAlmacenDestino.BoundText = mo_farmMovimiento.IdAlmacenDestino
   mo_cmbTipoDocum.BoundText = mo_farmMovimiento.DocumentoIdtipo
   txtNdocum.Text = mo_farmMovimiento.DocumentoNumero
   txtObservaciones.Text = mo_farmMovimiento.Observaciones
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDelMovimiento("idEstadoMovimiento=" & mo_farmMovimiento.idEstadoMovimiento)
   txtFregistro.Text = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   txtHoraRegistro.Text = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
   ml_idUsuarioCreo = mo_farmMovimiento.idUsuario
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.movNumero = ml_movNumero
   grdProductos.CargaProductosPorMovNumero
   '
   lcMovNumeroSalida = ""
   Set oRsTmp1 = mo_ReglasFarmacia.FarmaciaFiltraTodosMovimientos(" and dbo.farmMovimiento.idAlmacenOrigen=" & _
                                                    mo_cmbAlmacenOrigen.BoundText & _
                                                    " and dbo.farmMovimiento.MovTipo='S' " & _
                                                    " and dbo.farmMovimiento.documentoIdTipo=" & mo_cmbTipoDocum.BoundText & _
                                                    " and dbo.farmMovimiento.idTipoConcepto=" & mo_cmbConceptos.BoundText & _
                                                    " and dbo.farmMovimiento.documentoNumero='" & txtNdocum.Text & "'", oConexion)
   If oRsTmp1.RecordCount > 0 Then
      lcMovNumeroSalida = oRsTmp1!movNumero
      grdProductos.CargaProductosPorLotes lcMovNumeroSalida
   End If
   '
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   If mo_farmMovimiento.idEstadoMovimiento = 0 Then
      btnAceptar.Enabled = False
   End If
   '******permiso a Modificar documento con Fecha Anterior a la actual
   Dim mo_PermisosFacturacion As New PermisosFacturacion
   Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
   Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
   If mo_PermisosFacturacion.ActualizaFechaDocumentoES = False And mi_Opcion <> sghConsultar Then
      If CDate(lcBuscaParametro.RetornaFechaServidorSQL) <> CDate(txtFregistro.Text) Then
         MsgBox "No tiene ACCESO a Modificar/Anular una NS" & Chr(13) & " de una Fecha Registro diferente a la actual", vbExclamation, Me.Caption
         btnAceptar.Enabled = False
      End If
   End If
   '
   oConexion.Close
   Set mo_PermisosFacturacion = Nothing
   Set mo_ReglasSeguridad = Nothing
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   
   grdProductos.MuestraItemsDelPrimerPaquete
End Sub

Sub DeshabilitaCabecera()
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmDestino, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocum, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbConcepto, False
  
End Sub
Private Sub btnCancelar_Click()
     Me.Visible = False
     LimpiarVariablesDeMemoria
End Sub
Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenOrigen.BoundText)) = True Then
      btnCancelar_Click
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
        'SCCQ 14/10/2020 Cambio28 Inicio
        'Antes:  If ValidarDatosObligatorios() Then
        If ValidarDatosObligatorios("A") Then
        'SCCQ 14/10/2020 Cambio28 Fin
            CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
'                If MsgBox("Se agregó correctamente la Nota de Salida N° " + txtNotaSalida.Text + Chr(13) + Chr(13) + "Desea Imprimir el Documento ?", vbQuestion + vbYesNo, "") = vbYes Then
'                   ml_idUsuarioCreo = ml_idUsuario
'                   ImprimeDocumento
'                End If
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   Case sghModificar
        MsgBox "Hay problemas al MODIFICAR,        debe ANULAR y luego AGREGAR", vbInformation, Me.Caption
        Exit Sub
        'SCCQ 14/10/2020 Cambio28 Inicio
        'Antes:  If ValidarDatosObligatorios() Then
        If ValidarDatosObligatorios("M") Then
        'SCCQ 14/10/2020 Cambio28 Fin
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
                If MsgBox("Se Modificó correctamente la Nota de Salida N° " + txtNotaSalida.Text + Chr(13) + Chr(13) + "Desea Imprimir el Documento ?", vbQuestion + vbYesNo, "") = vbYes Then
                   ml_idUsuarioCreo = ml_idUsuario
                   ImprimeDocumento
                End If
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If AnularNS() Then
                MsgBox " Se anuló la Nota de Salida N° " + txtNotaSalida.Text, vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub
'SCCQ 14/10/2020 Cambio28 Inicio
'Antes:  Function ValidarDatosObligatorios() As Boolean
Function ValidarDatosObligatorios(modo As String) As Boolean
'SCCQ 14/10/2020 Cambio28 Fin
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If cmbAlmOrigen.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
   ElseIf cmbConcepto.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Concepto" + Chr(13)
       cmbConcepto.SetFocus
   ElseIf cmbAlmDestino.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Destino" + Chr(13)
       cmbAlmDestino.SetFocus
   ElseIf mo_cmbAlmacenOrigen.BoundText = mo_cmbAlmacenDestino.BoundText Then
       ms_MensajeError = ms_MensajeError + "El Almacén Origen y Destino deben ser DIFERENTES" + Chr(13)
   ElseIf cmbTipoDocum.Text <> "" Then
   'SCCQ 14/10/2020 Cambio28 Inicio
        If modo = "M" Then 'Modifica
   'SCCQ 14/10/2020 Cambio28 Fin
            If txtNdocum.Text = "" Then
              ' ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° Documento" + Chr(13)
               'txtNdocum.SetFocus
            End If
   'SCCQ 14/10/2020 Cambio28 Inicio
        End If
   'SCCQ 14/10/2020 Cambio28 Fin
   End If
   'SCCQ 14/10/2020 Cambio28 Inicio
'   If mi_Opcion = sghAgregar And txtNdocum.Text <> "" Then
'      Dim oRsTmp As New ADODB.Recordset
'      Set oRsTmp = mo_ReglasFarmacia.farmMovimientoSeleccionarPorTipoYnumeroDocumento(txtNdocum.Text, Val(mo_cmbTipoDocum.BoundText))
'      oRsTmp.Filter = "idEstadoMovimiento=1"
'      If oRsTmp.RecordCount > 0 Then
'         ms_MensajeError = ms_MensajeError + "El Número de Documento: " & txtNdocum.Text & "   EXISTE en NS: " & Trim(oRsTmp.Fields!movNumero) & "     Fecha: " & oRsTmp.Fields!fechaCreacion & Chr(13)
'      End If
'      oRsTmp.Close
'      Set oRsTmp = Nothing
'   End If
   'SCCQ 14/10/2020 Cambio28 Fin
   lnTotalDocumento = grdProductos.DevuelveTotal
   Set mRs_Productos = grdProductos.DevuelveProductos
   Set mRs_ProductosLotes = grdProductos.DevuelveProductosLotes
   mRs_ProductosLotes.Filter = "lote=''"
   If mRs_ProductosLotes.RecordCount > 0 Then
      ms_MensajeError = ms_MensajeError + "Existen ITEMS del PAQUETE que no tienen LOTES" + Chr(13)
   End If
   mRs_ProductosLotes.Filter = ""
   If mRs_Productos.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        Dim LdFechaMinimaDespacho As Date, lnIdProductoPaquete As Long, lnIdProducto As Long, lnCantidad1 As Long
        Dim lnCantidad2 As Long, lcItem As String
        If mo_cmbConceptos.BoundText = "5" Then
           LdFechaMinimaDespacho = Date - 1000  'Devolucion por Vencimiento
        Else
           LdFechaMinimaDespacho = CDate(txtFregistro.Text) + Val(lcBuscaParametro.SeleccionaFilaParametro(220))
        End If
        mRs_Productos.MoveFirst
        Do While Not mRs_Productos.EOF
           If Trim(mRs_Productos.Fields!codigo) = "" Or Trim(mRs_Productos.Fields!nombreProducto) = "" Then
              mRs_Productos.Delete
              mRs_Productos.Update
           ElseIf mRs_Productos.Fields!Cantidad <= 0 Then
              ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas" + Chr(13)
           ElseIf mRs_Productos!Precio <= 0 Then
               ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas con el Precio" + Chr(13)
           ElseIf mRs_Productos!FechaVencimiento < LdFechaMinimaDespacho Then
               ms_MensajeError = ms_MensajeError + "La F.Vencimiento mínima de despacho es: " & LdFechaMinimaDespacho & " para " & Trim(mRs_Productos.Fields!codigo) & " - " & Trim(mRs_Productos.Fields!nombreProducto) & Chr(13)
           Else
               mRs_ProductosLotes.Sort = "idProductoPaquete,IdProducto"
               mRs_ProductosLotes.MoveFirst
               Do While Not mRs_ProductosLotes.EOF
                  lnIdProductoPaquete = mRs_ProductosLotes!IdProductoPaquete
                  lnIdProducto = mRs_ProductosLotes!idProducto
                  lnCantidad2 = mRs_ProductosLotes!cantidadPqte * mRs_Productos!Cantidad
                  lnCantidad1 = 0
                  lcItem = mRs_ProductosLotes!codigo & " " & mRs_ProductosLotes!nombreProducto
                  Do While Not mRs_ProductosLotes.EOF And lnIdProductoPaquete = mRs_ProductosLotes!IdProductoPaquete And lnIdProducto = mRs_ProductosLotes!idProducto
                     lnCantidad1 = lnCantidad1 + mRs_ProductosLotes!Cantidad
                     mRs_ProductosLotes.MoveNext
                     If mRs_ProductosLotes.EOF Then
                        Exit Do
                     End If
                  Loop
                  If lnCantidad2 <> lnCantidad1 Then
                     ms_MensajeError = ms_MensajeError + "Problemas de SALDOS para " & lcItem & Chr(13)
                  End If
               Loop
           End If
           mRs_Productos.MoveNext
        Loop
   End If
   mRs_ProductosLotes.Filter = ""
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With mo_farmMovimiento
            .DocumentoIdtipo = Val(mo_cmbTipoDocum.BoundText)                   '10
            .DocumentoNumero = txtNdocum.Text                                   'inven-2014
            .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL       'igual
            .IdAlmacenDestino = Val(mo_cmbAlmacenDestino.BoundText)             '0
            .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)               '8
            .idEstadoMovimiento = sghEstadoTabla.sghRegistrado                  'igual
            .idTipoConcepto = Val(mo_cmbConceptos.BoundText)                    '20
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoSalida
            .Observaciones = txtObservaciones.Text
            .Total = lnTotalDocumento
            
        End With
   Case sghModificar
        With mo_farmMovimiento
            .DocumentoNumero = txtNdocum.Text
            .Observaciones = txtObservaciones.Text
            .IdUsuarioAuditoria = ml_idUsuario
            .Total = lnTotalDocumento
            '.FechaCreacion = txtFregistro.Text
        End With
   Case sghEliminar
        With mo_farmMovimiento
            .fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .idEstadoMovimiento = sghEstadoTabla.sghAnulado    'Anulado
            .IdUsuarioAuditoria = ml_idUsuario
        End With
   End Select
End Sub
Function AgregarDatos() As Boolean
    Dim lbAgregarDatos As Boolean
    '*********  graba tabla RELMOD  ***************
    'If lbDocumentoEsAutomatico = True Then
'       Dim oRsTmp As New ADODB.Recordset
'       Dim lcFiltro As String
'       lcFiltro = "tipoAlmacen='" & oRsAlmacenOrigen.Fields!idTipoLocales & "' and tipoMov='S' and tipoSuministro='" & oRsAlmacenOrigen.Fields!idTipoSuministro & "' and DocumentoId=" & mo_cmbTipoDocum.BoundText
'       Set oRsTmp = mo_ReglasFarmacia.FarmRelModDevuelveSegunFiltro(lcFiltro)
'       If oRsTmp.RecordCount = 0 Then
'          AgregarDatos = False
'       Else
'          mo_ReglasFarmacia.FarmRelModActualizaSegunFiltro lcFiltro, txtNdocum.Text
'       End If
'       oRsTmp.Close
'       Set oRsTmp = Nothing
   ' End If
    '
    'SCCQ 21/10/2020 Cambio28 Inicio
     If lbDocumentoEsAutomatico = True Then
        lbAgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaSalida_NumDocAutomatico(oRsAlmacenOrigen.Fields!idTipoLocales, oRsAlmacenOrigen.Fields!idTipoSuministro, CLng(mo_cmbTipoDocum.BoundText), mo_farmMovimiento, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        txtNdocum.Text = mo_farmMovimiento.DocumentoNumero
     Else
    'SCCQ 21/10/2020 Cambio28 Fin
     lbAgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaSalida(mo_farmMovimiento, mRs_ProductosLotes, mo_lnIdTablaLISTBARITEMS, _
                                                               mo_lcNombrePc)
     'SCCQ 21/10/2020 Cambio28 Inicio
     End If
     'SCCQ 21/10/2020 Cambio28 Fin
    txtNotaSalida.Text = mo_farmMovimiento.movNumero
    If lbAgregarDatos = True Then
    'If GeneraNIenFormaAutomatica(lbAgregarDatos) Then
        With mo_farmMovimiento
            .MovTipo = lcConstanteMovimientoEntrada
            .IdAlmacenDestino = Val(mo_cmbAlmacenOrigen.BoundText)
            .IdAlmacenOrigen = 0
        End With
        With mo_farmMovimientoNotaIngreso
            .MovTipo = lcConstanteMovimientoEntrada
            .DocumentoFechaRecepcion = mo_farmMovimiento.fechaCreacion
            
        End With
        With oDoProveedores
        End With
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, _
                                                                  oDoProveedores, mRs_Productos, 0, mo_lnIdTablaLISTBARITEMS, _
                                                                  mo_lcNombrePc)
        MsgBox "Se creó Nota de Salida en forma automática", vbInformation, Me.Caption
    End If
    
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
    AgregarDatos = lbAgregarDatos
End Function
Function ModificarDatos() As Boolean
    
    
    Dim lbModificarDatos As Boolean
    Dim lbModificarDatosNI As Boolean
    Dim lnTotal As Double
    lbModificarDatos = mo_ReglasFarmacia.ModificaDatosDeNotaSalida(mo_farmMovimiento, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    If GeneraNIenFormaAutomatica(lbModificarDatos) Then
        Dim oRsTmp As New Recordset
        Dim oConexion As New ADODB.Connection
        Dim oMovimiento As New farmMovimiento
        Dim oMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
        '
        oConexion.Open SIGHEntidades.CadenaConexion
        Set oMovimiento.Conexion = oConexion
        Set oMovimientoNotaIngreso.Conexion = oConexion
        '
        lnTotal = mo_farmMovimiento.Total
        Set oRsTmp = mo_ReglasFarmacia.farmMovimientoSeleccionarPorTipoYnumeroDocumento(mo_farmMovimiento.DocumentoNumero, mo_farmMovimiento.DocumentoIdtipo)
        oRsTmp.Filter = "movTipo='E' and idAlmacenDestino=" & mo_farmMovimiento.IdAlmacenDestino
        If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            '
            With mo_farmMovimiento
                .MovTipo = lcConstanteMovimientoEntrada
                .movNumero = oRsTmp.Fields!movNumero
            End With
            If Not oMovimiento.SeleccionarPorId(mo_farmMovimiento) Then
               MsgBox "Fallo en Nota de Ingreso automática" & Chr(13) & oMovimiento.MensajeError
               Exit Function
            End If
            mo_farmMovimiento.Total = lnTotal
            '
            With mo_farmMovimientoNotaIngreso
                .MovTipo = lcConstanteMovimientoEntrada
                .movNumero = mo_farmMovimiento.movNumero
            End With
            If Not oMovimientoNotaIngreso.SeleccionarPorId(mo_farmMovimientoNotaIngreso) Then
               MsgBox "Fallo en Nota de Ingreso automática" & Chr(13) & oMovimientoNotaIngreso.MensajeError
               Exit Function
            End If
            With oDoProveedores
            End With
            lbModificarDatosNI = mo_ReglasFarmacia.ModificaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
            MsgBox "Se actualizó Nota de Ingreso en forma automática", vbInformation, Me.Caption
        End If
        Set oRsTmp = Nothing
        Set oConexion = Nothing
        Set oMovimiento = Nothing
        Set oMovimientoNotaIngreso = Nothing
    Else
        ms_MensajeError = mo_ReglasFarmacia.MensajeError
    End If
    ModificarDatos = lbModificarDatos
End Function
Function AnularNS() As Boolean
     AnularNS = mo_ReglasFarmacia.AnulaNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, 0, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
     If AnularNS = True Then
            With mo_farmMovimiento
                .MovTipo = lcConstanteMovimientoSalida
                .movNumero = lcMovNumeroSalida
                .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
            End With
            AnularNS = mo_ReglasFarmacia.AnulaNotaSalida(mo_farmMovimiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, 0)
     End If

End Function






Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtNdocum_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNdocum

End Sub

Private Sub txtNdocum_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
    End If

End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservaciones

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

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbConceptos = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbAlmacenDestino = Nothing
    Set mo_cmbTipoDocum = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oRsConceptos = Nothing
    Set oRsAlmacenOrigen = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mo_farmMovimiento = Nothing
    Set mo_farmMovimientoNotaIngreso = Nothing
    Set oDoProveedores = Nothing
End Sub

'*****Genera NI en forma automática para:
'*****DISTRIBUCION del ALMACEN ESPECIALIZADO: crea automaticamente NI hacia alguna Farmacia
'*****DEVOLUCIONES de la FARMACIA: crea automaticamente NI hacia el ALMACEN ESPECIALIZADO
'*****DISTRIBUCION de la FARMACIA: crea automaticamente hacia alguna farmacia
'*****DONACIONES del ALMACEN ESPECIALIZADO: hacia una de las Farmacias
Function GeneraNIenFormaAutomatica(lbRealizoMantenimiento As Boolean) As Boolean
         GeneraNIenFormaAutomatica = False
         If lbRealizoMantenimiento = True And (Val(mo_cmbConceptos.BoundText) = 4 And lcTipoLocalesAlmOrigen = "A") Or (Val(mo_cmbConceptos.BoundText) >= 4 And Val(mo_cmbConceptos.BoundText) <= 7 And lcTipoLocalesAlmOrigen = "F") Or (mo_cmbConceptos.BoundText = 3 And lcTipoLocalesAlmDestino = "F") Then
            GeneraNIenFormaAutomatica = True
         End If
 End Function

VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form RegistroComprobante 
   Caption         =   "Registro de Comprobante de Pago"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   7620
      TabIndex        =   34
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox cmbIdTipoComprobante 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   41
         Top             =   240
         Width           =   2625
      End
      Begin VB.TextBox txtNroSerie 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   36
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox txtNroDocumento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   35
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc:"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Cambio:"
         Height          =   195
         Left            =   1080
         TabIndex        =   39
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   38
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1980
         TabIndex        =   37
         Top             =   840
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo [F5]"
      DisabledPicture =   "RegistroComprobante.frx":0000
      DownPicture     =   "RegistroComprobante.frx":03E9
      Height          =   700
      Left            =   7560
      Picture         =   "RegistroComprobante.frx":07F5
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8160
      Width           =   1365
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir [F3]"
      Height          =   705
      Left            =   6240
      Picture         =   "RegistroComprobante.frx":0C01
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8160
      Width           =   1245
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Aceptar (F2)"
      DisabledPicture =   "RegistroComprobante.frx":10DA
      DownPicture     =   "RegistroComprobante.frx":153A
      Height          =   700
      Left            =   4800
      Picture         =   "RegistroComprobante.frx":19AF
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8160
      Width           =   1365
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "RegistroComprobante.frx":1E24
      DownPicture     =   "RegistroComprobante.frx":22E8
      Height          =   700
      Left            =   9000
      Picture         =   "RegistroComprobante.frx":27D4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resumen"
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   7080
      Width           =   11415
      Begin VB.TextBox txtMontoRecibidoSoles 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5400
         TabIndex        =   18
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblMontoVueltoSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8940
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblMontoFaltanteSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7140
         TabIndex        =   20
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblMontoFacturadoSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   16
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "SOLES"
         Height          =   255
         Left            =   2820
         TabIndex        =   25
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Vuelto (S/.)"
         Height          =   195
         Left            =   9300
         TabIndex        =   23
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Faltan (S/.)"
         Height          =   195
         Left            =   7500
         TabIndex        =   21
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Recibido (S/.)"
         Height          =   195
         Left            =   5640
         TabIndex        =   19
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Monto Facturado (S/.)"
         Height          =   195
         Left            =   3540
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame fraItems 
      Height          =   5445
      Left            =   60
      TabIndex        =   8
      Top             =   1620
      Width           =   11415
      Begin VB.CommandButton cmdQuitarItems 
         Caption         =   "Quitar[F7]"
         DisabledPicture =   "RegistroComprobante.frx":2CC0
         DownPicture     =   "RegistroComprobante.frx":304B
         Height          =   615
         Left            =   10320
         Picture         =   "RegistroComprobante.frx":33DE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   900
         Width           =   1005
      End
      Begin VB.CommandButton cmdAgregarItem 
         Caption         =   "Agregar[F6]"
         DisabledPicture =   "RegistroComprobante.frx":376F
         DownPicture     =   "RegistroComprobante.frx":3B58
         Height          =   615
         Left            =   10320
         Picture         =   "RegistroComprobante.frx":3F64
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1005
      End
      Begin UltraGrid.SSUltraGrid grdItems 
         Height          =   4665
         Left            =   60
         TabIndex        =   24
         Top             =   180
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   8229
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
         Caption         =   "Items Comprobante"
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         TabIndex        =   14
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar (S/.):"
         Height          =   195
         Left            =   7740
         TabIndex        =   13
         Top             =   5040
         Width           =   1365
      End
      Begin VB.Label lblIGV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         TabIndex        =   12
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblLabelIGV 
         AutoSize        =   -1  'True
         Caption         =   "IGV 0%"
         Height          =   195
         Left            =   5640
         TabIndex        =   11
         Top             =   5040
         Width           =   825
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFEBD9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4260
         TabIndex        =   10
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sub Total (S/.):"
         Height          =   195
         Left            =   3120
         TabIndex        =   9
         Top             =   5040
         Width           =   1095
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
      Height          =   1635
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   7515
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
         Left            =   2820
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   540
         Width           =   4575
      End
      Begin VB.TextBox txtIdNroHistoria 
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
         Left            =   1140
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   540
         Width           =   1320
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Medicamentos"
         Height          =   255
         Index           =   1
         Left            =   3900
         TabIndex        =   33
         Top             =   180
         Width           =   1395
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Servicios"
         Height          =   255
         Index           =   0
         Left            =   2220
         TabIndex        =   32
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtRUC 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   1260
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtRazonSocial 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   900
         Width           =   6225
      End
      Begin VB.CommandButton cmdCuentaAtencion 
         Caption         =   "..."
         Height          =   315
         Left            =   2460
         TabIndex        =   2
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Nº historia"
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
         Left            =   120
         TabIndex        =   0
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   990
      End
   End
End
Attribute VB_Name = "RegistroComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD-04 Todo el archivo

Option Explicit

Const ID_TIPO_MONEDA_SOLES = 1
Const ID_TIPO_MONEDA_DOLAR = 2

Const ID_TIPO_COMPROBANTE_FACTURA = 2

Dim ml_IdComprobantePago As Long
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean

Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision

Dim mo_CajaComprobantesPago As New DOCajaComprobantesPago
Dim mo_ItemsAFacturar As New Collection
Dim mo_ItemsDinero As New Collection

Dim mo_cmbIdTipoComprobante As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New ListaDespleglable

Dim mrs_ComprobantesDetalle As ADODB.Recordset
Attribute mrs_ComprobantesDetalle.VB_VarHelpID = -1
'Dim mrs_FormaPago As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic

Dim md_PorcentajeIGV  As Double
Dim md_PorcentajeIGVDefault As Double

Dim md_TipoCambioDolar As Double

Dim mo_CajaLoteActual As New DOCajaLote
Dim mo_CajaCajaActual As New DOCajaCaja

Dim ml_IdCuentaAtencionActual As Long

Dim bCalculandoSubTotales As Boolean
Property Set CajaLoteActual(oValue As DOCajaLote)
   Set mo_CajaLoteActual = oValue
   'Ubicamos la caja en función del Lote
   Set mo_CajaCajaActual = mo_AdminCaja.CajaSeleccionarPorId(mo_CajaLoteActual.IdCaja)
End Property
Property Get CajaLoteActual() As DOCajaLote
   Set CajaLoteActual = mo_CajaLoteActual
End Property

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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdComprobantePago(lValue As Long)
   ml_IdComprobantePago = lValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
    
    CargarDatosGenerales
        
    Select Case mi_Opcion
        Case sghAgregar
            NuevoComprobante
        Case sghModificar
            CargarDatosALosControles
        Case sghConsultar
            CargarDatosALosControles
        Case sghEliminar
            CargarDatosALosControles
    End Select
End Sub


Private Sub cmbIdTipoComprobante_Click()
    Dim oCajaNroDoc As New DOCajaNroDocumento
    oCajaNroDoc.IdCaja = CajaLoteActual.IdCaja
    oCajaNroDoc.IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
    
    If mo_AdminCaja.ObtenerSiguienteNumeroDocumento(oCajaNroDoc) Then
        txtNroSerie = oCajaNroDoc.NroSerie
        txtNroDocumento = oCajaNroDoc.NroDocumento
    Else
        txtNroSerie = ""
        txtNroDocumento = ""
    End If
End Sub

Private Sub cmbIdTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoComprobante
    AdministrarKeyPreview KeyCode
End Sub

'Private Sub cmdAgregarDinero_Click()
'    With mrs_FormaPago
'        .AddNew
'        If Me.grdDinero.ValueLists("TipoFormaPago").ValueListItems.Count > 0 Then
'            .Fields!IdTipoFormaPago = Me.grdDinero.ValueLists("TipoFormaPago").ValueListItems(0).DataValue
'        End If
'        If Me.grdDinero.ValueLists("TipoMoneda").ValueListItems.Count > 0 Then
'            .Fields!IdTipoMoneda = Me.grdDinero.ValueLists("TipoMoneda").ValueListItems(0).DataValue
'        End If
'        .Update
'    End With
'    grdDinero.SetFocus
'    CalcularSubTotalesDinero
'End Sub

Private Sub cmdAgregarItem_Click()
    Dim oRegDetalleComprobante As New RegistroDetalleComprobante
    oRegDetalleComprobante.Show vbModal
    'Agregamos el detalle seleccionado
    Dim oDetalleSeleccionado As DOCajaComprobantesDetalle
    Set oDetalleSeleccionado = oRegDetalleComprobante.GetComprobantesDetalleSeleccionado
    If oDetalleSeleccionado Is Nothing Then
        Exit Sub
    End If
    Dim bFound As Boolean
    bFound = False
    'Buscamos si ya existe el producto para sumar las cantidades
    If mrs_ComprobantesDetalle.EOF = False And mrs_ComprobantesDetalle.BOF = False Then
        Do While Not mrs_ComprobantesDetalle.EOF
            If Not IsNull(mrs_ComprobantesDetalle.Fields!IdProducto) Then
                If mrs_ComprobantesDetalle.Fields!IdProducto = oDetalleSeleccionado.IdProducto Then
                    mrs_ComprobantesDetalle.Fields!cantidad = mrs_ComprobantesDetalle.Fields!cantidad + oDetalleSeleccionado.cantidad
                    mrs_ComprobantesDetalle.Fields!SubTotalPagado = mrs_ComprobantesDetalle.Fields!cantidad * mrs_ComprobantesDetalle.Fields!precioUnitario
                    bFound = True
                    Exit Do
                End If
            End If
            mrs_ComprobantesDetalle.MoveNext
        Loop
    End If
    If Not bFound Then
        With mrs_ComprobantesDetalle
            .AddNew
            .Fields!IdComprobanteDetalle = Null
            .Fields!CheckSeleccionado = True
            .Fields!TipoDetalle = oDetalleSeleccionado.TipoDetalle
            .Fields!CodigoProducto = oDetalleSeleccionado.CodigoProducto
            .Fields!IdProducto = oDetalleSeleccionado.IdProducto
            .Fields!Producto = oDetalleSeleccionado.NombreProducto
            .Fields!cantidad = oDetalleSeleccionado.cantidad
            .Fields!precioUnitario = oDetalleSeleccionado.precioUnitario
            .Fields!SubTotalExonerado = 0
            .Fields!SubTotalPagado = Round(oDetalleSeleccionado.cantidad * oDetalleSeleccionado.precioUnitario, 2)
        End With
    End If
    CalcularSubTotalesItems
End Sub

Private Sub cmdCuentaAtencion_Click()
    Dim oFrm As New PacientesBusqueda
    oFrm.Caption = "Seleccionar el paciente para el cual se desea la cuenta"
    oFrm.TipoFiltro = sghFiltrarTodos
    oFrm.Show vbModal
    If oFrm.BotonPresionado = sghAceptar Then
        
        Dim oDOPaciente As doPaciente
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oFrm.IdRegistroSeleccionado)
        ObtenerCuentaAtencionPorHistoriaClinica oDOPaciente.NroHistoriaClinica
    End If
End Sub

Private Sub cmdNuevo_Click()
    NuevoComprobante
End Sub

'Private Sub cmdQuitarDinero_Click()
'    On Error Resume Next
'    With mrs_FormaPago
'        If Not .EOF And Not .BOF Then
'           .Delete
'           .Update
'        End If
'        .MoveFirst
'    End With
''    CalcularSubTotalesDinero
'End Sub

Private Sub cmdQuitarItems_Click()
    On Error Resume Next
    With mrs_ComprobantesDetalle
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
        .MoveFirst
    End With
    CalcularSubTotalesItems
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
    GenerarRecordsetTemporal
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Comprobante de Pago"
    Case sghModificar
        Me.Caption = "Modificar Comprobante de Pago"
    Case sghConsultar
        Me.Caption = "Consultar Comprobante de Pago"
    Case sghEliminar
        Me.Caption = "Eliminar Comprobante de Pago"
    End Select
    CargarComboBoxes
    CargarDatosAlFormulario
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    
    SeleccionarTipoDocumento
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CajaComprobantesPago
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
           cmdGrabar.Value = True
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdSalir.Value = True
        Case vbKeyF2
            Me.cmdGrabar.Value = True
        Case vbKeyF3
            Me.cmdImprimir.Value = True
         Case vbKeyF5
            Me.cmdNuevo.Value = True
         Case vbKeyF6
            Me.cmdAgregarItem.Value = True
         Case vbKeyF7
            Me.cmdQuitarItems.Value = True
    End Select
End Sub
Private Sub cmdGrabar_Click()
   
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    NuevoComprobante
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
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
Dim bFound As Boolean

    ValidarDatosObligatorios = False
    
    If Me.txtNroSerie.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nº de Serie" + Chr(13)
    End If
    If Me.txtNroDocumento.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nº de Documento" + Chr(13)
    End If
    If mo_cmbIdTipoComprobante.BoundText = "" Then
        sMensaje = sMensaje + "Ingrese el tipo de Comprobante" + Chr(13)
    End If
    If mo_cmbIdTipoComprobante.BoundText = ID_TIPO_COMPROBANTE_FACTURA Then
        If Me.txtRUC = "" Then
            sMensaje = sMensaje + "Ingrese el RUC para la Factura" + Chr(13)
        End If
    Else
         Me.txtRUC = ""
    End If
    If Trim(Me.txtRazonSocial.Text) = "" Then
        sMensaje = sMensaje + "Ingrese la Razón Social" + Chr(13)
    End If
    
    bFound = False
    If mrs_ComprobantesDetalle.EOF = False And mrs_ComprobantesDetalle.BOF = False Then
    mrs_ComprobantesDetalle.MoveFirst
    Do Until mrs_ComprobantesDetalle.EOF
        If mrs_ComprobantesDetalle.Fields!CheckSeleccionado Then
            bFound = True
            Exit Do
        End If
        mrs_ComprobantesDetalle.MoveNext
    Loop
    End If
    If Not bFound Then
        sMensaje = sMensaje + "Ingrese los Items a Facturar" + Chr(13)
    End If
    
    If Me.lblMontoFaltanteSoles <> "" Then
        If CCurrency(lblMontoFaltanteSoles) > 0 Then
            sMensaje = sMensaje + "No puede registrar un comprobante con faltante de dinero" + Chr(13)
        End If
    End If
    
    If sMensaje <> "" Then
         MsgBox sMensaje, vbExclamation, Me.Caption
         Exit Function
    End If
    
    ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean

    ValidarReglas = False
   
    If mi_Opcion = sghAgregar Then
    
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

    With mo_CajaComprobantesPago
        .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
        .NroSerie = Me.txtNroSerie
        .NroDocumento = Me.txtNroDocumento
        .IdCuentaAtencion = ml_IdCuentaAtencionActual
        .IdCuentaAtencion = 0
        .RazonSocial = Me.txtRazonSocial
        .Observaciones = ""
        .IdGestionCaja = Me.CajaLoteActual.IdLote
        .IdUsuarioAuditoria = ml_IdUsuario
        
        .subtotal = CCurrency(Me.lblSubTotal)
        .IGV = CCurrency(Me.lblIGV)
        .Total = CCurrency(Me.lblTotal)
    End With
    
    'mo_CajaCajaActual.NroSerie = mo_CajaComprobantesPago.NroSerie
    'mo_CajaCajaActual.NroComprobante = mo_CajaComprobantesPago.NroDocumento
    
    '------------------------------------
    'Cargamos los Items a Facturar
    '------------------------------------
    Set mo_ItemsAFacturar = New Collection
    Dim oItemFactura As DOCajaComprobantesDetalle
    Dim SubTotalPorPagar As Double
    If Not (mrs_ComprobantesDetalle.BOF And mrs_ComprobantesDetalle.EOF) Then
        mrs_ComprobantesDetalle.MoveFirst
        Do While Not mrs_ComprobantesDetalle.EOF
            If mrs_ComprobantesDetalle!CheckSeleccionado Then
                Set oItemFactura = New DOCajaComprobantesDetalle
                oItemFactura.IdFacturacionDetalle = IIf(IsNull(mrs_ComprobantesDetalle!IdFacturacionDetalle), 0, mrs_ComprobantesDetalle!IdFacturacionDetalle)
                oItemFactura.IdProducto = mrs_ComprobantesDetalle!IdProducto
                oItemFactura.TipoDetalle = mrs_ComprobantesDetalle!TipoDetalle
                oItemFactura.cantidad = mrs_ComprobantesDetalle!cantidad
                oItemFactura.precioUnitario = mrs_ComprobantesDetalle!precioUnitario
                oItemFactura.SubTotalPagado = mrs_ComprobantesDetalle!SubTotalPagado
                oItemFactura.SubTotalExonerado = mrs_ComprobantesDetalle!SubTotalExonerado
                oItemFactura.SubTotalPagadoACuenta = mrs_ComprobantesDetalle!SubTotalPagadoACuenta
                SubTotalPorPagar = Round(oItemFactura.cantidad * oItemFactura.precioUnitario, 2) - oItemFactura.SubTotalExonerado - oItemFactura.SubTotalPagadoACuenta
                If oItemFactura.SubTotalPagado < SubTotalPorPagar Then
                    oItemFactura.EsPagoACuenta = 1
                Else
                    oItemFactura.EsPagoACuenta = 0
                End If
                oItemFactura.IdUsuarioAuditoria = ml_IdUsuario
                
                mo_ItemsAFacturar.Add oItemFactura
            End If
            mrs_ComprobantesDetalle.MoveNext
        Loop
        mrs_ComprobantesDetalle.MoveFirst
    End If
    
    '------------------------------------
    'Cargamos los Items de Dinero
    '------------------------------------
    Set mo_ItemsDinero = New Collection
    Dim oItemDinero  As New DOCajaFormaPagoComprobante
    
    oItemDinero.IdTipoFormaPago = 1
    oItemDinero.IdTipoMoneda = 1
    oItemDinero.Importe = CCurrency(Me.txtMontoRecibidoSoles)
    oItemDinero.IdUsuarioAuditoria = ml_IdUsuario
    oItemDinero.TipoCambio = 1
    oItemDinero.TotalSoles = oItemDinero.Importe
    
    mo_ItemsDinero.Add oItemDinero
    
'    If Not (mrs_FormaPago.BOF And mrs_FormaPago.EOF) Then
'        mrs_FormaPago.MoveFirst
'        Do While Not mrs_FormaPago.EOF
'            Set oItemDinero = New DOCajaFormaPagoComprobante
'            oItemDinero.IdTipoFormaPago = mrs_FormaPago!IdTipoFormaPago
'            oItemDinero.IdTipoMoneda = mrs_FormaPago!IdTipoMoneda
'            oItemDinero.Importe = mrs_FormaPago!Importe
'            oItemDinero.IdUsuarioAuditoria = ml_IdUsuario
'            If oItemDinero.IdTipoMoneda = ID_TIPO_MONEDA_DOLAR Then
'                oItemDinero.TipoCambio = md_TipoCambioDolar
'            Else
'                oItemDinero.TipoCambio = 1
'            End If
'            oItemDinero.TotalSoles = IIf(IsNull(mrs_FormaPago!Importe), 0, mrs_FormaPago!Importe) * oItemDinero.TipoCambio
'
'            mo_ItemsDinero.Add oItemDinero
'            mrs_FormaPago.MoveNext
'        Loop
'    End If
    
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminCaja.ComprobantePagoAgregar(mo_CajaComprobantesPago, mo_ItemsAFacturar, mo_ItemsDinero, mo_CajaCajaActual)
    
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminCaja.ComprobantePagoModificar(mo_CajaComprobantesPago, mo_ItemsAFacturar, mo_ItemsDinero)
       
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = mo_AdminCaja.ComprobantePagoEliminar(mo_CajaComprobantesPago)
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
Dim oDOCajaComprobantesPago  As DOCajaComprobantesPago

       Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(Me.IdComprobantePago)
       
       If mo_AdminCaja.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbCritical, Me.Caption
            mb_ExistenDatos = False
            Exit Sub
       End If
       
       If Not oDOCajaComprobantesPago Is Nothing Then
            With oDOCajaComprobantesPago
                mo_cmbIdTipoComprobante.BoundText = .IdTipoComprobante
                txtNroSerie = .NroSerie
                txtNroDocumento = .NroDocumento
                txtRazonSocial = .RazonSocial
                'txtIdCuentaAtencion = IIf(.IdCuentaAtencion = 0, "", CStr(.IdCuentaAtencion))
                txtRUC = .RUC
                lblSubTotal = Format(.subtotal, "0.00")
                lblIGV = Format(.IGV, "0.00")
                lblTotal = Format(.Total, "0.00")
                
                Set mo_CajaComprobantesPago = oDOCajaComprobantesPago
                mb_ExistenDatos = True
            End With
            '-------------------------------------
            'Cargamos del Items de la Factura
            '-------------------------------------
            Dim rsDetalle As New Recordset
            Dim oDOCajaDetalle As New DOCajaComprobantesDetalle
            oDOCajaDetalle.IdComprobantePago = oDOCajaComprobantesPago.IdComprobantePago
            Set rsDetalle = mo_AdminCaja.CajaComprobantesDetalle(oDOCajaDetalle)
            Do While Not rsDetalle.EOF
                With mrs_ComprobantesDetalle
                    .AddNew
                    .Fields!CheckSeleccionado = True
                    .Fields!IdFacturacionDetalle = rsDetalle!IdFacturacionDetalle
                    .Fields!TipoDetalle = rsDetalle!TipoDetalle
                    .Fields!CodigoProducto = rsDetalle!Codigo
                    .Fields!IdProducto = rsDetalle!IdProducto
                    .Fields!Producto = rsDetalle!Producto
                    .Fields!cantidad = rsDetalle!cantidad
                    .Fields!precioUnitario = rsDetalle!precioUnitario
                    .Fields!SubTotalExonerado = rsDetalle!SubTotalExonerado
                    '.Fields!SubTotalPagadoACuenta = rsDetalle!SubTotalPagadoACuenta
                    .Fields!SubTotalPagadoACuenta = 0
                    .Fields!SubTotalPagado = rsDetalle!SubTotalPagado
                
                End With
                rsDetalle.MoveNext
            Loop
            rsDetalle.Close
            mo_Apariencia.ConfigurarFilasBiColores Me.grdItems, SIGHComun.GrillaConFilasBicolor
            '-------------------------------------
            
            '-------------------------------------
            'Cargamos las Formas de Pago de la Factura
            '-------------------------------------
            Dim rsFormaPago As New Recordset
            Dim oDOCajaFormaPago As New DOCajaFormaPagoComprobante
            oDOCajaFormaPago.IdComprobantePago = oDOCajaComprobantesPago.IdComprobantePago
            Set rsFormaPago = mo_AdminCaja.CajaFormaPagoComprobante(oDOCajaFormaPago)
            Do While Not rsFormaPago.EOF
                Me.txtMontoRecibidoSoles.Text = rsFormaPago!Importe
'                With mrs_FormaPago
'                    .AddNew
'                    .Fields!IdFormaPago = rsFormaPago!IdFormaPago
'                    .Fields!IdTipoFormaPago = rsFormaPago!IdTipoFormaPago
'                    .Fields!Importe = rsFormaPago!Importe
'                    .Fields!IdTipoMoneda = rsFormaPago!IdTipoMoneda
'                End With
                rsFormaPago.MoveNext
            Loop
            rsFormaPago.Close
'            mo_Apariencia.ConfigurarFilasBiColores Me.grdDinero, SIGHCOmun.GrillaConFilasBicolor
            '-------------------------------------
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       
    CalcularSubTotalesItems
'    CalcularSubTotalesDinero
   
    cmdImprimir.Enabled = False
    cmdNuevo.Enabled = False
    fraDatosGenerales.Enabled = False
    fraItems.Enabled = False
'    fraDinero.Enabled = False
    Select Case mi_Opcion
        Case sghAgregar
            cmdNuevo.Enabled = True
            cmdImprimir.Enabled = True
            
            fraDatosGenerales.Enabled = True
            fraItems.Enabled = True
'            fraDinero.Enabled = True
        Case sghModificar
            cmdImprimir.Enabled = True
        
            fraDatosGenerales.Enabled = True
            fraItems.Enabled = True
'            fraDinero.Enabled = True
        Case sghEliminar
        Case sghConsultar
            Me.cmdGrabar.Enabled = False
    End Select
   
End Sub
Sub CargarComboBoxes()
    Dim sSQL As String
    
    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()

    mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
    mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub

Private Sub grdDinero_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If bCalculandoSubTotales Then Exit Sub
'    CalcularSubTotalesDinero
End Sub


'Private Sub grdDinero_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    grdDinero.Bands(0).Columns("IdFormaPago").Hidden = True
'
'    Dim rs As New Recordset
'
'    Set rs = mo_AdminCaja.CajaTiposFormasPago()
'    With grdDinero.ValueLists.Add("TipoFormaPago").ValueListItems
'        Do Until rs.EOF
'            .Add Trim(Str(rs.Fields!IdTipoFormaPago)), rs.Fields!Descripcion
'            rs.MoveNext
'        Loop
'    End With
'    rs.Close
'    Set rs = mo_AdminCaja.CajaTiposMoneda
'    With grdDinero.ValueLists.Add("TipoMoneda").ValueListItems
'        Do Until rs.EOF
'            .Add Trim(Str(rs.Fields!IdTipoMoneda)), rs.Fields!Descripcion
'            rs.MoveNext
'        Loop
'    End With
'    rs.Close
'    grdDinero.Bands(0).Columns("IdTipoFormaPago").Header.Caption = "Forma Pago"
'    grdDinero.Bands(0).Columns("IdTipoFormaPago").Width = 2500
'    grdDinero.Bands(0).Columns("IdTipoFormaPago").ValueList = "TipoFormaPago"
'    grdDinero.Bands(0).Columns("IdTipoFormaPago").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
'
'    grdDinero.Bands(0).Columns("IdTipoMoneda").Header.Caption = "Moneda"
'    grdDinero.Bands(0).Columns("IdTipoMoneda").Width = 3000
'    grdDinero.Bands(0).Columns("IdTipoMoneda").ValueList = "TipoMoneda"
'    grdDinero.Bands(0).Columns("IdTipoMoneda").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
'
'    grdDinero.Bands(0).Columns("Importe").Header.Caption = "Importe"
'    grdDinero.Bands(0).Columns("Importe").Width = 1500
'
'End Sub

Private Sub grdDinero_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode.Value, cmdGrabar
    AdministrarKeyPreview KeyCode.Value
End Sub

Private Sub grdItems_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    Dim SubTotalPorPagar As Double
    If bCalculandoSubTotales Then Exit Sub
    If Cell.Column.Style <> ssStyleCheckBox Then
        On Error Resume Next
        With mrs_ComprobantesDetalle
            If Not .EOF And Not .BOF Then
                bCalculandoSubTotales = True
                SubTotalPorPagar = Round(Round(.Fields!cantidad * .Fields!precioUnitario, 2) - .Fields!SubTotalExonerado - .Fields!SubTotalPagadoACuenta, 2)
                If Cell.Column.Key = "Cantidad" Or Cell.Column.Key = "CheckSeleccionado" Then
                    .Fields!SubTotalPagado = Round(SubTotalPorPagar, 2)
                    .Update
                Else
                    If Cell.Value > SubTotalPorPagar Then
                        MsgBox "El valor por pagar no puede ser mayor que " & SubTotalPorPagar, vbExclamation, Me.Caption
                        .Fields!SubTotalPagado = SubTotalPorPagar
                        .Update
                    End If
                End If
                bCalculandoSubTotales = False
            End If
        End With
    End If
    CalcularSubTotalesItems
End Sub


Private Sub grdItems_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (Cell.Column.Key = "Cantidad" Or Cell.Column.Key = "CheckSeleccionado" Or Cell.Column.Key = "SubTotalPagado") Then
        Cancel = True
    Else
        If Cell.Column.Key <> "CheckSeleccionado" Then
            If mrs_ComprobantesDetalle.EOF = False And mrs_ComprobantesDetalle.BOF = False Then
                If Not IsNull(mrs_ComprobantesDetalle.Fields!IdFacturacionDetalle) Then
                    If Cell.Column.Key = "Cantidad" Then
                        Cancel = True
                    End If
                Else
                    If Cell.Column.Key = "SubTotalPagado" Then
                        Cancel = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub grdItems_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdItems.Bands(0).Columns("IdComprobanteDetalle").Hidden = True
    grdItems.Bands(0).Columns("IdFacturacionDetalle").Hidden = True
    grdItems.Bands(0).Columns("TipoDetalle").Hidden = True
    grdItems.Bands(0).Columns("IdProducto").Hidden = True
    
    grdItems.Bands(0).Columns("CheckSeleccionado").Header.Caption = "X"
    grdItems.Bands(0).Columns("CheckSeleccionado").Style = ssStyleCheckBox
    grdItems.Bands(0).Columns("CheckSeleccionado").Width = 300
    
    grdItems.Bands(0).Columns("CodigoProducto").Header.Caption = "Código"
    grdItems.Bands(0).Columns("CodigoProducto").Width = 600

    grdItems.Bands(0).Columns("Producto").Header.Caption = "Descripción"
    grdItems.Bands(0).Columns("Producto").Width = 3500

    grdItems.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    grdItems.Bands(0).Columns("Cantidad").Width = 800

    grdItems.Bands(0).Columns("PrecioUnitario").Header.Caption = "Cost.Unit"
    grdItems.Bands(0).Columns("PrecioUnitario").Width = 900
    
    grdItems.Bands(0).Columns("SubTotalExonerado").Header.Caption = "Exonerado"
    grdItems.Bands(0).Columns("SubTotalExonerado").Width = 1000
    
    grdItems.Bands(0).Columns("SubTotalPagadoACuenta").Header.Caption = "PagadoACuenta"
    grdItems.Bands(0).Columns("SubTotalPagadoACuenta").Width = 1400
    
    grdItems.Bands(0).Columns("SubTotalPagado").Header.Caption = "Por Pagar"
    grdItems.Bands(0).Columns("SubTotalPagado").Width = 1300
End Sub
Sub CargarDatosComprobantesDetalle()
Dim rsDetalle As New Recordset
Dim oCompDetalle As New DOCajaComprobantesDetalle
    oCompDetalle.IdComprobantePago = ml_IdComprobantePago

    Set rsDetalle = mo_AdminCaja.CajaComprobantesDetalle(oCompDetalle)
    Do While Not rsDetalle.EOF
        With mrs_ComprobantesDetalle
            .AddNew
            .Fields!IdComprobanteDetalle = rsDetalle!IdComprobanteDetalle
            .Fields!IdFacturacionDetalle = rsDetalle!IdFacturacionDetalle
            .Fields!CodigoProducto = rsDetalle!CodigoProducto
            .Fields!IdProducto = rsDetalle!IdProducto
            .Fields!Producto = rsDetalle!Producto
            .Fields!cantidad = rsDetalle!cantidad
            .Fields!precioUnitario = rsDetalle!precioUnitario
            .Fields!SubTotalExonerado = rsDetalle!SubTotalExonerado
            .Fields!SubTotalPagado = rsDetalle!SubTotalPagado
        End With
        rsDetalle.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores Me.grdItems, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub GenerarRecordsetTemporal()
    Set mrs_ComprobantesDetalle = New Recordset
    With mrs_ComprobantesDetalle
          .Fields.Append "CheckSeleccionado", adBoolean, 4, adFldIsNullable
          .Fields.Append "TipoDetalle", adVarChar, 4, adFldIsNullable
          .Fields.Append "IdFacturacionDetalle", adInteger, 4, adFldIsNullable
          .Fields.Append "IdComprobanteDetalle", adInteger, 4, adFldIsNullable
          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
          .Fields.Append "CodigoProducto", adVarChar, 20, adFldIsNullable
          .Fields.Append "Producto", adVarChar, 200, adFldIsNullable
          .Fields.Append "Cantidad", adCurrency, 8, adFldIsNullable
          .Fields.Append "PrecioUnitario", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalExonerado", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalPagadoACuenta", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalPagado", adCurrency, 8, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdItems.DataSource = mrs_ComprobantesDetalle
    
'    Set mrs_FormaPago = New Recordset
'    With mrs_FormaPago
'          .Fields.Append "IdFormaPago", adInteger, 4, adFldIsNullable
'          .Fields.Append "IdTipoFormaPago", adInteger, 4, adFldIsNullable
'          .Fields.Append "Importe", adCurrency, 8, adFldIsNullable
'          .Fields.Append "IdTipoMoneda", adInteger, 4, adFldIsNullable
'
'          .LockType = adLockOptimistic
'          .Open
'    End With
'    Set Me.grdDinero.DataSource = mrs_FormaPago
End Sub
Private Sub CalcularSubTotalesItems()
    Dim dSubTotal As Double
    Dim dImpuesto As Double
    Dim dTotal As Double
    
    bCalculandoSubTotales = True
        
    dSubTotal = 0
    
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = mrs_ComprobantesDetalle.Clone(adLockReadOnly)
    
    If Not (rsTemp.BOF And rsTemp.EOF) Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If rsTemp.Fields!CheckSeleccionado Then
                dSubTotal = dSubTotal + rsTemp.Fields!SubTotalPagado
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.MoveFirst
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    bCalculandoSubTotales = False
    
'    Me.lblSubTotal = Format(dSubTotal, "0.00")
'    Me.lblIGV = Format(dSubTotal * md_PorcentajeIGV, "0.00")
'    Me.lblTotal = Format(dSubTotal * (1 + md_PorcentajeIGV), "0.00")
'    Me.lblMontoFacturadoSoles = Me.lblTotal
    
    Me.lblTotal = Format(dSubTotal, "0.00")
    Me.lblSubTotal = Format(dSubTotal / (1 + md_PorcentajeIGV), "0.00")
    Me.lblIGV = CCurrency(lblTotal) - CCurrency(lblSubTotal)
    
    Me.lblMontoFacturadoSoles = Me.lblTotal
    
    CalcularVuelto
End Sub
'Private Sub CalcularSubTotalesDinero()
'    Dim dSubTotalSoles As Double
'    Dim dSubTotalDolar As Double
'    dSubTotalSoles = 0
'    dSubTotalDolar = 0
'    bCalculandoSubTotales = True
'
'
'    Dim rsTemp As ADODB.Recordset
'    Set rsTemp = mrs_FormaPago.Clone(adLockReadOnly)
'
'    If Not (rsTemp.BOF And rsTemp.EOF) Then
'        rsTemp.MoveFirst
'        Do While Not rsTemp.EOF
'            If rsTemp.Fields!IdTipoMoneda = ID_TIPO_MONEDA_DOLAR Then
'                dSubTotalDolar = dSubTotalDolar + IIf(IsNull(rsTemp.Fields!Importe), 0, rsTemp.Fields!Importe)
'            Else
'                dSubTotalSoles = dSubTotalSoles + IIf(IsNull(rsTemp.Fields!Importe), 0, rsTemp.Fields!Importe)
'            End If
'            rsTemp.MoveNext
'        Loop
'        rsTemp.MoveFirst
'    End If
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'    bCalculandoSubTotales = False
'    lblMontoRecibidoDolares = Format(dSubTotalDolar, "0.00")
'    lblMontoRecibidoSoles = Format(dSubTotalSoles, "0.00")
'
'    CalcularVuelto
'End Sub
Private Sub CalcularVuelto()
    Dim dTotalRecibidoSoles As Currency
    Dim dTotalFacturadoSoles As Currency

    
    dTotalRecibidoSoles = CCurrency(Me.txtMontoRecibidoSoles)
    dTotalFacturadoSoles = CCurrency(Me.lblMontoFacturadoSoles)
    If dTotalFacturadoSoles > dTotalRecibidoSoles Then
        lblMontoFaltanteSoles = Format(dTotalFacturadoSoles - dTotalRecibidoSoles, "0.00")
        lblMontoVueltoSoles = Format(0, "0.00")
    Else
        lblMontoFaltanteSoles = Format(0, "0.00")
        lblMontoVueltoSoles = Format(dTotalRecibidoSoles - dTotalFacturadoSoles, "0.00")
    End If
End Sub
Private Sub CargarDatosGenerales()
    Dim oTipoMoneda As New DOCajaTiposMoneda
    oTipoMoneda.IdTipoMoneda = ID_TIPO_MONEDA_DOLAR
    md_TipoCambioDolar = mo_AdminCaja.CajaTipoCambioActualMoneda(oTipoMoneda)
    lblTipoCambio.Caption = Format(md_TipoCambioDolar, "0.00")
    ConfigurarIGV
End Sub
Private Sub ConfigurarIGV()
    If Me.optFiltro(0).Value Then
        md_PorcentajeIGV = 0
        Me.lblLabelIGV.Caption = "IGV " & Format(0, "0.00") & "%"
    Else
        md_PorcentajeIGV = Round(mo_AdminCaja.ImpuestoIGV / 100#, 2)
        Me.lblLabelIGV.Caption = "IGV " & Format(mo_AdminCaja.ImpuestoIGV, "0.00") & "%"
    End If
End Sub

Private Sub grdItems_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode.Value, cmdGrabar
    AdministrarKeyPreview KeyCode.Value

End Sub

Private Sub optFiltro_Click(Index As Integer)
    SeleccionarTipoDocumento
    ConfigurarIGV
    ObtenerDetallesFacturablesCuentaAtencion "CUENTA", ml_IdCuentaAtencionActual, ObtenerTipoFiltro
End Sub

Private Sub txtIdNroHistoria_GotFocus()
    txtIdNroHistoria.Tag = Trim(txtIdNroHistoria.Text)
End Sub

Private Sub txtIdNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRazonSocial
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdNroHistoria_LostFocus()
    If Trim(txtIdNroHistoria.Text) = "" Then
        Exit Sub
    End If
    If txtIdNroHistoria.Tag <> Trim(txtIdNroHistoria.Text) Then
        ObtenerCuentaAtencionPorHistoriaClinica Val(txtIdNroHistoria.Text)
    End If
End Sub
Sub ObtenerCuentaAtencionPorHistoriaClinica(NroHistoriaClinica As Long)
    Dim rsCuentasAtencion As New ADODB.Recordset
    Dim iCount As Integer

    ml_IdCuentaAtencionActual = 0
    Set rsCuentasAtencion = mo_AdminCaja.ObtenerCuentasAtencionPorHistoriaClinica(NroHistoriaClinica)
    iCount = 0
    Do While Not rsCuentasAtencion.EOF
        iCount = iCount + 1
        ml_IdCuentaAtencionActual = rsCuentasAtencion!IdCuentaAtencion
        rsCuentasAtencion.MoveNext
    Loop
    If iCount > 1 Then
        'Levantamos el formulario para seleccionar la cuenta de atención
        Dim oFrmCuentasAtencion As New CuentasAtencionSeleccionar
        Set oFrmCuentasAtencion.DataSource = rsCuentasAtencion
        oFrmCuentasAtencion.Show vbModal
        If oFrmCuentasAtencion.BotonPresionado = sghCancelar Then
            ml_IdCuentaAtencionActual = 0
        Else
            ml_IdCuentaAtencionActual = oFrmCuentasAtencion.IdRegistroSeleccionado
        End If
    End If
    ObtenerDetallesFacturablesCuentaAtencion "CUENTA", ml_IdCuentaAtencionActual, ObtenerTipoFiltro
End Sub

Private Sub txtMontoRecibidoSoles_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMontoRecibidoSoles
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtMontoRecibidoSoles_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtMontoRecibidoSoles_LostFocus()
    CalcularVuelto
End Sub

Private Sub txtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRazonSocial
    AdministrarKeyPreview KeyCode
End Sub
Private Sub NuevoComprobante()
    ml_IdComprobantePago = 0
    mi_Opcion = sghOpciones.sghAgregar
    If cmbIdTipoComprobante.ListCount > 0 Then
        cmbIdTipoComprobante.ListIndex = 0
    End If
    cmbIdTipoGenHistoriaClinica.ListIndex = -1
    Me.txtIdNroHistoria.Text = ""
    ml_IdCuentaAtencionActual = 0
    Me.txtRazonSocial = ""
    Me.txtRUC = ""
    Me.txtMontoRecibidoSoles = ""
    
    cmbIdTipoComprobante_Click
    
    GenerarRecordsetTemporal
    CalcularSubTotalesItems
'    CalcularSubTotalesDinero
    CalcularVuelto
    
End Sub
Private Function CCurrency(sValor As String) As Currency
    If Trim(sValor) = "" Then
        CCurrency = 0
    Else
        CCurrency = CCur(sValor)
    End If
End Function

Private Sub ObtenerDetallesFacturablesCuentaAtencion(TipoFuente As String, IdPacienteCuenta As Long, TipoFiltro As String)
Dim rsDetalle As New Recordset
Dim oPaciente As New doPaciente

    GenerarRecordsetTemporal
    mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
    If TipoFuente = "PACIENTE" Then
        Set rsDetalle = mo_AdminCaja.ObtenerDetallesFacturablesPorPaciente(IdPacienteCuenta, TipoFiltro)
        oPaciente.IdPaciente = IdPacienteCuenta
    Else
        Set rsDetalle = mo_AdminCaja.ObtenerDetallesFacturablesPorCuenta(IdPacienteCuenta, TipoFiltro)
    End If
    Do While Not rsDetalle.EOF
        With mrs_ComprobantesDetalle
            .AddNew
            .Fields!CheckSeleccionado = True
            .Fields!TipoDetalle = rsDetalle!TipoDetalle
            .Fields!IdFacturacionDetalle = rsDetalle!IdFacturacionDetalle
            .Fields!CodigoProducto = rsDetalle!Codigo
            .Fields!IdProducto = rsDetalle!IdProducto
            .Fields!Producto = rsDetalle!Producto
            .Fields!cantidad = rsDetalle!cantidad
            .Fields!precioUnitario = rsDetalle!precioUnitario
            .Fields!SubTotalExonerado = rsDetalle!SubTotalExonerado
            .Fields!SubTotalPagadoACuenta = rsDetalle!SubTotalPagadoACuenta
            .Fields!SubTotalPagado = Round((rsDetalle!cantidad * rsDetalle!precioUnitario) - (rsDetalle!SubTotalExonerado + rsDetalle!SubTotalPagadoACuenta), 2)
            
            oPaciente.IdPaciente = rsDetalle!IdPaciente
            
            'Me.txtIdCuentaAtencion.Text = CStr(IIf(IsNull(rsDetalle!IdCuentaAtencion), 0, rsDetalle!IdCuentaAtencion))
            Me.txtIdNroHistoria.Text = rsDetalle!NroHistoriaClinica
            mo_cmbIdTipoGenHistoriaClinica.BoundText = rsDetalle!IdTipoNumeracion
        End With
        rsDetalle.MoveNext
    Loop
    If mrs_ComprobantesDetalle.RecordCount > 0 Then
        mrs_ComprobantesDetalle.MoveFirst
    End If
    
    rsDetalle.Close
    Me.txtRazonSocial.Text = ""
    If mo_AdminCaja.ObtenerPacientePorId(oPaciente) Then
        Me.txtRazonSocial.Text = oPaciente.ApellidoPaterno & " " & oPaciente.ApellidoMaterno & " " & oPaciente.PrimerNombre & " " & oPaciente.SegundoNombre
    End If
    mo_Apariencia.ConfigurarFilasBiColores Me.grdItems, SIGHComun.GrillaConFilasBicolor
    CalcularSubTotalesItems
End Sub
Function ObtenerTipoFiltro() As String
    If Me.optFiltro(0).Value Then
        ObtenerTipoFiltro = "SERVICIO"
    ElseIf Me.optFiltro(1).Value Then
        ObtenerTipoFiltro = "MEDICAMENTO"
'    ElseIf Me.optFiltro(2).Value Then
'        ObtenerTipoFiltro = "AMBOS"
    End If
End Function
Sub SeleccionarTipoDocumento()
    If Me.optFiltro(0).Value Then   'Servicios
        mo_cmbIdTipoComprobante.BoundText = "5"     'Recibo
    ElseIf Me.optFiltro(1).Value Then   'Medicamentos
        mo_cmbIdTipoComprobante.BoundText = "1"     'Boleta
    End If
End Sub

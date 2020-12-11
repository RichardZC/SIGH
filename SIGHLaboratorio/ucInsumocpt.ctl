VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucInsumocpt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   ScaleHeight     =   5730
   ScaleWidth      =   11850
   Begin VB.CheckBox chkTodosNinguno 
      Caption         =   "Todos/ninguno"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8580
      TabIndex        =   7
      Top             =   5415
      Width           =   1590
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   3201
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   2445
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   4313
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BorderStyle     =   5
      ScrollBars      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pruebas a realizar"
   End
   Begin UltraGrid.SSUltraGrid grdInsumos 
      Height          =   5325
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   9393
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      ScrollBars      =   2
      BorderStyleCaption=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "ucInsumocpt.ctx":0000
      CaptionAppearance=   "ucInsumocpt.ctx":003C
      Caption         =   "Insumos"
   End
   Begin Threed.SSOption optPorCodigo 
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   5400
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reg.Código"
   End
   Begin Threed.SSOption optPorDescripcion 
      Height          =   255
      Left            =   6810
      TabIndex        =   6
      Top             =   5400
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reg.Descripción"
      Value           =   -1
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   3
      Top             =   5400
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Teclas de ayuda: <F10>=Agregar  Pruebas         <Supr>=Eliminar Registro        "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Menu mnuProductos 
      Caption         =   "Elija"
      Begin VB.Menu mnuAgregarInsumo 
         Caption         =   "Agregar Insumo"
      End
      Begin VB.Menu mnuAgregarServicio 
         Caption         =   "Agregar CPT"
      End
   End
End
Attribute VB_Name = "ucInsumocpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registro de Procedimientos
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Public Event Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim gridInfra As New sighentidades.GridInfragistic
Dim mo_PermisosFacturacion As New PermisosFacturacion
Dim ms_TipoProducto As sghTipoProducto
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idOrden As Long
Dim ml_idCuentaAtencion As Long
Dim mb_CargandoProductos As Boolean
Dim ms_Opcion As sghOpciones
Dim mrs_FacturacionProductos As New Recordset
Dim mrs_FacturacionInsumos As New Recordset
Dim mo_DOAtencion As DOAtencion
Dim ml_idUsuario As Long
Dim ml_IdPuntoCarga As Long
Dim ms_EstadosFacturacion As String
Dim ms_TiposFinanciamiento As String
Dim ml_IdEstadoOrden As Long

'edicion de la grilla
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_DocumentoYaRegistradoEnSeguros As Boolean
Dim ml_PermiteAgregarItems As Boolean
Dim ml_idOrdenPago As Long
Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
Dim ml_BoletaNumero As String
Dim ml_BoletaSerie As String
Dim ml_IdMovimiento As Long
Dim ml_HabilitaIngresoDePrecio As Boolean
Dim ml_PermiteVerColumnaCantidadFallada As Boolean
Dim ml_NoPermiteCargarCantidadFallada As Boolean
Dim lbEstoyEnGridCPT As Boolean
Dim lnIdProductoCPT  As Long
Dim ml_ParametroHoras As Integer
Dim lnMaximaCantidadExamen As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lbTieneResultadoAutomatico As Boolean

Property Let TieneResultadoAutomatico(lValue As Boolean)
    lbTieneResultadoAutomatico = lValue
End Property
Property Get TieneResultadoAutomatico() As Boolean
    TieneResultadoAutomatico = lbTieneResultadoAutomatico
End Property

Property Let HabilitaIngresoDePrecio(lValue As Boolean)
  ml_HabilitaIngresoDePrecio = lValue
End Property

Property Let IdMovimiento(lValue As Long)
  ml_IdMovimiento = lValue
End Property

Property Let BoletaSerie(lValue As String)
  ml_BoletaSerie = lValue
End Property

Property Let BoletaNumero(lValue As String)
  ml_BoletaNumero = lValue
End Property

Property Let IdOrdenPago(lValue As Long)
  ml_idOrdenPago = lValue
End Property

Property Get IdOrdenPago() As Long
  IdOrdenPago = ml_idOrdenPago
End Property

Property Let PermiteAgregarItems(lValue As Boolean)
  ml_PermiteAgregarItems = lValue
  EditableColumnasDelGrid
End Property

Sub EditableColumnasDelGrid()
  On Error Resume Next
  If ml_PermiteAgregarItems = True Then
     grdProductos.Bands(0).Columns("Codigo").Activation = ssActivationAllowEdit
     grdProductos.Bands(0).Columns("NombreProducto").Activation = ssActivationAllowEdit
     grdProductos.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
  Else
     grdProductos.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("NombreProducto").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
  End If

End Sub

Property Let PermiteVerColumnaCantidadFallada(lValue As Boolean)
  ml_PermiteVerColumnaCantidadFallada = lValue
End Property

Property Let NoPermiteCargarCantidadFallada(lValue As Boolean)
  ml_NoPermiteCargarCantidadFallada = lValue
End Property

Property Let DocumentoYaRegistradoEnSeguros(lValue As Boolean)
  ml_DocumentoYaRegistradoEnSeguros = lValue
End Property

Property Let idOrden(lValue As Long)
  ml_idOrden = lValue
End Property
Property Get idOrden() As Long
  idOrden = ml_idOrden
End Property

Property Let IdEstadoOrden(lValue As Long)
  ml_IdEstadoOrden = lValue
  Select Case ms_TipoProducto
    Case sghServicio
      HabilitarMenuSegunEstadoOrden ml_IdEstadoOrden
    Case sghbien
      HabilitarMenuSegunEstadoOrden ml_IdEstadoOrden
  End Select
End Property

Property Get IdEstadoOrden() As Long
  IdEstadoOrden = ml_IdEstadoOrden
End Property

Property Let idCuentaAtencion(lValue As Long)
  ml_idCuentaAtencion = lValue
End Property

Property Get idCuentaAtencion() As Long
  idCuentaAtencion = ml_idCuentaAtencion
End Property

Property Set Atencion(oValue As DOAtencion)
  Set mo_DOAtencion = oValue
End Property

Property Get Atencion() As DOAtencion
  Set Atencion = mo_DOAtencion
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
  idUsuario = ml_idUsuario
End Property

Property Let TipoProducto(iTipo As sghTipoProducto)
  ms_TipoProducto = iTipo
  Select Case ms_TipoProducto
    Case sghServicio
        UserControl.mnuAgregarServicio.Caption = "Agregar Servicio"
    Case sghbien
        UserControl.mnuAgregarServicio.Caption = "Agregar Bien Insumo"
  End Select
End Property

Property Get TipoProducto() As sghTipoProducto
  TipoProducto = ms_TipoProducto
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
  ml_IdTipoFinanciamiento = lValue
  If mrs_FacturacionProductos.RecordCount > 0 Then
    mrs_FacturacionProductos.MoveFirst
    Do While Not mrs_FacturacionProductos.EOF
      mrs_FacturacionProductos.Fields!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
      mrs_FacturacionProductos.Update
      mrs_FacturacionProductos.MoveNext
    Loop
  End If
End Property

Property Get IdTipoFinanciamiento() As Long
  IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Property Let Opcion(oValue As sghOpciones)
  ms_Opcion = oValue
End Property

Property Get Opcion() As sghOpciones
  Opcion = ms_Opcion
End Property

Property Set FacturacionProductos(oValue As Recordset)
  Set mrs_FacturacionProductos = oValue
End Property

Property Get FacturacionProductos() As Recordset
  'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
  'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
  Set FacturacionProductos = mrs_FacturacionProductos.Clone()
End Property

Property Set FacturacionInsumos(oValue As Recordset)
  Set mrs_FacturacionInsumos = oValue
End Property

Property Get FacturacionInsumos() As Recordset
  'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
  'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
  Set FacturacionInsumos = mrs_FacturacionInsumos.Clone()
End Property

Property Set ProductosEliminados(oValue As Collection)
  Set mo_ProductosEliminados = oValue
End Property

Property Get ProductosEliminados() As Collection
  Set ProductosEliminados = mo_ProductosEliminados
End Property

Property Let IdPuntoCarga(lValue As Long)
  ml_IdPuntoCarga = lValue
End Property

Property Get IdPuntoCarga() As Long
  IdPuntoCarga = ml_IdPuntoCarga
End Property

Property Let EstadosFacturacion(sValue As String)
  ms_EstadosFacturacion = sValue
End Property

Property Get EstadosFacturacion() As String
  EstadosFacturacion = ms_EstadosFacturacion
End Property

Property Let TiposFinanciamiento(sValue As String)
  ms_TiposFinanciamiento = sValue
End Property

Property Get TiposFinanciamiento() As String
  TiposFinanciamiento = ms_TiposFinanciamiento
End Property

Sub Inicializar()
  ml_DocumentoYaRegistradoEnSeguros = False
  
  Set mrs_FacturacionProductos = New Recordset
  Set mrs_FacturacionInsumos = New Recordset
  GenerarRecordsetProductos
  ms_EstadosFacturacion = ""
  Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
  
  UserControl.mnuAgregarServicio.Enabled = True  'mo_PermisosFacturacion.AgregarServicios
  
  UserControl.grdProductos.DialogStrings(ssDeleteRow) = "Ud. ha seleccionado una fila para borrarla" + Chr(13) + "Seleccione <Si> para borrar la fila o <No> para Cancelar "
  lnMaximoNroItems = BuscarMaximoItemsEnParametros()
  ml_PermiteAgregarItems = True
  lbEstoyEnGridCPT = True
  lnMaximaCantidadExamen = Val(lcBuscaParametro.SeleccionaFilaParametro(293))   'debb-06-03-2012
End Sub

Function BuscarMaximoItemsEnParametros() As Long
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
  Set lcBuscaParametro = Nothing
End Function

Sub AgregaProducto()
  On Error GoTo ErrAgrega
  mb_CargandoProductos = True
  With mrs_FacturacionProductos
    .AddNew
    .Fields!IdFacturacionProducto = 0
    .Fields!idProducto = 0
    .Fields!Codigo = ""
    .Fields!NombreProducto = ""
    .Fields!Cantidad = 1
    .Fields!PrecioUnitario = 0
    .Fields!TotalPorPagar = 0
    .Fields!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    .Fields!IdPuntoCarga = ml_IdPuntoCarga
    If Not mo_DOAtencion Is Nothing Then .Fields!IdAtencion = mo_DOAtencion.IdAtencion
    .Fields!EstadoLocal = "A"   'Agregar
    .Fields!FechaAutorizaPendiente = 0
    .Fields!IdUsuarioAutorizaPendiente = 0
    
    Select Case ml_IdTipoFinanciamiento
      Case 2, 3, 4
        .Fields!IdEstadoFacturacion = 4
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
      Case Else
        .Fields!IdEstadoFacturacion = 1
        .Fields!FechaAutorizaSeguro = 0
        .Fields!IdUsuarioAutorizaSeguro = 0
    End Select
    
    .Fields!IdFuenteFinanciamiento = 1
    .Fields!IdServicioInternamiento = 0
    .Fields!IdUsuarioAuditoria = ml_idUsuario
    .Fields!IdComprobantePago = 0
    .Fields!IdComprobantePagoDevolucion = 0
    .Fields!idOrden = ml_idOrden
  End With
  mb_CargandoProductos = False
  
  Totalizar
  mb_FilaEditable = True
ErrAgrega:
End Sub

Sub AgregaProductoInsumo()
  On Error GoTo ErrAddP
  With mrs_FacturacionInsumos
    .AddNew
    .Fields!IdFacturacionProducto = 0
    .Fields!idProducto = 0
    .Fields!Codigo = ""
    .Fields!NombreProducto = ""
    .Fields!Cantidad = 1
    .Fields!PrecioUnitario = 0
    .Fields!TotalPorPagar = 0
    .Fields!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    .Fields!IdPuntoCarga = ml_IdPuntoCarga
    If Not mo_DOAtencion Is Nothing Then .Fields!IdAtencion = mo_DOAtencion.IdAtencion
    .Fields!EstadoLocal = "A"   'Agregar
    .Fields!FechaAutorizaPendiente = 0
    .Fields!IdUsuarioAutorizaPendiente = 0
        
    Select Case ml_IdTipoFinanciamiento
      Case 2, 3, 4
        .Fields!IdEstadoFacturacion = 4
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
      Case Else
        .Fields!IdEstadoFacturacion = 1
        .Fields!FechaAutorizaSeguro = 0
        .Fields!IdUsuarioAutorizaSeguro = 0
    End Select
    
    .Fields!IdFuenteFinanciamiento = 1
    .Fields!IdServicioInternamiento = 0
    .Fields!IdUsuarioAuditoria = ml_idUsuario
    .Fields!IdComprobantePago = 0
    .Fields!IdComprobantePagoDevolucion = 0
    .Fields!idOrden = ml_idOrden
    .Fields!idProductoCPT = lnIdProductoCPT
  End With
  grdInsumos.PerformAction ssKeyActionActivateCell
  grdInsumos.PerformAction ssKeyActionEnterEditMode
  mrs_FacturacionInsumos.Filter = "idProductoCPT=" & lnIdProductoCPT
  If mrs_FacturacionInsumos.RecordCount > 0 Then mrs_FacturacionInsumos.MoveFirst
ErrAddP:
End Sub

Sub CargaProductosPorIdOrden()
  Dim rs As Recordset
  Select Case ms_TipoProducto
    Case sghServicio
      If ml_IdTipoFinanciamiento = 5 Or ml_IdTipoFinanciamiento = 1 Then
        Set rs = mo_ReglasFacturacion.FacturacionServicioPagosFiltraPorIdOrden(ml_idOrden)
      Else
        Set rs = mo_ReglasFacturacion.FacturacionServicioFinanciamientosFiltraPorIdOrden(ml_idOrden)
      End If
      CargarItemsALaGrillaS rs
      CargarItemsALaGrillaCPT rs, False
  End Select
End Sub

Sub CargaProductosPorIdOrdenPago()
  Dim rs As Recordset
  Select Case ms_TipoProducto
    Case sghServicio
      Set rs = mo_ReglasFacturacion.FacturacionServicioPagosFiltraPorIdOrdenPago(ml_idOrdenPago)
      CargarItemsALaGrillaS rs
    Case sghbien
  End Select
End Sub

Sub CargaProductosPorIdMovimiento()
  Dim rs As Recordset
  Set rs = mo_ReglasLaboratorio.LabMovimientoDetalleSeleccionarPorIdMovimiento(ml_IdMovimiento)
  CargarItemsALaGrillaS rs
  Set rs = mo_ReglasLaboratorio.LabMovimientoCPTSeleccionarPorIdMovimiento(ml_IdMovimiento)
  CargarItemsALaGrillaCPT rs, False
  If Not (mrs_FacturacionProductos.EOF = True And mrs_FacturacionProductos.BOF = True) Then mrs_FacturacionProductos.MoveFirst
  grdProductos_Click
End Sub

Sub CargarItemsALaGrillaS(rs As Recordset)
  Dim oRsTmp1 As New Recordset
  mb_CargandoProductos = True
  Do While Not rs.EOF
    Set oRsTmp1 = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorCodigo, rs!Codigo, ml_IdTipoFinanciamiento, ml_IdPuntoCarga, sghSoloInsumos)
    If oRsTmp1.RecordCount > 0 Then
      mrs_FacturacionInsumos.AddNew
      mrs_FacturacionInsumos!idProducto = rs!idProducto
      mrs_FacturacionInsumos!Codigo = rs!Codigo
      mrs_FacturacionInsumos!NombreProducto = rs!Nombre
      mrs_FacturacionInsumos!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
      mrs_FacturacionInsumos!Cantidad = rs!Cantidad
      mrs_FacturacionInsumos!PrecioUnitario = 1
      mrs_FacturacionInsumos!TotalPorPagar = 1
      
      If ml_NoPermiteCargarCantidadFallada = False Then mrs_FacturacionInsumos!cantidadFallada = rs!cantidadFallada
      mrs_FacturacionInsumos!idProductoCPT = rs!idProductoCPT
    End If
    rs.MoveNext
  Loop
  mb_CargandoProductos = False
  Set grdInsumos.DataSource = mrs_FacturacionInsumos
End Sub

Sub CargarItemsALaGrillaCPT(rs As Recordset, lbSeCargaDesdeBoletaOpcionAgregar As Boolean)
  Dim oRsTmp1 As New Recordset
  Dim lbPrimeraVez As Boolean
  Dim lnPrecio As Double
  Dim lnImporte As Double
  Dim rsCatHos As New Recordset
  On Error Resume Next
  mb_CargandoProductos = True
  lbPrimeraVez = True
  Do While Not rs.EOF
    Set oRsTmp1 = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorCodigo, rs!Codigo, ml_IdTipoFinanciamiento, ml_IdPuntoCarga, sghSoloCPT)
    If oRsTmp1.RecordCount > 0 Then
      'debb-05/04/2011
      If lbSeCargaDesdeBoletaOpcionAgregar = True Then
          If rs!exoneraciones > 0 Then
             lnImporte = Round((rs!Importe * rs!totalBoleta) / (rs!exoneraciones + rs!totalBoleta), 2)
             lnPrecio = Round(lnImporte / rs!Cantidad, 2)
          Else
             lnPrecio = rs!precio
             lnImporte = rs!Importe
          End If
      Else
          lnPrecio = rs!precio
          lnImporte = rs!Importe
      End If
      '
      Set rsCatHos = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(rs!idProductoCPT, ml_IdTipoFinanciamiento)
      mrs_FacturacionProductos.AddNew
      mrs_FacturacionProductos!idProducto = rs!idProductoCPT
      mrs_FacturacionProductos!Codigo = rs!Codigo
      mrs_FacturacionProductos!NombreProducto = rs!Nombre
      mrs_FacturacionProductos!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
      mrs_FacturacionProductos!Cantidad = rs!Cantidad
      mrs_FacturacionProductos!PrecioUnitario = lnPrecio
      mrs_FacturacionProductos!TotalPorPagar = lnImporte
      mrs_FacturacionProductos!SeUsaSinPrecio = IIf(IsNull(rsCatHos!SeUsaSinPrecio), False, rsCatHos!SeUsaSinPrecio)
      mrs_FacturacionProductos!idOrden = rs!idOrden
      mrs_FacturacionProductos!CantidadSinEditar = rs!Cantidad
      If lbSeCargaDesdeBoletaOpcionAgregar = True Then
         mrs_FacturacionProductos!ResultadoAutomatico = IIf(rsCatHos!LabResultadoAutomatico = 1, True, False)
      Else
        If IsNull(rs!ResultadoAutomatico) Then
           mrs_FacturacionProductos!ResultadoAutomatico = False
        Else
           mrs_FacturacionProductos!ResultadoAutomatico = IIf(rs!ResultadoAutomatico = 1, True, False)
        End If
      End If
    End If
    rs.MoveNext
  Loop
  mb_CargandoProductos = False
  Set rsCatHos = Nothing
  Totalizar
  Set grdProductos.DataSource = mrs_FacturacionProductos
End Sub

Sub HabilitarMenuSegunEstadoOrden(IdEstadoOrden As Long)
  Select Case IdEstadoOrden
    Case 1
      HabilitarMenus True
    Case 4
      HabilitarMenus False
    Case 9
      HabilitarMenus False
  End Select
End Sub

Sub HabilitarMenus(Estado As Boolean)
  UserControl.mnuAgregarServicio.Enabled = Estado
End Sub

Function DevuelveTotalPagar() As Double
  Dim rsProductos As New Recordset
  Dim dTotalPagado As Double
  Set rsProductos = mrs_FacturacionProductos.Clone
  dTotalPagado = 0
  If rsProductos.RecordCount > 0 Then
    rsProductos.MoveFirst
    Do While Not rsProductos.EOF
      dTotalPagado = dTotalPagado + rsProductos!TotalPorPagar
      rsProductos.MoveNext
    Loop
  End If
  DevuelveTotalPagar = dTotalPagado
End Function

Sub Totalizar()
  Dim dSubTotal As Double
  Dim lIdEstadoFacturacion As Long
  Dim lIdProducto As Long
  Dim rsProductos As New Recordset
  Dim dTotalExonerado As Double
  Dim dTotalPagoACuenta As Double
  Dim dTotalIngresado As Double
  Dim dTotalPendientePago As Double
  Dim dTotalPagado As Double
  Dim dTotalPorDevolver As Double
  Dim dTotalDevuelto As Double
  Dim dTotalAnulado As Double
  
  dTotalExonerado = 0
  dTotalPagoACuenta = 0
  dTotalIngresado = 0
  dTotalPendientePago = 0
  dTotalPagado = 0
  dTotalPorDevolver = 0
  dTotalDevuelto = 0
  dTotalAnulado = 0
  
  lbTieneResultadoAutomatico = False
  
  Set rsProductos = mrs_FacturacionProductos.Clone
  
  If rsProductos.RecordCount = 0 Then Exit Sub
  
  If Not (rsProductos.EOF = True And rsProductos.BOF = True) Then
    rsProductos.MoveFirst
    Do While Not rsProductos.EOF
      If rsProductos!ResultadoAutomatico = True Then
         lbTieneResultadoAutomatico = True
      End If
      dSubTotal = rsProductos!TotalPorPagar
      lIdEstadoFacturacion = rsProductos!IdEstadoFacturacion
      lIdProducto = rsProductos!idProducto
      dTotalIngresado = dTotalIngresado + dSubTotal
      Select Case lIdEstadoFacturacion
        Case 1
          Select Case lIdProducto
            Case 4692
              dTotalExonerado = dTotalExonerado + dSubTotal
            Case Else
              ' dTotalIngresado = dTotalIngresado + dSubTotal
          End Select
        Case 3
          dTotalPendientePago = dTotalPendientePago + dSubTotal
        Case 4
          Select Case lIdProducto
            Case 4691
              dTotalPagoACuenta = dTotalPagoACuenta + dSubTotal
            Case Else
              dTotalPagado = dTotalPagado + dSubTotal
          End Select
        Case 5
          dTotalPorDevolver = dTotalPorDevolver + dSubTotal
        Case 6
          dTotalDevuelto = dTotalDevuelto + dSubTotal
        Case 9
          dTotalAnulado = dTotalAnulado + dSubTotal
      End Select
      rsProductos.MoveNext
    Loop
  End If
  
  RaiseEvent Totalizado(dTotalIngresado, dTotalPendientePago, dTotalPagoACuenta, dTotalExonerado, dTotalPagado, dTotalPorDevolver, dTotalDevuelto, dTotalAnulado)
  lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")
End Sub

Private Sub chkTodosNinguno_Click()
    On Error GoTo ErrTod
    mrs_FacturacionProductos.MoveFirst
    Do While Not mrs_FacturacionProductos.EOF
       If chkTodosNinguno.Value = 1 Then
          mrs_FacturacionProductos!ResultadoAutomatico = True
       Else
          mrs_FacturacionProductos!ResultadoAutomatico = False
       End If
       mrs_FacturacionProductos.MoveNext
    Loop
ErrTod:
End Sub

Private Sub grdInsumos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
   Totalizar
End Sub

Private Sub grdInsumos_AfterRowsDeleted()
  'Set grdInsumos.DataSource = mrs_FacturacionInsumos
  grdProductos_Click
End Sub

Private Sub grdInsumos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
  If ml_PermiteAgregarItems = True Then
    'Si la fila es editable y estamos en la celda de codigo se completa los datos
    'del producto
    Select Case grdInsumos.ActiveCell.Column.Key
      Case "Codigo"
        'oRow.Cells("Codigo").Value = Right("000000" & Trim(oRow.Cells("Codigo").Value), 6)
        ConfigurarProductoPorCodigo grdInsumos
      Case "Cantidad"
        'RecalcularSubTotal grdProductos
    End Select
  End If
End Sub

Private Sub grdInsumos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  lbEstoyEnGridCPT = False
  InicializarLaGrilla grdInsumos
End Sub

Private Sub grdInsumos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
  lbEstoyEnGridCPT = False
  OnKeyDown grdInsumos, KeyCode
End Sub

Private Sub grdInsumos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
  lbEstoyEnGridCPT = False
  OnKeyPress grdInsumos, KeyAscii
End Sub

Private Sub grdInsumos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuProductos
End Sub

Private Sub grdProductos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If Cell.Column.Key = "Cantidad" Then
        If Cell.Row.Cells("Cantidad").Value > lnMaximaCantidadExamen Then          'debb-06-03-2012
           If MsgBox("La cantidad registrada es mayor que el máximo permitido: " & Trim(Str(lnMaximaCantidadExamen)) & Chr(13) & "¿Realmente desea registrar esa cantidad?", vbQuestion + vbYesNo, "Imágenes") = vbNo Then
              Cell.Row.Cells("Cantidad").Value = 0
           End If
        End If
    End If
    Totalizar
End Sub

'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
  If ml_PermiteAgregarItems = True Then
    If mb_CargandoProductos Then Exit Sub
  End If
End Sub

Private Sub grdProductos_AfterRowsDeleted()
'  If ml_PermiteAgregarItems = True Then
    If ml_ultimoProductoEliminado > 0 Then
      mo_ProductosEliminados.Add ml_ultimoProductoEliminado
      ml_ultimoProductoEliminado = 0
      Totalizar
    Else
      Totalizar
      Set grdProductos.DataSource = mrs_FacturacionProductos
    End If
    If mrs_FacturacionInsumos.RecordCount > 0 Then
      mrs_FacturacionInsumos.MoveFirst
      Do While Not mrs_FacturacionInsumos.EOF
        mrs_FacturacionInsumos.Delete
        mrs_FacturacionInsumos.Update
        mrs_FacturacionInsumos.MoveNext
      Loop
    End If
  'Else
    
  'End If
End Sub

Private Sub grdProductos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
  If ml_PermiteAgregarItems = True Then
    If mb_FilaEditable Then
      'Si la fila es editable y estamos en la celda de codigo se completa los datos
      'del producto
      Select Case grdProductos.ActiveCell.Column.Key
        Case "Codigo"
          'oRow.Cells("Codigo").Value = Right("000000" & Trim(oRow.Cells("Codigo").Value), 6)
          ConfigurarProductoPorCodigo grdProductos
        Case "Cantidad"
          RecalcularSubTotal grdProductos
        Case "TipoFinanciamiento"
        Case "EstadoFacturacion"
      End Select
    End If
  End If
End Sub

Private Sub grdProductos_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
  If ml_PermiteAgregarItems = True Then
    'Si la fila no es editable, cancela cualquier cambio en la fila
    If Not mb_FilaEditable Then
      Cancel = True
      Exit Sub
    End If
  End If
End Sub

Private Sub grdProductos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
  If ml_PermiteAgregarItems = True Then
    mb_FilaEditable = True
  End If
End Sub

Private Sub grdProductos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
  'If ml_PermiteAgregarItems = True Then
    'Si ya esta pagado cancela la eliminacion
    If ml_IdEstadoOrden = 9 Then
      MsgBox "No se puede eliminar ITEMS de una orden que ya ha sido ANULADA", vbExclamation, "Facturación"
      Cancel = True
      Exit Sub
    End If
        
    If ml_IdEstadoOrden = 4 Then
      MsgBox "No se puede eliminar ITEMS de una orden que ya ha sido PAGADA", vbExclamation, "Facturación"
      Cancel = True
      Exit Sub
    End If
        
    If Rows.Item(0).Cells("EstadoLocal").Value = "L" And Rows.Item(0).Cells("idestadofacturacion").Value = 4 Then
      Cancel = True
    Else
      ml_ultimoProductoEliminado = 0
      ml_ultimoProductoEliminado = Val(Rows.Item(0).Cells("IdFacturacionProducto").Value)
    End If
  'Else
   ' Cancel = True
  'end If
End Sub

Private Sub grdProductos_Click()
  On Error GoTo ErrCPT
  If Not mrs_FacturacionProductos.EOF Then
    grdInsumos.Caption = "Insumos para el CPT: (" & mrs_FacturacionProductos.Fields!Codigo & ") " & mrs_FacturacionProductos.Fields!NombreProducto
    lnIdProductoCPT = mrs_FacturacionProductos.Fields!idProducto
    mrs_FacturacionInsumos.Filter = "idProductoCPT=" & lnIdProductoCPT
    If mrs_FacturacionInsumos.RecordCount > 0 Then mrs_FacturacionInsumos.MoveFirst
  End If
  
ErrCPT:
  'AgregaProductoInsumo
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
  lbEstoyEnGridCPT = True
  InicializarLaGrilla grdProductos
  EditableColumnasDelGrid
End Sub

Private Sub grdProductos_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
  On Error Resume Next
  If ml_PermiteAgregarItems = True Then ModificarColorDeFila Row
End Sub

Sub ModificarColorDeFila(ByVal Row As UltraGrid.SSRow)
  Select Case Row.Cells("IdProducto").Value
    Case 4691
      Row.Appearance.ForeColor = &HC7613F
    Case 4692
      Row.Appearance.ForeColor = &H16CD32
    Case 4693
      Row.Appearance.ForeColor = &H3049FA
  End Select
End Sub

Private Sub grdProductos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    lbEstoyEnGridCPT = True
    OnKeyDown grdProductos, KeyCode
    If KeyCode = vbKeyF2 Then
       Dim lnKeyCode As Integer
       lnKeyCode = KeyCode
       RaiseEvent SePresionoTeclaEspecial(lnKeyCode)
    End If
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
  lbEstoyEnGridCPT = True
  OnKeyPress grdProductos, KeyAscii
End Sub

Private Sub grdProductos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If ml_PermiteAgregarItems = True Then
    If Button = 2 Then PopupMenu mnuProductos
  End If
End Sub
Sub RecalcularSubTotal(oGrilla As SSUltraGrid)
  Dim oRow As SSRow
  Dim dValorAntesDe As Double
  
  Set oRow = oGrilla.ActiveCell.Row
  dValorAntesDe = CDbl(oRow.Cells("TotalPorPagar").Value)
  oRow.Cells("TotalPorPagar").Value = CDbl(oRow.Cells("PrecioUnitario").Value) * Val(oRow.Cells("Cantidad").Value)
  If dValorAntesDe - CDbl(oRow.Cells("TotalPorPagar").Value) <> 0 Then
    If oRow.Cells("EstadoLocal").Value = "A" Then
      'Si recen ha sido agregado lo deja como agregado
    End If
    If oRow.Cells("EstadoLocal").Value = "L" Then
      'Si ya estuvo en la base de datos, lo marca como modificado
      oRow.Cells("EstadoLocal").Value = "M"   'Modificado
    End If
  End If
  Totalizar
End Sub

Sub ConfigurarProductoPorCodigo(oGrilla As SSUltraGrid)
  Dim rs As Recordset
  Dim oRow As SSRow

  Set oRow = oGrilla.ActiveCell.Row
    
  If IsNull(oRow.Cells("codigo").Value) Or IsNull(oRow.Cells("idtipofinanciamiento").Value) Or oRow.Cells("codigo").Value = "" Then Exit Sub
  If ms_TipoProducto = sghbien Then oRow.Cells("codigo").Value = Right("0000000000" & oRow.Cells("codigo").Value, 5)
  Select Case ms_TipoProducto
    Case sghServicio
      If lbEstoyEnGridCPT = True Then
        Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorCodigo, oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, ml_IdPuntoCarga, sghSoloCPT)
      Else
        'Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorCodigo, oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, ml_IdPuntoCarga, sghSoloInsumos)
        Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga(sghPorCodigo, oRow.Cells("codigo").Value, ml_IdPuntoCarga, sghSoloInsumos)
      End If
  End Select
    
  If rs.RecordCount > 0 Then
    If rs.Fields("idproducto").Value <> 4691 Then
      'Busca si ya existe el producto
      If Not ItemYaExiste(rs.Fields("idproducto").Value) Then
       Dim fechaDespacho As Date
       fechaDespacho = DateAdd("h", ml_ParametroHoras, Now)
       If idCuentaAtencion = 0 Then GoTo Sigue1
       If BuscaItemEnDia(idCuentaAtencion, Format(fechaDespacho, sighentidades.DevuelveFechaSoloFormato_DMY_HMS), rs.Fields("idproducto").Value) = True Then
Sigue1:
        oRow.Cells("IdFacturacionProducto").Value = 0
        oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
        oRow.Cells("NombreProducto").Value = rs.Fields("Nombre").Value
        If lbEstoyEnGridCPT = True Then
          oRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
          oRow.Cells("TotalPorPagar").Value = rs.Fields("preciounitario").Value
        Else
          oRow.Cells("preciounitario").Value = 1
          oRow.Cells("TotalPorPagar").Value = 1
        End If
        oRow.Cells("cantidad").Value = 1
        If lbEstoyEnGridCPT = True Then
           oRow.Cells("SeUsaSinPrecio").Value = IIf(IsNull(rs.Fields("SeUsaSinPrecio").Value), False, rs.Fields("SeUsaSinPrecio").Value)
        End If
        If rs!LabResultadoAutomatico = 1 And lbEstoyEnGridCPT = True Then
           oRow.Cells("ResultadoAutomatico").Value = True
        End If
        
        
       End If
      End If
    End If
  End If
End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As Long) As Boolean
  Dim lbExiste As Boolean
  Dim oRsTmp As New ADODB.Recordset
  If lbEstoyEnGridCPT = True Then
    Set oRsTmp = mrs_FacturacionProductos.Clone
    ItemYaExiste = False
    If oRsTmp.RecordCount > 0 Then
      oRsTmp.MoveFirst
      oRsTmp.Find "idProducto=" & lnIdProducto
      If Not oRsTmp.EOF Then
        ItemYaExiste = True
        MsgBox "La prueba que desea agregar ya está en el listado.", vbInformation, "SIGH "
      End If
    End If
    oRsTmp.Close
  Else
    Set oRsTmp = mrs_FacturacionInsumos.Clone
    ItemYaExiste = False
    If oRsTmp.RecordCount > 0 Then
      oRsTmp.MoveFirst
      Do While Not oRsTmp.EOF
        If oRsTmp.Fields!idProducto = lnIdProducto And oRsTmp.Fields!idProductoCPT = lnIdProductoCPT Then
          ItemYaExiste = True
          MsgBox "Este insumo ya está registrado", vbInformation, "Facturación"
          Exit Do
        End If
        oRsTmp.MoveNext
      Loop
    End If
    oRsTmp.Close
  End If
End Function

Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)
  If ml_PermiteAgregarItems = True Or lbEstoyEnGridCPT = True Then
    If Not oGrilla.ActiveCell Is Nothing Then
      Select Case oGrilla.ActiveCell.Column.Key
        Case "Cantidad"
          Select Case Val(Chr(KeyCode))
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
              
            Case Else
              KeyCode = 0
          End Select
        Case "NombreProducto"
          Select Case KeyCode
            Case vbKeyBack
            Case vbKeyEscape
              Set grillaBusqueda.DataSource = Nothing
              grillaBusqueda.Visible = False
            'Case vbKeyReturn
              'SendKeys "{TAB}"
            Case vbKeyDown, vbKeyUp
              On Error Resume Next
              grillaBusqueda.SetFocus
            Case vbKeyLeft
            
          End Select
      End Select
    End If
  End If
  Select Case KeyCode
    Case vbKeyF10
         mnuAgregarServicio_Click
        If optPorCodigo.Value = True Then
           grdProductosFocusColumna "codigo"
        Else
           grdProductosFocusColumna "NombreProducto"
        End If
    Case vbKeyF11
      'mnuAgregarInsumo_Click
  End Select
End Sub

Sub OnKeyPress(oGrilla As SSUltraGrid, KeyAscii As UltraGrid.SSReturnShort)
  If ml_PermiteAgregarItems = True Or lbEstoyEnGridCPT = True Then
    'Si la fila no es editable, cancela cualquier cambio en la fila
    'If Not mb_FilaEditable Then
    '     Exit Sub
    ' End If
    
    If oGrilla.ActiveCell Is Nothing Then Exit Sub
    
    If oGrilla.ActiveCell.Column.Key = "Codigo" And KeyAscii = 13 Then
      SendKeys "{Tab}"
      If Trim(oGrilla.ActiveCell.GetText) <> "" Then
      SendKeys "{Tab}"
      'SendKeys "{Tab}"
      End If
      Exit Sub
    End If
    If oGrilla.ActiveCell.Column.Key = "Cantidad" Then
      If KeyAscii = 13 Then
        If ml_HabilitaIngresoDePrecio = True Then
          SendKeys "{Tab}"
        Else
          If lbEstoyEnGridCPT = True Then
            mnuAgregarServicio_Click
            grdProductos_Click
            mnuAgregarInsumo_Click
            If optPorCodigo.Value = True Then
               grdProductosFocusColumna "codigo"
            Else
               grdProductosFocusColumna "NombreProducto"
            End If
            
          Else
            mnuAgregarInsumo_Click
          End If
        End If
      End If
      Exit Sub
    End If
    If oGrilla.ActiveCell.Column.Key = "PrecioUnitario" Then
      If KeyAscii = 13 Then
        If ml_HabilitaIngresoDePrecio = True Then
          mnuAgregarServicio_Click
        Else
          SendKeys "{Tab}"
        End If
      End If
      Exit Sub
    End If
    
    If oGrilla.ActiveCell.Column.Key = "NombreProducto" Then
      Select Case KeyAscii
        Case vbKeyEscape
          If Trim(oGrilla.ActiveCell.GetText) = "" Then
            grillaBusqueda.Visible = False
            Set grillaBusqueda.DataSource = Nothing
          End If
        Case vbKeyReturn
        Case vbKeyDown
        Case vbKeyLeft
        Case Else
          Dim lIdTipoFinanciamiento As Long
          Dim sNombre As String
          Select Case KeyAscii
            Case vbKeyBack
              sNombre = oGrilla.ActiveCell.GetText
            Case Else
              sNombre = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
          End Select
          
          lIdTipoFinanciamiento = oGrilla.ActiveCell.Row.Cells("IdTipoFinanciamiento").Value
          Dim rs As New Recordset
                    
          Select Case ms_TipoProducto
            Case sghServicio
              If lbEstoyEnGridCPT = True Then
                Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga, sghSoloCPT)
              Else
                'Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga, sghSoloInsumos)
                Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga(sghPorDescripcion, sNombre, ml_IdPuntoCarga, sghSoloInsumos)
              End If
            Case Else
                        
          End Select
          Set grillaBusqueda.DataSource = rs
          grillaBusqueda.Left = oGrilla.Left
          If lbEstoyEnGridCPT = True Then
            If mrs_FacturacionProductos.RecordCount < 4 Then
               grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.Rect.Bottom * Screen.TwipsPerPixelY
            Else
               grillaBusqueda.Top = 0
            End If
          Else
            grillaBusqueda.Top = grdProductos.Top
          End If
          grillaBusqueda.Visible = True
          grillaBusqueda.Enabled = True
                    
      End Select
    End If
  End If
End Sub

'WILLIAM CASTRO
Sub GenerarRecordsetProductos()
  With mrs_FacturacionProductos
    .Fields.Append "IdFacturacionProducto", adInteger
    .Fields.Append "IdProducto", adInteger
    .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
    .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
    .Fields.Append "IdTipoFinanciamiento", adInteger
    .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
    .Fields.Append "Poliza", adVarChar, 255
    .Fields.Append "TipoFinanciamiento", adVarChar, 255
    .Fields.Append "Cantidad", adInteger
    .Fields.Append "PrecioUnitario", adCurrency
    .Fields.Append "TotalPorPagar", adCurrency
    .Fields.Append "ResultadoAutomatico", adBoolean
    .Fields.Append "CantidadFallada", adInteger
    .Fields.Append "IdEstadoFacturacion", adInteger
    .Fields.Append "IdPuntoCarga", adInteger
    .Fields.Append "IdAtencion", adInteger, , adFldIsNullable
    .Fields.Append "IdCajero", adInteger, , adFldIsNullable
    .Fields.Append "FechaAutorizaPendiente", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "FechaAutorizaSeguro", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "FechaAutorizaDevolucion", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "FechaCajero", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "IdUsuarioAutorizaPendiente", adInteger, , adFldIsNullable
    .Fields.Append "IdUsuarioAutorizaSeguro", adInteger, , adFldIsNullable
    .Fields.Append "IdUsuarioAutorizaDevolucion", adInteger, , adFldIsNullable
    .Fields.Append "IdServicioInternamiento", adInteger, , adFldIsNullable
    .Fields.Append "IdUsuarioAuditoria", adInteger, , adFldIsNullable
    .Fields.Append "EstadoLocal", adVarChar, 1
    .Fields.Append "IdComprobantePago", adInteger, , adFldIsNullable
    .Fields.Append "IdComprobantePagoDevolucion", adInteger, , adFldIsNullable
    .Fields.Append "IdOrden", adInteger
    .Fields.Append "movTipo", adVarChar, 1, adFldIsNullable
    .Fields.Append "movNumero", adVarChar, 9, adFldIsNullable
    .Fields.Append "SeUsaSinPrecio", adBoolean
    .Fields.Append "CantidadSinEditar", adInteger
    .Fields.Append "ObsReceta", adVarChar, 300, adFldIsNullable
    
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  With mrs_FacturacionInsumos
    .Fields.Append "IdFacturacionProducto", adInteger
    .Fields.Append "IdProducto", adInteger
    .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
    .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
    .Fields.Append "IdTipoFinanciamiento", adInteger
    .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
    .Fields.Append "Poliza", adVarChar, 255
    .Fields.Append "TipoFinanciamiento", adVarChar, 255
    .Fields.Append "Cantidad", adInteger
    .Fields.Append "CantidadFallada", adInteger
    .Fields.Append "PrecioUnitario", adCurrency
    .Fields.Append "TotalPorPagar", adCurrency
    .Fields.Append "IdEstadoFacturacion", adInteger
    .Fields.Append "IdPuntoCarga", adInteger
    .Fields.Append "IdAtencion", adInteger, , adFldIsNullable
    .Fields.Append "IdCajero", adInteger, , adFldIsNullable
    .Fields.Append "FechaAutorizaPendiente", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "FechaAutorizaSeguro", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "FechaAutorizaDevolucion", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "FechaCajero", adDBTimeStamp, , adFldIsNullable
    .Fields.Append "IdUsuarioAutorizaPendiente", adInteger, , adFldIsNullable
    .Fields.Append "IdUsuarioAutorizaSeguro", adInteger, , adFldIsNullable
    .Fields.Append "IdUsuarioAutorizaDevolucion", adInteger, , adFldIsNullable
    .Fields.Append "IdServicioInternamiento", adInteger, , adFldIsNullable
    .Fields.Append "IdUsuarioAuditoria", adInteger, , adFldIsNullable
    .Fields.Append "EstadoLocal", adVarChar, 1
    .Fields.Append "IdComprobantePago", adInteger, , adFldIsNullable
    .Fields.Append "IdComprobantePagoDevolucion", adInteger, , adFldIsNullable
    .Fields.Append "IdOrden", adInteger
    .Fields.Append "movTipo", adVarChar, 1, adFldIsNullable
    .Fields.Append "movNumero", adVarChar, 9, adFldIsNullable
    .Fields.Append "IdProductoCPT", adInteger
    .Fields.Append "SeUsaSinPrecio", adBoolean
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
  oGrilla.Bands(0).Columns("IdFacturacionProducto").Hidden = True
  oGrilla.Bands(0).Columns("SeUsaSinPrecio").Hidden = True
  oGrilla.Bands(0).Columns("idProducto").Hidden = True
  oGrilla.Bands(0).Columns("TipoFinanciamiento").Hidden = True
  oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
  oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Hidden = True
  oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Hidden = True
  oGrilla.Bands(0).Columns("IdUsuarioAutorizaPendiente").Hidden = True
  oGrilla.Bands(0).Columns("IdUsuarioAutorizaSeguro").Hidden = True
  oGrilla.Bands(0).Columns("IdFuenteFinanciamiento").Hidden = True
  oGrilla.Bands(0).Columns("IdServicioInternamiento").Hidden = True
  oGrilla.Bands(0).Columns("IdUsuarioAuditoria").Hidden = True
  oGrilla.Bands(0).Columns("Poliza").Hidden = True
  oGrilla.Bands(0).Columns("EstadoLocal").Hidden = True
  oGrilla.Bands(0).Columns("IdCajero").Hidden = True
  oGrilla.Bands(0).Columns("FechaCajero").Hidden = True
  oGrilla.Bands(0).Columns("IdUsuarioAutorizaDevolucion").Hidden = True
  oGrilla.Bands(0).Columns("FechaAutorizaDevolucion").Hidden = True
  oGrilla.Bands(0).Columns("IdComprobantePago").Hidden = True
  oGrilla.Bands(0).Columns("IdComprobantePagoDevolucion").Hidden = True
  oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Hidden = True
  oGrilla.Bands(0).Columns("CantidadFallada").Hidden = True
  oGrilla.Bands(0).Columns("idOrden").Hidden = True
  oGrilla.Bands(0).Columns("movTipo").Hidden = True
  oGrilla.Bands(0).Columns("movNumero").Hidden = True
  
  
  oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
  oGrilla.Bands(0).Columns("Codigo").Width = 750
  oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationAllowEdit
  
  oGrilla.Bands(0).Columns("NombreProducto").Header.Caption = "Descripción"
  If lbEstoyEnGridCPT = True Then
    oGrilla.Bands(0).Columns("NombreProducto").Width = 7200
    oGrilla.Bands(0).Columns("CantidadSinEditar").Hidden = True
    oGrilla.Bands(0).Columns("obsReceta").Width = 7200
  Else
    oGrilla.Bands(0).Columns("NombreProducto").Width = 9700
  End If
  oGrilla.Bands(0).Columns("NombreProducto").Activation = ssActivationAllowEdit
    
  oGrilla.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
  oGrilla.Bands(0).Columns("Cantidad").Format = "###0"
  oGrilla.Bands(0).Columns("Cantidad").Width = 1000
  oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
    
  oGrilla.Bands(0).Columns("preciounitario").Header.Caption = "P.U.(S/.)"
  oGrilla.Bands(0).Columns("preciounitario").Format = "#0.000"
  oGrilla.Bands(0).Columns("preciounitario").Width = "1000"
  
  oGrilla.Bands(0).Columns("TotalPorPagar").Header.Caption = "Sub Total"
  oGrilla.Bands(0).Columns("TotalPorPagar").Format = "#0.00"
  oGrilla.Bands(0).Columns("TotalPorPagar").Activation = ssActivationActivateNoEdit
  oGrilla.Bands(0).Columns("TotalPorPagar").Width = 1200
  
  oGrilla.Bands(0).Columns("IdEstadoFacturacion").Width = 1500
  oGrilla.Bands(0).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
  oGrilla.Bands(0).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList
  oGrilla.Bands(0).Columns("IdEstadoFacturacion").Hidden = True

  oGrilla.Bands(0).Columns("idPuntoCarga").Header.Caption = "Puntos de carga"
  oGrilla.Bands(0).Columns("idPuntoCarga").Width = 1500
  oGrilla.Bands(0).Columns("idPuntoCarga").Style = ssStyleDropDownList
  oGrilla.Bands(0).Columns("idPuntoCarga").Hidden = True
  
  oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Width = 2500
  oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Aut. Pend."
  oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Format = sighentidades.DevuelveFechaSoloFormato_DMY_HM

  oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Width = 2500
  oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Header.Caption = "Fec. Aut. Seguro."
  oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Format = sighentidades.DevuelveFechaSoloFormato_DMY_HM
    
  'Configura Values List
  SeteaListaEstado oGrilla, oGrilla.Bands(0).Columns("idEstadoFacturacion")
  SeteaListaTipoFinanciamiento oGrilla, oGrilla.Bands(0).Columns("IdTipoFinanciamiento")
  SeteaPuntosDeCarga oGrilla, oGrilla.Bands(0).Columns("idPuntoCarga")
  If ml_HabilitaIngresoDePrecio = True Then
    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
  Else
    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
  End If
  If ml_PermiteVerColumnaCantidadFallada = True Then
    'oGrilla.Bands(0).Columns("cantidadFallada").Hidden = False
    oGrilla.Bands(0).Columns("cantidadFallada").Hidden = True
  Else
    oGrilla.Bands(0).Columns("cantidadFallada").Hidden = True
  End If
  oGrilla.Bands(0).Columns("idPuntoCarga").Activation = ssActivationActivateNoEdit
  oGrilla.Bands(0).Columns("idEstadoFacturacion").Activation = ssActivationActivateNoEdit
  oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Activation = ssActivationActivateNoEdit
    
ConfigEstilo:
  gridInfra.ConfigurarFilasBiColores oGrilla, sighentidades.GrillaConFilasBicolor
End Sub

Sub SeteaListaTipoFinanciamiento(oGrilla As SSUltraGrid, oColumn As SSColumn)
  Dim rs As New ADODB.Recordset
  Dim I As Integer
  Dim oValueTF As SSValueList
    
  If Not oGrilla.ValueLists.Exists("listaTipoFinanciamiento") Then
    Set oValueTF = oGrilla.ValueLists.Add("listaTipoFinanciamiento")
    Set rs = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarTodos
    Do While Not rs.EOF
      If rs!IdTipoFinanciamiento <> 0 Then oValueTF.ValueListItems.Add Val(rs!IdTipoFinanciamiento), Trim(rs!descripcion)
      rs.MoveNext
    Loop
    rs.Close
  Else
    Set oValueTF = oGrilla.ValueLists.Item("listaTipoFinanciamiento")
  End If
  Set oColumn.ValueList = oValueTF
End Sub

Sub SeteaPuntosDeCarga(oGrilla As SSUltraGrid, oColumn As SSColumn)
  Dim rs As New ADODB.Recordset
  Dim I As Integer
  Dim oValuePC As SSValueList
    
  If Not oGrilla.ValueLists.Exists("listaPuntosCarga") Then
    Set oValuePC = oGrilla.ValueLists.Add("listaPuntosCarga")
    Set rs = mo_reglasComunes.SeleccionarPuntosDeCarga()
    Do While Not rs.EOF
      If rs!IdPuntoCarga <> 0 Then oValuePC.ValueListItems.Add Val(rs!IdPuntoCarga), Trim(rs!descripcion)
      rs.MoveNext
    Loop
    rs.Close
  Else
    Set oValuePC = oGrilla.ValueLists.Item("listaPuntosCarga")
  End If
  Set oColumn.ValueList = oValuePC
End Sub

Sub SeteaListaEstado(oGrilla As SSUltraGrid, oColumn As SSColumn)
  Dim rs As ADODB.Recordset
  Dim I As Integer
  Dim oValueEstado As SSValueList
    
  If Not oGrilla.ValueLists.Exists("listaEstadoFacturacion") Then
    Set oValueEstado = oGrilla.ValueLists.Add("listaEstadoFacturacion")
    Set rs = mo_ReglasFacturacion.EstadosFacturacionObtenerTodos
    Do While Not rs.EOF
      oValueEstado.ValueListItems.Add Val(rs!IdEstadoFacturacion), Trim(rs!descripcion)
      rs.MoveNext
    Loop
    rs.Close
  Else
    Set oValueEstado = oGrilla.ValueLists.Item("listaEstadoFacturacion")
  End If
  Set oColumn.ValueList = oValueEstado
End Sub

Private Sub grillaBusqueda_Click()
  'RefrescarDatos
End Sub

Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  InicializarLaGrillaBusqueda grillaBusqueda
  gridInfra.ConfigurarFilasBiColores grillaBusqueda, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
  oGrilla.Bands(0).Columns("IdProducto").Hidden = True
  If lbEstoyEnGridCPT = True Then
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    'oGrilla.Bands(0).Columns("preciounitario").Hidden = True
  End If
  oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
  oGrilla.Bands(0).Columns("Codigo").Width = 800
  oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
  oGrilla.Bands(0).Columns("Nombre").Width = 7000
  oGrilla.Bands(0).Columns("idPuntoCarga").Hidden = True
  oGrilla.Bands(0).Columns("NombreProducto").Hidden = True
  oGrilla.Bands(0).Columns("SeUsaSinPrecio").Hidden = True
  oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
  oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
  oGrilla.Bands(0).Columns("preciounitario").Format = "0.00"
  oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
  gridInfra.ConfigurarFilasBiColores oGrilla, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub grillaBusqueda_DblClick()
  Dim fila As New Record
  Dim fechaDespacho As Date
  Dim rsTmp1 As ADODB.Recordset
    
  Set rsTmp1 = grillaBusqueda.DataSource
  If rsTmp1.EOF = True And rsTmp1.BOF = True Then Exit Sub
    
  If ItemYaExiste(grillaBusqueda.ActiveRow.Cells("idproducto").Value) Then
    If lbEstoyEnGridCPT = True Then
      grdProductos.ActiveRow.Cells("codigo").Value = ""
      grdProductos.ActiveRow.Cells("idproducto").Value = 0
      grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
    Else
      grdInsumos.ActiveRow.Cells("codigo").Value = ""
      grdInsumos.ActiveRow.Cells("idproducto").Value = 0
      grdInsumos.ActiveRow.Cells("NombreProducto").Value = ""
    End If
  Else
    fechaDespacho = DateAdd("h", ml_ParametroHoras, Now)
    If idCuentaAtencion = 0 Then GoTo Sigue
    If BuscaItemEnDia(idCuentaAtencion, Format(fechaDespacho, sighentidades.DevuelveFechaSoloFormato_DMY_HMS), grillaBusqueda.ActiveRow.Cells("idproducto").Value) = True Then
Sigue:
      RefrescarDatos
      Set grillaBusqueda.DataSource = Nothing
      grillaBusqueda.Visible = False
      If lbEstoyEnGridCPT = False Then grdInsumos.SetFocus
      SendKeys "{Tab}"
    End If
'    mnuAgregarServicio_Click
    'SendKeys "{Tab}"
  End If
End Sub

Sub RefrescarDatos()
  Dim fila As New Record
  Dim lnPrecioUnitario  As Double
  If Not grillaBusqueda.ActiveRow Is Nothing Then
    If lbEstoyEnGridCPT = True Then
      'lnPrecioUnitario = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
      lnPrecioUnitario = 0
      oDoCatalogoServicioHosp.PrecioUnitario = 0
      Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(grillaBusqueda.ActiveRow.Cells("idproducto").Value, ml_IdTipoFinanciamiento)
      If oDoCatalogoServicioHosp.PrecioUnitario = 0 And oDoCatalogoServicioHosp.SeUsaSinPrecio = False Then
        MsgBox "Ese Producto no tiene precio para el TIPO DE  FINANCIAMIENTO", vbExclamation, "Facturación"
      Else
        lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
      End If
      'If lnPrecioUnitario > 0 Then
        grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
        grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
        grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
        grdProductos.ActiveRow.Cells("preciounitario").Value = lnPrecioUnitario
        grdProductos.ActiveRow.Cells("TotalPorPagar").Value = lnPrecioUnitario
        grdProductos.ActiveRow.Cells("cantidad").Value = 1
        grdProductos.ActiveRow.Cells("idestadofacturacion").Value = 1
        If Not IsNull(grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value) Then
           grdProductos.ActiveRow.Cells("SeUsaSinPrecio").Value = grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value
        End If
        If grillaBusqueda.ActiveRow.Cells("LabResultadoAutomatico").Value = 1 Then
           grdProductos.ActiveRow.Cells("ResultadoAutomatico").Value = True
        End If
        Totalizar
      'End If
    Else
      grdInsumos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
      grdInsumos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
      grdInsumos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
      grdInsumos.ActiveRow.Cells("preciounitario").Value = 1
      grdInsumos.ActiveRow.Cells("TotalPorPagar").Value = 1
      grdInsumos.ActiveRow.Cells("cantidad").Value = 1
      grdInsumos.ActiveRow.Cells("idestadofacturacion").Value = 1
    End If
  End If
End Sub

Private Sub grillaBusqueda_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
    Case vbKeyReturn
        grillaBusqueda_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
  End Select
End Sub

Private Sub mnuAgregarInsumo_Click()
  'If lnIdProductoCPT > 0 Then
  '  grdInsumos.SetFocus
  '  SendKeys "{Tab}"
  '  AgregaProductoInsumo
  'End If
End Sub

Private Sub mnuAgregarServicio_Click()
  If ml_PermiteAgregarItems = True Then
    'grdProductos.SetFocus
    'SendKeys "{Tab}"
    AgregaProducto
  End If
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  grdProductos.Top = 0
  grdProductos.Left = 10
  grdProductos.Width = UserControl.Width - 120
  grdProductos.Height = UserControl.Height - Label1.Height  '1350 'UserControl.Height - UserControl.Label1.Height - 5
  grdInsumos.Top = grdProductos.Top + grdProductos.Height + 30
  grdInsumos.Left = 10
  grdInsumos.Width = UserControl.Width - 120
  grdInsumos.Height = UserControl.Height - grdProductos.Height - UserControl.Label1.Height - 5
  Label1.Top = UserControl.Height - UserControl.Label1.Height
  lblTotal.Top = UserControl.Height - UserControl.Label1.Height
  optPorCodigo.Top = UserControl.Height - UserControl.Label1.Height - 30
  optPorDescripcion.Top = UserControl.Height - UserControl.Label1.Height - 30
  chkTodosNinguno.Top = optPorDescripcion.Top + optPorDescripcion.Width + 30
  ml_ParametroHoras = -1 * Val(lcBuscaParametro.SeleccionaFilaParametro(248))
End Sub

Sub LimpiarGrilla()
  If mrs_FacturacionProductos Is Nothing Then Exit Sub
  Set grdProductos.DataSource = Nothing
  If mrs_FacturacionProductos.RecordCount > 0 Then
    mrs_FacturacionProductos.MoveFirst
    Do While Not mrs_FacturacionProductos.EOF
      mrs_FacturacionProductos.Delete
      mrs_FacturacionProductos.Update
      mrs_FacturacionProductos.MoveNext
    Loop
  End If
  
  If Not (mrs_FacturacionInsumos Is Nothing) Then
    Set grdInsumos.DataSource = Nothing
    If mrs_FacturacionInsumos.RecordCount > 0 Then
      mrs_FacturacionInsumos.MoveFirst
      Do While Not mrs_FacturacionInsumos.EOF
        mrs_FacturacionInsumos.Delete
        mrs_FacturacionInsumos.Update
        mrs_FacturacionInsumos.MoveNext
      Loop
    End If
  End If
  
  ml_idOrden = -1000  'Esto es aproposito para que obtenga solo la estructura del recordset
  CargaProductosPorIdOrden
  grillaBusqueda.Visible = False
End Sub

Sub TabEnDescripcion()
    On Error Resume Next
    mnuAgregarServicio_Click
    If optPorCodigo.Value = True Then
       grdProductosFocusColumna "codigo"
    Else
       grdProductosFocusColumna "NombreProducto"
    End If
End Sub

'============================== Adams BONILLA MAGALLANES ==============================
'====== Verifica si ITEM ya fue registrado el mismo día (al momento de registro) ======
'======================================================================================
Function BuscaItemEnDia(idCuentaAtencion As Long, fechaDespacho As Date, idProductoCPT As Long) As Boolean
  Dim oRsTmp As New ADODB.Recordset
  BuscaItemEnDia = False
  If lbEstoyEnGridCPT = True Then
    Set oRsTmp = mo_ReglasLaboratorio.LaboratorioYaRegistroPrueba(idCuentaAtencion, fechaDespacho, idProductoCPT)
    If oRsTmp.RecordCount > 0 Then
      If MsgBox("La prueba que desea agregar ya fue agregada al paciente en las últimas " & (-1 * (ml_ParametroHoras)) & " horas." & Chr(13) & "Esta seguro que quiere agregar la prueba de todas maneras", vbQuestion + vbOKCancel, "SIGH ") = vbOK Then BuscaItemEnDia = True
    Else
      BuscaItemEnDia = True
    End If
    oRsTmp.Close
  End If
End Function

Sub grdProductosFocusColumna(lcNombreColumna As String)
    With grdProductos
        'scroll the column into view
        .ActiveColScrollRegion.ScrollColumnIntoView .Bands(0).Columns(lcNombreColumna), True
        If Not .ActiveRow Is Nothing Then
            'if there is an activerow then activate the cell from this column
            Set .ActiveCell = .ActiveRow.Cells(lcNombreColumna)
            .ActiveCell.Selected = True
        End If
        'give the grid focus
        .SetFocus
        .PerformAction ssKeyActionActivateCell
        .PerformAction ssKeyActionEnterEditMode
    End With
End Sub

Sub CargaProductosPorIdCita(lnIdCitaSI As Long)
    Dim rs As Recordset
    Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
    Set rs = mo_ReglasImagenes.SiCitasDetallePorIdentificador(lnIdCitaSI)
    CargaProductosPorIdReceta rs
    Set mo_ReglasImagenes = Nothing
End Sub
'<Agregado por: WABG el: 11/29/2020-12:44:07 en el equipo: SISGALENPLUS-PC><CAMBIO 44>
Sub CargarItemsALaGrillaPaquete(rs As Recordset)

  Dim oRsTmp1 As New Recordset
  Dim lnSubTotal As Double
  mb_CargandoProductos = True
  Do While Not rs.EOF
    Set oRsTmp1 = rs
    If oRsTmp1.RecordCount > 0 Then
    'codigo/descripcion/cantidad/p.u./subtotal=totalporpagar/resultadoautomatico/ObsReceta
      mrs_FacturacionProductos.AddNew
      mrs_FacturacionProductos!idProducto = rs!idProducto
      mrs_FacturacionProductos!Codigo = rs!Codigo
      mrs_FacturacionProductos!NombreProducto = rs!Nombre
      mrs_FacturacionProductos!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
      mrs_FacturacionProductos!Cantidad = 1
      mrs_FacturacionProductos!PrecioUnitario = rs!PrecioUnitario
      mrs_FacturacionProductos!TotalPorPagar = rs!PrecioUnitario
      mrs_FacturacionProductos!ResultadoAutomatico = IIf(rs!LabResultadoAutomatico = 1, True, False)
    End If
     rs.MoveNext
   Loop
    mb_CargandoProductos = False
    
    Set rs = Nothing
    Totalizar
  Set grdProductos.DataSource = mrs_FacturacionProductos
 
 
End Sub
'</Agregado por: WABG el: 11/29/2020-12:44:07 en el equipo: SISGALENPLUS-PC><CAMBIO 44>

'Actualizado 16092014 Frank
Sub CargaProductosPorIdReceta(rs As Recordset)
    Dim rsCatHos As New Recordset
    Dim lbNuevoProducto As Boolean
    mb_CargandoProductos = True
    
    Do While Not rs.EOF
        Set rsCatHos = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(rs!idItem, ml_IdTipoFinanciamiento)
        
            lbNuevoProducto = True
            If mrs_FacturacionProductos.RecordCount > 0 Then
               mrs_FacturacionProductos.MoveFirst
               mrs_FacturacionProductos.Find "idProducto=" & rs!idItem
               If Not mrs_FacturacionProductos.EOF Then
                  lbNuevoProducto = False
               End If
            End If
            If lbNuevoProducto = True Then
                mrs_FacturacionProductos.AddNew
                mrs_FacturacionProductos!idProducto = rs!idItem
                mrs_FacturacionProductos!Codigo = rs!Codigo
                mrs_FacturacionProductos!NombreProducto = rs!Nombre
                mrs_FacturacionProductos!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
                mrs_FacturacionProductos!Cantidad = rs!CantidadPedida
                mrs_FacturacionProductos!PrecioUnitario = rs!precio
                mrs_FacturacionProductos!TotalPorPagar = rs!Total
                If rsCatHos.RecordCount > 0 Then
                   mrs_FacturacionProductos!SeUsaSinPrecio = IIf(IsNull(rsCatHos!SeUsaSinPrecio), False, rsCatHos!SeUsaSinPrecio)
                   If rsCatHos!LabResultadoAutomatico = 1 Then
                      mrs_FacturacionProductos!ResultadoAutomatico = True
                   End If
                End If
                mrs_FacturacionProductos!idOrden = 0 'rs!IdOrden
                mrs_FacturacionProductos!ObsReceta = rs!observaciones
            End If

        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
    End If
    Set rsCatHos = Nothing
    Totalizar
    Set grdProductos.DataSource = mrs_FacturacionProductos
End Sub
Sub CargaObservacionesDeReceta(lnIdReceta As Long, oConexion As Connection)
    If lnIdReceta > 0 Then
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_reglasComunes.RecetaDetalleSeleccioarPorIdReceta(lnIdReceta, oConexion)
        oRsTmp1.Filter = "observaciones<>''"
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If Not IsNull(oRsTmp1!observaciones) Then
                 mrs_FacturacionProductos.MoveFirst
                 mrs_FacturacionProductos.Find "idProducto=" & oRsTmp1!idItem
                 If Not mrs_FacturacionProductos.EOF Then
                    mrs_FacturacionProductos!ObsReceta = oRsTmp1!observaciones
                    mrs_FacturacionProductos.Update
                 End If
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
    End If
End Sub

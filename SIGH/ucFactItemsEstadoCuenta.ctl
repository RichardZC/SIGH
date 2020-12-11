VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFactItemsEstadoCuenta 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   LockControls    =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   11685
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5685
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10028
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Productos"
   End
   Begin VB.Menu mnuProductos 
      Caption         =   "mnuProductos"
      Begin VB.Menu mnuAgregarServicio 
         Caption         =   "Agregar servicio"
      End
      Begin VB.Menu mnuAgregarExoneracion 
         Caption         =   "Agregar exoneración"
      End
      Begin VB.Menu mnuAgregarPagoACuenta 
         Caption         =   "Agregar pago a cuenta"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutorizaPacienteNormal 
         Caption         =   "Paciente Normal"
      End
      Begin VB.Menu mnuAutorizarSIS 
         Caption         =   "Autorizado por SIS"
      End
      Begin VB.Menu mnuAutorizarSOAT 
         Caption         =   "Autorizado por SOAT"
      End
      Begin VB.Menu mnuAutorizarConvenio 
         Caption         =   "Autorizado por Convenio"
      End
      Begin VB.Menu mnuAutorizarPendientePago 
         Caption         =   "Autorizar pendiente de pago"
      End
      Begin VB.Menu mnuAutorizarDevolucion 
         Caption         =   "Autorizar devolución"
      End
   End
End
Attribute VB_Name = "ucFactItemsEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mostras Procedimientos/Medicamentos de una cuenta
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Public Event Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, _
                        TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, _
                        dTotalDevuelto As Double, dTotalAnulado As Double, lbTieneExoneracion As Boolean)
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim gridInfra As New GridInfragistic
Dim mo_PermisosFacturacion As New PermisosFacturacion

Dim ms_TipoProducto As sghTipoProducto
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idOrden As Long
Dim ml_idCuentaAtencion As Long
Dim mb_CargandoProductos As Boolean
Dim ms_Opcion As sghOpciones
Dim mrs_FacturacionProductos As New Recordset
Dim mo_DoAtencion As DOAtencion
Dim ml_idUsuario As Long
Dim ml_IdPuntoCarga As Long
Dim ms_EstadosFacturacion As String
Dim ms_TiposFinanciamiento As String
Dim ml_AgruparPor As Long
Dim ml_idUsuarioConPermisoEnSISoEXOoSOAT As Long

'edicion de la grilla
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim ml_EstadoCuentaAtencion As Long
Dim lcIdServiciosEstanciaEnParametros As String
Dim ms_MovNumero As String
Dim lnIdPagosACuenta     As Long
Dim ms_MensajeError As String
Dim lnIdDevoluciones As Long
Dim lcSql As String
Dim oConexion As New ADODB.Connection
Dim mb_ProcesoEnElServidor As Boolean
Dim ml_Paciente As String
Dim ml_IdPaciente As Long
Dim ml_idTipoSexo As Long
Dim ml_lbPuedeVerResultados As Boolean
Dim lbTieneDerechoExoneraSIS As Boolean
Dim lbTieneQueGrabarAntesDeImprimir As Boolean


Property Get TieneQueGrabarAntesDeImprimir() As Long
    TieneQueGrabarAntesDeImprimir = lbTieneQueGrabarAntesDeImprimir
End Property


Property Let TieneDerechoExoneraSIS(lValue As Boolean)
  lbTieneDerechoExoneraSIS = lValue
End Property


Property Let PuedeVerResultados(lValue As Boolean)
  ml_lbPuedeVerResultados = lValue
End Property
Property Let idPaciente(lValue As Long)
  ml_IdPaciente = lValue
End Property
Property Let Paciente(lValue As String)
  ml_Paciente = lValue
End Property
Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property




Property Let ProcesoEnElServidor(sValue As Boolean)
    mb_ProcesoEnElServidor = sValue
End Property


Property Let movNumero(sValue As String)
    ms_MovNumero = sValue
End Property
Property Let idUsuarioConPermisoEnSISoEXOoSOAT(lValue As Long)
    ml_idUsuarioConPermisoEnSISoEXOoSOAT = lValue
End Property
Property Get idUsuarioConPermisoEnSISoEXOoSOAT() As Long
    idUsuarioConPermisoEnSISoEXOoSOAT = ml_idUsuarioConPermisoEnSISoEXOoSOAT
End Property

Property Let IdOrden(lValue As Long)
    ml_idOrden = lValue
End Property
Property Get IdOrden() As Long
    IdOrden = ml_idOrden
End Property

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
    idCuentaAtencion = ml_idCuentaAtencion
End Property

Property Let EstadoCuentaAtencion(lValue As Long)
    ml_EstadoCuentaAtencion = lValue
End Property
Property Get EstadoCuentaAtencion() As Long
    EstadoCuentaAtencion = ml_EstadoCuentaAtencion
End Property

Property Set Atencion(oValue As DOAtencion)
    Set mo_DoAtencion = oValue
End Property
Property Get Atencion() As DOAtencion
    Set Atencion = mo_DoAtencion
End Property

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property

Property Let TipoProducto(iTipo As sghTipoProducto)
    ms_TipoProducto = iTipo
End Property

Property Get TipoProducto() As sghTipoProducto
    TipoProducto = ms_TipoProducto
End Property

Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Get idTipoFinanciamiento() As Long
    idTipoFinanciamiento = ml_IdTipoFinanciamiento
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

Property Set ProductosEliminados(oValue As Collection)
    Set mo_ProductosEliminados = oValue
End Property

Property Get ProductosEliminados() As Collection
    Set ProductosEliminados = mo_ProductosEliminados
End Property

Property Let idPuntoCarga(lValue As Long)
    ml_IdPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
    idPuntoCarga = ml_IdPuntoCarga
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

Property Let AgruparPor(lTipo As Long)
    ml_AgruparPor = lTipo
End Property

Property Get AgruparPor() As Long
    AgruparPor = ml_AgruparPor
End Property


Sub Inicializar()
    If oConexion.State = 0 Then
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 150
    End If
    Set mrs_FacturacionProductos = New Recordset
    GenerarRecordsetProductos
    
    ms_EstadosFacturacion = ""
    Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
    UserControl.mnuAgregarServicio.Enabled = mo_PermisosFacturacion.AgregarServicios
    UserControl.mnuAgregarExoneracion.Enabled = mo_PermisosFacturacion.AgregarExoneraciones
    UserControl.mnuAutorizarSIS.Enabled = mo_PermisosFacturacion.AutorizarSIS
    UserControl.mnuAutorizarSOAT.Enabled = mo_PermisosFacturacion.AutorizarSOAT
    UserControl.mnuAutorizarPendientePago.Enabled = mo_PermisosFacturacion.AutorizarPendientesDePago
    UserControl.mnuAutorizarConvenio.Enabled = mo_PermisosFacturacion.AutorizarConvenios
    UserControl.mnuAutorizarDevolucion.Enabled = mo_PermisosFacturacion.AutorizarDevoluciones
    '
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    lcIdServiciosEstanciaEnParametros = lcBuscaParametro.SeleccionaFilaParametro(202) & "."
    lcIdServiciosEstanciaEnParametros = lcIdServiciosEstanciaEnParametros & lcBuscaParametro.SeleccionaFilaParametro(203) & "."
    lcIdServiciosEstanciaEnParametros = lcIdServiciosEstanciaEnParametros & lcBuscaParametro.SeleccionaFilaParametro(204)
    '
    lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
    lnIdDevoluciones = Val(lcBuscaParametro.SeleccionaFilaParametro(265))
    wxParametro511 = lcBuscaParametro.SeleccionaFilaParametro(511)
    wxParametro514 = lcBuscaParametro.SeleccionaFilaParametro(514)
    wxParametro554 = lcBuscaParametro.SeleccionaFilaParametro(554)
    wxparametro563 = lcBuscaParametro.SeleccionaFilaParametro(563)
    wxparametro564 = lcBuscaParametro.SeleccionaFilaParametro(564)
    wxparametro565 = lcBuscaParametro.SeleccionaFilaParametro(565)
    wxparametro566 = lcBuscaParametro.SeleccionaFilaParametro(566)
End Sub

Sub CargaProductosPorIdCuentaAtencion(ByRef lnTotalPagoSeguro As Double, ByRef lnTotalPagoDelPaciente As Double, ByRef lnTotalizaPagosDelPacienteConSeguro As Double, ByRef oRsCuentaCabecera As Recordset, ByRef oRsCuentaDetalle As Recordset, lnIdTipoConceptoFarmaciaPlanActual As Integer, ByRef lnTotalApagar As Double)
    lnTotalPagoSeguro = 0: lnTotalPagoDelPaciente = 0: lnTotalizaPagosDelPacienteConSeguro = 0: lnTotalApagar = 0
    If ml_idCuentaAtencion <= 0 Then
       Exit Sub
    End If
    Dim rs As New Recordset
    Select Case ms_TipoProducto
    Case sghServicio
            Set rs = mo_ReglasFacturacion.FacturacionServicioDespachoXcuenta(ml_idCuentaAtencion)
            rs.Filter = "idEstadoFacturacion<>9 and idEstadoFacturacion<>12"
            CargarItemsALaGrillaS_menorTiempo rs, lnTotalPagoSeguro, lnTotalPagoDelPaciente, _
                                              lnTotalizaPagosDelPacienteConSeguro, oRsCuentaCabecera, _
                                              oRsCuentaDetalle, lnIdTipoConceptoFarmaciaPlanActual, _
                                              lnTotalApagar
    Case sghbien
            Set rs = mo_ReglasFarmacia.farmMovimientoVentasDetalleXcuenta(ml_idCuentaAtencion)
            CargarItemsALaGrillaB_menorTiempo rs, lnTotalPagoSeguro, lnTotalPagoDelPaciente, _
                                              lnTotalizaPagosDelPacienteConSeguro, oRsCuentaCabecera, _
                                              oRsCuentaDetalle, lnIdTipoConceptoFarmaciaPlanActual, _
                                              lnTotalApagar
    End Select
    If ml_AgruparPor = 1 Then
        Totalizar
    End If
    Set grdProductos.DataSource = mrs_FacturacionProductos
End Sub

Sub CargaProductosPorIdOrden()
    Dim rs As Recordset
    Select Case ms_TipoProducto
    Case sghServicio
        Set rs = mo_ReglasFacturacion.FacturacionServicioDespachoFiltraPorIdOrden(ml_idOrden, oConexion)
        rs.Filter = "idEstadoFacturacion<>9"
        CargarItemsALaGrillaS rs
    Case sghbien
    End Select
    If ml_AgruparPor = 1 Then
        Totalizar
    End If
    Set grdProductos.DataSource = mrs_FacturacionProductos
End Sub

Sub CargaProductosPorMovNumero()
    Dim rs As Recordset
    Select Case ms_TipoProducto
    Case sghServicio
    Case sghbien
        Set rs = mo_ReglasFacturacion.FarmMovimientoVentasDetalleSeleccionarPorMovNumero(ms_MovNumero, "S", oConexion)
        CargarItemsALaGrillaB rs
    End Select
    If ml_AgruparPor = 1 Then
        Totalizar
    End If
    Set grdProductos.DataSource = mrs_FacturacionProductos
End Sub

Sub CargarItemsALaGrillaB(rs As Recordset)
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
          mrs_FacturacionProductos.Delete
          mrs_FacturacionProductos.MoveNext
       Loop
    End If
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Dim oRs As New Recordset
            Dim lnCantidadSIS As Long: Dim lnPrecioSIS As Double: Dim lnImporteSIS As Double
            Dim lnCantidadSOAT As Long: Dim lnPrecioSOAT As Double: Dim lnImporteSOAT As Double
            Dim lnImporteEXO As Double: Dim ldFechaAutorizaSeguro As String: Dim lnIdUsuarioAutoriza As Long
            Dim lnIdTipoFinanciamiento As Long: Dim lcEsConvenio As String: Dim lnPrecioCONV As Double
            Dim lnCantidadPagar As Long: Dim lnPrecioPagar As Double: Dim lnTotalPagar As Double
            Dim lnIdEstadoFacturacion As Long: Dim lnIdComprobantePago As Long: Dim lcDocumentoPago As String
            Dim lnCantidadDev As Long: Dim lnIdComprobDev As Long: Dim lnIdEstadoDev As Long
            Dim lcFechaAutDev As String: Dim lnIdUsuarioAutDev As Long
            Dim LcMovNumero As String: Dim LcMovTipo As String
            Dim lnIdOrden As Long
            Dim oDoComprobantesPago As New DOCajaComprobantesPago
            Dim lnCantidadConv As Long: Dim lnImporteConv As Double
            Dim lnIdTipoConceptoFarmacia As Long
            Dim lnIdFuenteFinanciamiento As Long
            Dim lnPrecioDespacho As Double
            Dim lcProcedencia As String
            Dim lnImporteEnBoleta As Double
            Dim lnComoSeTrabajaEnEstadoCuenta As sghComoSeTrabajaEnEstadoCuentaLosSeguros
            Dim lnComoSeTrabajaEnEstadoCuenta1 As Long
            '
            lnComoSeTrabajaEnEstadoCuenta1 = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion)
            rs.MoveFirst
            LcMovTipo = "S": LcMovNumero = rs.Fields!movNumero
            Do While Not rs.EOF
                'Si tiene algun Seguro
                lnCantidadSIS = 0: lnPrecioSIS = 0: lnImporteSIS = 0
                lnCantidadSOAT = 0: lnPrecioSOAT = 0: lnImporteSOAT = 0
                lnImporteEXO = 0: ldFechaAutorizaSeguro = "": lnIdUsuarioAutoriza = 0
                lnIdTipoFinanciamiento = 0: lnIdFuenteFinanciamiento = 0
                lcEsConvenio = "No": lnPrecioCONV = 0
                lnCantidadConv = 0: lnImporteConv = 0: lnIdTipoConceptoFarmacia = 0
                lnIdEstadoFacturacion = 0: lnComoSeTrabajaEnEstadoCuenta = sghTrabajaNinguno
                'debb 01/02/2011
                Set oRs = mo_ReglasFacturacion.FacturacionBienesFinanciamSeleccionarPorIdProducto(rs.Fields!movNumero, LcMovTipo, rs.Fields!idProducto, oConexion)
                If oRs.RecordCount > 0 Then
                   oRs.MoveFirst
                   lnIdEstadoFacturacion = rs.Fields!idEstadoMovimiento
                   Do While Not oRs.EOF
                        If oRs.Fields!idTipoFinanciamiento > 0 Then
                           lnComoSeTrabajaEnEstadoCuenta = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(oRs.Fields!idTipoFinanciamiento, oConexion)
                        End If
                        If oRs.Fields!IdFuenteFinanciamiento > 0 Then
                           lnIdTipoConceptoFarmacia = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(oRs.Fields!IdFuenteFinanciamiento, oConexion)
                        End If
                        Select Case lnComoSeTrabajaEnEstadoCuenta
                        Case sghTrabajaSeguroSIS
                             lnCantidadSIS = oRs.Fields!CantidadFinanciada: lnPrecioSIS = oRs.Fields!precioFinanciado
                             lnImporteSIS = oRs.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                             lnIdEstadoFacturacion = oRs.Fields!idestadofacturacion
                        Case sghTrabajaSeguroSOAT
                             lnCantidadSOAT = oRs.Fields!CantidadFinanciada: lnPrecioSOAT = oRs.Fields!precioFinanciado
                             lnImporteSOAT = oRs.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                             lnIdEstadoFacturacion = oRs.Fields!idestadofacturacion
                        Case sghTrabajaSeguroConvenios
                             lnCantidadPagar = oRs.Fields!CantidadFinanciada: lnTotalPagar = oRs.Fields!TotalFinanciado
                             lnPrecioCONV = oRs.Fields!precioFinanciado: lcEsConvenio = "Si"
                             lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnCantidadConv = oRs.Fields!CantidadFinanciada
                             lnImporteConv = oRs.Fields!TotalFinanciado
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                             lnIdEstadoFacturacion = oRs.Fields!idestadofacturacion
                        Case Else           'exoneraciones/particular hospitalizado
                             lnImporteEXO = oRs.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                        End Select
                        oRs.MoveNext
                   Loop
                Else
                   lnIdTipoFinanciamiento = rs.Fields!idTipoFinanciamiento
                End If
                oRs.Close
                'Pagos
                lnCantidadPagar = 0: lnPrecioPagar = 0: lnTotalPagar = 0: lnIdOrden = 0
                lnIdComprobantePago = 0: lcDocumentoPago = "": lnImporteEnBoleta = 0
                'debb 01/02/2011
                Set oRs = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorIdProducto(rs.Fields!movNumero, LcMovTipo, rs.Fields!idProducto, oConexion)
                If oRs.RecordCount > 0 Then
                    oRs.MoveLast
                    'If oRs.Fields!IdEstadoFacturacion = 1 Then
                        lnCantidadPagar = oRs.Fields!CantidadPagar: lnPrecioPagar = oRs.Fields!PrecioVenta
                        lnTotalPagar = oRs.Fields!TotalPagar - lnImporteEXO: lnIdEstadoFacturacion = oRs.Fields!idestadofacturacion
                        lnIdComprobantePago = IIf(IsNull(oRs.Fields!IdComprobantePago), 0, oRs.Fields!IdComprobantePago)
                        lnIdOrden = oRs.Fields!IdOrden
                        If lnIdComprobantePago > 0 Then
                            lnImporteEnBoleta = oRs.Fields!TotalPagar
                            Set oDoComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePago, oConexion)
                            lcDocumentoPago = Trim(oDoComprobantesPago.nroSerie) + "-" + Trim(oDoComprobantesPago.nrodocumento)
                            lnTotalPagar = 0
                        End If
                    'End If
                End If
                oRs.Close
                'Devoluciones
                lnCantidadDev = 0: lnIdComprobDev = 0: lnIdEstadoDev = 0: lcFechaAutDev = ""
                lnIdUsuarioAutDev = 0
                Set oRs = mo_ReglasFacturacion.FacturacionBienesDevolucionesSeleccionarPorIdProductoConexion(oConexion, rs.Fields!movNumero, LcMovTipo, rs.Fields!idProducto)
                If oRs.RecordCount > 0 Then
                    'lnCantidadDev = oRs.Fields!CantidadAdevolver
                    lnIdComprobDev = IIf(IsNull(oRs.Fields!IdComprobantePago), 0, oRs.Fields!IdComprobantePago)
                    lnIdEstadoDev = oRs.Fields!idEstadoDevolucion: lcFechaAutDev = oRs.Fields!FechaAutoriza
                    lnIdUsuarioAutDev = oRs.Fields!IdUsuarioAutoriza
                    Do While Not oRs.EOF
                       lnCantidadDev = lnCantidadDev + oRs.Fields!CantidadAdevolver
                       oRs.MoveNext
                    Loop
                End If
                oRs.Close
                '
                If lnPrecioPagar = 0 Then
                    Set oRs = mo_ReglasComunes.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, 1)
                    lnPrecioDespacho = 0
                    If oRs.RecordCount > 0 Then
                       lnPrecioDespacho = oRs.Fields!PrecioUnitario
                    End If
                Else
                    lnPrecioDespacho = lnPrecioPagar
                End If
                'Actualiza Precio del SEGURO, en caso sea igual a CERO
                Select Case lnComoSeTrabajaEnEstadoCuenta1
                Case sghTrabajaSeguroSIS
                     If lnPrecioSIS = 0 Then
                          Set oRs = mo_ReglasComunes.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, ml_IdTipoFinanciamiento)
                          If oRs.RecordCount > 0 Then
                             lnPrecioSIS = oRs.Fields!PrecioUnitario
                          End If
                          oRs.Close
                     End If
                     If lnPrecioDespacho = 0 Then
                        lnPrecioDespacho = lnPrecioSIS
                     End If
                Case sghTrabajaSeguroSOAT
                     If lnPrecioSOAT = 0 Then
                          Set oRs = mo_ReglasComunes.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, ml_IdTipoFinanciamiento)
                          If oRs.RecordCount > 0 Then
                             lnPrecioSOAT = oRs.Fields!PrecioUnitario
                          End If
                          oRs.Close
                     End If
                     If lnPrecioDespacho = 0 Then
                        lnPrecioDespacho = lnPrecioSOAT
                     End If
                Case sghTrabajaSeguroConvenios
                     If lnPrecioCONV = 0 Then
                          Set oRs = mo_ReglasComunes.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, ml_IdTipoFinanciamiento)
                          If oRs.RecordCount > 0 Then
                             lnPrecioCONV = oRs.Fields!PrecioUnitario
                          End If
                          oRs.Close
                     End If
                     If lnPrecioDespacho = 0 Then
                        lnPrecioDespacho = lnPrecioCONV
                     End If
                End Select
                '
                mrs_FacturacionProductos.AddNew
                mrs_FacturacionProductos!movNumero = rs!movNumero
                mrs_FacturacionProductos!MovTipo = "S"
                mrs_FacturacionProductos!idProducto = rs!idProducto
                mrs_FacturacionProductos!Codigo = rs!Codigo
                mrs_FacturacionProductos!NombreProducto = rs!nombre
                mrs_FacturacionProductos!CantidadPagar = (rs!Cantidad - lnCantidadDev) 'cantidad inicial (no varia)....menos Cantidad Devuelta(NI)
                mrs_FacturacionProductos!PrecioUnitario = lnPrecioDespacho  'precio de venta
                mrs_FacturacionProductos!TotalPagar = Round(lnPrecioDespacho * (rs!Cantidad - lnCantidadDev), 2)   '....menos Cantidad Devuelta (NI)
                mrs_FacturacionProductos!CantidadSIS = lnCantidadSIS
                mrs_FacturacionProductos!precioSIS = lnPrecioSIS
                mrs_FacturacionProductos!ImporteSIS = lnImporteSIS
                mrs_FacturacionProductos!CantidadSOAT = lnCantidadSOAT
                mrs_FacturacionProductos!PrecioSOAT = lnPrecioSOAT
                mrs_FacturacionProductos!ImporteSOAT = lnImporteSOAT
                mrs_FacturacionProductos!importeEXO = lnImporteEXO
                mrs_FacturacionProductos!idPuntoCarga = 5
                mrs_FacturacionProductos!idestadofacturacion = lnIdEstadoFacturacion    'IIf(lnIdUsuarioAutoriza > 0, rs.Fields!idEstadoMovimiento, lnIdEstadoFacturacion)  '
                mrs_FacturacionProductos!Cantidad = lnCantidadPagar 'cantidad a pagar en caja (varia)
                mrs_FacturacionProductos!TotalPorPagar = lnTotalPagar  '(a pagar en caja)
                mrs_FacturacionProductos!IdComprobantePago = lnIdComprobantePago
                mrs_FacturacionProductos!IdOrden = lnIdOrden
                mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = lnIdUsuarioAutoriza
                If lnIdTipoConceptoFarmacia > 0 Then
                   mrs_FacturacionProductos!FechaAutorizaSeguro = ldFechaAutorizaSeguro
                End If
                mrs_FacturacionProductos!IdUsuarioAutorizaDevolucion = lnIdUsuarioAutDev
                mrs_FacturacionProductos!FechaAutorizaDevolucion = IIf(lnCantidadDev = 0, 0, lcFechaAutDev)
                mrs_FacturacionProductos!IdComprobantePagoDevolucion = lnIdComprobDev
                mrs_FacturacionProductos!NroComprobante = IIf(lcDocumentoPago = "", rs!DocumentoNumero, lcDocumentoPago)  'si ya se PAGO muestra BOLETA sino muestra TICKET
                mrs_FacturacionProductos!idTipoFinanciamiento = lnIdTipoFinanciamiento
                mrs_FacturacionProductos!precioCONV = lnPrecioCONV
                mrs_FacturacionProductos!esConvenio = lcEsConvenio
                mrs_FacturacionProductos!FechaOrden = rs!fechacreacion
                mrs_FacturacionProductos!NombreServicio = rs!dalmacen
                mrs_FacturacionProductos!cantidadConv = lnCantidadConv
                mrs_FacturacionProductos!ImporteConv = lnImporteConv
                mrs_FacturacionProductos!idTipoConceptoFarmacia = lnIdTipoConceptoFarmacia
                mrs_FacturacionProductos!IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
                If Not IsNull(rs!idServicioPaciente) Then
                    mrs_FacturacionProductos!ServicioDeEstancia = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(rs!idServicioPaciente)
                    mrs_FacturacionProductos!idServicioDeEstancia = rs!idServicioPaciente
                End If
                mrs_FacturacionProductos!CantidadDevuelta = lnCantidadDev
                mrs_FacturacionProductos!nrodocumento = rs!DocumentoNumero
                If ml_AgruparPor = 3 Or ml_AgruparPor = 5 Then
                   mrs_FacturacionProductos!descripcion = rs!dfinanciamiento
                End If
                mrs_FacturacionProductos!ImporteEnBoleta = lnImporteEnBoleta
                mrs_FacturacionProductos!nroDcto = rs!DocumentoNumero
                mrs_FacturacionProductos!ComoSeTrabajaEnEstadoCuenta = lnComoSeTrabajaEnEstadoCuenta
                mrs_FacturacionProductos!FechaDespacho = rs!fechacreacion
                rs.MoveNext
            Loop
            Set oRs = Nothing
        End If
    End If
End Sub

Sub CargarItemsALaGrillaS(rs As Recordset)
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
          mrs_FacturacionProductos.Delete
          mrs_FacturacionProductos.MoveNext
       Loop
    End If
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Dim oRs As New Recordset
            Dim lnCantidadSIS As Long: Dim lnPrecioSIS As Double: Dim lnImporteSIS As Double
            Dim lnCantidadSOAT As Long: Dim lnPrecioSOAT As Double: Dim lnImporteSOAT As Double
            Dim lnImporteEXO As Double: Dim ldFechaAutorizaSeguro As String: Dim lnIdUsuarioAutoriza As Long
            Dim lnIdTipoFinanciamiento As Long: Dim lcEsConvenio As String: Dim lnPrecioCONV As Double
            Dim lnCantidadPagar As Long: Dim lnPrecioPagar As Double: Dim lnTotalPagar As Double
            Dim lnIdEstadoFacturacion As Long: Dim lnIdComprobantePago As Long: Dim lcDocumentoPago As String
            Dim lnCantidadDev As Long: Dim lnIdComprobDev As Long: Dim lnIdEstadoDev As Long
            Dim lcFechaAutDev As String: Dim lnIdUsuarioAutDev As Long
            Dim LcMovNumero As String: Dim LcMovTipo As String
            Dim lnIdOrden As Long
            Dim oDoComprobantesPago As New DOCajaComprobantesPago
            Dim lnCantidadConv As Long: Dim lnImporteConv As Double
            Dim lnIdTipoConceptoFarmacia As Long
            Dim lnIdFuenteFinanciamiento As Long
            Dim lnPrecioDespacho As Double
            Dim lnImporteEnBoleta As Double
            Dim lcNroDcto As String
            Dim lnIdOrdenPago As Long
            Dim lnComoSeTrabajaEnEstadoCuenta As sghComoSeTrabajaEnEstadoCuentaLosSeguros
            Dim lnComoSeTrabajaEnEstadoCuenta1 As Long
            Dim lbElMovimientoNoEstaAnulado As Boolean
            '
            lnComoSeTrabajaEnEstadoCuenta1 = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion)
            rs.MoveFirst
            Do While Not rs.EOF
                'Si tiene algun Seguro
                lnCantidadSIS = 0: lnPrecioSIS = 0: lnImporteSIS = 0
                lnCantidadSOAT = 0: lnPrecioSOAT = 0: lnImporteSOAT = 0
                lnImporteEXO = 0: ldFechaAutorizaSeguro = "": lnIdUsuarioAutoriza = 0
                lnIdTipoFinanciamiento = 0: lnIdFuenteFinanciamiento = 0
                lcEsConvenio = "No": lnPrecioCONV = 0
                lnCantidadConv = 0: lnImporteConv = 0: lnIdTipoConceptoFarmacia = 0
                lnIdEstadoFacturacion = 0: lnComoSeTrabajaEnEstadoCuenta = sghTrabajaNinguno
                'debb 01/02/2011
                Set oRs = mo_ReglasFacturacion.FacturacionServicioFinanciamPorIdOrdenIdProducto(oConexion, rs.Fields!IdOrden, rs.Fields!idProducto)
                If oRs.RecordCount > 0 Then
                   
                   oRs.MoveFirst
                   lnIdEstadoFacturacion = rs!idestadofacturacion
                   Do While Not oRs.EOF
                        If oRs.Fields!idTipoFinanciamiento > 0 Then
                           lnComoSeTrabajaEnEstadoCuenta = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(oRs.Fields!idTipoFinanciamiento, oConexion)
                        End If
                        If oRs.Fields!IdFuenteFinanciamiento > 0 Then
                           lnIdTipoConceptoFarmacia = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(oRs.Fields!IdFuenteFinanciamiento, oConexion)
                        End If
                        Select Case lnComoSeTrabajaEnEstadoCuenta
                        Case sghTrabajaSeguroSIS
                             lnCantidadSIS = oRs.Fields!CantidadFinanciada: lnPrecioSIS = oRs.Fields!precioFinanciado
                             lnImporteSIS = oRs.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                             lnIdEstadoFacturacion = oRs!idestadofacturacion
                        Case sghTrabajaSeguroSOAT
                             lnCantidadSOAT = oRs.Fields!CantidadFinanciada: lnPrecioSOAT = oRs.Fields!precioFinanciado
                             lnImporteSOAT = oRs.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                             lnIdEstadoFacturacion = oRs!idestadofacturacion
                        Case sghTrabajaSeguroConvenios
                             lnPrecioCONV = oRs.Fields!precioFinanciado: lcEsConvenio = "Si"
                             lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnCantidadConv = oRs.Fields!CantidadFinanciada
                             lnImporteConv = oRs.Fields!TotalFinanciado
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                             lnIdEstadoFacturacion = oRs!idestadofacturacion
                        Case Else           'exoneraciones/particular hospitalizado
                             lnImporteEXO = oRs.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRs.Fields!idTipoFinanciamiento
                             ldFechaAutorizaSeguro = oRs.Fields!FechaAutoriza
                             lnIdUsuarioAutoriza = oRs.Fields!IdUsuarioAutoriza
                             lnIdFuenteFinanciamiento = oRs.Fields!IdFuenteFinanciamiento
                        End Select
                        oRs.MoveNext
                   Loop
                Else
                   lnIdTipoFinanciamiento = rs.Fields!idTipoFinanciamiento
                End If
                oRs.Close
                'Pagos
                lnCantidadPagar = 0: lnPrecioPagar = 0: lnTotalPagar = 0: lnIdOrden = 0
                lnIdComprobantePago = 0: lcDocumentoPago = "": lnImporteEnBoleta = 0
                lnIdOrdenPago = 0
                'debb 01/02/2011
                Set oRs = mo_ReglasFacturacion.FacturacionServicioPagosSeleccionarPorIdOrdenIdProducto(rs.Fields!IdOrden, rs.Fields!idProducto, oConexion)
                If oRs.RecordCount > 0 Then
                    oRs.MoveLast
                    'If oRs.Fields!IdEstadoFacturacion = 1 Then
                        lnCantidadPagar = oRs.Fields!Cantidad: lnPrecioPagar = oRs.Fields!Precio
                        lnTotalPagar = oRs.Fields!Total - lnImporteEXO: lnIdEstadoFacturacion = oRs.Fields!idestadofacturacion
                        lnIdComprobantePago = IIf(IsNull(oRs.Fields!IdComprobantePago), 0, oRs.Fields!IdComprobantePago)
                        lnIdOrden = rs.Fields!IdOrden
                        lnIdOrdenPago = oRs.Fields!IdOrdenPago
                        If lnIdComprobantePago > 0 Then
                            lnImporteEnBoleta = oRs.Fields!Total
                            Set oDoComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePago, oConexion)
                            lcDocumentoPago = Trim(oDoComprobantesPago.nroSerie) + "-" + Trim(oDoComprobantesPago.nrodocumento)
                            lnTotalPagar = 0
                        End If
                    'End If
                End If
                oRs.Close
                'Devoluciones
                lnCantidadDev = 0: lnIdComprobDev = 0: lnIdEstadoDev = 0: lcFechaAutDev = ""
                lnIdUsuarioAutDev = 0
                Set oRs = mo_ReglasFacturacion.FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProductoConexion(oConexion, rs.Fields!IdOrden, rs.Fields!idProducto)
                If oRs.RecordCount > 0 Then
                    lnCantidadDev = oRs.Fields!CantidadAdevolver: lnIdComprobDev = oRs.Fields!IdComprobantePago
                    lnIdEstadoDev = oRs.Fields!idEstadoDevolucion: lcFechaAutDev = oRs.Fields!FechaAutoriza
                    lnIdUsuarioAutDev = oRs.Fields!IdUsuarioAutoriza
                End If
                oRs.Close
                'Actualiza Precios Contado
                If lnPrecioPagar = 0 Then
                    Set oRs = mo_ReglasComunes.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, 1)
                    lnPrecioDespacho = 0
                    If oRs.RecordCount > 0 Then
                       If rs!idPuntoCarga = sghPtoCargaAdmisionHospitalizacion Then
                          'Estancia hospitalaria
                          lnPrecioDespacho = oRs.Fields!PrecioUnitario / 24
                       Else
                          lnPrecioDespacho = oRs.Fields!PrecioUnitario
                       End If
                    End If
                    oRs.Close
                Else
                    lnPrecioDespacho = lnPrecioPagar
                End If
                'Nro de Documento, para mostrar en Pantalla e Impresora
                lbElMovimientoNoEstaAnulado = True
                lcNroDcto = Trim(Str(rs!IdOrden))
                Select Case rs!idPuntoCarga
                Case sghPtoCargaEcogGeneral, sghPtoCargaRayosX, sghPtoCargaTomografia, sghPtoCargaEcogObstetrica                      'Imagenes
                     'debb 01/02/2011
                     Set oRs = mo_ReglasImagenes.ImagMovimientoImagenesXidOrden(rs!IdOrden, oConexion)
                     If oRs.RecordCount > 0 Then
                        lcNroDcto = Trim(Str(oRs.Fields!IdMovimiento))
                     Else
                        lbElMovimientoNoEstaAnulado = False
                     End If
                     oRs.Close
                Case sghPtoCargaPatologiaClinica, sghPtoCargaAnatomiaPatologica1, sghPtoCargaBancoSangre1    'Laboratorio
                     'debb 01/02/2011
                     Set oRs = mo_ReglasLaboratorio.LabMovimientoLaboratorioXidOrden(rs!IdOrden, oConexion)
                     If oRs.RecordCount > 0 Then
                        lcNroDcto = Trim(Str(oRs.Fields!IdMovimiento))
                     Else
                        lbElMovimientoNoEstaAnulado = False
                     End If
                     oRs.Close
                Case sghPtoCargaCaja          'se genero en CAJA
                    If lcNroDcto <> "" Then
                       lcNroDcto = lcDocumentoPago
                    End If
                End Select
                'Actualiza Precio del SEGURO, en caso sea igual a CERO
                Select Case lnComoSeTrabajaEnEstadoCuenta1
                Case sghTrabajaSeguroSIS
                     If lnPrecioSIS = 0 Then
                          Set oRs = mo_ReglasComunes.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, ml_IdTipoFinanciamiento)
                          If oRs.RecordCount > 0 Then
                             lnPrecioSIS = oRs.Fields!PrecioUnitario
                          End If
                          oRs.Close
                     End If
                     If lnPrecioDespacho = 0 Then
                        lnPrecioDespacho = lnPrecioSIS
                     End If
                Case sghTrabajaSeguroSOAT
                     If lnPrecioSOAT = 0 Then
                          Set oRs = mo_ReglasComunes.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, ml_IdTipoFinanciamiento)
                          If oRs.RecordCount > 0 Then
                             lnPrecioSOAT = oRs.Fields!PrecioUnitario
                          End If
                          oRs.Close
                     End If
                     If lnPrecioDespacho = 0 Then
                        lnPrecioDespacho = lnPrecioSOAT
                     End If
                Case sghTrabajaSeguroConvenios
                     If lnPrecioCONV = 0 Then
                          Set oRs = mo_ReglasComunes.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oConexion, rs!idProducto, ml_IdTipoFinanciamiento)
                          If oRs.RecordCount > 0 Then
                             lnPrecioCONV = oRs.Fields!PrecioUnitario
                          End If
                          oRs.Close
                     End If
                     If lnPrecioDespacho = 0 Then
                        lnPrecioDespacho = lnPrecioCONV
                     End If
                End Select
                '
If lcNroDcto = "9008" Then
ms_MensajeError = ""
End If
                If lbElMovimientoNoEstaAnulado = True Then
                    mrs_FacturacionProductos.AddNew
                    mrs_FacturacionProductos!idProducto = rs!idProducto
                    mrs_FacturacionProductos!Codigo = rs!Codigo
                    mrs_FacturacionProductos!NombreProducto = rs!nombre
                    mrs_FacturacionProductos!CantidadPagar = rs!Cantidad  'cantidad inicial (no varia)
                    mrs_FacturacionProductos!PrecioUnitario = lnPrecioDespacho    'rs!precio  'precio de venta
                    mrs_FacturacionProductos!TotalPagar = Round(rs!Cantidad * lnPrecioDespacho, 2)    'rs!Total
                    mrs_FacturacionProductos!CantidadSIS = lnCantidadSIS
                    mrs_FacturacionProductos!precioSIS = lnPrecioSIS
                    mrs_FacturacionProductos!ImporteSIS = lnImporteSIS
                    mrs_FacturacionProductos!CantidadSOAT = lnCantidadSOAT
                    mrs_FacturacionProductos!PrecioSOAT = lnPrecioSOAT
                    mrs_FacturacionProductos!ImporteSOAT = lnImporteSOAT
                    mrs_FacturacionProductos!importeEXO = lnImporteEXO
                    mrs_FacturacionProductos!idPuntoCarga = rs!idPuntoCarga
                    mrs_FacturacionProductos!idestadofacturacion = lnIdEstadoFacturacion      ' IIf(lnIdUsuarioAutoriza > 0, rs!IdEstadoFacturacion, lnIdEstadoFacturacion)
                    mrs_FacturacionProductos!Cantidad = lnCantidadPagar 'cantidad a pagar en caja (varia)
                    mrs_FacturacionProductos!TotalPorPagar = lnTotalPagar  '(a pagar en caja)
                    mrs_FacturacionProductos!IdComprobantePago = lnIdComprobantePago
                    mrs_FacturacionProductos!IdOrden = rs!IdOrden
                    mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = lnIdUsuarioAutoriza
                    If lnIdTipoConceptoFarmacia > 0 Then
                       mrs_FacturacionProductos!FechaAutorizaSeguro = ldFechaAutorizaSeguro
                    End If
                    mrs_FacturacionProductos!IdUsuarioAutorizaDevolucion = lnIdUsuarioAutDev
                    mrs_FacturacionProductos!FechaAutorizaDevolucion = IIf(lnCantidadDev = 0, 0, lcFechaAutDev)
                    mrs_FacturacionProductos!IdComprobantePagoDevolucion = lnIdComprobDev
                    mrs_FacturacionProductos!NroComprobante = IIf(lcDocumentoPago = "", "", lcDocumentoPago)  'si ya se PAGO muestra BOLETA sino muestra TICKET
                    mrs_FacturacionProductos!idTipoFinanciamiento = lnIdTipoFinanciamiento
                    mrs_FacturacionProductos!precioCONV = lnPrecioCONV
                    mrs_FacturacionProductos!esConvenio = lcEsConvenio
                    mrs_FacturacionProductos!FechaOrden = rs!fechacreacion
                    mrs_FacturacionProductos!cantidadConv = lnCantidadConv
                    mrs_FacturacionProductos!ImporteConv = lnImporteConv
                    mrs_FacturacionProductos!idTipoConceptoFarmacia = lnIdTipoConceptoFarmacia
                    mrs_FacturacionProductos!IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
                    If Not IsNull(rs!idServicioPaciente) Then
                        mrs_FacturacionProductos!ServicioDeEstancia = IIf(IsNull(rs!idServicioPaciente), ".", mo_ReglasFacturacion.BuscaServicioActualDelPaciente(rs!idServicioPaciente))
                        mrs_FacturacionProductos!idServicioDeEstancia = IIf(IsNull(rs!idServicioPaciente), 0, rs!idServicioPaciente)
                    End If
                    If ml_AgruparPor = 3 Or ml_AgruparPor = 5 Then
                       mrs_FacturacionProductos!descripcion = rs!dfinanciamiento
                    End If
                    mrs_FacturacionProductos!ImporteEnBoleta = lnImporteEnBoleta
                    mrs_FacturacionProductos!nroDcto = lcNroDcto
                    mrs_FacturacionProductos!ComoSeTrabajaEnEstadoCuenta = lnComoSeTrabajaEnEstadoCuenta
                    mrs_FacturacionProductos!IdOrdenPago = lnIdOrdenPago
                    mrs_FacturacionProductos!FechaDespacho = rs!FechaDespacho
                End If
                rs.MoveNext
            Loop
            '
            Set oRs = Nothing
        End If
    End If
End Sub



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
Dim lnTotalImporteEXO As Double

    dTotalExonerado = 0
    dTotalPagoACuenta = 0
    dTotalIngresado = 0
    dTotalPendientePago = 0
    dTotalPagado = 0
    dTotalPorDevolver = 0
    dTotalDevuelto = 0
    dTotalAnulado = 0
    
    Set rsProductos = mrs_FacturacionProductos.Clone
    
    If rsProductos.RecordCount = 0 Then
        Exit Sub
    End If
    
    If Not (rsProductos.EOF And rsProductos.BOF) Then
        rsProductos.MoveFirst
        Do While Not rsProductos.EOF
            lnTotalImporteEXO = lnTotalImporteEXO + rsProductos!importeEXO
            dSubTotal = rsProductos!TotalPagar
            lIdEstadoFacturacion = rsProductos!idestadofacturacion
            lIdProducto = rsProductos!idProducto
            
            Select Case lIdEstadoFacturacion
            Case 1
                Select Case lIdProducto
                Case 4692
                    dTotalExonerado = dTotalExonerado + dSubTotal
                Case Else
                    dTotalIngresado = dTotalIngresado + dSubTotal
                End Select
            Case 3
                dTotalPendientePago = dTotalPendientePago + dSubTotal
            Case 4
                Select Case lIdProducto
                Case lnIdPagosACuenta
                    'se entiende que este estado de cuenta es solo para las cuentas de atencion pendientes
                    'If IsNull(rsProductos!IdComprobantePago) Then
                        dTotalPagoACuenta = dTotalPagoACuenta + dSubTotal
                    'End If
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
    
    RaiseEvent Totalizado(dTotalIngresado, dTotalPendientePago, dTotalPagoACuenta, dTotalExonerado, dTotalPagado, _
                          dTotalPorDevolver, dTotalDevuelto, dTotalAnulado, IIf(lnTotalImporteEXO = 0, False, True))
        


End Sub




Private Sub grdProductos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If mb_FilaEditable = False Then
       Exit Sub
    End If
    If Cell.Column.Key = "CantidadSIS" Or Cell.Column.Key = "CantidadSOAT" Or Cell.Column.Key = "ImporteEXO" Then
        Dim lnDctos As Double
        Dim lnCant As Long
        If Cell.Column.Key = "ImporteEXO" Then
            lnCant = 0
            If Cell.Row.Cells("CantidadSIS").Value > 0 And Cell.Row.Cells("PrecioSIS").Value > 0 Then
               lnCant = lnCant + CDbl(Cell.Row.Cells("CantidadSIS").Value)
            End If
            If Cell.Row.Cells("CantidadSOAT").Value > 0 And Cell.Row.Cells("PrecioSOAT").Value > 0 Then
               lnCant = lnCant + CDbl(Cell.Row.Cells("cantidadSOAT").Value)
            End If
            lnCant = Cell.Row.Cells("CantidadPagar").Value - lnCant
            lnDctos = lnCant * Cell.Row.Cells("preciounitario").Value
            If Cell.Row.Cells("ImporteEXO").Value > Round(lnDctos, 2) Then
                MsgBox "El Descuento (EXONERACION) = " & Trim(Str(Cell.Row.Cells("ImporteEXO").Value)) & Chr(13) & "    pasa del Total a Pagar", vbInformation, "Mensaje"
                Select Case grdProductos.ActiveCell.Column.Key
                Case "ImporteEXO"
                      Cell.Row.Cells("ImporteEXO").Value = 0
                End Select
            Else
                Cell.Row.Cells("TotalPorPagar").Value = Round(lnDctos, 2) - Cell.Row.Cells("ImporteEXO").Value
            End If
        Else
            lnCant = 0
            If Cell.Row.Cells("CantidadSIS").Value > 0 And Cell.Row.Cells("PrecioSIS").Value > 0 Then
               lnCant = lnCant + CDbl(Cell.Row.Cells("CantidadSIS").Value)
            End If
            If Cell.Row.Cells("CantidadSOAT").Value > 0 And Cell.Row.Cells("PrecioSOAT").Value > 0 Then
               lnCant = lnCant + CDbl(Cell.Row.Cells("cantidadSOAT").Value)
            End If
            If Cell.Row.Cells("CantidadPagar").Value < lnCant Then
                MsgBox "La cantidad autorizada (SIS, SOAT) = " & Trim(Str(lnCant)) & Chr(13) & "    pasa de la Cantidad= " & Trim(Str(Cell.Row.Cells("Cantidad").Value)), vbInformation, "Mensaje"
                Select Case grdProductos.ActiveCell.Column.Key
                Case "CantidadSIS"
                      Cell.Row.Cells("CantidadSIS").Value = 0
                Case "PrecioSIS"
                      Cell.Row.Cells("PrecioSIS").Value = 0
                Case "CantidadSOAT"
                      Cell.Row.Cells("CantidadSOAT").Value = 0
                Case "PrecioSOAT"
                      Cell.Row.Cells("PrecioSOAT").Value = 0
                End Select
            Else
                If Cell.Row.Cells("idEstadoFacturacion").Value = 1 Then
                    Select Case grdProductos.ActiveCell.Column.Key
                    Case "CantidadSIS", "PrecioSIS"
                       Cell.Row.Cells("ImporteSIS").Value = Cell.Row.Cells("CantidadSIS").Value * Cell.Row.Cells("PrecioSIS").Value
                    Case "CantidadSOAT", "PrecioSOAT"
                       Cell.Row.Cells("ImporteSOAT").Value = Cell.Row.Cells("CantidadSOAT").Value * Cell.Row.Cells("PrecioSOAT").Value
                    End Select
                    Cell.Row.Cells("Cantidad").Value = Cell.Row.Cells("CantidadPagar").Value - lnCant
                    Cell.Row.Cells("TotalPorPagar").Value = Cell.Row.Cells("preciounitario").Value * Cell.Row.Cells("Cantidad").Value
                End If
            End If
        End If
        Totalizar
    End If

End Sub

Private Sub grdProductos_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
        Select Case Cell.Row.Cells("idEstadoFacturacion").Value
        Case 1, 16   'Registrado, conPreventa
             Cell.Row.Cells("importeSIS").Activation = ssActivationAllowEdit
             Cell.Row.Cells("importeSOAT").Activation = ssActivationAllowEdit
             Cell.Row.Cells("importeEXO").Activation = ssActivationAllowEdit
             Cell.Row.Cells("precioSIS").Activation = ssActivationActivateNoEdit
             Cell.Row.Cells("precioSOAT").Activation = ssActivationActivateNoEdit
             mb_FilaEditable = True
        Case Else 'Pagado, Anulado, Devuelto  (4,9,6)
             Cell.Row.Cells("cantidadSIS").Activation = ssActivationActivateNoEdit
             Cell.Row.Cells("precioSIS").Activation = ssActivationActivateNoEdit
             Cell.Row.Cells("cantidadSOAT").Activation = ssActivationActivateNoEdit
             Cell.Row.Cells("precioSOAT").Activation = ssActivationActivateNoEdit
             Cell.Row.Cells("importeEXO").Activation = ssActivationActivateNoEdit
             mb_FilaEditable = False
        End Select

End Sub



Private Sub grdProductos_DblClick()
  On Error GoTo ErrPrdCli
  If ml_lbPuedeVerResultados = True Then
        If mrs_FacturacionProductos!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or _
                  mrs_FacturacionProductos!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2 Or _
                  mrs_FacturacionProductos!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or _
                  mrs_FacturacionProductos!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Or _
                  mrs_FacturacionProductos!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Then
              
              
              '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
              Dim oRsTmp1 As New Recordset
              Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & mrs_FacturacionProductos!Codigo & "'")
              If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.Close
                   Dim oResultadoXitems As New SIGHLaboratorio.ResultadoXitems
                   oResultadoXitems.IdOrden = mrs_FacturacionProductos!IdOrden
                   oResultadoXitems.idProductoCpt = mrs_FacturacionProductos!idProducto
                   'oResultadoXitems.idUsuario = ml_idUsuario
                   oResultadoXitems.NoMuestraBotonGrabar = True
                   'oResultadoXitems.lcNombrePc = mo_lcNombrePc
                   'oResultadoXitems.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
                   'oResultadoXitems.idTipoSexo = lnIdTipoSexo
                   'If SIGHEntidades.EsFecha(Format(ldFechaNacimiento, "dd/mm/yyyy"), "DD/MM/AAAA") Then
                   '   oResultadoXitems.FechaNacimiento = ldFechaNacimiento
                   'End If
                   oResultadoXitems.MostrarFormulario
                   Set oResultadoXitems = Nothing
                   Set oRsTmp1 = Nothing
                   Exit Sub
              End If
              oRsTmp1.Close
              Set oRsTmp1 = Nothing
              '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
              
              
              Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
              oMuestraResultado.MuestraResultadoDelExamen mrs_FacturacionProductos!Codigo, mrs_FacturacionProductos!NombreProducto, _
                                                          ml_Paciente, mrs_FacturacionProductos!IdOrden, ml_IdPaciente, "", _
                                                          0, 0, ml_idTipoSexo, True
              Set oMuestraResultado = Nothing
        End If
  End If
ErrPrdCli:
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    InicializarLaGrilla grdProductos
End Sub

Private Sub grdProductos_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    On Error Resume Next
    ModificarColorDeFila Row
End Sub

Sub ModificarColorDeFila(ByVal Row As UltraGrid.SSRow)
        
        Select Case Row.Cells("IdProducto").Value
        Case lnIdPagosACuenta
            Row.Appearance.ForeColor = &HC7613F
        Case 4692
            Row.Appearance.ForeColor = &H16CD32
        Case 4693
            Row.Appearance.ForeColor = &H3049FA
        Case lnIdDevoluciones
            Row.Appearance.ForeColor = vbGreen
        End Select

End Sub

Sub RecalcularSubTotal(oGrilla As SSUltraGrid)
Dim oRow As SSRow
Dim dValorAntesDe As Double: Dim lnDctos As Double

    Set oRow = oGrilla.ActiveCell.Row
    
    dValorAntesDe = CDbl(oRow.Cells("TotalPorPagar").Value)
    
    lnDctos = 0
    If Not IsNull(oRow.Cells("ImporteSIS").Value) Then
       lnDctos = lnDctos + CDbl(oRow.Cells("ImporteSIS").Value)
    End If
    If Not IsNull(oRow.Cells("ImporteSOAT").Value) Then
       lnDctos = lnDctos + CDbl(oRow.Cells("ImporteSOAT").Value)
    End If
    If Not IsNull(oRow.Cells("ImporteEXO").Value) Then
       lnDctos = lnDctos + CDbl(oRow.Cells("ImporteEXO").Value)
    End If
    oRow.Cells("TotalPorPagar").Value = CDbl(oRow.Cells("totalPagar").Value) - lnDctos
    
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
    
    Select Case ms_TipoProducto
    Case sghServicio
        Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigo(oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, oConexion)
    Case sghbien
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, oConexion)
    Case Else
        Exit Sub
    End Select
    
    If rs.RecordCount = 1 Then
        oRow.Cells("IdFacturacionProducto").Value = 0
        oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
        oRow.Cells("NombreProducto").Value = rs.Fields("NombreProducto").Value
        oRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
        oRow.Cells("TotalPorPagar").Value = rs.Fields("preciounitario").Value
        oRow.Cells("cantidad").Value = 1
    End If
 '   oConexion.Close
 '   Set oConexion = Nothing
End Sub


Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
Dim lBanda As Long
Dim idUsuarioConPermisoEnSISoEXOoSOATconf As sghComoSeTrabajaEnEstadoCuentaLosSeguros
    If idUsuarioConPermisoEnSISoEXOoSOAT = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion) Or (idUsuarioConPermisoEnSISoEXOoSOAT = 9 And mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion) = sghTrabajaParticular) Then
       idUsuarioConPermisoEnSISoEXOoSOATconf = idUsuarioConPermisoEnSISoEXOoSOAT
    Else
       idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaNinguno
       If ml_IdTipoFinanciamiento = sghTipoFinanciamiento.sghSis And lbTieneDerechoExoneraSIS = True Then
          idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaSeguroSIS
       End If
    End If
    
    If ml_AgruparPor > 1 Then
        Select Case ml_AgruparPor
        Case 2
            oGrilla.Bands(0).Columns("IdOrden").Width = 1000
            oGrilla.Bands(0).Columns("IdOrden").Header.Caption = "N° Orden"
        Case 3
            oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Hidden = True
            oGrilla.Bands(0).Columns("Descripcion").Width = 2000
            oGrilla.Bands(0).Columns("Descripcion").Header.Caption = "Producto/Plan"
        Case 4
            oGrilla.Bands(0).Columns("IdAtencion").Width = 1250
            oGrilla.Bands(0).Columns("IdAtencion").Header.Caption = "N° Atención"
        Case 5
            oGrilla.Bands(0).Columns("idPuntoCarga").Hidden = True
            oGrilla.Bands(0).Columns("Descripcion").Width = 800        '2000
            oGrilla.Bands(0).Columns("Descripcion").Header.Caption = "Punto Carga"
        End Select
    End If
    
    lBanda = IIf(ml_AgruparPor > 1, 1, 0)
    
    On Error Resume Next
    oGrilla.Bands(lBanda).Columns("IdFacturacionProducto").Hidden = True
    
    oGrilla.Bands(lBanda).Columns("IdProducto").Hidden = True
    'oGrilla.Bands(lBanda).Columns("IdTipoFinanciamiento").Hidden = True
    oGrilla.Bands(lBanda).Columns("IdAtencion").Hidden = True
    oGrilla.Bands(lBanda).Columns("FechaAutorizaPendiente").Hidden = True
    oGrilla.Bands(lBanda).Columns("FechaAutorizaSeguro").Hidden = True
    oGrilla.Bands(lBanda).Columns("IdUsuarioAutorizaPendiente").Hidden = True
    oGrilla.Bands(lBanda).Columns("IdUsuarioAutorizaSeguro").Hidden = True
    oGrilla.Bands(lBanda).Columns("IdFuenteFinanciamiento").Hidden = True
    oGrilla.Bands(lBanda).Columns("IdComprobantePago").Hidden = True
    oGrilla.Bands(lBanda).Columns("IdComprobantePagoDevolucion").Hidden = True
    oGrilla.Bands(lBanda).Columns("ImporteEnBoleta").Hidden = True
    
    
    If ms_TipoProducto = sghServicio Then
        oGrilla.Bands(lBanda).Columns("IdServicioInternamiento").Hidden = True
    
        oGrilla.Bands(lBanda).Columns("NombreServicio").Header.Caption = "Serv. Internamiento"
        oGrilla.Bands(lBanda).Columns("NombreServicio").Width = 3000
        oGrilla.Bands(lBanda).Columns("NombreServicio").Activation = ssActivationActivateNoEdit
    
    End If
    
    
    oGrilla.Bands(lBanda).Columns("IdUsuarioAutorizaDevolucion").Hidden = True
    oGrilla.Bands(lBanda).Columns("FechaAutorizaDevolucion").Hidden = True
        
    oGrilla.Bands(lBanda).Columns("Codigo").Header.Caption = "Codigo"
    oGrilla.Bands(lBanda).Columns("Codigo").Width = 600    '750
    oGrilla.Bands(lBanda).Columns("Codigo").Activation = ssActivationAllowEdit
    
    oGrilla.Bands(0).Columns("nroDcto").Header.Caption = "N° Dcto"
    oGrilla.Bands(0).Columns("nroDcto").Width = 600
    oGrilla.Bands(0).Columns("nroDcto").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(lBanda).Columns("NombreProducto").Header.Caption = "Descripción"
    oGrilla.Bands(lBanda).Columns("NombreProducto").Width = 3000      '4300
    oGrilla.Bands(lBanda).Columns("NombreProducto").Activation = ssActivationAllowEdit
    
    oGrilla.Bands(lBanda).Columns("IdTipoFinanciamiento").Width = 2500
    oGrilla.Bands(lBanda).Columns("IdTipoFinanciamiento").Header.Caption = "Producto/Plan"
    oGrilla.Bands(lBanda).Columns("IdTipoFinanciamiento").Style = ssStyleDropDownList
    
    oGrilla.Bands(lBanda).Columns("CantidadPagar").Header.Caption = "Cantidad"
    oGrilla.Bands(lBanda).Columns("CantidadPagar").Format = "###0"
    oGrilla.Bands(lBanda).Columns("CantidadPagar").Width = 600
    oGrilla.Bands(lBanda).Columns("CantidadPagar").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(lBanda).Columns("preciounitario").Header.Caption = "Pr.Unit"
    oGrilla.Bands(lBanda).Columns("preciounitario").Format = "#0.0000"
    oGrilla.Bands(lBanda).Columns("preciounitario").Width = 800
    
    oGrilla.Bands(lBanda).Columns("TotalPagar").Header.Caption = "SubTotal"
    oGrilla.Bands(lBanda).Columns("TotalPagar").Format = "#0.00"
    oGrilla.Bands(lBanda).Columns("TotalPagar").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("TotalPagar").Width = 900
  
    
    oGrilla.Bands(lBanda).Columns("IdEstadoFacturacion").Width = 1500
    oGrilla.Bands(lBanda).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
    oGrilla.Bands(lBanda).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList

    oGrilla.Bands(lBanda).Columns("idPuntoCarga").Header.Caption = "Puntos de carga"
    oGrilla.Bands(lBanda).Columns("idPuntoCarga").Width = 500   '1500
    oGrilla.Bands(lBanda).Columns("idPuntoCarga").Style = ssStyleDropDownList

    oGrilla.Bands(lBanda).Columns("FechaAutorizaPendiente").Width = 2500
    oGrilla.Bands(lBanda).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Aut. Pend."
    oGrilla.Bands(lBanda).Columns("FechaAutorizaPendiente").Format = sighEntidades.DevuelveFechaSoloFormato_DMY_HM

    oGrilla.Bands(lBanda).Columns("FechaAutorizaSeguro").Width = 2500
    oGrilla.Bands(lBanda).Columns("FechaAutorizaSeguro").Header.Caption = "Fec. Aut. Seguro."
    oGrilla.Bands(lBanda).Columns("FechaAutorizaSeguro").Format = sighEntidades.DevuelveFechaSoloFormato_DMY_HM
    
   
    oGrilla.Bands(lBanda).Columns("IdOrden").Width = 1000
    oGrilla.Bands(lBanda).Columns("IdOrden").Header.Caption = "N° Orden"
    
    oGrilla.Bands(lBanda).Columns("FechaOrden").Width = 1200     '1300
    oGrilla.Bands(lBanda).Columns("FechaOrden").Header.Caption = "F.Atención"
    
    
    'Configura Values List
    SeteaListaEstado oGrilla, oGrilla.Bands(lBanda).Columns("idEstadoFacturacion")
    SeteaListaTipoFinanciamiento oGrilla, oGrilla.Bands(lBanda).Columns("IdTipoFinanciamiento")
    SeteaPuntosDeCarga oGrilla, oGrilla.Bands(lBanda).Columns("idPuntoCarga")

    oGrilla.Bands(lBanda).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("idPuntoCarga").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("idEstadoFacturacion").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("IdTipoFinanciamiento").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("NombreProducto").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("IdOrden").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("FechaOrden").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("NroComprobante").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(lBanda).Columns("Cantidad").Header.Caption = "Cant.Pagar"
    oGrilla.Bands(lBanda).Columns("Cantidad").Width = 700
    
    'oGrilla.Bands(lBanda).Columns("PrecioUnitario").Header.Caption = "Cant.Pagar"
    
    oGrilla.Bands(lBanda).Columns("TotalPorPagar").Header.Caption = "Tot.Pagar"
    oGrilla.Bands(lBanda).Columns("TotalPorPagar").Width = 900

    
    oGrilla.Bands(0).Columns("CantidadSIS").Header.Caption = "Can.SIS"
    oGrilla.Bands(0).Columns("PrecioSIS").Header.Caption = "Pr.SIS"
    oGrilla.Bands(0).Columns("ImporteSIS").Header.Caption = "Imp.SIS"
    oGrilla.Bands(lBanda).Columns("cantidadSIS").Width = 500
    oGrilla.Bands(lBanda).Columns("precioSIS").Width = 700
    oGrilla.Bands(lBanda).Columns("ImporteSIS").Width = 800
    oGrilla.Bands(lBanda).Columns("cantidadSIS").Format = "#0"
    oGrilla.Bands(lBanda).Columns("precioSIS").Format = "#0.00"
    oGrilla.Bands(lBanda).Columns("ImporteSIS").Format = "#0.00"
    If idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaSeguroSIS Then
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("precioSIS").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("precioSIS").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("precioSIS").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Hidden = True
       '
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Hidden = True
       '
       oGrilla.Bands(lBanda).Columns("EsConvenio").Hidden = True
       oGrilla.Bands(lBanda).Columns("CantidadConv").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioConv").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteConv").Hidden = True
    Else
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("precioSIS").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Activation = ssActivationActivateNoEdit
    End If
    
    oGrilla.Bands(0).Columns("CantidadSOAT").Header.Caption = "Can.SOAT"
    oGrilla.Bands(0).Columns("PrecioSOAT").Header.Caption = "Pr.SOAT"
    oGrilla.Bands(0).Columns("ImporteSOAT").Header.Caption = "Imp.SOAT"
    oGrilla.Bands(lBanda).Columns("CantidadSOAT").Width = 500
    oGrilla.Bands(lBanda).Columns("PrecioSOAT").Width = 700
    oGrilla.Bands(lBanda).Columns("ImporteSOAT").Width = 800
    oGrilla.Bands(lBanda).Columns("cantidadSOAT").Format = "#0"
    oGrilla.Bands(lBanda).Columns("precioSOAT").Format = "#0.00"
    oGrilla.Bands(lBanda).Columns("ImporteSOAT").Format = "#0.00"
    If idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaSeguroSOAT Then
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("precioSOAT").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("precioSOAT").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("precioSOAT").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Hidden = True
       oGrilla.Bands(lBanda).Columns("EsConvenio").Hidden = True
       oGrilla.Bands(lBanda).Columns("CantidadConv").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioConv").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteConv").Hidden = True
    Else
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("precioSOAT").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Activation = ssActivationActivateNoEdit
    End If
    
    oGrilla.Bands(0).Columns("Importeexo").Header.Caption = "Imp.EXO"
    oGrilla.Bands(lBanda).Columns("ImporteEXO").Width = 900
    oGrilla.Bands(lBanda).Columns("ImporteEXO").Format = "#0.00"
    If idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaServicioSocial Then
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Width = 500
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Width = 500
       oGrilla.Bands(lBanda).Columns("ImporteConv").Width = 500
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("EsConvenio").Hidden = True
       oGrilla.Bands(lBanda).Columns("CantidadConv").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioConv").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteConv").Hidden = True
    ElseIf idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaSeguroSIS And lbTieneDerechoExoneraSIS = True Then  'debb-25/10/2016
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Header.Appearance.ForeColor = vbWhite
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Activation = ssActivationAllowEdit
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Hidden = False
       oGrilla.Bands(lBanda).Columns("CantidadPagar").Width = 400
       oGrilla.Bands(lBanda).Columns("preciounitario").Format = "#0.00"
       oGrilla.Bands(lBanda).Columns("preciounitario").Width = 400
       oGrilla.Bands(lBanda).Columns("TotalPagar").Format = "#0.00"
       oGrilla.Bands(lBanda).Columns("TotalPagar").Width = 500
    Else
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Activation = ssActivationActivateNoEdit
    End If
    
    If idUsuarioConPermisoEnSISoEXOoSOATconf = sghTrabajaSeguroConvenios Then   'Convenios
       oGrilla.Bands(lBanda).Columns("CantidadConv").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("PrecioConv").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("ImporteConv").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("EsConvenio").Header.Appearance.BackColor = vbRed
       oGrilla.Bands(lBanda).Columns("CantidadConv").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("PrecioConv").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("ImporteConv").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("cantidad").Format = "#0"
       oGrilla.Bands(lBanda).Columns("precioConv").Format = "#0.00"
       oGrilla.Bands(lBanda).Columns("totalPorPagar").Format = "#0.00"
       oGrilla.Bands(0).Columns("CantidadConv").Header.Caption = "Cant.Conv"
       oGrilla.Bands(0).Columns("PrecioConv").Header.Caption = "Prec.Conv"
       oGrilla.Bands(0).Columns("TotalConv").Header.Caption = "Tot.Conv"
       oGrilla.Bands(lBanda).Columns("cantidadPagar").Hidden = True
       oGrilla.Bands(lBanda).Columns("PrecioUnitario").Hidden = True
       oGrilla.Bands(lBanda).Columns("TotalPagar").Hidden = True
       oGrilla.Bands(lBanda).Columns("cantidadSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteSOAT").Hidden = True
       oGrilla.Bands(lBanda).Columns("cantidadSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("precioSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteSIS").Hidden = True
       oGrilla.Bands(lBanda).Columns("ImporteEXO").Hidden = True
    Else
       oGrilla.Bands(lBanda).Columns("EsConvenio").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("PrecioConv").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("CantidadConv").Activation = ssActivationActivateNoEdit
       oGrilla.Bands(lBanda).Columns("ImporteConv").Activation = ssActivationActivateNoEdit
    End If
    
    oGrilla.Bands(lBanda).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(lBanda).Columns("TotalPorPagar").Activation = ssActivationActivateNoEdit
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
    
End Sub

Sub SeteaListaTipoFinanciamiento(oGrilla As SSUltraGrid, oColumn As SSColumn)
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim oValueTF As SSValueList
    
    If Not oGrilla.ValueLists.Exists("listaTipoFinanciamiento") Then
        Set oValueTF = oGrilla.ValueLists.Add("listaTipoFinanciamiento")
        Set rs = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarTodos
        Do While Not rs.EOF
            If rs!idTipoFinanciamiento <> 0 Then
                oValueTF.ValueListItems.Add Val(rs!idTipoFinanciamiento), Trim(rs!descripcion)
            End If
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
Dim i As Integer
Dim oValuePC As SSValueList
    
    If Not oGrilla.ValueLists.Exists("listaPuntosCarga") Then
        Set oValuePC = oGrilla.ValueLists.Add("listaPuntosCarga")
        Set rs = mo_ReglasComunes.SeleccionarPuntosDeCarga()
        Do While Not rs.EOF
            If rs!idPuntoCarga <> 0 Then
                oValuePC.ValueListItems.Add Val(rs!idPuntoCarga), Trim(rs!descripcion)
            End If
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
Dim i As Integer
Dim oValueEstado As SSValueList
    
    If Not oGrilla.ValueLists.Exists("listaEstadoFacturacion") Then
        Set oValueEstado = oGrilla.ValueLists.Add("listaEstadoFacturacion")
        Set rs = mo_ReglasFacturacion.EstadosFacturacionObtenerTodos
        Do While Not rs.EOF
            oValueEstado.ValueListItems.Add Val(rs!idestadofacturacion), Trim(rs!descripcion)
            rs.MoveNext
        Loop
        rs.Close
    Else
        Set oValueEstado = oGrilla.ValueLists.Item("listaEstadoFacturacion")
    End If
     
    Set oColumn.ValueList = oValueEstado
    
End Sub



Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height
   
End Sub

Sub LimpiarGrilla()
        If oConexion.State = 0 Then
            oConexion.Open sighEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 150
        End If
        Set grdProductos.DataSource = Nothing
        ml_idCuentaAtencion = -1000
        Dim lnTotalPagoSeguro As Double, lnTotalPagoDelPaciente As Double, lnTotalizaSeguros As Double
        Dim oRsCuentaCabecera As New Recordset, oRsCuentaDetalle As New Recordset
        Dim lnTipoConCeptoFarmacia As Integer, lnTotalApagar As Double
        CargaProductosPorIdCuentaAtencion lnTotalPagoSeguro, lnTotalPagoDelPaciente, lnTotalizaSeguros, oRsCuentaCabecera, oRsCuentaDetalle, lnTipoConCeptoFarmacia, lnTotalApagar

End Sub



Sub GenerarRecordsetProductos()
    Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
    oGenerarRecordsetProductos.GenerarRecordsetProductos mrs_FacturacionProductos
    
    'Set grdProductos.DataSource = mrs_FacturacionProductos
    
End Sub

'***************daniel barrantes**************
'***************PRORRATEA el TOTAL POR EXONERAR en cada ITEM
'***************
Sub ActualizaExoneracionesPorPorcentaje(lbExoneraTodos As Boolean)
    Dim lnDctos As Double: Dim lnTotalApagar As Double
    Dim lnCant As Long
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
          If mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Or mrs_FacturacionProductos.Fields!idestadofacturacion = sghConPreVenta Then
             lnCant = 0
             lnCant = lnCant + mrs_FacturacionProductos.Fields!CantidadSIS
             lnCant = lnCant + mrs_FacturacionProductos.Fields!CantidadSOAT
             lnCant = mrs_FacturacionProductos.Fields!CantidadPagar - lnCant
             lnDctos = mrs_FacturacionProductos!PrecioUnitario * mrs_FacturacionProductos.Fields!CantidadPagar
             If lbExoneraTodos = True Then
                mrs_FacturacionProductos.Fields!importeEXO = lnDctos
                mrs_FacturacionProductos.Update
                mrs_FacturacionProductos.Fields!TotalPorPagar = 0
                mrs_FacturacionProductos.Update
             Else
                mrs_FacturacionProductos.Fields!importeEXO = 0
                mrs_FacturacionProductos.Update
                mrs_FacturacionProductos.Fields!TotalPorPagar = lnDctos
                mrs_FacturacionProductos.Update
             End If
          End If
          mrs_FacturacionProductos.MoveNext
       Loop
       Set grdProductos.DataSource = mrs_FacturacionProductos
    End If
End Sub

'***************daniel barrantes**************
'***************Actualiza PRECIOS/IMPORTES para aquellos ITEMS que tienen
'***************CONVENIO MINSA-FOSPOLIS
Sub ActualizaConvenioEnTodosItems()
    If mrs_FacturacionProductos.RecordCount > 0 Then
       Dim lnPrecioCONV As Double: Dim lnPrecioNormal As Double
       Dim rs As Recordset
       Dim lcMensaje As String
       lcMensaje = ""
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
          If mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Then
             Select Case ms_TipoProducto
                Case sghServicio
                    Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigoDEBB(mrs_FacturacionProductos.Fields!Codigo, 4, 1)
                Case sghbien
                    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigoDEBB(mrs_FacturacionProductos.Fields!Codigo, 4, 1)
             End Select
             lnPrecioNormal = 0: lnPrecioCONV = 0
             If rs.RecordCount > 0 Then
                lnPrecioCONV = rs.Fields!PrecioUnitario
             Else
                lcMensaje = lcMensaje + Trim(mrs_FacturacionProductos.Fields!Codigo) & " - " & Trim(mrs_FacturacionProductos.Fields!NombreProducto) & " (NO TIENE CONVENIO)" & Chr(13)
                rs.Close
                Select Case ms_TipoProducto
                   Case sghServicio
                       Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigoDEBB(mrs_FacturacionProductos.Fields!Codigo, 1, 1)
                   Case sghbien
                       Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigoDEBB(mrs_FacturacionProductos.Fields!Codigo, 1, 1)
                End Select
                If rs.RecordCount > 0 Then
                   lnPrecioNormal = rs.Fields!PrecioUnitario
                End If
             End If
             If lnPrecioCONV > 0 Then
                mrs_FacturacionProductos.Fields!esConvenio = "Si"
                mrs_FacturacionProductos.Fields!precioCONV = lnPrecioCONV
                mrs_FacturacionProductos.Fields!cantidadConv = mrs_FacturacionProductos.Fields!CantidadPagar
                mrs_FacturacionProductos.Fields!ImporteConv = mrs_FacturacionProductos.Fields!CantidadPagar * lnPrecioCONV
                'mrs_FacturacionProductos.Fields!totalPorPagar = lnPrecioCONV * mrs_FacturacionProductos.Fields!cantidad
             Else
                mrs_FacturacionProductos.Fields!esConvenio = "No"
                mrs_FacturacionProductos.Fields!precioCONV = lnPrecioNormal
                mrs_FacturacionProductos.Fields!cantidadConv = 0
                mrs_FacturacionProductos.Fields!ImporteConv = 0
                mrs_FacturacionProductos.Fields!TotalPorPagar = lnPrecioNormal * mrs_FacturacionProductos.Fields!Cantidad
             End If
             mrs_FacturacionProductos.Update
          End If
          mrs_FacturacionProductos.MoveNext
       Loop
      ' Set grdProductos.DataSource = mrs_FacturacionProductos
       If lcMensaje <> "" Then
          MsgBox lcMensaje, vbInformation, "Mensaje"
       End If
    End If
End Sub


'***************daniel barrantes**************
'***************Actualiza PRECIOS para aquellos ITEMS que tienen
'***************SIS, SOAT
Sub ActualizaPreciosImportesEnTodosItemsParaSisSoat(lnTipoFinanciamiento As Long)
    If lnTipoFinanciamiento = 2 Or lnTipoFinanciamiento = 3 Then '2=SIS, 3=SOAT
        If mrs_FacturacionProductos.RecordCount > 0 Then
           Dim lnPrecioNormal As Double
           Dim rs As Recordset
           Dim lcMensaje As String
           lcMensaje = ""
           mrs_FacturacionProductos.MoveFirst
           Do While Not mrs_FacturacionProductos.EOF
              If mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Then
                 Select Case ms_TipoProducto
                    Case sghServicio
                        Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigo(mrs_FacturacionProductos.Fields!Codigo, lnTipoFinanciamiento, oConexion)
                    Case sghbien
                        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(mrs_FacturacionProductos.Fields!Codigo, lnTipoFinanciamiento, oConexion)
                 End Select
                 lnPrecioNormal = 0
                 If rs.RecordCount > 0 Then
                    lnPrecioNormal = rs.Fields!PrecioUnitario
                 Else
                    lcMensaje = lcMensaje + Trim(mrs_FacturacionProductos.Fields!Codigo) & " - " & Trim(mrs_FacturacionProductos.Fields!NombreProducto) & " (NO SE TRABAJARA)" & Chr(13)
                 End If
                 rs.Close
                 Select Case lnTipoFinanciamiento
                 Case 2                       'SIS
                       'mrs_FacturacionProductos.Fields!PrecioSIS = lnPrecioNormal
                       mrs_FacturacionProductos!ImporteSIS = mrs_FacturacionProductos.Fields!CantidadSIS * mrs_FacturacionProductos.Fields!precioSIS
                       mrs_FacturacionProductos.Update
                 Case 3                        'SOAT
                       'mrs_FacturacionProductos.Fields!PrecioSOAT = lnPrecioNormal
                       mrs_FacturacionProductos!ImporteSOAT = mrs_FacturacionProductos.Fields!CantidadSOAT * mrs_FacturacionProductos.Fields!PrecioSOAT
                       mrs_FacturacionProductos.Update
                 End Select
              End If
              mrs_FacturacionProductos.MoveNext
           Loop
           If lcMensaje <> "" Then
              MsgBox lcMensaje, vbInformation, "Mensaje"
           End If
        End If
    End If
End Sub

Function TotalizaPagoDelPaciente() As Double
    TotalizaPagoDelPaciente = 0
    If mrs_FacturacionProductos.RecordCount > 0 Then
       Dim lnTotal As Double
       mrs_FacturacionProductos.MoveFirst
       lnTotal = 0
       Do While Not mrs_FacturacionProductos.EOF
          If (mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Or mrs_FacturacionProductos.Fields!idestadofacturacion = sghConPreVenta) And (mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion) = sghTrabajaParticular Or mrs_FacturacionProductos.Fields!idTipoFinanciamiento = sghTrabajaServicioSocial) Then
               lnTotal = lnTotal + mrs_FacturacionProductos.Fields!TotalPorPagar
          End If
          mrs_FacturacionProductos.MoveNext
       Loop
       TotalizaPagoDelPaciente = lnTotal
    End If
End Function

Function TotalizaPagoDeSeguros() As Double
    TotalizaPagoDeSeguros = 0
    If mrs_FacturacionProductos.RecordCount > 0 Then
       Dim lnTotal As Double
       
       lnTotal = 0
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
            Select Case mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion)
            Case sghTrabajaSeguroSIS
                lnTotal = lnTotal + mrs_FacturacionProductos.Fields!ImporteSIS
            Case sghTrabajaSeguroSOAT
                lnTotal = lnTotal + mrs_FacturacionProductos.Fields!ImporteSOAT
            Case sghTrabajaSeguroConvenios
                lnTotal = lnTotal + mrs_FacturacionProductos.Fields!ImporteConv
            End Select
            mrs_FacturacionProductos.MoveNext
       Loop
       TotalizaPagoDeSeguros = lnTotal
    End If
End Function

Function TotalizaPagosDelPacienteConSeguro()
    TotalizaPagosDelPacienteConSeguro = 0
    If mrs_FacturacionProductos.RecordCount > 0 Then
       Dim lnTotal As Double
       lnTotal = 0
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
            lnTotal = lnTotal + mrs_FacturacionProductos.Fields!TotalPorPagar
            mrs_FacturacionProductos.MoveNext
       Loop
       TotalizaPagosDelPacienteConSeguro = lnTotal
    End If
End Function

Function TotalizaAdelantos() As Double
    TotalizaAdelantos = 0
    If mrs_FacturacionProductos.RecordCount > 0 Then
       Dim lnTotal As Double
       lnTotal = 0
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
            If mrs_FacturacionProductos.Fields!idProducto = lnIdPagosACuenta Then
               lnTotal = lnTotal + mrs_FacturacionProductos.Fields!Cantidad
            End If
            mrs_FacturacionProductos.MoveNext
       Loop
       TotalizaAdelantos = lnTotal
    End If
End Function

Sub CargaTodaLaCantidadPedidaHaciaCantidadSisSoat(lnTipoFinanciamiento As Long, lbCargaTodaLaCantidadPedida As Boolean)
    If lnTipoFinanciamiento = 2 Or lnTipoFinanciamiento = 3 Then '2=SIS, 3=SOAT
        If mrs_FacturacionProductos.RecordCount > 0 Then
           Dim lnPrecioNormal As Double
           Dim rs As Recordset
           Dim lcMensaje As String
           lcMensaje = ""
           mrs_FacturacionProductos.MoveFirst
           Do While Not mrs_FacturacionProductos.EOF
              'If mrs_FacturacionProductos!cantidadPagar > 0 Then
              If mrs_FacturacionProductos.Fields!idestadofacturacion = 1 And mrs_FacturacionProductos!CantidadPagar > 0 Then
                 Select Case lnTipoFinanciamiento
                 Case 2                       'SIS
                       If lbCargaTodaLaCantidadPedida = True And mrs_FacturacionProductos.Fields!precioSIS > 0 Then
                            mrs_FacturacionProductos.Fields!CantidadSIS = mrs_FacturacionProductos.Fields!CantidadPagar
                            mrs_FacturacionProductos!ImporteSIS = mrs_FacturacionProductos.Fields!CantidadPagar * mrs_FacturacionProductos.Fields!precioSIS
                            mrs_FacturacionProductos!Cantidad = 0
                            mrs_FacturacionProductos!TotalPorPagar = 0
                       Else
                            mrs_FacturacionProductos.Fields!CantidadSIS = 0
                            mrs_FacturacionProductos!ImporteSIS = 0
                            mrs_FacturacionProductos!Cantidad = mrs_FacturacionProductos.Fields!CantidadPagar
                            mrs_FacturacionProductos!TotalPorPagar = mrs_FacturacionProductos.Fields!CantidadPagar * mrs_FacturacionProductos.Fields!PrecioUnitario
                       End If
                       mrs_FacturacionProductos.Update
                 Case 3                        'SOAT
                       If lbCargaTodaLaCantidadPedida = True And mrs_FacturacionProductos.Fields!PrecioSOAT > 0 Then
                            mrs_FacturacionProductos.Fields!CantidadSOAT = mrs_FacturacionProductos.Fields!Cantidad
                            mrs_FacturacionProductos!ImporteSOAT = mrs_FacturacionProductos.Fields!Cantidad * mrs_FacturacionProductos.Fields!PrecioSOAT
                            mrs_FacturacionProductos!Cantidad = 0
                            mrs_FacturacionProductos!TotalPorPagar = 0
                       Else
                            mrs_FacturacionProductos.Fields!CantidadSOAT = 0
                            mrs_FacturacionProductos!ImporteSOAT = 0
                            mrs_FacturacionProductos!Cantidad = mrs_FacturacionProductos.Fields!CantidadPagar
                            mrs_FacturacionProductos!TotalPorPagar = mrs_FacturacionProductos.Fields!CantidadPagar * mrs_FacturacionProductos.Fields!PrecioUnitario
                       End If
                       mrs_FacturacionProductos.Update
                 End Select
              End If
              mrs_FacturacionProductos.MoveNext
           Loop
        End If
    End If
End Sub























Sub ActualizaEstadoAtencionEnGridServiciosYfarmacia(lnIdEstadoCuenta As Long, lcDocumentoReembolso As String)
    On Error Resume Next
    mrs_FacturacionProductos.MoveFirst
    Do While Not mrs_FacturacionProductos.EOF
       mrs_FacturacionProductos.Fields!idestadofacturacion = lnIdEstadoCuenta
       mrs_FacturacionProductos.Fields!docReembolso = lcDocumentoReembolso
       mrs_FacturacionProductos.Update
       mrs_FacturacionProductos.MoveNext
    Loop
End Sub



Function CPTesPAQUETEdisminuyeMedicamentosInsumos(oRsFarmaciaOriginal As Recordset, oRsServicios As Recordset, _
                                                                   oConexion As Connection) As Recordset
    On Error GoTo ErrCptEsPte
    Dim oRsTmpCab As New Recordset
    Dim oRsTmpDet As New Recordset
    Dim oRsFarmacia As New Recordset
    Dim lbTieneUnCaso As Boolean, lnCantidadNew As Long, lnCAntidadPqte As Long
    Set oRsFarmacia = HCigualDNI_DevuelveRsConHistoriaOCHOdigitos(oRsFarmaciaOriginal, "nroHistoriaClinica")
    Set CPTesPAQUETEdisminuyeMedicamentosInsumos = oRsFarmacia
    lbTieneUnCaso = False
    Set oRsTmpCab = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro("not cpt is null")
    If oRsTmpCab.RecordCount > 0 Then
       oRsServicios.MoveFirst
       Do While Not oRsServicios.EOF
          oRsTmpCab.MoveFirst
          oRsTmpCab.Find "cpt='" & Trim(oRsServicios!Codigo) & "'"
          If Not oRsTmpCab.EOF Then
             Set oRsTmpDet = mo_ReglasFacturacion.FacturacionCatalogoPaqueteFarmSeleccionarXid(oRsTmpCab!idFactPaquete)
             If oRsTmpDet.RecordCount > 0 Then
                oRsTmpDet.MoveFirst
                Do While Not oRsTmpDet.EOF
                   lnCAntidadPqte = oRsTmpDet!Cantidad
                   oRsFarmacia.Filter = "idProducto=" & oRsTmpDet!idProducto
                   If oRsFarmacia.RecordCount > 0 Then
                      oRsFarmacia.MoveFirst
                      Do While Not oRsFarmacia.EOF
                         lbTieneUnCaso = True
                         lnCantidadNew = oRsFarmacia!CantidadFinanciada - lnCAntidadPqte
                         If lnCantidadNew < 0 Then
                            lnCAntidadPqte = lnCAntidadPqte - oRsFarmacia!CantidadFinanciada
                            oRsFarmacia.Delete
                         Else
                            oRsFarmacia!CantidadFinanciada = lnCAntidadPqte
                            oRsFarmacia!TotalFinanciado = Round(lnCAntidadPqte * oRsFarmacia!precioFinanciado, 2)
                            Exit Do
                         End If
                         oRsFarmacia.MoveNext
                      Loop
                   End If
                   oRsTmpDet.MoveNext
                Loop
                oRsFarmacia.Filter = ""
             End If
             oRsTmpDet.Close
          End If
          oRsServicios.MoveNext
       Loop
    End If
    oRsTmpCab.Close
    If lbTieneUnCaso = True Then
       Set CPTesPAQUETEdisminuyeMedicamentosInsumos = oRsFarmacia
    End If
ErrCptEsPte:
    Set oRsTmpCab = Nothing
    Set oRsTmpDet = Nothing
End Function


Sub CargarItemsALaGrillaB_menorTiempo(rs As Recordset, ByRef lnTotalPagoSeguro As Double, ByRef lnTotalPagoDelPaciente As Double, _
                                      ByRef lnTotalizaPagosDelPacienteConSeguro As Double, ByRef oRsCuentaCabecera As Recordset, _
                                      ByRef oRsCuentaDetalle As Recordset, lnIdTipoConceptoFarmaciaPlanActual As Integer, _
                                      ByRef lnTotalApagar As Double)
    'Limpia temporal
    Set mrs_FacturacionProductos = Nothing
    GenerarRecordsetProductos
    '
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Dim oRs As New Recordset
            Dim oRsCatalogo As New Recordset
            Dim oRsFinanciamientos As New Recordset
            Dim oRsPagos As New Recordset
            Dim oRsDevoluciones As New Recordset
            Dim oRecetas As New Recordset
            Dim oRsFinanciamientosServ As New Recordset
            Dim lnCantidadSIS As Long: Dim lnPrecioSIS As Double: Dim lnImporteSIS As Double
            Dim lnCantidadSOAT As Long: Dim lnPrecioSOAT As Double: Dim lnImporteSOAT As Double
            Dim lnImporteEXO As Double: Dim ldFechaAutorizaSeguro As String: Dim lnIdUsuarioAutoriza As Long
            Dim lnIdTipoFinanciamiento As Long: Dim lcEsConvenio As String: Dim lnPrecioCONV As Double
            Dim lnCantidadPagar As Long: Dim lnPrecioPagar As Double: Dim lnTotalPagar As Double
            Dim lnIdEstadoFacturacion As Long: Dim lnIdComprobantePago As Long: Dim lcDocumentoPago As String
            Dim lnCantidadDev As Long: Dim lnIdComprobDev As Long: Dim lnIdEstadoDev As Long
            Dim lcFechaAutDev As String: Dim lnIdUsuarioAutDev As Long
            Dim LcMovNumero As String: Dim LcMovTipo As String
            Dim lnIdOrden As Long
            Dim oDoComprobantesPago As New DOCajaComprobantesPago
            Dim lnCantidadConv As Long: Dim lnImporteConv As Double
            Dim lnIdTipoConceptoFarmacia As Long
            Dim lnIdFuenteFinanciamiento As Long
            Dim lnPrecioDespacho As Double
            Dim lcProcedencia As String
            Dim lnImporteEnBoleta As Double
            Dim lnComoSeTrabajaEnEstadoCuenta As sghComoSeTrabajaEnEstadoCuentaLosSeguros
            Dim lcMovNumero111 As String, lcMovTipo111 As String
            Dim lnComoSeTrabajaEnEstadoCuenta1 As Long, lnIdOrden111 As Long
            Dim lnImpo As Double, lnPrec As Double, lnCant As Long, lcTexto As String, lcLlave As String
            Dim lbNuevo As Boolean, lnReceta As Long, lbGeneraReciboPago As Boolean

            lnTotalPagoSeguro = 0: lnTotalPagoDelPaciente = 0
            lnTotalizaPagosDelPacienteConSeguro = 0: lnTotalApagar = 0
            '
            Set oRsFinanciamientos = mo_ReglasFarmacia.FacturacionBienesFinanciamientosXcuenta(ml_idCuentaAtencion, oConexion)
            '
            Set oRsPagos = mo_ReglasFarmacia.FacturacionBienesPagosXCuenta(ml_idCuentaAtencion, oConexion)
            '
            Set oRsDevoluciones = mo_ReglasFarmacia.FacturacionBienesDevolucionesXcuenta(ml_idCuentaAtencion, oConexion)
            '
            Set oRecetas = mo_ReglasComunes.RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion(ml_idCuentaAtencion, False)
            '
            If wxParametro514 <> "S" Then
                Set oRsCatalogo = mo_ReglasComunes.CatalogoBienesInsumosHospSeleccionarTodos(oConexion)
            Else
                Set oRsCatalogo = mo_ReglasComunes.CatalogoBienesInsumosHospPorCuenta(oConexion, ml_idCuentaAtencion)
            End If
            '
            lnComoSeTrabajaEnEstadoCuenta1 = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion)
            'Existen CPT que ya viene incluidos Medicamentos/Insumos - hay que eliminarlos (inicio)
            lbGeneraReciboPago = False
            If sghTipoFinanciamiento.sghSis = ml_IdTipoFinanciamiento Then
               lbGeneraReciboPago = False
            ElseIf lnComoSeTrabajaEnEstadoCuenta = sghTrabajaSeguroSOAT Or lnComoSeTrabajaEnEstadoCuenta = sghTrabajaSeguroConvenios Then
               lbGeneraReciboPago = mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(ml_IdTipoFinanciamiento, oConexion)
            End If
            If wxParametro554 = "S" And lbGeneraReciboPago = False Then
                Set oRsFinanciamientosServ = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXcuentaConexion(ml_idCuentaAtencion, oConexion)
                oRsFinanciamientosServ.Filter = "idEstadoFacturacion<>9"
                oRsFinanciamientos.Filter = "idEstadoFacturacion<>9"
                If oRsFinanciamientos.RecordCount > 0 And oRsFinanciamientosServ.RecordCount > 0 Then
                   Set oRsFinanciamientos = CPTesPAQUETEdisminuyeMedicamentosInsumos(oRsFinanciamientos, oRsFinanciamientosServ, oConexion)
                Else
                   oRsFinanciamientos.Filter = ""
                End If
                oRsFinanciamientosServ.Close
            End If
            'Existen CPT que ya viene incluidos Medicamentos/Insumos - hay que eliminarlos (fin)
            '
            rs.MoveFirst
            LcMovTipo = "S": LcMovNumero = rs.Fields!movNumero
            Do While Not rs.EOF
                '
                lcMovNumero111 = rs.Fields!movNumero
                Do While Not rs.EOF And lcMovNumero111 = rs.Fields!movNumero
                    'Si tiene algun Seguro
                    lnCantidadSIS = 0: lnPrecioSIS = 0: lnImporteSIS = 0
                    lnCantidadSOAT = 0: lnPrecioSOAT = 0: lnImporteSOAT = 0
                    lnImporteEXO = 0: ldFechaAutorizaSeguro = "": lnIdUsuarioAutoriza = 0
                    lnIdTipoFinanciamiento = 0: lnIdFuenteFinanciamiento = 0
                    lcEsConvenio = "No": lnPrecioCONV = 0
                    lnCantidadConv = 0: lnImporteConv = 0: lnIdTipoConceptoFarmacia = 0
                    lnIdEstadoFacturacion = 0: lnComoSeTrabajaEnEstadoCuenta = sghTrabajaNinguno
                    'financiamientos
                    oRsFinanciamientos.Filter = "MovNumero = '" & lcMovNumero111 & _
                                                 "' and MovTipo='" & LcMovTipo & "'" & _
                                                 " and idProducto=" & rs.Fields!idProducto
                    If oRsFinanciamientos.RecordCount > 0 Then
                       oRsFinanciamientos.MoveFirst
                       lnIdEstadoFacturacion = rs.Fields!idEstadoMovimiento
                       Do While Not oRsFinanciamientos.EOF
                            If oRsFinanciamientos.Fields!idTipoFinanciamiento > 0 Then
                               lnComoSeTrabajaEnEstadoCuenta = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(oRsFinanciamientos.Fields!idTipoFinanciamiento, oConexion)
                            End If
                            If oRsFinanciamientos.Fields!IdFuenteFinanciamiento > 0 Then
                               lnIdTipoConceptoFarmacia = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(oRsFinanciamientos.Fields!IdFuenteFinanciamiento, oConexion)
                            End If
                            Select Case lnComoSeTrabajaEnEstadoCuenta
                            Case sghTrabajaSeguroSIS
                                 lnCantidadSIS = oRsFinanciamientos.Fields!CantidadFinanciada: lnPrecioSIS = oRsFinanciamientos.Fields!precioFinanciado
                                 lnImporteSIS = oRsFinanciamientos.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                                 lnIdEstadoFacturacion = oRsFinanciamientos.Fields!idestadofacturacion
                            Case sghTrabajaSeguroSOAT
                                 lnCantidadSOAT = oRsFinanciamientos.Fields!CantidadFinanciada: lnPrecioSOAT = oRsFinanciamientos.Fields!precioFinanciado
                                 lnImporteSOAT = oRsFinanciamientos.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                                 lnIdEstadoFacturacion = oRsFinanciamientos.Fields!idestadofacturacion
                            Case sghTrabajaSeguroConvenios
                                 lnCantidadPagar = oRsFinanciamientos.Fields!CantidadFinanciada: lnTotalPagar = oRsFinanciamientos.Fields!TotalFinanciado
                                 lnPrecioCONV = oRsFinanciamientos.Fields!precioFinanciado: lcEsConvenio = "Si"
                                 lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnCantidadConv = oRsFinanciamientos.Fields!CantidadFinanciada
                                 lnImporteConv = oRsFinanciamientos.Fields!TotalFinanciado
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                                 lnIdEstadoFacturacion = oRsFinanciamientos.Fields!idestadofacturacion
                            Case Else           'exoneraciones/particular hospitalizado
                                 lnImporteEXO = oRsFinanciamientos.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                            End Select
                            oRsFinanciamientos.MoveNext
                       Loop
                    Else
                       lnIdTipoFinanciamiento = rs.Fields!idTipoFinanciamiento
                    End If
                    'Pagos
                    lnCantidadPagar = 0: lnPrecioPagar = 0: lnTotalPagar = 0: lnIdOrden = 0
                    lnIdComprobantePago = 0: lcDocumentoPago = "": lnImporteEnBoleta = 0
                    oRsPagos.Filter = "MovNumero = '" & lcMovNumero111 & "' and MovTipo='" & LcMovTipo & "'" & _
                                       " and idProducto=" & rs.Fields!idProducto
                    If oRsPagos.RecordCount > 0 Then
                        oRsPagos.MoveLast
                        'If oRsPagos.Fields!IdEstadoFacturacion = 1 Then
                            lnCantidadPagar = oRsPagos.Fields!CantidadPagar: lnPrecioPagar = oRsPagos.Fields!PrecioVenta
                            lnTotalPagar = oRsPagos.Fields!TotalPagar - lnImporteEXO: lnIdEstadoFacturacion = oRsPagos.Fields!idestadofacturacion
                            lnIdComprobantePago = IIf(IsNull(oRsPagos.Fields!IdComprobantePago), 0, oRsPagos.Fields!IdComprobantePago)
                            lnIdOrden = oRsPagos.Fields!IdOrden
                            If lnIdComprobantePago > 0 Then
                                lnImporteEnBoleta = oRsPagos.Fields!TotalPagar
                                Set oDoComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePago, oConexion)
                                lcDocumentoPago = Trim(oDoComprobantesPago.nroSerie) + "-" + Trim(oDoComprobantesPago.nrodocumento)
                                lnTotalPagar = 0
                            End If
                        'End If
                    End If
                    'Devoluciones
                    lnCantidadDev = 0: lnIdComprobDev = 0: lnIdEstadoDev = 0: lcFechaAutDev = ""
                    lnIdUsuarioAutDev = 0
                    oRsDevoluciones.Filter = "MovNumero= '" & lcMovNumero111 & "' and MovTipo='S'" & _
                                             " and idProducto=" & rs.Fields!idProducto
                    If oRsDevoluciones.RecordCount > 0 Then
                        lnIdComprobDev = IIf(IsNull(oRsDevoluciones.Fields!IdComprobantePago), 0, oRsDevoluciones.Fields!IdComprobantePago)
                        lnIdEstadoDev = oRsDevoluciones.Fields!idEstadoDevolucion: lcFechaAutDev = oRsDevoluciones.Fields!FechaAutoriza
                        lnIdUsuarioAutDev = oRsDevoluciones.Fields!IdUsuarioAutoriza
                        Do While Not oRsDevoluciones.EOF
                           lnCantidadDev = lnCantidadDev + oRsDevoluciones.Fields!CantidadAdevolver
                           oRsDevoluciones.MoveNext
                        Loop
                    End If
                    '
                    If lnPrecioPagar = 0 Then
                        oRsCatalogo.Filter = "idProducto=" & rs.Fields!idProducto & " and IdTipoFinanciamiento=1"
                        lnPrecioDespacho = 0
                        If oRsCatalogo.RecordCount > 0 Then
                           lnPrecioDespacho = oRsCatalogo.Fields!PrecioUnitario
                        End If
                    Else
                        lnPrecioDespacho = lnPrecioPagar
                    End If
                    'Actualiza Precio del SEGURO, en caso sea igual a CERO
                    Select Case lnComoSeTrabajaEnEstadoCuenta1
                    Case sghTrabajaSeguroSIS
                         If lnPrecioSIS = 0 Then
                              oRsCatalogo.Filter = "idProducto=" & rs.Fields!idProducto & " and IdTipoFinanciamiento=" & ml_IdTipoFinanciamiento
                              If oRsCatalogo.RecordCount > 0 Then
                                 lnPrecioSIS = oRsCatalogo.Fields!PrecioUnitario
                              End If

                         End If
                         If lnPrecioDespacho = 0 Then
                            lnPrecioDespacho = lnPrecioSIS
                         End If
                    Case sghTrabajaSeguroSOAT
                         If lnPrecioSOAT = 0 Then
                              oRsCatalogo.Filter = "idProducto=" & rs.Fields!idProducto & " and IdTipoFinanciamiento=" & ml_IdTipoFinanciamiento
                              If oRsCatalogo.RecordCount > 0 Then
                                 lnPrecioSOAT = oRsCatalogo.Fields!PrecioUnitario
                              End If
                         End If
                         If lnPrecioDespacho = 0 Then
                            lnPrecioDespacho = lnPrecioSOAT
                         End If
                    Case sghTrabajaSeguroConvenios
                         If lnPrecioCONV = 0 Then
                              oRsCatalogo.Filter = "idProducto=" & rs.Fields!idProducto & " and IdTipoFinanciamiento=" & ml_IdTipoFinanciamiento
                              If oRsCatalogo.RecordCount > 0 Then
                                 lnPrecioCONV = oRsCatalogo.Fields!PrecioUnitario
                              End If
                         End If
                         If lnPrecioDespacho = 0 Then
                            lnPrecioDespacho = lnPrecioCONV
                         End If
                    End Select
                    '
                    lnReceta = 0
                    oRecetas.Filter = "codigo='" & rs!Codigo & "' and documentoDespacho='" & rs!DocumentoNumero & "'"
                    If oRecetas.RecordCount > 0 Then
                       lnReceta = oRecetas!idReceta
                    End If
                    '
                    mrs_FacturacionProductos.AddNew
                    mrs_FacturacionProductos!movNumero = rs!movNumero
                    mrs_FacturacionProductos!MovTipo = "S"
                    mrs_FacturacionProductos!idProducto = rs!idProducto
                    mrs_FacturacionProductos!Codigo = rs!Codigo
                    mrs_FacturacionProductos!NombreProducto = rs!nombre
                    mrs_FacturacionProductos!CantidadPagar = (rs!Cantidad - lnCantidadDev) 'cantidad inicial (no varia)....menos Cantidad Devuelta(NI)
                    mrs_FacturacionProductos!PrecioUnitario = lnPrecioDespacho  'precio de venta
                    mrs_FacturacionProductos!TotalPagar = Round(lnPrecioDespacho * (rs!Cantidad - lnCantidadDev), 2)   '....menos Cantidad Devuelta (NI)
                    mrs_FacturacionProductos!CantidadSIS = lnCantidadSIS
                    mrs_FacturacionProductos!precioSIS = lnPrecioSIS
                    mrs_FacturacionProductos!ImporteSIS = lnImporteSIS
                    mrs_FacturacionProductos!CantidadSOAT = lnCantidadSOAT
                    mrs_FacturacionProductos!PrecioSOAT = lnPrecioSOAT
                    mrs_FacturacionProductos!ImporteSOAT = lnImporteSOAT
                    mrs_FacturacionProductos!importeEXO = lnImporteEXO
                    mrs_FacturacionProductos!idPuntoCarga = 5
                    mrs_FacturacionProductos!idestadofacturacion = lnIdEstadoFacturacion    'IIf(lnIdUsuarioAutoriza > 0, rs.Fields!idEstadoMovimiento, lnIdEstadoFacturacion)  '
                    mrs_FacturacionProductos!Cantidad = lnCantidadPagar 'cantidad a pagar en caja (varia)
                    mrs_FacturacionProductos!TotalPorPagar = lnTotalPagar  '(a pagar en caja)
                    mrs_FacturacionProductos!IdComprobantePago = lnIdComprobantePago
                    mrs_FacturacionProductos!IdOrden = lnIdOrden
                    mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = lnIdUsuarioAutoriza
                    If lnIdTipoConceptoFarmacia > 0 Then
                       mrs_FacturacionProductos!FechaAutorizaSeguro = ldFechaAutorizaSeguro
                    End If
                    mrs_FacturacionProductos!IdUsuarioAutorizaDevolucion = lnIdUsuarioAutDev
                    mrs_FacturacionProductos!FechaAutorizaDevolucion = IIf(lnCantidadDev = 0, 0, lcFechaAutDev)
                    mrs_FacturacionProductos!IdComprobantePagoDevolucion = lnIdComprobDev
                    mrs_FacturacionProductos!NroComprobante = IIf(lcDocumentoPago = "", rs!DocumentoNumero, lcDocumentoPago)  'si ya se PAGO muestra BOLETA sino muestra TICKET
                    mrs_FacturacionProductos!idTipoFinanciamiento = lnIdTipoFinanciamiento
                    mrs_FacturacionProductos!precioCONV = lnPrecioCONV
                    mrs_FacturacionProductos!esConvenio = lcEsConvenio
                    mrs_FacturacionProductos!FechaOrden = rs!fechacreacion
                    mrs_FacturacionProductos!NombreServicio = rs!dalmacen
                    mrs_FacturacionProductos!cantidadConv = lnCantidadConv
                    mrs_FacturacionProductos!ImporteConv = lnImporteConv
                    mrs_FacturacionProductos!idTipoConceptoFarmacia = lnIdTipoConceptoFarmacia
                    mrs_FacturacionProductos!IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
                    If Not IsNull(rs!idServicioPaciente) Then
                        mrs_FacturacionProductos!ServicioDeEstancia = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(rs!idServicioPaciente)
                        mrs_FacturacionProductos!idServicioDeEstancia = rs!idServicioPaciente
                    End If
                    mrs_FacturacionProductos!CantidadDevuelta = lnCantidadDev
                    mrs_FacturacionProductos!nrodocumento = rs!DocumentoNumero
                    If ml_AgruparPor = 3 Or ml_AgruparPor = 5 Then
                       mrs_FacturacionProductos!descripcion = rs!dfinanciamiento
                    End If
                    mrs_FacturacionProductos!ImporteEnBoleta = lnImporteEnBoleta
                    mrs_FacturacionProductos!nroDcto = rs!DocumentoNumero
                    mrs_FacturacionProductos!ComoSeTrabajaEnEstadoCuenta = lnComoSeTrabajaEnEstadoCuenta
                    mrs_FacturacionProductos!FechaDespacho = rs!fechacreacion
                    mrs_FacturacionProductos!esPaquete = IIf(IsNull(rs!esPaquete), False, rs!esPaquete)
                    mrs_FacturacionProductos!Receta = lnReceta
                    'Function TotalizaPagoDelPaciente()
                    If (mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Or _
                          mrs_FacturacionProductos.Fields!idestadofacturacion = sghConPreVenta) And _
                          (lnComoSeTrabajaEnEstadoCuenta1 = sghTrabajaParticular Or _
                           mrs_FacturacionProductos.Fields!idTipoFinanciamiento = sghTrabajaServicioSocial) Then
                      lnTotalPagoDelPaciente = lnTotalPagoDelPaciente + mrs_FacturacionProductos.Fields!TotalPorPagar
                    End If
                    'Function TotalizaPagoDeSeguros()
                    Select Case lnComoSeTrabajaEnEstadoCuenta1
                    Case sghTrabajaSeguroSIS
                        lnTotalPagoSeguro = lnTotalPagoSeguro + mrs_FacturacionProductos.Fields!ImporteSIS
                    Case sghTrabajaSeguroSOAT
                        lnTotalPagoSeguro = lnTotalPagoSeguro + mrs_FacturacionProductos.Fields!ImporteSOAT
                    Case sghTrabajaSeguroConvenios
                        lnTotalPagoSeguro = lnTotalPagoSeguro + mrs_FacturacionProductos.Fields!ImporteConv
                    End Select
                    'TotalizaPagosDelPacienteConSeguro()
                    lnTotalizaPagosDelPacienteConSeguro = lnTotalizaPagosDelPacienteConSeguro + mrs_FacturacionProductos.Fields!TotalPorPagar
                    'Resumen-Farmacia
                    
                    If mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Or mrs_FacturacionProductos.Fields!idestadofacturacion = 4 Then
                        Set oRs = mo_ReglasComunes.FactPuntosCargaSeleccionarPorId(mrs_FacturacionProductos.Fields!idPuntoCarga, oConexion)
                        lcTexto = ""
                        If oRs.RecordCount > 0 Then
                           lcTexto = Trim(oRs.Fields!descripcion)
                        End If
                        oRs.Close
                        Select Case lnIdTipoConceptoFarmaciaPlanActual
                        Case sghTipoConceptoFarmacia.sghTipoConceptoSIS
                             lnCant = mrs_FacturacionProductos.Fields!CantidadSIS
                             lnPrec = mrs_FacturacionProductos.Fields!precioSIS
                             lnImpo = mrs_FacturacionProductos.Fields!ImporteSIS
                        Case sghTipoConceptoFarmacia.sghTipoConceptoSOAT
                             lnCant = mrs_FacturacionProductos.Fields!CantidadSOAT
                             lnPrec = mrs_FacturacionProductos.Fields!PrecioSOAT
                             lnImpo = mrs_FacturacionProductos.Fields!ImporteSOAT
                        Case sghTipoConceptoFarmacia.sghTipoConceptoConvenios
                             lnCant = mrs_FacturacionProductos.Fields!cantidadConv
                             lnPrec = mrs_FacturacionProductos.Fields!precioCONV
                             lnImpo = mrs_FacturacionProductos.Fields!ImporteConv
                        Case Else
                             lnCant = mrs_FacturacionProductos.Fields!Cantidad
                             lnPrec = mrs_FacturacionProductos.Fields!PrecioUnitario
                             lnImpo = mrs_FacturacionProductos.Fields!TotalPorPagar
                        End Select
                        lcLlave = lcTexto & " - " & mrs_FacturacionProductos.Fields!FechaOrden
                        lbNuevo = True
                        If oRsCuentaCabecera.RecordCount > 0 Then
                           oRsCuentaCabecera.MoveFirst
                           oRsCuentaCabecera.Find "llave='" & lcLlave & "'"
                           If Not oRsCuentaCabecera.EOF Then
                              lbNuevo = False
                           End If
                        End If
                        If lbNuevo Then
                              oRsCuentaCabecera.AddNew
                              oRsCuentaCabecera.Fields!llave = lcLlave
                              oRsCuentaCabecera.Fields!puntoDeCarga = lcTexto
                              oRsCuentaCabecera.Fields!fecha = mrs_FacturacionProductos.Fields!FechaDespacho
                              oRsCuentaCabecera.Fields!Servicio = mrs_FacturacionProductos.Fields!ServicioDeEstancia
                              oRsCuentaCabecera.Fields!Importe = lnImpo
                              oRsCuentaCabecera.Fields!nrodocumento = mrs_FacturacionProductos.Fields!nrodocumento
                        Else
                              oRsCuentaCabecera.Fields!Importe = oRsCuentaCabecera.Fields!Importe + lnImpo
                        End If
                        oRsCuentaCabecera.Update
                        oRsCuentaDetalle.AddNew
                        oRsCuentaDetalle.Fields!llave = lcLlave
                        oRsCuentaDetalle.Fields!Codigo = mrs_FacturacionProductos.Fields!Codigo
                        oRsCuentaDetalle.Fields!descripcion = Left(mrs_FacturacionProductos.Fields!NombreProducto, 150)
                        oRsCuentaDetalle.Fields!Cantidad = lnCant
                        oRsCuentaDetalle.Fields!Precio = lnPrec
                        oRsCuentaDetalle.Fields!Importe = lnImpo
                        oRsCuentaDetalle.Fields!CantDevuelta = mrs_FacturacionProductos.Fields!CantidadDevuelta
                        If mrs_FacturacionProductos.Fields!idestadofacturacion = 4 Then
                           oRsCuentaDetalle.Fields!nrodocumento = mrs_FacturacionProductos.Fields!NroComprobante
                        End If
                        oRsCuentaDetalle.Update
                        lnTotalApagar = lnTotalApagar + lnImpo
                    End If
                    '
                    rs.MoveNext
                    If rs.EOF Then
                       Exit Sub
                    End If
                Loop
            Loop
            Set oRs = Nothing
            oRsFinanciamientos.Close: Set oRsFinanciamientos = Nothing
            oRsPagos.Close: Set oRsPagos = Nothing
            oRsDevoluciones.Close: Set oRsDevoluciones = Nothing
            oRsCatalogo.Close: Set oRsCatalogo = Nothing
            Set oRecetas = Nothing
            Set oRsFinanciamientosServ = Nothing
        End If
    End If
End Sub

Sub CargarItemsALaGrillaS_menorTiempo(rs As Recordset, ByRef lnTotalPagoSeguro As Double, ByRef lnTotalPagoDelPaciente As Double, ByRef lnTotalizaPagosDelPacienteConSeguro As Double, ByRef oRsCuentaCabecera As Recordset, ByRef oRsCuentaDetalle As Recordset, lnIdTipoConceptoFarmaciaPlanActual As Integer, ByRef lnTotalApagar As Double)
    'Limpia temporal
    Set mrs_FacturacionProductos = Nothing
    GenerarRecordsetProductos
    '
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Dim oRs As New Recordset
            Dim oRsCatalogo As New Recordset
            Dim oRsFinanciamientos As New Recordset
            Dim oRsPagos As New Recordset
            Dim oRsDevoluciones As New Recordset
            Dim oRecetas As New Recordset
            Dim lnCantidadSIS As Long: Dim lnPrecioSIS As Double: Dim lnImporteSIS As Double
            Dim lnCantidadSOAT As Long: Dim lnPrecioSOAT As Double: Dim lnImporteSOAT As Double
            Dim lnImporteEXO As Double: Dim ldFechaAutorizaSeguro As String: Dim lnIdUsuarioAutoriza As Long
            Dim lnIdTipoFinanciamiento As Long: Dim lcEsConvenio As String: Dim lnPrecioCONV As Double
            Dim lnCantidadPagar As Long: Dim lnPrecioPagar As Double: Dim lnTotalPagar As Double
            Dim lnIdEstadoFacturacion As Long: Dim lnIdComprobantePago As Long: Dim lcDocumentoPago As String
            Dim lnCantidadDev As Long: Dim lnIdComprobDev As Long: Dim lnIdEstadoDev As Long
            Dim lcFechaAutDev As String: Dim lnIdUsuarioAutDev As Long
            Dim LcMovNumero As String: Dim LcMovTipo As String
            Dim lnIdOrden As Long
            Dim oDoComprobantesPago As New DOCajaComprobantesPago
            Dim lnCantidadConv As Long: Dim lnImporteConv As Double
            Dim lnIdTipoConceptoFarmacia As Long
            Dim lnIdFuenteFinanciamiento As Long
            Dim lnPrecioDespacho As Double
            Dim lnImporteEnBoleta As Double
            Dim lcNroDcto As String
            Dim lnIdOrdenPago As Long
            Dim lnComoSeTrabajaEnEstadoCuenta As sghComoSeTrabajaEnEstadoCuentaLosSeguros
            Dim lbElMovimientoNoEstaAnulado As Boolean
            Dim lnComoSeTrabajaEnEstadoCuenta1 As Long, lnIdOrden111 As Long
            Dim lnImpo As Double, lnPrec As Double, lnCant As Long, lcTexto As String, lcLlave As String
            Dim lbNuevo As Boolean, lnReceta As Long, lcObservacionesCaja As String, lcIdProductoEPS As String
            lcIdProductoEPS = "/" & wxparametro563 & "/" & wxparametro564 & "/" & wxparametro565 & "/" & wxparametro566 & "/"
            '
            lnTotalPagoSeguro = 0: lnTotalPagoDelPaciente = 0: lnTotalizaPagosDelPacienteConSeguro = 0: lnTotalApagar = 0
            '
            Set oRsFinanciamientos = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXcuentaConexion(ml_idCuentaAtencion, oConexion)
            '
            Set oRsPagos = mo_ReglasFacturacion.FacturacionServicioPagosXcuentaYconexion(ml_idCuentaAtencion, oConexion)
            '
            Set oRsDevoluciones = mo_ReglasFacturacion.FacturacionServicioDevolucionesXcuenta(ml_idCuentaAtencion, oConexion)
            '
            Set oRecetas = mo_ReglasComunes.RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion(ml_idCuentaAtencion, False)
            '
            If wxParametro514 <> "S" Then
               Set oRsCatalogo = mo_ReglasComunes.CatalogoServiciosSeleccionarTodos(oConexion)
            Else
               Set oRsCatalogo = mo_ReglasComunes.CatalogoServiciosHospPorCuenta(oConexion, ml_idCuentaAtencion)
            End If
            '
            lnComoSeTrabajaEnEstadoCuenta1 = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(ml_IdTipoFinanciamiento, oConexion)
            rs.MoveFirst
            lbTieneQueGrabarAntesDeImprimir = False
            Do While Not rs.EOF
                lnIdOrden111 = rs.Fields!IdOrden
                Do While Not rs.EOF And lnIdOrden111 = rs.Fields!IdOrden
                    'Si tiene algun Seguro
                    lnCantidadSIS = 0: lnPrecioSIS = 0: lnImporteSIS = 0
                    lnCantidadSOAT = 0: lnPrecioSOAT = 0: lnImporteSOAT = 0
                    lnImporteEXO = 0: ldFechaAutorizaSeguro = "": lnIdUsuarioAutoriza = 0
                    lnIdTipoFinanciamiento = 0: lnIdFuenteFinanciamiento = 0
                    lcEsConvenio = "No": lnPrecioCONV = 0
                    lnCantidadConv = 0: lnImporteConv = 0: lnIdTipoConceptoFarmacia = 0
                    lnIdEstadoFacturacion = 0: lnComoSeTrabajaEnEstadoCuenta = sghTrabajaNinguno
                    oRsFinanciamientos.Filter = "idOrden = " & rs.Fields!IdOrden & _
                                                " and idProducto=" & rs.Fields!idProducto
                    If oRsFinanciamientos.RecordCount > 0 Then
                       oRsFinanciamientos.MoveFirst
                       lnIdEstadoFacturacion = rs!idestadofacturacion
                       Do While Not oRsFinanciamientos.EOF
                            If oRsFinanciamientos.Fields!idTipoFinanciamiento > 0 Then
                               lnComoSeTrabajaEnEstadoCuenta = mo_ReglasFacturacion.TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(oRsFinanciamientos.Fields!idTipoFinanciamiento, oConexion)
                            End If
                            If oRsFinanciamientos.Fields!IdFuenteFinanciamiento > 0 Then
                               lnIdTipoConceptoFarmacia = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(oRsFinanciamientos.Fields!IdFuenteFinanciamiento, oConexion)
                            End If
                            Select Case lnComoSeTrabajaEnEstadoCuenta
                            Case sghTrabajaSeguroSIS
                                 lnCantidadSIS = oRsFinanciamientos.Fields!CantidadFinanciada: lnPrecioSIS = oRsFinanciamientos.Fields!precioFinanciado
                                 lnImporteSIS = oRsFinanciamientos.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                                 lnIdEstadoFacturacion = oRsFinanciamientos!idestadofacturacion
                            Case sghTrabajaSeguroSOAT
                                 lnCantidadSOAT = oRsFinanciamientos.Fields!CantidadFinanciada: lnPrecioSOAT = oRsFinanciamientos.Fields!precioFinanciado
                                 lnImporteSOAT = oRsFinanciamientos.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                                 lnIdEstadoFacturacion = oRsFinanciamientos!idestadofacturacion
                            Case sghTrabajaSeguroConvenios
                                 lnPrecioCONV = oRsFinanciamientos.Fields!precioFinanciado: lcEsConvenio = "Si"
                                 lnImporteConv = oRsFinanciamientos.Fields!TotalFinanciado
                                 lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnCantidadConv = oRsFinanciamientos.Fields!CantidadFinanciada
                                 
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                                 lnIdEstadoFacturacion = oRsFinanciamientos!idestadofacturacion
                            Case Else           'exoneraciones/particular hospitalizado
                                 lnImporteEXO = oRsFinanciamientos.Fields!TotalFinanciado: lnIdTipoFinanciamiento = oRsFinanciamientos.Fields!idTipoFinanciamiento
                                 ldFechaAutorizaSeguro = oRsFinanciamientos.Fields!FechaAutoriza
                                 lnIdUsuarioAutoriza = oRsFinanciamientos.Fields!IdUsuarioAutoriza
                                 lnIdFuenteFinanciamiento = oRsFinanciamientos.Fields!IdFuenteFinanciamiento
                            End Select
                            oRsFinanciamientos.MoveNext
                       Loop
                    Else
                       lnIdTipoFinanciamiento = rs.Fields!idTipoFinanciamiento
                    End If
                    'Pagos
                    lnCantidadPagar = 0: lnPrecioPagar = 0: lnTotalPagar = 0: lnIdOrden = 0
                    lnIdComprobantePago = 0: lcDocumentoPago = "": lnImporteEnBoleta = 0
                    lnIdOrdenPago = 0: lcObservacionesCaja = ""
                    oRsPagos.Filter = "idOrden = " & rs.Fields!IdOrden & _
                                      " and idProducto=" & rs.Fields!idProducto
                    If oRsPagos.RecordCount > 0 Then
                        oRsPagos.MoveLast
                        'If oRsPagos.Fields!IdEstadoFacturacion = 1 Then
                            lnCantidadPagar = oRsPagos.Fields!Cantidad: lnPrecioPagar = oRsPagos.Fields!Precio
                            lnTotalPagar = oRsPagos.Fields!Total - lnImporteEXO: lnIdEstadoFacturacion = oRsPagos.Fields!idestadofacturacion
                            lnIdComprobantePago = IIf(IsNull(oRsPagos.Fields!IdComprobantePago), 0, oRsPagos.Fields!IdComprobantePago)
                            lnIdOrden = rs.Fields!IdOrden
                            lnIdOrdenPago = oRsPagos.Fields!IdOrdenPago
                            If lnIdComprobantePago > 0 Then
                                lnImporteEnBoleta = oRsPagos.Fields!Total
                                Set oDoComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePago, oConexion)
                                lcDocumentoPago = Trim(oDoComprobantesPago.nroSerie) + "-" + Trim(oDoComprobantesPago.nrodocumento)
                                lnTotalPagar = 0
                                If InStr(lcIdProductoEPS, Trim(Str(rs!idProducto))) > 0 Then
                                   lcObservacionesCaja = oDoComprobantesPago.Observaciones
                                End If
                            End If
                        'End If
                    End If
                    'Devoluciones
                    lnCantidadDev = 0: lnIdComprobDev = 0: lnIdEstadoDev = 0: lcFechaAutDev = ""
                    lnIdUsuarioAutDev = 0
                    oRsDevoluciones.Filter = "idOrden = " & rs.Fields!IdOrden & _
                                             " and idProducto=" & rs.Fields!idProducto
                    If oRsDevoluciones.RecordCount > 0 Then
                        lnCantidadDev = oRsDevoluciones.Fields!CantidadAdevolver: lnIdComprobDev = oRsDevoluciones.Fields!IdComprobantePago
                        lnIdEstadoDev = oRsDevoluciones.Fields!idEstadoDevolucion: lcFechaAutDev = oRsDevoluciones.Fields!FechaAutoriza
                        lnIdUsuarioAutDev = oRsDevoluciones.Fields!IdUsuarioAutoriza
                    End If
                    'Actualiza Precios Contado
                    If lnPrecioPagar = 0 Then
                        oRsCatalogo.Filter = "idProducto=" & rs!idProducto & " and IdTipoFinanciamiento=1"
                        lnPrecioDespacho = 0
                        If oRsCatalogo.RecordCount > 0 Then
                           If rs!idPuntoCarga = sghPtoCargaAdmisionHospitalizacion Then
                              'Estancia hospitalaria
                              If wxParametro511 = "S" Then
                                 lnPrecioDespacho = oRsCatalogo.Fields!PrecioUnitario
                              Else
                                 lnPrecioDespacho = oRsCatalogo.Fields!PrecioUnitario / 24
                              End If
                           Else
                              lnPrecioDespacho = oRsCatalogo.Fields!PrecioUnitario
                           End If
                        End If
                    Else
                        lnPrecioDespacho = lnPrecioPagar
                    End If
                    'Nro de Documento, para mostrar en Pantalla e Impresora
                    lbElMovimientoNoEstaAnulado = True
                    lcNroDcto = Trim(Str(rs!IdOrden))
                    Select Case rs!idPuntoCarga
                    Case sghPtoCargaEcogGeneral, sghPtoCargaRayosX, sghPtoCargaTomografia, sghPtoCargaEcogObstetrica                      'Imagenes
                         Set oRs = mo_ReglasImagenes.ImagMovimientoImagenesXidOrden(rs!IdOrden, oConexion)
                         If oRs.RecordCount > 0 Then
                            lcNroDcto = Trim(Str(oRs.Fields!IdMovimiento))
                         Else
                            lbElMovimientoNoEstaAnulado = False
                         End If
                         oRs.Close
                    Case sghPtoCargaPatologiaClinica, sghPtoCargaAnatomiaPatologica1, sghPtoCargaBancoSangre1    'Laboratorio
                         Set oRs = mo_ReglasLaboratorio.LabMovimientoLaboratorioXidOrden(rs!IdOrden, oConexion)
                         If oRs.RecordCount > 0 Then
                            lcNroDcto = Trim(Str(oRs.Fields!IdMovimiento))
                         Else
                            lbElMovimientoNoEstaAnulado = False
                         End If
                         oRs.Close
                    Case sghPtoCargaCaja          'se genero en CAJA
                        If lcNroDcto <> "" Then
                           lcNroDcto = lcDocumentoPago
                        End If
                    End Select
                    'Actualiza Precio del SEGURO, en caso sea igual a CERO
                    Select Case lnComoSeTrabajaEnEstadoCuenta1
                    Case sghTrabajaSeguroSIS
                         If lnPrecioSIS = 0 Then
                              oRsCatalogo.Filter = "idProducto=" & rs!idProducto & " and IdTipoFinanciamiento=" & ml_IdTipoFinanciamiento
                              If oRsCatalogo.RecordCount > 0 Then
                                 lnPrecioSIS = oRsCatalogo.Fields!PrecioUnitario
                              End If
                         End If
                         If lnPrecioDespacho = 0 Then
                            lnPrecioDespacho = lnPrecioSIS
                         End If
                         'debb-25/10/2016
                         If lnPrecioSIS = 0 And lnTotalPagar = 0 Then
                              lnTotalPagar = lnCantidadSIS * lnPrecioDespacho
                              If lnCantidadPagar = 0 Then lnCantidadPagar = lnCantidadSIS
                              lnCantidadSIS = 0
                              If lnIdOrdenPago = 0 Then lbTieneQueGrabarAntesDeImprimir = True
                         End If
'                         If lnPrecioSIS = 0 Or lnCantidadSIS = 0 Then
'                            lnCantidadSIS = 0
'                            lnImporteSIS = 0
'                            lnPrecioSIS = 0
'                         End If
                    Case sghTrabajaSeguroSOAT
                         If lnPrecioSOAT = 0 Then
                              oRsCatalogo.Filter = "idProducto=" & rs!idProducto & " and IdTipoFinanciamiento=" & ml_IdTipoFinanciamiento
                              If oRsCatalogo.RecordCount > 0 Then
                                 lnPrecioSOAT = oRsCatalogo.Fields!PrecioUnitario
                              End If
                         End If
                         If lnPrecioDespacho = 0 Then
                            lnPrecioDespacho = lnPrecioSOAT
                         End If
                         'debb-25/10/2016
                         If lnPrecioSOAT = 0 And lnTotalPagar = 0 Then
                              lnTotalPagar = lnCantidadSOAT * lnPrecioDespacho
                              If lnCantidadPagar = 0 Then lnCantidadPagar = lnCantidadSOAT
                              lnCantidadSOAT = 0
                              If lnIdOrdenPago = 0 Then lbTieneQueGrabarAntesDeImprimir = True
                         End If
                    Case sghTrabajaSeguroConvenios
                         If lnPrecioCONV = 0 Then
                              oRsCatalogo.Filter = "idProducto=" & rs!idProducto & " and IdTipoFinanciamiento=" & ml_IdTipoFinanciamiento
                              If oRsCatalogo.RecordCount > 0 Then
                                 lnPrecioCONV = oRsCatalogo.Fields!PrecioUnitario
                              End If
                         End If
                         If lnPrecioDespacho = 0 Then
                            lnPrecioDespacho = lnPrecioCONV
                         End If
                         'debb-25/10/2016
                         If lnPrecioCONV = 0 And lnTotalPagar = 0 Then
                              lnTotalPagar = lnCantidadConv * lnPrecioDespacho
                              If lnCantidadPagar = 0 Then lnCantidadPagar = lnCantidadConv
                              lnCantidadConv = 0
                              If lnIdOrdenPago = 0 Then lbTieneQueGrabarAntesDeImprimir = True
                         End If
                    End Select
                    '
                    If lbElMovimientoNoEstaAnulado = True Then
                        '
                        lnReceta = 0
                        oRecetas.Filter = "codigo='" & rs!Codigo & "' and documentoDespacho='" & lcNroDcto & "'"
                        If oRecetas.RecordCount > 0 Then
                           lnReceta = oRecetas!idReceta
                        End If
                        '
                        mrs_FacturacionProductos.AddNew
                        mrs_FacturacionProductos!idProducto = rs!idProducto
                        mrs_FacturacionProductos!Codigo = rs!Codigo
                        mrs_FacturacionProductos!NombreProducto = Trim(rs!nombre) & lcObservacionesCaja
                        mrs_FacturacionProductos!CantidadPagar = rs!Cantidad  'cantidad inicial (no varia)
                        mrs_FacturacionProductos!PrecioUnitario = lnPrecioDespacho    'rs!precio  'precio de venta
                        mrs_FacturacionProductos!TotalPagar = Round(rs!Cantidad * lnPrecioDespacho, 2)    'rs!Total
                        mrs_FacturacionProductos!CantidadSIS = lnCantidadSIS
                        mrs_FacturacionProductos!precioSIS = lnPrecioSIS
                        mrs_FacturacionProductos!ImporteSIS = lnImporteSIS
                        mrs_FacturacionProductos!CantidadSOAT = lnCantidadSOAT
                        mrs_FacturacionProductos!PrecioSOAT = lnPrecioSOAT
                        mrs_FacturacionProductos!ImporteSOAT = lnImporteSOAT
                        mrs_FacturacionProductos!importeEXO = lnImporteEXO
                        mrs_FacturacionProductos!idPuntoCarga = rs!idPuntoCarga
                        mrs_FacturacionProductos!idestadofacturacion = lnIdEstadoFacturacion      ' IIf(lnIdUsuarioAutoriza > 0, rs!IdEstadoFacturacion, lnIdEstadoFacturacion)
                        mrs_FacturacionProductos!Cantidad = lnCantidadPagar 'cantidad a pagar en caja (varia)
                        mrs_FacturacionProductos!TotalPorPagar = lnTotalPagar  '(a pagar en caja)
                        
                        mrs_FacturacionProductos!IdComprobantePago = lnIdComprobantePago
                        mrs_FacturacionProductos!IdOrden = rs!IdOrden
                        mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = lnIdUsuarioAutoriza
                        If lnIdTipoConceptoFarmacia > 0 Then
                           mrs_FacturacionProductos!FechaAutorizaSeguro = ldFechaAutorizaSeguro
                        End If
                        mrs_FacturacionProductos!IdUsuarioAutorizaDevolucion = lnIdUsuarioAutDev
                        mrs_FacturacionProductos!FechaAutorizaDevolucion = IIf(lnCantidadDev = 0, 0, lcFechaAutDev)
                        mrs_FacturacionProductos!IdComprobantePagoDevolucion = lnIdComprobDev
                        mrs_FacturacionProductos!NroComprobante = IIf(lcDocumentoPago = "", "", lcDocumentoPago)  'si ya se PAGO muestra BOLETA sino muestra TICKET
                        mrs_FacturacionProductos!idTipoFinanciamiento = lnIdTipoFinanciamiento
                        mrs_FacturacionProductos!precioCONV = lnPrecioCONV
                        mrs_FacturacionProductos!esConvenio = lcEsConvenio
                        mrs_FacturacionProductos!FechaOrden = rs!fechacreacion
                        mrs_FacturacionProductos!cantidadConv = lnCantidadConv
                        mrs_FacturacionProductos!ImporteConv = lnImporteConv
                        mrs_FacturacionProductos!idTipoConceptoFarmacia = lnIdTipoConceptoFarmacia
                        mrs_FacturacionProductos!IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
                        If Not IsNull(rs!idServicioPaciente) Then
                            mrs_FacturacionProductos!ServicioDeEstancia = IIf(IsNull(rs!idServicioPaciente), ".", mo_ReglasFacturacion.BuscaServicioActualDelPaciente(rs!idServicioPaciente))
                            mrs_FacturacionProductos!idServicioDeEstancia = IIf(IsNull(rs!idServicioPaciente), 0, rs!idServicioPaciente)
                        End If
                        If ml_AgruparPor = 3 Or ml_AgruparPor = 5 Then
                           mrs_FacturacionProductos!descripcion = rs!dfinanciamiento
                        End If
                        mrs_FacturacionProductos!ImporteEnBoleta = lnImporteEnBoleta
                        mrs_FacturacionProductos!nroDcto = lcNroDcto
                        mrs_FacturacionProductos!ComoSeTrabajaEnEstadoCuenta = lnComoSeTrabajaEnEstadoCuenta
                        mrs_FacturacionProductos!IdOrdenPago = lnIdOrdenPago
                        mrs_FacturacionProductos!FechaDespacho = rs!FechaDespacho
                        mrs_FacturacionProductos!Receta = lnReceta
                    End If
                    'Function TotalizaPagoDelPaciente()
                    If (mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Or _
                          mrs_FacturacionProductos.Fields!idestadofacturacion = sghConPreVenta) And _
                          (lnComoSeTrabajaEnEstadoCuenta1 = sghTrabajaParticular Or _
                           mrs_FacturacionProductos.Fields!idTipoFinanciamiento = sghTrabajaServicioSocial) Then
                      lnTotalPagoDelPaciente = lnTotalPagoDelPaciente + mrs_FacturacionProductos.Fields!TotalPorPagar
                    End If
                    'Function TotalizaPagoDeSeguros()
                    Select Case lnComoSeTrabajaEnEstadoCuenta1
                    Case sghTrabajaSeguroSIS
                        lnTotalPagoSeguro = lnTotalPagoSeguro + mrs_FacturacionProductos.Fields!ImporteSIS
                    Case sghTrabajaSeguroSOAT
                        lnTotalPagoSeguro = lnTotalPagoSeguro + mrs_FacturacionProductos.Fields!ImporteSOAT
                    Case sghTrabajaSeguroConvenios
                        lnTotalPagoSeguro = lnTotalPagoSeguro + mrs_FacturacionProductos.Fields!ImporteConv
                    End Select
                    'TotalizaPagosDelPacienteConSeguro()
                    lnTotalizaPagosDelPacienteConSeguro = lnTotalizaPagosDelPacienteConSeguro + mrs_FacturacionProductos.Fields!TotalPorPagar
                    '***Resumen-Servicios
                    If mrs_FacturacionProductos.Fields!idestadofacturacion = 1 Or mrs_FacturacionProductos.Fields!idestadofacturacion = 4 Or mrs_FacturacionProductos.Fields!idestadofacturacion = sghConPreVenta Then
                        Set oRs = mo_ReglasComunes.FactPuntosCargaSeleccionarPorId(mrs_FacturacionProductos.Fields!idPuntoCarga, oConexion)
                        lcTexto = ""
                        If oRs.RecordCount > 0 Then
                           lcTexto = Trim(oRs.Fields!descripcion)
                        End If
                        oRs.Close
                        Select Case lnIdTipoConceptoFarmaciaPlanActual
                        Case sghTipoConceptoFarmacia.sghTipoConceptoSIS
                             lnCant = mrs_FacturacionProductos.Fields!CantidadSIS
                             lnPrec = mrs_FacturacionProductos.Fields!precioSIS
                             lnImpo = mrs_FacturacionProductos.Fields!ImporteSIS
                        Case sghTipoConceptoFarmacia.sghTipoConceptoSOAT
                             lnCant = mrs_FacturacionProductos.Fields!CantidadSOAT
                             lnPrec = mrs_FacturacionProductos.Fields!PrecioSOAT
                             lnImpo = mrs_FacturacionProductos.Fields!ImporteSOAT
                        Case sghTipoConceptoFarmacia.sghTipoConceptoConvenios
                             lnCant = mrs_FacturacionProductos.Fields!cantidadConv
                             lnPrec = mrs_FacturacionProductos.Fields!precioCONV
                             lnImpo = mrs_FacturacionProductos.Fields!ImporteConv
                        Case Else
                             lnCant = mrs_FacturacionProductos.Fields!Cantidad
                             lnPrec = mrs_FacturacionProductos.Fields!PrecioUnitario
                             lnImpo = mrs_FacturacionProductos.Fields!TotalPorPagar
                        End Select
                        lcLlave = lcTexto & " - " & mrs_FacturacionProductos.Fields!FechaOrden & " - " & mrs_FacturacionProductos.Fields!nroDcto
                        lbNuevo = True
                        If oRsCuentaCabecera.RecordCount > 0 Then
                           oRsCuentaCabecera.MoveFirst
                           oRsCuentaCabecera.Find "llave='" & lcLlave & "'"
                           If Not oRsCuentaCabecera.EOF Then
                              lbNuevo = False
                           End If
                        End If
                        If lbNuevo Then
                              oRsCuentaCabecera.AddNew
                              oRsCuentaCabecera.Fields!llave = lcLlave
                              oRsCuentaCabecera.Fields!puntoDeCarga = lcTexto
                              oRsCuentaCabecera.Fields!fecha = mrs_FacturacionProductos.Fields!FechaDespacho
                              oRsCuentaCabecera.Fields!Servicio = mrs_FacturacionProductos.Fields!ServicioDeEstancia
                              oRsCuentaCabecera.Fields!Importe = lnImpo
                              oRsCuentaCabecera.Fields!nrodocumento = mrs_FacturacionProductos.Fields!IdOrden
                        Else
                              oRsCuentaCabecera.Fields!Importe = oRsCuentaCabecera.Fields!Importe + lnImpo
                        End If
                        oRsCuentaCabecera.Update
                        oRsCuentaDetalle.AddNew
                        oRsCuentaDetalle.Fields!llave = lcLlave
                        oRsCuentaDetalle.Fields!Codigo = mrs_FacturacionProductos.Fields!Codigo
                        oRsCuentaDetalle.Fields!descripcion = Left(mrs_FacturacionProductos.Fields!NombreProducto, 50)
                        oRsCuentaDetalle.Fields!Cantidad = lnCant
                        oRsCuentaDetalle.Fields!Precio = lnPrec
                        oRsCuentaDetalle.Fields!Importe = lnImpo
                        If mrs_FacturacionProductos.Fields!idestadofacturacion = 4 Then
                           oRsCuentaDetalle.Fields!nrodocumento = mrs_FacturacionProductos.Fields!NroComprobante
                        End If
                        oRsCuentaDetalle.Update
                        lnTotalApagar = lnTotalApagar + lnImpo
                    End If
                    '
                    rs.MoveNext
                    If rs.EOF Then
                       Exit Do
                    End If
                Loop
            Loop
            '
            Set oRs = Nothing
            oRsFinanciamientos.Close: Set oRsFinanciamientos = Nothing
            oRsPagos.Close: Set oRsPagos = Nothing
            oRsDevoluciones.Close: Set oRsDevoluciones = Nothing
            oRsCatalogo.Close: Set oRsCatalogo = Nothing
            Set oRecetas = Nothing
        End If
    End If
End Sub



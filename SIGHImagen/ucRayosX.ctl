VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucRayosX 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   ScaleHeight     =   5730
   ScaleWidth      =   11685
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   1695
      Left            =   150
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   2990
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BorderStyle     =   9
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
      Height          =   5265
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   9287
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
      Caption         =   "Productos"
   End
   Begin VB.Label Label1 
      Caption         =   "Teclas de ayuda: <F10> = Agregar         <Supr>  = Eliminar "
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
      Left            =   90
      TabIndex        =   3
      Top             =   5400
      Width           =   6165
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
      Height          =   240
      Left            =   11055
      TabIndex        =   2
      Top             =   5400
      Width           =   555
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
Attribute VB_Name = "ucRayosX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim gridInfra As New GridInfragistic
Dim mo_PermisosFacturacion As New PermisosFacturacion

Dim ms_TipoProducto As sghTipoProducto
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idOrden As Long
Dim ml_idCuentaAtencion As Long
Dim mb_CargandoProductos As Boolean
Dim ms_Opcion As sghOpciones
Dim mrs_FacturacionProductos As New Recordset
Dim mo_DOAtencion As DOAtencion
Dim ml_IdUsuario As Long
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

Property Let IdOrdenPago(lValue As Long)
    ml_idOrdenPago = lValue
End Property
Property Get IdOrdenPago() As Long
    IdOrdenPago = ml_idOrdenPago
End Property

Property Let PermiteAgregarItems(lValue As Boolean)
    ml_PermiteAgregarItems = lValue
End Property


Property Let DocumentoYaRegistradoEnSeguros(lValue As Boolean)
    ml_DocumentoYaRegistradoEnSeguros = lValue
End Property

Property Let IdOrden(lValue As Long)
    ml_idOrden = lValue
End Property
Property Get IdOrden() As Long
    IdOrden = ml_idOrden
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

Property Let IdUsuario(lValue As Long)
    ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
    IdUsuario = ml_IdUsuario
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
    GenerarRecordsetProductos
    ms_EstadosFacturacion = ""
    Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_IdUsuario)
    
    
    UserControl.mnuAgregarServicio.Enabled = mo_PermisosFacturacion.AgregarServicios
    UserControl.mnuAgregarExoneracion.Enabled = False 'mo_PermisosFacturacion.AgregarExoneraciones
    UserControl.mnuAutorizarSIS.Enabled = False   'mo_PermisosFacturacion.AutorizarSIS
    UserControl.mnuAutorizarSOAT.Enabled = False  ' mo_PermisosFacturacion.AutorizarSOAT
    UserControl.mnuAutorizarPendientePago.Enabled = False 'mo_PermisosFacturacion.AutorizarPendientesDePago
    UserControl.mnuAutorizarConvenio.Enabled = False 'mo_PermisosFacturacion.AutorizarConvenios
    UserControl.mnuAutorizarDevolucion.Enabled = False 'mo_PermisosFacturacion.AutorizarDevoluciones

    UserControl.grdProductos.DialogStrings(ssDeleteRow) = "Ud. ha seleccionado una fila para borrarla" + Chr(13) + "Seleccione <Si> para borrar la fila o <No> para Cancelar "
    lnMaximoNroItems = BuscarMaximoItemsEnParametros()
    ml_PermiteAgregarItems = True
End Sub

Function BuscarMaximoItemsEnParametros() As Long
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
        Set lcBuscaParametro = Nothing
End Function

Sub AgregaProducto()
        
    If mrs_FacturacionProductos.RecordCount >= lnMaximoNroItems Then
       MsgBox "Solo se permite registrar hasta " & Trim(Str(lnMaximoNroItems)) & " Items", vbExclamation, "Facturación"
       Exit Sub
    End If
    
'    grdProductos.SetFocus
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!IdProducto = 0
        .Fields!codigo = ""
        .Fields!NombreProducto = ""
        .Fields!cantidad = 1
        .Fields!precioUnitario = 0
        .Fields!totalPorPagar = 0
        .Fields!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
        .Fields!IdPuntoCarga = ml_IdPuntoCarga
        If Not mo_DOAtencion Is Nothing Then
            .Fields!idAtencion = mo_DOAtencion.idAtencion
        End If
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        
        Select Case ml_IdTipoFinanciamiento
        Case 2, 3, 4
            .Fields!idEstadoFacturacion = 4
            .Fields!FechaAutorizaSeguro = Now
            .Fields!IdUsuarioAutorizaSeguro = ml_IdUsuario
        Case Else
            .Fields!idEstadoFacturacion = 1
            .Fields!FechaAutorizaSeguro = 0
            .Fields!IdUsuarioAutorizaSeguro = 0
        End Select
        
        .Fields!IdFuenteFinanciamiento = 1
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_IdUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!IdOrden = ml_idOrden
        
    End With
    mb_CargandoProductos = False
    
    Totalizar
    
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode

    
End Sub

Sub AgregaExoneracion()
        
    mb_CargandoProductos = True
    
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!IdProducto = 4692
        .Fields!codigo = "F00002"
        .Fields!NombreProducto = "Exoneracion"
        .Fields!cantidad = 1
        .Fields!precioUnitario = 1
        .Fields!totalPorPagar = 0
        .Fields!IdTipoFinanciamiento = 9
        .Fields!IdPuntoCarga = ml_IdPuntoCarga
        If Not mo_DOAtencion Is Nothing Then
            .Fields!idAtencion = mo_DOAtencion.idAtencion
        Else
            .Fields!idAtencion = 0
        End If
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idEstadoFacturacion = 1
        .Fields!FechaAutorizaSeguro = 0
        .Fields!IdUsuarioAutorizaSeguro = 0
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_IdUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        
    End With
    mb_CargandoProductos = False

    mb_FilaEditable = True
    
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    
    ModificarColorDeFila grdProductos.ActiveRow
    
End Sub

Sub AgregaPagoACuenta()
        
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!IdProducto = 4691
        .Fields!codigo = "F00001"
        .Fields!NombreProducto = "Pago a cuenta"
        .Fields!cantidad = 1
        .Fields!precioUnitario = 1
        .Fields!totalPorPagar = 0
        .Fields!IdTipoFinanciamiento = 1
        .Fields!IdPuntoCarga = ml_IdPuntoCarga
        On Error Resume Next
        .Fields!idAtencion = mo_DOAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idEstadoFacturacion = 1
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_IdUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_IdUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        
    End With
    mb_CargandoProductos = False

    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    
End Sub

Sub AgregaDevolucion()
        
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!IdProducto = 4693
        .Fields!codigo = "F00001"
        .Fields!NombreProducto = "Devolución"
        .Fields!cantidad = 1
        .Fields!precioUnitario = -1
        .Fields!totalPorPagar = 0
        .Fields!IdTipoFinanciamiento = 0
        .Fields!IdPuntoCarga = ml_IdPuntoCarga
        .Fields!idAtencion = mo_DOAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idEstadoFacturacion = 4
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_IdUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_IdUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        
    End With
    mb_CargandoProductos = False
    
    
    
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    
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
    Case sghbien
        Set rs = mo_ReglasFacturacion.FacturacionBienInsumoPorOrdenAtencion(ml_idOrden, ms_EstadosFacturacion, ms_TiposFinanciamiento)
        CargarItemsALaGrillaB rs
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
Sub CargaDespachosPorIdOrden()
Dim rs As Recordset
Dim rs1 As Recordset
    Select Case ms_TipoProducto
    Case sghServicio
           Set rs = mo_ReglasFacturacion.FacturacionServicioDespachoFiltraPorIdOrden(ml_idOrden)
           CargarItemsALaGrillaS rs
    Case sghbien
    End Select
    
End Sub

Sub CargaDevolucionesPorIdComprobante(lIdComprobante As Long)
Dim rs As Recordset
    
    Select Case ms_TipoProducto
    Case sghServicio
        Set rs = mo_ReglasFacturacion.FacturacionServicioDevolucionPorIdComprobante(lIdComprobante)
        CargarItemsALaGrillaS rs
    Case sghbien
        Set rs = mo_ReglasFacturacion.FacturacionBienInsumoDevolucionPorIdComprobante(lIdComprobante)
        CargarItemsALaGrillaB rs
    End Select
    
    
    
End Sub

Sub CargarItemsALaGrillaB(rs As Recordset)
    mb_CargandoProductos = True
    Do While Not rs.EOF
        mrs_FacturacionProductos.AddNew
        mrs_FacturacionProductos!movNumero = rs!movNumero
        mrs_FacturacionProductos!movTipo = rs!movTipo
        mrs_FacturacionProductos!IdProducto = rs!IdProducto
        mrs_FacturacionProductos!codigo = rs!codigo
        mrs_FacturacionProductos!NombreProducto = rs!NombreProducto
      '  mrs_FacturacionProductos!IdTipoFinanciamiento = rs!IdTipoFinanciamiento
       ' mrs_FacturacionProductos!tipoFinanciamiento = rs!IdTipoFinanciamiento
        mrs_FacturacionProductos!cantidad = rs!cantidad
        mrs_FacturacionProductos!precioUnitario = rs!precioUnitario
        mrs_FacturacionProductos!totalPorPagar = rs!totalPorPagar
        mrs_FacturacionProductos!idEstadoFacturacion = rs!idEstadoFacturacion
        mrs_FacturacionProductos!IdPuntoCarga = rs!IdPuntoCarga
        'mrs_FacturacionProductos!idAtencion = rs!idAtencion
        'mrs_FacturacionProductos!FechaAutorizaPendiente = rs!FechaAutorizaPendiente
        'mrs_FacturacionProductos!FechaAutorizaSeguro = rs!FechaAutorizaSeguro
        'mrs_FacturacionProductos!FechaAutorizaDevolucion = rs!FechaAutorizaDevolucion
        'mrs_FacturacionProductos!IdUsuarioAutorizaPendiente = rs!IdUsuarioAutorizaPendiente
        'mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = rs!IdUsuarioAutorizaSeguro
        'mrs_FacturacionProductos!idUsuarioAutorizadevolucion = rs!idUsuarioAutorizadevolucion
        'mrs_FacturacionProductos!IdFuenteFinanciamiento = rs!IdFuenteFinanciamiento
        mrs_FacturacionProductos!IdComprobantePago = rs!IdComprobantePago
        'mrs_FacturacionProductos!IdComprobantePagoDevolucion = rs!IdComprobantePagoDevolucion
        mrs_FacturacionProductos!IdOrden = rs!IdOrden
        mrs_FacturacionProductos!IdCajero = rs!IdCajero
        'mrs_FacturacionProductos!FechaCajero = rs!FechaCajero
        mrs_FacturacionProductos!EstadoLocal = "L" 'Estado Leido de la BD
        '************se va a modificar CANTIDADES de un Documento generado en SEGUROS (INICIO)
'        If ml_DocumentoYaRegistradoEnSeguros = True Then
'           Select Case rs!IdTipoFinanciamiento
'           Case 2    'SIS
'                mrs_FacturacionProductos!Cantidad = rs!CantidadSIS
'                mrs_FacturacionProductos!TotalporPagar = rs!PrecioUnitario * rs!CantidadSIS
'           Case 3    'SOAT
'                mrs_FacturacionProductos!Cantidad = rs!CantidadSOAT
'                mrs_FacturacionProductos!TotalporPagar = rs!PrecioUnitario * rs!CantidadSOAT
'           End Select
'           mrs_FacturacionProductos!EstadoLocal = "M"
'        End If
        '************se va a modificar CANTIDADES de un Documento generado en SEGUROS (FIN)
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    
    Totalizar
    
    Set grdProductos.DataSource = mrs_FacturacionProductos

End Sub


Sub CargarItemsALaGrillaS(rs As Recordset)
    mb_CargandoProductos = True
    Do While Not rs.EOF
        mrs_FacturacionProductos.AddNew
        'mrs_FacturacionProductos!IdFacturacionProducto = rs!IdFacturacionProducto
        mrs_FacturacionProductos!IdProducto = rs!IdProducto
        mrs_FacturacionProductos!codigo = rs!codigo
        mrs_FacturacionProductos!NombreProducto = rs!Nombre
        mrs_FacturacionProductos!IdTipoFinanciamiento = ml_IdTipoFinanciamiento
        'mrs_FacturacionProductos!tipoFinanciamiento = rs!IdTipoFinanciamiento
        mrs_FacturacionProductos!cantidad = rs!cantidad
        mrs_FacturacionProductos!precioUnitario = rs!precio
        mrs_FacturacionProductos!totalPorPagar = rs!Total
        'mrs_FacturacionProductos!IdEstadoFacturacion = rs!IdEstadoFacturacion
        'mrs_FacturacionProductos!IdPuntoCarga = rs!IdPuntoCarga
        'mrs_FacturacionProductos!idAtencion = rs!idAtencion
        'mrs_FacturacionProductos!FechaAutorizaPendiente = rs!FechaAutorizaPendiente
        'mrs_FacturacionProductos!FechaAutorizaSeguro = rs!FechaAutorizaSeguro
        'mrs_FacturacionProductos!FechaAutorizaDevolucion = rs!FechaAutorizaDevolucion
        'mrs_FacturacionProductos!IdUsuarioAutorizaPendiente = rs!IdUsuarioAutorizaPendiente
        'mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = rs!IdUsuarioAutorizaSeguro
        'mrs_FacturacionProductos!idUsuarioAutorizadevolucion = rs!idUsuarioAutorizadevolucion
        'mrs_FacturacionProductos!IdFuenteFinanciamiento = rs!IdFuenteFinanciamiento
        'mrs_FacturacionProductos!IdComprobantePago = rs!IdComprobantePago
        'mrs_FacturacionProductos!IdComprobantePagoDevolucion = rs!IdComprobantePagoDevolucion
        'mrs_FacturacionProductos!IdOrden = rs!IdOrden
        'mrs_FacturacionProductos!IdCajero = rs!IdCajero
        'mrs_FacturacionProductos!FechaCajero = rs!FechaCajero
'        Select Case ms_TipoProducto
'        Case sghServicio
'            mrs_FacturacionProductos!IdServicioInternamiento = rs!IdServicioInternamiento
'        Case sghbien
'        End Select
'        mrs_FacturacionProductos!EstadoLocal = "L" 'Estado Leido de la BD
'        '************se va a modificar CANTIDADES de un Documento generado en SEGUROS (INICIO)
'        If ml_DocumentoYaRegistradoEnSeguros = True Then
'           Select Case rs!IdTipoFinanciamiento
'           Case 2    'SIS
'                mrs_FacturacionProductos!Cantidad = rs!CantidadSIS
'                mrs_FacturacionProductos!totalPorPagar = rs!PrecioUnitario * rs!CantidadSIS
'           Case 3    'SOAT
'                mrs_FacturacionProductos!Cantidad = rs!CantidadSOAT
'                mrs_FacturacionProductos!totalPorPagar = rs!PrecioUnitario * rs!CantidadSOAT
'           End Select
'           mrs_FacturacionProductos!EstadoLocal = "M"
'        End If
        '************se va a modificar CANTIDADES de un Documento generado en SEGUROS (FIN)
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    
    Totalizar
    
    Set grdProductos.DataSource = mrs_FacturacionProductos

End Sub


Sub HabilitarMenuSegunEstadoOrden(IdEstadoOrden As Long)

    Select Case IdEstadoOrden
    Case 1
        HabilitarMenus True
    Case 4
        HabilitarMenus False
        UserControl.mnuAutorizarDevolucion.Enabled = True   'Esto debe estar habilitado para poder autorizar devoluciones
    Case 9
        HabilitarMenus False
    End Select

End Sub

Sub HabilitarMenus(estado As Boolean)

        UserControl.mnuAgregarPagoACuenta.Enabled = estado
        UserControl.mnuAgregarExoneracion.Enabled = estado
        UserControl.mnuAgregarServicio.Enabled = estado
        UserControl.mnuAutorizaPacienteNormal.Enabled = estado
        UserControl.mnuAutorizarConvenio.Enabled = estado
        UserControl.mnuAutorizarDevolucion.Enabled = estado
        UserControl.mnuAutorizarPendientePago.Enabled = estado
        UserControl.mnuAutorizarSIS.Enabled = estado
        UserControl.mnuAutorizarSOAT.Enabled = estado

End Sub

'Sub CargaProductosPorIdCuentaAtencion()
'Dim rs As Recordset
'
'    Select Case ms_TipoProducto
'    Case sghServicio
'        Set rs = mo_ReglasFacturacion.FacturacionServicioPorCuentaAtencion(ml_IdCuentaAtencion, ms_EstadosFacturacion, ms_TiposFinanciamiento, 0)
'    Case sghBien
'        Set rs = mo_ReglasFacturacion.FacturacionBienInsumoPorCuentaAtencion(ml_IdCuentaAtencion, ms_EstadosFacturacion, ms_TiposFinanciamiento, 0)
'    End Select
'
'    mb_CargandoProductos = True
'    Do While Not rs.EOF
'        mrs_FacturacionProductos.AddNew
'        mrs_FacturacionProductos!IdFacturacionProducto = rs!IdFacturacionProducto
'        mrs_FacturacionProductos!IdProducto = rs!IdProducto
'        mrs_FacturacionProductos!Codigo = rs!Codigo
'        mrs_FacturacionProductos!NombreProducto = rs!NombreProducto
'        mrs_FacturacionProductos!IdTipoFinanciamiento = rs!IdTipoFinanciamiento
'        mrs_FacturacionProductos!tipoFinanciamiento = rs!IdTipoFinanciamiento
'        mrs_FacturacionProductos!Cantidad = rs!Cantidad
'        mrs_FacturacionProductos!PrecioUnitario = rs!PrecioUnitario
'        mrs_FacturacionProductos!TotalPorPagar = rs!TotalPorPagar
'        mrs_FacturacionProductos!IdEstadoFacturacion = rs!IdEstadoFacturacion
'        mrs_FacturacionProductos!IdPuntoCarga = rs!IdPuntoCarga
'        mrs_FacturacionProductos!IdAtencion = rs!IdAtencion
'        mrs_FacturacionProductos!FechaAutorizaPendiente = rs!FechaAutorizaPendiente
'        mrs_FacturacionProductos!FechaAutorizaSeguro = rs!FechaAutorizaSeguro
'        mrs_FacturacionProductos!FechaAutorizaDevolucion = rs!FechaAutorizaSeguro
'        mrs_FacturacionProductos!IdUsuarioAutorizaPendiente = rs!IdUsuarioAutorizaPendiente
'        mrs_FacturacionProductos!IdUsuarioAutorizaSeguro = rs!IdUsuarioAutorizaSeguro
'        mrs_FacturacionProductos!IdUsuarioAutorizaDevolucion = rs!IdUsuarioAutorizaDevolucion
'        mrs_FacturacionProductos!IdFuenteFinanciamiento = rs!IdFuenteFinanciamiento
'        mrs_FacturacionProductos!IdComprobantePago = rs!ComprobantePago
'        mrs_FacturacionProductos!IdComprobantePagoDevolucion = rs!IdComprobantePagoDevolucion
'
'        Select Case ms_TipoProducto
'        Case sghServicio
'            mrs_FacturacionProductos!IdServicioInternamiento = rs!IdServicioInternamiento
'        Case sghBien
'        End Select
'
'        mrs_FacturacionProductos!EstadoLocal = "L" 'Estado Leido de la BD
'        rs.MoveNext
'    Loop
'    mb_CargandoProductos = False
'
'    Totalizar
'
'    Set grdProductos.DataSource = mrs_FacturacionProductos
'
'    Set grdProductos.DataSource = rs
'
'
'End Sub
'

Function DevuelveTotalPagar() As Double
Dim rsProductos As New Recordset
Dim dTotalPagado As Double
    Set rsProductos = mrs_FacturacionProductos.Clone
    dTotalPagado = 0
    If rsProductos.RecordCount > 0 Then
        rsProductos.MoveFirst
        Do While Not rsProductos.EOF
           dTotalPagado = dTotalPagado + rsProductos!totalPorPagar
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
    
    Set rsProductos = mrs_FacturacionProductos.Clone
    
    If rsProductos.RecordCount = 0 Then
        Exit Sub
    End If
    
    If Not (rsProductos.EOF And rsProductos.BOF) Then
        rsProductos.MoveFirst
        Do While Not rsProductos.EOF
        
            dSubTotal = rsProductos!totalPorPagar
            lIdEstadoFacturacion = rsProductos!idEstadoFacturacion
            lIdProducto = rsProductos!IdProducto
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

'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
    If ml_PermiteAgregarItems = True Then
        If mb_CargandoProductos Then
            Exit Sub
        End If
    End If
End Sub


Private Sub grdProductos_AfterRowsDeleted()
    If ml_PermiteAgregarItems = True Then
        If ml_ultimoProductoEliminado > 0 Then
            mo_ProductosEliminados.Add ml_ultimoProductoEliminado
            ml_ultimoProductoEliminado = 0
            Totalizar
        Else
            Totalizar
            Set grdProductos.DataSource = mrs_FacturacionProductos
        End If
    End If
End Sub

Private Sub grdProductos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
      If ml_PermiteAgregarItems = True Then
        If mb_FilaEditable Then
            'Si la fila es editable y estamos en la celda de codigo se completa los datos
            'del producto
            Select Case grdProductos.ActiveCell.Column.Key
            Case "Codigo"
'                oRow.Cells("Codigo").Value = Right("000000" & Trim(oRow.Cells("Codigo").Value), 6)
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
    If ml_PermiteAgregarItems = True Then
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
    End If
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdProductos
End Sub


Private Sub grdProductos_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    On Error Resume Next
    If ml_PermiteAgregarItems = True Then
       ModificarColorDeFila Row
    End If
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
    OnKeyDown grdProductos, KeyCode
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    OnKeyPress grdProductos, KeyAscii
End Sub

Private Sub grdProductos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ml_PermiteAgregarItems = True Then
        If Button = 2 Then
            PopupMenu mnuProductos
        End If
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
    
    If IsNull(oRow.Cells("codigo").Value) Or IsNull(oRow.Cells("idtipofinanciamiento").Value) Or oRow.Cells("codigo").Value = "" Then
        Exit Sub
    End If
    If ms_TipoProducto = sghbien Then
       oRow.Cells("codigo").Value = Right("0000000000" & oRow.Cells("codigo").Value, 5)
       
    End If
    Select Case ms_TipoProducto
    Case sghServicio
        Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigodebb(oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, ml_IdPuntoCarga)
    Case sghbien
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, ml_IdPuntoCarga)
    Case Else
        Exit Sub
    End Select
    
    If rs.RecordCount > 0 Then
       If rs.Fields("idproducto").Value <> 4691 Then
            'Busca si ya existe el producto
            If Not ItemYaExiste(rs.Fields("idproducto").Value) Then
                oRow.Cells("IdFacturacionProducto").Value = 0
                oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
                oRow.Cells("NombreProducto").Value = rs.Fields("NombreProducto").Value
                oRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
                oRow.Cells("TotalPorPagar").Value = rs.Fields("preciounitario").Value
                oRow.Cells("cantidad").Value = 1
            End If
       End If
    End If

End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mrs_FacturacionProductos.Clone
        ItemYaExiste = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           oRsTmp.Find "idProducto=" & lnIdProducto
           If Not oRsTmp.EOF Then
              ItemYaExiste = True
              MsgBox "Este producto ya está registrado", vbInformation, "Facturación"
           End If
        End If
        oRsTmp.Close
End Function

Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)
     If ml_PermiteAgregarItems = True Then
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
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grillaBusqueda.SetFocus
                Case vbKeyLeft
                End Select
            End Select
        End If
        Select Case KeyCode
        Case vbKeyF10
            mnuAgregarServicio_Click
        End Select
     End If
End Sub

Sub OnKeyPress(oGrilla As SSUltraGrid, KeyAscii As UltraGrid.SSReturnShort)
    If ml_PermiteAgregarItems = True Then
            'Si la fila no es editable, cancela cualquier cambio en la fila
            If Not mb_FilaEditable Then
                Exit Sub
            End If
            
            If oGrilla.ActiveCell Is Nothing Then
                Exit Sub
            End If
    
            If oGrilla.ActiveCell.Column.Key = "Codigo" And KeyAscii = 13 Then
                SendKeys "{Tab}"
                If Trim(oGrilla.ActiveCell.GetText) <> "" Then
                SendKeys "{Tab}"
                SendKeys "{Tab}"
                End If
                Exit Sub
            End If
            If oGrilla.ActiveCell.Column.Key = "Cantidad" Then
                If KeyAscii = 13 Then
                   'SendKeys "{Tab}"
                   mnuAgregarServicio_Click
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
                        Set rs = mo_AdminCaja.ServiciosFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga)
                    Case sghbien
                        Set rs = mo_AdminCaja.BienesFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga)
                    Case Else
                        
                    End Select
                    Set grillaBusqueda.DataSource = rs
                    grillaBusqueda.Left = oGrilla.Left
                    If mrs_FacturacionProductos.RecordCount < 7 Then
                       grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.Rect.Bottom * Screen.TwipsPerPixelY
                    Else
                       grillaBusqueda.Top = 0
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
          
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    'Set grdProductos.DataSource = mrs_FacturacionProductos
    
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    
    oGrilla.Bands(0).Columns("IdFacturacionProducto").Hidden = True
    
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
        
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
    oGrilla.Bands(0).Columns("Codigo").Width = 750
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationAllowEdit
    
    oGrilla.Bands(0).Columns("NombreProducto").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("NombreProducto").Width = 9000
    oGrilla.Bands(0).Columns("NombreProducto").Activation = ssActivationAllowEdit
    
'    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Width = 2  '2500
'    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Header.Caption = "Tipo Financiamiento"
'    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Style = ssStyleDropDownList
    'oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Hidden = True
    
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

    oGrilla.Bands(0).Columns("idPuntoCarga").Header.Caption = "Puntos de carga"
    oGrilla.Bands(0).Columns("idPuntoCarga").Width = 1500
    oGrilla.Bands(0).Columns("idPuntoCarga").Style = ssStyleDropDownList

    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Width = 2500
    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Aut. Pend."
    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Format = "dd/MM/yyyy hh:mm"

    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Width = 2500
    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Header.Caption = "Fec. Aut. Seguro."
    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Format = "dd/MM/yyyy hh:mm"
    
    'Configura Values List
    SeteaListaEstado oGrilla, oGrilla.Bands(0).Columns("idEstadoFacturacion")
    SeteaListaTipoFinanciamiento oGrilla, oGrilla.Bands(0).Columns("IdTipoFinanciamiento")
    SeteaPuntosDeCarga oGrilla, oGrilla.Bands(0).Columns("idPuntoCarga")

    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("idPuntoCarga").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("idEstadoFacturacion").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Activation = ssActivationActivateNoEdit
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, sighcomun.GrillaConFilasBicolor
    
End Sub

Sub SeteaListaTipoFinanciamiento(oGrilla As SSUltraGrid, oColumn As SSColumn)
Dim rs As New ADODB.Recordset
Dim I As Integer
Dim oValueTF As SSValueList
    
    If Not oGrilla.ValueLists.Exists("listaTipoFinanciamiento") Then
        Set oValueTF = oGrilla.ValueLists.Add("listaTipoFinanciamiento")
        Set rs = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarTodos
        Do While Not rs.EOF
            If rs!IdTipoFinanciamiento <> 0 Then
                oValueTF.ValueListItems.Add Val(rs!IdTipoFinanciamiento), Trim(rs!Descripcion)
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
Dim I As Integer
Dim oValuePC As SSValueList
    
    If Not oGrilla.ValueLists.Exists("listaPuntosCarga") Then
        Set oValuePC = oGrilla.ValueLists.Add("listaPuntosCarga")
        Set rs = mo_ReglasComunes.SeleccionarPuntosDeCarga()
        Do While Not rs.EOF
            If rs!IdPuntoCarga <> 0 Then
                oValuePC.ValueListItems.Add Val(rs!IdPuntoCarga), Trim(rs!Descripcion)
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
Dim I As Integer
Dim oValueEstado As SSValueList
    
    If Not oGrilla.ValueLists.Exists("listaEstadoFacturacion") Then
        Set oValueEstado = oGrilla.ValueLists.Add("listaEstadoFacturacion")
        Set rs = mo_ReglasFacturacion.EstadosFacturacionObtenerTodos
        Do While Not rs.EOF
            oValueEstado.ValueListItems.Add Val(rs!idEstadoFacturacion), Trim(rs!Descripcion)
            rs.MoveNext
        Loop
        rs.Close
    Else
        Set oValueEstado = oGrilla.ValueLists.Item("listaEstadoFacturacion")
    End If
     
    Set oColumn.ValueList = oValueEstado
    
End Sub

Private Sub grillaBusqueda_Click()
'    RefrescarDatos
End Sub

Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, sighcomun.GrillaConFilasBicolor
End Sub
Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
   
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, sighcomun.GrillaConFilasBicolor
End Sub
Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
    
    If ItemYaExiste(grillaBusqueda.ActiveRow.Cells("idproducto").Value) Then
        grdProductos.ActiveRow.Cells("codigo").Value = ""
        grdProductos.ActiveRow.Cells("idproducto").Value = 0
        grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
    Else
        RefrescarDatos
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
        SendKeys "{Tab}"
       ' SendKeys "{Tab}"
    End If
End Sub
Sub RefrescarDatos()
Dim fila As New Record
Dim lnPrecioUnitario  As Double
    If Not grillaBusqueda.ActiveRow Is Nothing Then
        
            If ms_TipoProducto = sghbien Then
               grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
               grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
               grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
               grdProductos.ActiveRow.Cells("preciounitario").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
               grdProductos.ActiveRow.Cells("TotalPorPagar").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
               grdProductos.ActiveRow.Cells("cantidad").Value = 1
               grdProductos.ActiveRow.Cells("idestadofacturacion").Value = 1
                
            Else
               lnPrecioUnitario = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
               'If ml_IdTipoFinanciamiento <> 5 And ml_IdTipoFinanciamiento <> 1 Then
                    lnPrecioUnitario = 0
                    oDoCatalogoServicioHosp.precioUnitario = 0
                    Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(grillaBusqueda.ActiveRow.Cells("idproducto").Value, ml_IdTipoFinanciamiento)
                    If oDoCatalogoServicioHosp.precioUnitario > 0 Then
                        lnPrecioUnitario = oDoCatalogoServicioHosp.precioUnitario
                    Else
                        MsgBox "Ese Producto no tiene precio para el TIPO DE  FINANCIAMIENTO", vbExclamation, "Facturación"
                    End If
              ' End If
               If lnPrecioUnitario > 0 Then
                    grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
                    grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
                    grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
                    grdProductos.ActiveRow.Cells("preciounitario").Value = lnPrecioUnitario
                    grdProductos.ActiveRow.Cells("TotalPorPagar").Value = lnPrecioUnitario
                    grdProductos.ActiveRow.Cells("cantidad").Value = 1
                    grdProductos.ActiveRow.Cells("idestadofacturacion").Value = 1
               End If
            End If
            Totalizar
                
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

Private Sub mnuAgregarExoneracion_Click()
    AgregaExoneracion
End Sub

Private Sub mnuAgregarPagoACuenta_Click()
    AgregaPagoACuenta
End Sub

Private Sub mnuAgregarServicio_Click()
    SendKeys "{Tab}"
    AgregaProducto
End Sub

Private Sub mnuAutorizaPacienteNormal_Click()

    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If

    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 1   'Paciente Normal
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 1   'Ingresado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    
End Sub

Private Sub mnuAutorizarConvenio_Click()

    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 4   'Convenio
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaSeguro").Value = ml_IdUsuario
    grdProductos.ActiveRow.Cells("FechaAutorizaSeguro").Value = Now

End Sub

Private Sub mnuAutorizarDevolucion_Click()
    
    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    If grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 6 Then
        MsgBox "Este producto no se puede devolver, ya ha sido devuelto", vbInformation, "Facturación"
        Exit Sub
    End If
    
    
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 5   'Devolver
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaDevolucion").Value = ml_IdUsuario
    grdProductos.ActiveRow.Cells("FechaAutorizaDevolucion").Value = Now
    
    
End Sub

Private Sub mnuAutorizarPendientePago_Click()
    
    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    Select Case grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value
    Case 1
        grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 3   'Pagado
        grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
        grdProductos.ActiveRow.Cells("IdUsuarioAutorizaPendiente").Value = ml_IdUsuario
        grdProductos.ActiveRow.Cells("FechaAutorizaPendiente").Value = Now
    Case 2, 3, 4
        MsgBox "La autorización de pendientes de pago no aplica a seguros y convenios ", vbInformation, "Facturacion de servicios"
    End Select

End Sub

Private Sub mnuAutorizarSIS_Click()

    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 2   'SIS
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaSeguro").Value = ml_IdUsuario
    grdProductos.ActiveRow.Cells("FechaAutorizaSeguro").Value = Now
    
        
End Sub

Private Sub mnuAutorizarSOAT_Click()

    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 3   'SIS
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaSeguro").Value = ml_IdUsuario
    grdProductos.ActiveRow.Cells("FechaAutorizaSeguro").Value = Now
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 5
   
   Label1.Top = UserControl.Height - UserControl.Label1.Height
   lblTotal.Top = UserControl.Height - UserControl.Label1.Height + 60
   
End Sub

Sub LimpiarGrilla()

    
        If mrs_FacturacionProductos Is Nothing Then
            Exit Sub
        End If

        Set grdProductos.DataSource = Nothing

        If mrs_FacturacionProductos.RecordCount > 0 Then
            mrs_FacturacionProductos.MoveFirst
            Do While Not mrs_FacturacionProductos.EOF
                mrs_FacturacionProductos.Delete
                mrs_FacturacionProductos.Update
                mrs_FacturacionProductos.MoveNext
            Loop
        End If

        ml_idOrden = -1000  'Esto es aproposito para que obtenga solo la estructura del recordset
        CargaProductosPorIdOrden

End Sub


'***************daniel barrantes**************
'***************Registra la CANTIDAD a DEVOLVER en cada Item
'***************ya autorizada anteriormente
Sub ActualizaDevolucionAutorizada(oRs As Recordset)
    If oRs.RecordCount > 0 Then
       oRs.MoveFirst
       Do While Not oRs.EOF
          mrs_FacturacionProductos.MoveFirst
          mrs_FacturacionProductos.Find "idProducto=" & oRs.Fields!IdProducto
          If IsNull(oRs.Fields!cantidadDev) Or oRs.Fields!cantidadDev = 0 Then
             mrs_FacturacionProductos.Delete
          Else
          mrs_FacturacionProductos.Fields!cantidad = oRs.Fields!cantidadDev
          mrs_FacturacionProductos.Fields!IdTipoFinanciamiento = 1
          mrs_FacturacionProductos.Fields!totalPorPagar = oRs.Fields!cantidadDev * mrs_FacturacionProductos.Fields!precioUnitario
          mrs_FacturacionProductos.Fields!EstadoLocal = "A"
          End If
          mrs_FacturacionProductos.Update
          oRs.MoveNext
       Loop
       Totalizar
    End If
End Sub


Function OrdenRegistradaYaprobadaPorSisSoat() As Long   'Devuelve 2=sis, 3=soat, 0=ninguno
    OrdenRegistradaYaprobadaPorSisSoat = 0
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
          If mrs_FacturacionProductos.Fields!IdTipoFinanciamiento = 2 Or mrs_FacturacionProductos.Fields!IdTipoFinanciamiento = 3 Then
             OrdenRegistradaYaprobadaPorSisSoat = mrs_FacturacionProductos.Fields!IdTipoFinanciamiento
             Exit Function
          End If
          mrs_FacturacionProductos.MoveNext
       Loop
    End If
End Function

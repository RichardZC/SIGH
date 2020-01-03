VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucFacturacionIteIMG 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   ScaleHeight     =   5730
   ScaleWidth      =   11685
   Begin VB.Frame FraFiltroBusqueda 
      Height          =   1785
      Left            =   10290
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
      Begin VB.CommandButton cmdFiltroAdorno 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   90
         Width           =   1185
      End
      Begin VB.CommandButton cmdFiltraBusqueda 
         Caption         =   "Filtrar"
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
         Left            =   60
         TabIndex        =   8
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtFiltroBusqueda 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   630
         Width           =   1095
      End
   End
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
      Left            =   -30
      TabIndex        =   1
      Top             =   30
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
   Begin Threed.SSOption optPorCodigo 
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   5370
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reg.Por Código"
   End
   Begin Threed.SSOption optPorDescripcion 
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   5370
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reg. Por Descripción"
      Value           =   -1
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
      Width           =   5145
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
      Begin VB.Menu mnuCtaBancarias 
         Caption         =   "Agregar  Depósito de Garantía"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAgregarPagoACuenta 
         Caption         =   "Agregar  Pago a Cuenta"
      End
      Begin VB.Menu mnuAgregarServicio 
         Caption         =   "Agregar  Servicio"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAgregarExoneracion 
         Caption         =   "Agregar  Exoneración"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutorizaPacienteNormal 
         Caption         =   "Paciente Normal"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutorizarSIS 
         Caption         =   "Autorizado por SIS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutorizarSOAT 
         Caption         =   "Autorizado por SOAT"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutorizarConvenio 
         Caption         =   "Autorizado por Convenio"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutorizarPendientePago 
         Caption         =   "Autorizar pendiente de pago"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutorizarDevolucion 
         Caption         =   "Autorizar devolución"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddDevolucion 
         Caption         =   "Agregar Devolución"
      End
      Begin VB.Menu mnuOtrosAdm 
         Caption         =   "Otros ingresos Administrativos"
      End
      Begin VB.Menu mnuIngClinica 
         Caption         =   "Otros ingresos CLINICA"
      End
   End
End
Attribute VB_Name = "ucFacturacionIteIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para elegir Procedimientos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Public Event Totalizado(TotalIngresado As Double, TotalPendientePago As Double, TotalPagoACuenta As Double, TotalExonerado As Double, dTotalPagado As Double, dTotalPorDevolver As Double, dTotalDevuelto As Double, dTotalAnulado As Double)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
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
'mgaray201410f
Dim mrs_FacturacionProductosSinDatos As New Recordset
Dim mo_DoAtencion As DOAtencion
Dim ml_idUsuario As Long
Dim ml_IdPuntoCarga As Long
Dim ms_EstadosFacturacion As String
Dim ms_TiposFinanciamiento As String
Dim ml_IdEstadoOrden As Long

'Edicion de la grilla
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_DocumentoYaRegistradoEnSeguros As Boolean
Dim ml_PermiteAgregarItems As Boolean
Dim ml_idOrdenPago As Long
Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
Dim ml_IdPuntoCargaServicioHosp As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnIdPagosACuenta As Long
Dim lnIdDepositoGarantia As Long
Dim lnIdDevolucion As Long
Dim lnIdOtrosAdm As Long
Dim lnIdOtrosClinica As Long   'debb2014-d
Dim ml_FiltraCpt As sghFiltraCpt     'debb-24/03/2011
'mgaray201411a
Dim mb_MostrarColumnaLab As Boolean

'debb-24/03/2011
Property Let FiltraCpt(lValue As sghFiltraCpt)
    ml_FiltraCpt = lValue
End Property

Property Let MaximoNroItems(lValue As Long)
    lnMaximoNroItems = lValue
End Property

Property Let IdPuntoCargaServicioHosp(lValue As Long)
    ml_IdPuntoCargaServicioHosp = lValue
End Property

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
'mgaray201411a
Property Let MostrarColumnaLab(oValue As Boolean)
    mb_MostrarColumnaLab = oValue
End Property

Property Get MostrarColumnaLab() As Boolean
    MostrarColumnaLab = mb_MostrarColumnaLab
End Property
Private Sub Class_Initialize()
    mb_MostrarColumnaLab = False
End Sub


Sub Inicializar()
    ml_DocumentoYaRegistradoEnSeguros = False
    
    Set mrs_FacturacionProductos = New Recordset
    GenerarRecordsetProductos
    ms_EstadosFacturacion = ""
    Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
    
    
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
    '
    lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
    lnIdDepositoGarantia = Val(lcBuscaParametro.SeleccionaFilaParametro(254))
    lnIdDevolucion = Val(lcBuscaParametro.SeleccionaFilaParametro(265))
    lnIdOtrosAdm = Val(lcBuscaParametro.SeleccionaFilaParametro(266))
    lnIdOtrosClinica = Val(lcBuscaParametro.SeleccionaFilaParametro(8))    'debb2014-d
    '
    ml_FiltraCpt = sghMuestraTodosCpt
End Sub

'debb-18/05/2016
Function BuscarMaximoItemsEnParametros() As Long
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
        If UCase(lcBuscaParametro.SeleccionaFilaParametro(500)) = "S" Then
           BuscarMaximoItemsEnParametros = 500
        End If
        Set lcBuscaParametro = Nothing
End Function

Sub AgregaProducto()
    On Error GoTo ErrAProd
    If lnMaximoNroItems > 0 Then
        If mrs_FacturacionProductos.RecordCount >= lnMaximoNroItems Then
           MsgBox "Solo se permite registrar hasta " & Trim(Str(lnMaximoNroItems)) & " Items", vbExclamation, "Facturación"
           Exit Sub
        End If
    End If
    
'    grdProductos.SetFocus
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
        .Fields!idTipoFinanciamiento = ml_IdTipoFinanciamiento
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        If Not mo_DoAtencion Is Nothing Then
            .Fields!idAtencion = mo_DoAtencion.idAtencion
        End If
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        
        Select Case ml_IdTipoFinanciamiento
        Case 2, 3, 4
            .Fields!idestadofacturacion = 4
            .Fields!FechaAutorizaSeguro = Now
            .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        Case Else
            .Fields!idestadofacturacion = 1
            .Fields!FechaAutorizaSeguro = 0
            .Fields!IdUsuarioAutorizaSeguro = 0
        End Select
        
        .Fields!IdFuenteFinanciamiento = 1
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!IdOrden = ml_idOrden
        
    End With
    mb_CargandoProductos = False
    
    Totalizar
    
    mb_FilaEditable = True
    'grdProductos.PerformAction ssKeyActionActivateCell
    'grdProductos.PerformAction ssKeyActionEnterEditMode
ErrAProd:
    

End Sub

Sub AgregaExoneracion()
        
    mb_CargandoProductos = True
    
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!idProducto = 4692
        .Fields!Codigo = "F00002"
        .Fields!NombreProducto = "Exoneracion"
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = 1
        .Fields!TotalPorPagar = 0
        .Fields!idTipoFinanciamiento = 9
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        If Not mo_DoAtencion Is Nothing Then
            .Fields!idAtencion = mo_DoAtencion.idAtencion
        Else
            .Fields!idAtencion = 0
        End If
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 1
        .Fields!FechaAutorizaSeguro = 0
        .Fields!IdUsuarioAutorizaSeguro = 0
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
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
    Dim lcCodigoPrd As String, lcNombrePrd As String
    DevuelveCodigoYdescripcionSegunId lnIdPagosACuenta, lcCodigoPrd, lcNombrePrd
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!idProducto = lnIdPagosACuenta
        .Fields!Codigo = lcCodigoPrd
        .Fields!NombreProducto = lcNombrePrd
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = 1
        .Fields!TotalPorPagar = 0
        .Fields!idTipoFinanciamiento = 1
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        On Error Resume Next
        .Fields!idAtencion = mo_DoAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 1
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!PermiteEditarPrecio = True
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
        .Fields!idProducto = 4693
        .Fields!Codigo = "F00001"
        .Fields!NombreProducto = "Devolución"
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = -1
        .Fields!TotalPorPagar = 0
        .Fields!idTipoFinanciamiento = 0
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        .Fields!idAtencion = mo_DoAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 4
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        
    End With
    mb_CargandoProductos = False
    
    
    
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    
End Sub

Sub CargaProductosPorIdCitaSI(lnIdCitaSI As Long)
    Dim rs As Recordset
    Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
    Set rs = mo_ReglasImagenes.SiCitasDetallePorIdentificador(lnIdCitaSI)
    CargarItemsALaGrillaS rs
End Sub

Sub CargaProductosPorIdOrden()
Dim rs As Recordset
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Select Case ms_TipoProducto
    Case sghServicio
        'If ml_IdTipoFinanciamiento = 5 Or ml_IdTipoFinanciamiento = 1 Then
        If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(ml_IdTipoFinanciamiento, oConexion) Then
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

'debb-18/05/2016
Sub CargaProductosPorIdReceta(rs As Recordset)
    Dim rsCatHos As New Recordset, lcMensaje As String, lbContinar As Boolean
    Dim oRsTmp879 As New Recordset
    mb_CargandoProductos = True
    lcMensaje = ""
    Set oRsTmp879 = mo_reglasComunes.RecetaDetalleItemPorIdReceta(rs!idReceta)
    Do While Not rs.EOF
        If rs!Precio > 0 Then
            lbContinar = True
            If oRsTmp879.RecordCount > 0 Then
               oRsTmp879.MoveFirst
               oRsTmp879.Find "idItem=" & rs!idItem
               If Not oRsTmp879.EOF Then
                  lbContinar = False
               End If
            End If
            If lbContinar = True Then
                Set rsCatHos = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(rs!idItem, ml_IdTipoFinanciamiento)
                mrs_FacturacionProductos.AddNew
                mrs_FacturacionProductos!idProducto = rs!idItem
                mrs_FacturacionProductos!Codigo = rs!Codigo
                mrs_FacturacionProductos!NombreProducto = rs!nombre
                mrs_FacturacionProductos!idTipoFinanciamiento = ml_IdTipoFinanciamiento
                mrs_FacturacionProductos!Cantidad = rs!CantidadPedida
                mrs_FacturacionProductos!PrecioUnitario = rs!Precio
                mrs_FacturacionProductos!TotalPorPagar = rs!Total
                If rsCatHos.RecordCount > 0 Then
                   mrs_FacturacionProductos!SeUsaSinPrecio = IIf(IsNull(rsCatHos!SeUsaSinPrecio), False, rsCatHos!SeUsaSinPrecio)
                Else
                   mrs_FacturacionProductos!SeUsaSinPrecio = 0
                End If
                mrs_FacturacionProductos!IdFacturacionProducto = 1
                mrs_FacturacionProductos!poliza = rs!Observaciones
            End If
        Else
            lcMensaje = lcMensaje & "(No tiene precio) " & Trim(rs!Codigo) & "-" & rs!nombre & Chr(13)
        End If
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    Set rsCatHos = Nothing
    Set oRsTmp879 = Nothing
    Totalizar
    Set grdProductos.DataSource = mrs_FacturacionProductos
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
    End If
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, "Caja"
    End If
End Sub

Sub CargaDespachosPorIdOrden()
Dim rs As Recordset
Dim rs1 As Recordset
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Select Case ms_TipoProducto
    Case sghServicio
           Set rs = mo_ReglasFacturacion.FacturacionServicioDespachoFiltraPorIdOrden(ml_idOrden, oConexion)
           CargarItemsALaGrillaS rs
    Case sghbien
    End Select
    oConexion.Close
    Set oConexion = Nothing
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
        mrs_FacturacionProductos!MovTipo = rs!MovTipo
        mrs_FacturacionProductos!idProducto = rs!idProducto
        mrs_FacturacionProductos!Codigo = rs!Codigo
        mrs_FacturacionProductos!NombreProducto = rs!NombreProducto
      '  mrs_FacturacionProductos!IdTipoFinanciamiento = rs!IdTipoFinanciamiento
       ' mrs_FacturacionProductos!tipoFinanciamiento = rs!IdTipoFinanciamiento
        mrs_FacturacionProductos!Cantidad = rs!Cantidad
        mrs_FacturacionProductos!PrecioUnitario = rs!PrecioUnitario
        mrs_FacturacionProductos!TotalPorPagar = rs!TotalPorPagar
        mrs_FacturacionProductos!idestadofacturacion = rs!idestadofacturacion
        mrs_FacturacionProductos!idPuntoCarga = rs!idPuntoCarga
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
    Dim rsCatHos As New Recordset
    mb_CargandoProductos = True
    Do While Not rs.EOF
        Set rsCatHos = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(rs!idProducto, ml_IdTipoFinanciamiento)
        mrs_FacturacionProductos.AddNew
        mrs_FacturacionProductos!idProducto = rs!idProducto
        mrs_FacturacionProductos!Codigo = rs!Codigo
        mrs_FacturacionProductos!NombreProducto = rs!nombre
        'mgaray201411b
        If ExisteColumnaLabEnRs(rs) = True Then
            If ExisteColumnaLabEnRs(mrs_FacturacionProductos) = True Then
            mrs_FacturacionProductos!labConfHIS = rs!labConfHIS
            End If
        End If
        mrs_FacturacionProductos!idTipoFinanciamiento = ml_IdTipoFinanciamiento
        mrs_FacturacionProductos!Cantidad = rs!Cantidad
        mrs_FacturacionProductos!PrecioUnitario = rs!Precio
        mrs_FacturacionProductos!TotalPorPagar = rs!Total
        If rsCatHos.RecordCount > 0 Then
           mrs_FacturacionProductos!SeUsaSinPrecio = IIf(IsNull(rsCatHos!SeUsaSinPrecio), False, rsCatHos!SeUsaSinPrecio)
        Else
           mrs_FacturacionProductos!SeUsaSinPrecio = 0
        End If
        mrs_FacturacionProductos!CantidadSinEditar = rs!Cantidad
        mrs_FacturacionProductos!poliza = IIf(IsNull(rs!Observaciones), "", rs!Observaciones)
        mrs_FacturacionProductos!IdFacturacionProducto = 1
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
           rsProductos!TotalPorPagar = rsProductos!Cantidad * rsProductos!PrecioUnitario
           rsProductos.Update
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
    
    Set rsProductos = mrs_FacturacionProductos.Clone
    
    If rsProductos.RecordCount = 0 Then
        Exit Sub
    End If
    
    If Not (rsProductos.EOF And rsProductos.BOF) Then
        rsProductos.MoveFirst
        Do While Not rsProductos.EOF
        
            dSubTotal = rsProductos!TotalPorPagar
            lIdEstadoFacturacion = rsProductos!idestadofacturacion
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
                Case lnIdPagosACuenta
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
    grdProductos.Refresh
    RaiseEvent Totalizado(dTotalIngresado, dTotalPendientePago, dTotalPagoACuenta, dTotalExonerado, dTotalPagado, dTotalPorDevolver, dTotalDevuelto, dTotalAnulado)
    lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")

   
End Sub

'debb-26/04/2011
Private Sub cmdFiltraBusqueda_Click()
    If Trim(txtFiltroBusqueda.Text) <> "" Then
        Dim oRsFiltrar As New Recordset
        Set oRsFiltrar = grillaBusqueda.DataSource
        oRsFiltrar.Filter = "nombre like '%" & Trim(txtFiltroBusqueda.Text) & "%'"
        Set grillaBusqueda.DataSource = oRsFiltrar
        txtFiltroBusqueda.Text = ""
        grillaBusqueda.SetFocus
    End If
End Sub

Private Sub grdProductos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    Totalizar
End Sub

'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
  If ml_PermiteAgregarItems = True Then
    If mb_CargandoProductos Then Exit Sub
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
                'MsgBox grdProductos.ActiveCell.Row.Cells("idProducto").Value
                
                ConfigurarProductoPorCodigo grdProductos
'                If grdProductos.ActiveCell.Row.Cells("idProducto").Value = 5948 Then grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
            Case "NombreProducto"
               On Error Resume Next
              If grdProductos.ActiveCell.Row.Cells("NombreProducto").Value = "Otros Ingresos" Then grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
            Case "Cantidad"
                RecalcularSubTotal grdProductos
            Case "PrecioUnitario"
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
        If mrs_FacturacionProductos.Fields!PermiteEditarPrecio = False And grdProductos.ActiveCell.Column.Key = "PrecioUnitario" Then
            Cancel = True
            Exit Sub
        End If
        'mgaray201411a
        If grdProductos.ActiveCell.Column.Key = "labConfHIS" And Not (IsNull(Cell.Row.Cells("idProducto").Value)) Then
            Dim sLab As String
            If IsNull(NewValue) Then
                sLab = ""
            Else
                sLab = CStr(NewValue)
            End If
            If ItemLabYaExiste(Cell.Row.Cells("idProducto").Value, sLab, mrs_FacturacionProductos.Bookmark) Then
                Cancel = True
                Exit Sub
            End If
        End If
    Else
        If grdProductos.ActiveCell.Column.Key <> "IdFacturacionProducto" And grdProductos.ActiveCell.Column.Key <> "Poliza" Then
        Cancel = True
        End If
    End If
End Sub

Private Sub grdProductos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
  If ml_PermiteAgregarItems = True Then mb_FilaEditable = True
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
    Else
        Cancel = True
    End If
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
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
        Case lnIdPagosACuenta, lnIdDepositoGarantia
            Row.Appearance.ForeColor = &HC7613F
        Case 4692
            Row.Appearance.ForeColor = &H16CD32
        Case 4693
            Row.Appearance.ForeColor = &H3049FA
        Case lnIdDevolucion
            Row.Appearance.ForeColor = vbGreen
        End Select

End Sub

Private Sub grdProductos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    OnKeyDown grdProductos, KeyCode
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    OnKeyPress grdProductos, KeyAscii
End Sub


Private Sub grdProductos_KeyUp(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
       Dim lnKeyCode As Integer
       lnKeyCode = KeyCode
       RaiseEvent SePresionoTeclaEspecial(lnKeyCode)
    End If

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
        If ml_IdPuntoCargaServicioHosp > 0 Then
           Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorCodigo, oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, ml_IdPuntoCargaServicioHosp, sghSoloCPT)
        Else
           Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigoDEBB(oRow.Cells("codigo").Value, ml_IdTipoFinanciamiento, ml_IdPuntoCarga)
           'Exit Sub
        End If
    Case sghbien
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigoDEBB(oRow.Cells("codigo").Value, oRow.Cells("idtipofinanciamiento").Value, ml_IdPuntoCarga)
    Case Else
        Exit Sub
    End Select
    
    If rs.RecordCount > 0 Then
       'debb-14022011
       If rs.Fields("idproducto").Value <> lnIdPagosACuenta And rs.Fields("idproducto").Value <> lnIdOtrosAdm Then
       'debb-14022011
            'Busca si ya existe el producto
            'mgaray201411a
            Dim bExisteProducto As Boolean
'            If mb_MostrarColumnaLab = True Then
'                bExisteProducto = ItemLabYaExiste(rs.Fields("idproducto").Value, IIf(IsNull(mrs_FacturacionProductos.Fields!labConfHIS), "", mrs_FacturacionProductos.Fields!labConfHIS), mrs_FacturacionProductos.Bookmark)
'            Else
                bExisteProducto = ItemYaExiste(rs.Fields("idproducto").Value)
'            End If
            If Not bExisteProducto Then
                oRow.Cells("IdFacturacionProducto").Value = 0
                oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
                oRow.Cells("NombreProducto").Value = rs.Fields("NombreProducto").Value
                oRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
                oRow.Cells("TotalPorPagar").Value = rs.Fields("preciounitario").Value
                oRow.Cells("cantidad").Value = 1
                oRow.Cells("SeUsaSinPrecio").Value = IIf(IsNull(rs.Fields("SeUsaSinPrecio").Value), False, rs.Fields("SeUsaSinPrecio").Value)
                If rs.Fields("idproducto").Value = 5948 Then grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
                oRow.Cells("IdFacturacionProducto").Value = 1
            Else
                oRow.Cells("NombreProducto").Value = ""
            End If
       End If
    End If

End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As Long) As Boolean
        'debb-14022011
        On Error Resume Next
        If lnIdProducto > 0 Then
        'debb-14022011
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
        End If
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
                    FraFiltroBusqueda.Visible = False
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
            If optPorCodigo.Value = True Then
               grdProductosFocusColumna "codigo"
            Else
               grdProductosFocusColumna "NombreProducto"
            End If
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
                'SendKeys "{Tab}"
                End If
                Exit Sub
            End If
            If oGrilla.ActiveCell.Column.Key = "Poliza" Then
               If KeyAscii = 13 Then
                  SendKeys "{Tab}"
               End If
            End If
            If oGrilla.ActiveCell.Column.Key = "Cantidad" Then
                If KeyAscii = 13 Then
                   mnuAgregarServicio_Click
                   If optPorCodigo.Value = True Then
                      grdProductosFocusColumna "codigo"
                   Else
                      grdProductosFocusColumna "NombreProducto"
                   End If
                End If
                Exit Sub
            End If
    
            If oGrilla.ActiveCell.Column.Key = "PrecioUnitario" Then
                If KeyAscii = 13 Then
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
                        FraFiltroBusqueda.Visible = False
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
                        If ml_IdPuntoCargaServicioHosp > 0 Then
                           Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, ml_IdPuntoCargaServicioHosp, sghSoloCPT)
                        Else
                           'debb-24/03/2011
                           Select Case ml_FiltraCpt
                           Case sghCptSoloTomografia
                              Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, sghPtoCargaTomografia, sghSoloCPT)
                           Case sghCptSoloRayosX
                              Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, sghPtoCargaRayosX, sghSoloCPT)
                           Case sghCptSoloEcografiaObstetrica
                              Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, sghPtoCargaEcogObstetrica, sghSoloCPT)
                           Case sghCptSoloEcografiaGeneral
                              Set rs = mo_AdminCaja.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(sghPorDescripcion, sNombre, lIdTipoFinanciamiento, sghPtoCargaEcogGeneral, sghSoloCPT)
                           Case sghCptSoloLaboratorio
                              Set rs = mo_AdminCaja.ServiciosFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, sghPtoCargaPatologiaClinica)
                           Case Else
                               Set rs = mo_AdminCaja.ServiciosFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga)
                           End Select
                           'debb-24/03/2011
                        End If
                    Case sghbien
                        Set rs = mo_AdminCaja.BienesFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga)
                    Case Else
                        
                    End Select
                    Set grillaBusqueda.DataSource = rs
                    grillaBusqueda.Left = oGrilla.Left
                    FraFiltroBusqueda.Left = oGrilla.Left + grillaBusqueda.Width
                    If mrs_FacturacionProductos.RecordCount < 7 Then
                       If Not oGrilla.ActiveCell.GetUIElement Is Nothing Then 'Actualizado 10092014
                            grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                            FraFiltroBusqueda.Top = grillaBusqueda.Top - 70
                       End If
                    Else
                       grillaBusqueda.Top = 0
                       FraFiltroBusqueda.Top = 0
                    End If
                    grillaBusqueda.Visible = True
                    grillaBusqueda.Enabled = True
                    FraFiltroBusqueda.Visible = True
                    
                End Select
            End If
    End If
End Sub


'WILLIAM CASTRO
Sub GenerarRecordsetProductos()
    Set mrs_FacturacionProductos = DevuelveGenerarRecordsetProductos()
    
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    On Error GoTo ConfigEstilo
    
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
   
    oGrilla.Bands(0).Columns("EstadoLocal").Hidden = True
    oGrilla.Bands(0).Columns("IdCajero").Hidden = True
    oGrilla.Bands(0).Columns("FechaCajero").Hidden = True
    oGrilla.Bands(0).Columns("IdUsuarioAutorizaDevolucion").Hidden = True
    oGrilla.Bands(0).Columns("FechaAutorizaDevolucion").Hidden = True
    oGrilla.Bands(0).Columns("IdComprobantePago").Hidden = True
    oGrilla.Bands(0).Columns("IdComprobantePagoDevolucion").Hidden = True
    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Hidden = True
    oGrilla.Bands(0).Columns("IdEstadoFacturacion").Hidden = True
    oGrilla.Bands(0).Columns("idPuntoCarga").Hidden = True
    oGrilla.Bands(0).Columns("idOrden").Hidden = True
    oGrilla.Bands(0).Columns("MovTipo").Hidden = True
    oGrilla.Bands(0).Columns("MovNumero").Hidden = True
    oGrilla.Bands(0).Columns("SeUsaSinPrecio").Hidden = True
    oGrilla.Bands(0).Columns("PermiteEditarPrecio").Hidden = True
    oGrilla.Bands(0).Columns("PqteIdFactPaquete").Hidden = True
    oGrilla.Bands(0).Columns("PqteIdPuntoCarga").Hidden = True
    oGrilla.Bands(0).Columns("PqteIdEspecialidadServicio").Hidden = True
    oGrilla.Bands(0).Columns("PqteGrupo").Hidden = True
    oGrilla.Bands(0).Columns("CantidadSinEditar").Hidden = True
        
    oGrilla.Bands(0).Columns("IdFacturacionProducto").Header.Caption = "CantCita"
    oGrilla.Bands(0).Columns("IdFacturacionProducto").Width = 900
    oGrilla.Bands(0).Columns("IdFacturacionProducto").Activation = IIf(ms_Opcion = sghAgregar, ssActivationAllowEdit, ssActivationActivateNoEdit)
    oGrilla.Bands(0).Columns("IdFacturacionProducto").Header.Appearance.ForeColor = vbWhite
    oGrilla.Bands(0).Columns("IdFacturacionProducto").Header.Appearance.BackColor = vbRed
    
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
    oGrilla.Bands(0).Columns("Codigo").Width = 750
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationAllowEdit
    
    oGrilla.Bands(0).Columns("NombreProducto").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("NombreProducto").Width = 6000
    oGrilla.Bands(0).Columns("NombreProducto").Activation = ssActivationAllowEdit
    
    oGrilla.Bands(0).Columns("Poliza").Header.Caption = "Observaciones"
    oGrilla.Bands(0).Columns("Poliza").Width = 2500
    oGrilla.Bands(0).Columns("Poliza").Activation = ssActivationAllowEdit
    
    oGrilla.Bands(0).Columns("Poliza").Header.Appearance.ForeColor = vbWhite
    oGrilla.Bands(0).Columns("Poliza").Header.Appearance.BackColor = vbRed
    
    
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
    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Format = sighentidades.DevuelveFechaSoloFormato_DMY_HM

    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Width = 2500
    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Header.Caption = "Fec. Aut. Seguro."
    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Format = sighentidades.DevuelveFechaSoloFormato_DMY_HM
    'mgaray201411a
    If ExisteColumnaLab(oGrilla) = True Then
        oGrilla.Bands(0).Columns("labConfHIS").Hidden = Not mb_MostrarColumnaLab
        oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
        If mb_MostrarColumnaLab = True Then
            oGrilla.Bands(0).Columns("labConfHIS").Width = 700
        End If
        Call AsignarListaDeLabsEnGridaDiagnosticos(oGrilla, "labConfHIS")
    End If
    
    'Configura Values List
    SeteaListaEstado oGrilla, oGrilla.Bands(0).Columns("idEstadoFacturacion")
    SeteaListaTipoFinanciamiento oGrilla, oGrilla.Bands(0).Columns("IdTipoFinanciamiento")
    SeteaPuntosDeCarga oGrilla, oGrilla.Bands(0).Columns("idPuntoCarga")

    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("idPuntoCarga").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("idEstadoFacturacion").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("NumeroDeItem").Hidden = True          'debb-18/05/2016
    
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, sighentidades.GrillaConFilasBicolor
    
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
                oValueTF.ValueListItems.Add Val(rs!idTipoFinanciamiento), Trim(rs!Descripcion)
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
        Set rs = mo_reglasComunes.SeleccionarPuntosDeCarga()
        Do While Not rs.EOF
            If rs!idPuntoCarga <> 0 Then
                oValuePC.ValueListItems.Add Val(rs!idPuntoCarga), Trim(rs!Descripcion)
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
            oValueEstado.ValueListItems.Add Val(rs!idestadofacturacion), Trim(rs!Descripcion)
            rs.MoveNext
        Loop
        rs.Close
    Else
        Set oValueEstado = oGrilla.ValueLists.Item("listaEstadoFacturacion")
    End If
     
    Set oColumn.ValueList = oValueEstado
    
End Sub


Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, sighentidades.GrillaConFilasBicolor
End Sub
Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    oGrilla.Bands(0).Columns("SeUsaSinPrecio").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, sighentidades.GrillaConFilasBicolor
errInic:
End Sub
Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
Dim lnIdProductoBusqueda As Long
    'debb-hra-ya en version Polsalud
    On Error GoTo ErrGrillaBusqueda
    lnIdProductoBusqueda = grillaBusqueda.ActiveRow.Cells("idproducto").Value
    'mgaray201411a
    Dim bExisteProducto As Boolean
'    If mb_MostrarColumnaLab = True Then
'        bExisteProducto = ItemLabYaExiste(lnIdProductoBusqueda, IIf(IsNull(grdProductos.ActiveRow.Cells("labConfHIS").Value), "", grdProductos.ActiveRow.Cells("labConfHIS").Value), mrs_FacturacionProductos.Bookmark)
'    Else
        bExisteProducto = ItemYaExiste(lnIdProductoBusqueda)
'    End If
    If bExisteProducto Then
        grdProductos.ActiveRow.Cells("codigo").Value = ""
        grdProductos.ActiveRow.Cells("idproducto").Value = 0
        grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
    Else
        RefrescarDatos
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
        FraFiltroBusqueda.Visible = False
        SendKeys "{Tab}"
    End If
ErrGrillaBusqueda:
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
               grdProductos.ActiveRow.Cells("SeUsaSinPrecio").Value = IIf(IsNull(grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value), False, grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value)
               grdProductos.ActiveRow.Cells("IdFacturacionProducto").Value = 1
            Else
               lnPrecioUnitario = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
               lnPrecioUnitario = 0
               oDoCatalogoServicioHosp.PrecioUnitario = 0
               Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(grillaBusqueda.ActiveRow.Cells("idproducto").Value, ml_IdTipoFinanciamiento)
               If oDoCatalogoServicioHosp.PrecioUnitario = 0 And grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value = False Then
                    MsgBox "Ese Producto no tiene precio para el TIPO DE  FINANCIAMIENTO", vbExclamation, "Facturación"
               Else
                    lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                    grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
                    grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
                    grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
                    grdProductos.ActiveRow.Cells("preciounitario").Value = lnPrecioUnitario
                    grdProductos.ActiveRow.Cells("TotalPorPagar").Value = lnPrecioUnitario
                    grdProductos.ActiveRow.Cells("cantidad").Value = 1
                    grdProductos.ActiveRow.Cells("idestadofacturacion").Value = 1
                    grdProductos.ActiveRow.Cells("SeUsaSinPrecio").Value = IIf(IsNull(grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value), False, grillaBusqueda.ActiveRow.Cells("SeUsaSinPrecio").Value)
                    grdProductos.ActiveRow.Cells("IdFacturacionProducto").Value = 1
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
        FraFiltroBusqueda.Visible = False
    Case vbKeyReturn
        grillaBusqueda_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
    
End Sub

'debb-19/07/2016
Private Sub mnuAddDevolucion_Click()
    Dim lnRetornaConsumoPorCuenta As Double
    Dim oConexion As New Connection
    Dim lbContinuar99 As Boolean
    lbContinuar99 = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    If ml_idCuentaAtencion = 0 Then
        MsgBox "Las devoluciones de dinero se realizan sólo a Pacientes con N° Cuenta", vbInformation, "Mensaje"
    Else
        lnRetornaConsumoPorCuenta = mo_ReglasFacturacion.RetornaTotalPagosPendientesPorNroCuentadebb(ml_idCuentaAtencion, oConexion)
        If lnRetornaConsumoPorCuenta >= 0 Then
            'debb-01/12/2016 (inicio)
            'MsgBox "La suma de los CONSUMOS no sobrepasan los PAGOS A CUENTA", vbInformation, "Mensaje"
            If MsgBox("              La suma de los CONSUMOS no sobrepasan los PAGOS A CUENTA: " & Trim(Str(lnRetornaConsumoPorCuenta)) & Chr(13) & Chr(13) & _
                      "¿Elija el botón 'SI', si se trata de una DEVOLUCION al día siguiente de haber CANCELADO TODO", vbQuestion + vbYesNo, "Estado de Cuenta") <> vbNo Then
               lbContinuar99 = True
            End If
            'debb-01/12/2016 (fin)
        Else
            lbContinuar99 = True
            'AgregaDevolución Abs(lnRetornaConsumoPorCuenta)
            'grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
        End If
    End If
    If lbContinuar99 = True Then
        AgregaDevolución Abs(lnRetornaConsumoPorCuenta)
        grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub AgregaDevolución(lnConsumoXcuenta As Double)
    Dim lcCodigoPrd As String, lcNombrePrd As String
    DevuelveCodigoYdescripcionSegunId lnIdDevolucion, lcCodigoPrd, lcNombrePrd
    
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!idProducto = lnIdDevolucion
        .Fields!Codigo = lcCodigoPrd
        .Fields!NombreProducto = lcNombrePrd
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = lnConsumoXcuenta
        .Fields!TotalPorPagar = lnConsumoXcuenta
        .Fields!idTipoFinanciamiento = 1
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        On Error Resume Next
        .Fields!idAtencion = mo_DoAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 1
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!PermiteEditarPrecio = True
    End With
    mb_CargandoProductos = False
    Totalizar
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    

End Sub

Private Sub mnuAgregarExoneracion_Click()
    AgregaExoneracion
End Sub

Private Sub mnuAgregarPagoACuenta_Click()
    AgregaPagoACuenta
    grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
    grdProductos.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
End Sub

Private Sub mnuAgregarServicio_Click()
'    SendKeys "{Tab}"
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
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaSeguro").Value = ml_idUsuario
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
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaDevolucion").Value = ml_idUsuario
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
        grdProductos.ActiveRow.Cells("IdUsuarioAutorizaPendiente").Value = ml_idUsuario
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
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaSeguro").Value = ml_idUsuario
    grdProductos.ActiveRow.Cells("FechaAutorizaSeguro").Value = Now
    
        
End Sub

Private Sub mnuAutorizarSOAT_Click()

    If grdProductos.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 3   'SIS
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    grdProductos.ActiveRow.Cells("IdUsuarioAutorizaSeguro").Value = ml_idUsuario
    grdProductos.ActiveRow.Cells("FechaAutorizaSeguro").Value = Now
    
End Sub
'debb2014-d
Private Sub mnuIngClinica_Click()
    AgregaPagoOtrosClinica
    grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
End Sub

'debb2014-d
Sub AgregaPagoOtrosClinica()
    Dim lcCodigoPrd As String, lcNombrePrd As String
    DevuelveCodigoYdescripcionSegunId lnIdOtrosClinica, lcCodigoPrd, lcNombrePrd
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        'debb-14022011
        .Fields!idProducto = lnIdOtrosClinica
        'debb-14022011
        .Fields!Codigo = lcCodigoPrd
        .Fields!NombreProducto = lcNombrePrd
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = 1
        .Fields!TotalPorPagar = 0
        .Fields!idTipoFinanciamiento = 1
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        On Error Resume Next
        .Fields!idAtencion = mo_DoAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 1
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!PermiteEditarPrecio = True
    End With
    mb_CargandoProductos = False

    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode

End Sub


Private Sub txtFiltroBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdFiltraBusqueda_Click
    End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 5
   
   Label1.Top = UserControl.Height - UserControl.Label1.Height
   lblTotal.Top = UserControl.Height - UserControl.Label1.Height + 60
   optPorCodigo.Top = UserControl.Height - UserControl.Label1.Height
   optPorDescripcion.Top = UserControl.Height - UserControl.Label1.Height
   
End Sub

Sub LimpiarGrilla()
        On Error GoTo ErrLimpiar
        'mgaray201410f
'        If mrs_FacturacionProductosSinDatos.RecordCount > 0 Then
'            mrs_FacturacionProductosSinDatos.MoveFirst
'            Do While Not mrs_FacturacionProductosSinDatos.EOF
'                mrs_FacturacionProductosSinDatos.Delete
'                mrs_FacturacionProductosSinDatos.Update
'                mrs_FacturacionProductosSinDatos.MoveNext
'            Loop
'        End If
        Set mrs_FacturacionProductos = DevuelveGenerarRecordsetProductos
        Set grdProductos.DataSource = mrs_FacturacionProductos
        
        grillaBusqueda.Visible = False
        FraFiltroBusqueda.Visible = False
        grdProductos.Bands(0).Columns("cantidad").Activation = ssActivationAllowEdit
ErrLimpiar:
End Sub


'***************daniel barrantes**************
'***************Registra la CANTIDAD a DEVOLVER en cada Item
'***************ya autorizada anteriormente
Sub ActualizaDevolucionAutorizada(oRs As Recordset)
    If oRs.RecordCount > 0 Then
       oRs.MoveFirst
       Do While Not oRs.EOF
          mrs_FacturacionProductos.MoveFirst
          mrs_FacturacionProductos.Find "idProducto=" & oRs.Fields!idProducto
          If IsNull(oRs.Fields!cantidadDev) Or oRs.Fields!cantidadDev = 0 Then
             mrs_FacturacionProductos.Delete
          Else
          mrs_FacturacionProductos.Fields!Cantidad = oRs.Fields!cantidadDev
          mrs_FacturacionProductos.Fields!idTipoFinanciamiento = 1
          mrs_FacturacionProductos.Fields!TotalPorPagar = oRs.Fields!cantidadDev * mrs_FacturacionProductos.Fields!PrecioUnitario
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
          If mrs_FacturacionProductos.Fields!idTipoFinanciamiento = 2 Or mrs_FacturacionProductos.Fields!idTipoFinanciamiento = 3 Then
             OrdenRegistradaYaprobadaPorSisSoat = mrs_FacturacionProductos.Fields!idTipoFinanciamiento
             Exit Function
          End If
          mrs_FacturacionProductos.MoveNext
       Loop
    End If
End Function

Sub TabEnDescripcion()
    On Error Resume Next
'    grdProductos.SetFocus
'    SendKeys "{Tab}"
    If optPorCodigo.Value = True Then
       grdProductosFocusColumna "codigo"
    Else
       grdProductosFocusColumna "NombreProducto"
    End If
    
End Sub
Sub TabEnDescripcionParaFactuacion()
    On Error Resume Next
    'mgaray201410e
    If ExisteFilaNuevaEnListado() = False Then
        mnuAgregarServicio_Click
    End If
                   If optPorCodigo.Value = True Then
                      grdProductosFocusColumna "codigo"
                   Else
                      grdProductosFocusColumna "NombreProducto"
                   End If
End Sub


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

Private Sub mnuCtaBancarias_Click()
    AgregaPagoBanco
    grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
    grdProductos.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
End Sub

Sub AgregaPagoBanco()
    Dim lcCodigoPrd As String, lcNombrePrd As String
    DevuelveCodigoYdescripcionSegunId lnIdDepositoGarantia, lcCodigoPrd, lcNombrePrd
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!idProducto = lnIdDepositoGarantia
        .Fields!Codigo = lcCodigoPrd
        .Fields!NombreProducto = lcNombrePrd
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = 1
        .Fields!TotalPorPagar = 0
        .Fields!idTipoFinanciamiento = 1
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        On Error Resume Next
        .Fields!idAtencion = mo_DoAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 1
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!PermiteEditarPrecio = True
    End With
    mb_CargandoProductos = False
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    
End Sub

Function DevuelveNombreServicioSegunCodigo(lcCodigo As String) As String
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_reglasComunes.CatalogoServiciosSeleccionarPorCodigo(lcCodigo)
    If oRsTmp.RecordCount > 0 Then
       DevuelveNombreServicioSegunCodigo = oRsTmp.Fields!nombre
    Else
       DevuelveNombreServicioSegunCodigo = ""
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Sub DevuelveCodigoYdescripcionSegunId(lnIdProducto As Long, ByRef lcCodigo As String, ByRef lcDescripcion As String)
    Dim oRsTmp As New Recordset
    lcCodigo = ""
    lcDescripcion = ""
    Set oRsTmp = mo_reglasComunes.CatalogoServiciosSeleccionarXidentificador(lnIdProducto)
    If oRsTmp.RecordCount > 0 Then
       lcDescripcion = oRsTmp.Fields!nombre
       lcCodigo = oRsTmp.Fields!Codigo
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub


Public Sub PaqueteServicioAgregaProductos(lnIdPaquete As Long)
    LimpiarGrilla
    Dim oRsTmp As New Recordset
    Dim lnIdGrupo As Long, lnIdEspecialidadServicio As Long, lnIdProducto As Long
    Dim lcSql As String
    Set oRsTmp = mo_ReglasFacturacion.FacturacionCatalogoPaquetesXidPaquete(lnIdPaquete)
    If oRsTmp.RecordCount > 0 Then
       lnIdGrupo = 1
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          lnIdEspecialidadServicio = oRsTmp.Fields!idEspecialidadServicio
          Do While Not oRsTmp.EOF And lnIdEspecialidadServicio = oRsTmp.Fields!idEspecialidadServicio
                lnIdProducto = oRsTmp.Fields!idProducto
                With mrs_FacturacionProductos
                .AddNew
                .Fields!IdFacturacionProducto = 0
                .Fields!idProducto = oRsTmp.Fields!idProducto
                .Fields!Codigo = oRsTmp.Fields!Codigo
                .Fields!NombreProducto = oRsTmp.Fields!procedimiento
                .Fields!Cantidad = oRsTmp.Fields!Cantidad
                .Fields!PrecioUnitario = oRsTmp.Fields!Precio
                .Fields!TotalPorPagar = oRsTmp.Fields!Importe
                .Fields!idTipoFinanciamiento = 1
                .Fields!idPuntoCarga = ml_IdPuntoCarga
                On Error Resume Next
                .Fields!idAtencion = 0
                .Fields!EstadoLocal = "A"   'Agregar
                .Fields!FechaAutorizaPendiente = 0
                .Fields!IdUsuarioAutorizaPendiente = 0
                .Fields!idestadofacturacion = 1
                .Fields!FechaAutorizaSeguro = Now
                .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
                .Fields!IdFuenteFinanciamiento = 0
                .Fields!IdServicioInternamiento = 0
                .Fields!IdUsuarioAuditoria = ml_idUsuario
                .Fields!IdComprobantePago = 0
                .Fields!IdComprobantePagoDevolucion = 0
                .Fields!PqteIdFactPaquete = lnIdPaquete
                .Fields!PqteIdPuntoCarga = oRsTmp.Fields!idPuntoCarga
                .Fields!PqteIdEspecialidadServicio = oRsTmp.Fields!idEspecialidadServicio
                .Fields!PqteGrupo = lnIdGrupo
                End With
                oRsTmp.MoveNext
                If oRsTmp.EOF Then
                   Exit Do
                End If
                If lnIdProducto = oRsTmp.Fields!idProducto Then
                   lnIdGrupo = lnIdGrupo + 1
                End If
           Loop
           lnIdGrupo = lnIdGrupo + 1
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Totalizar
    'debb-18/05/2016 (inicio)
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
    End If
    'debb-18/05/2016 (fin)
End Sub

'***Otros Procedimientos Administrativos
Private Sub mnuOtrosAdm_Click()
    AgregaPagoOtrosAdm
    grdProductos.Bands(0).Columns("preciounitario").Activation = ssActivationAllowEdit
    grdProductos.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
End Sub
Sub AgregaPagoOtrosAdm()
    Dim lcCodigoPrd As String, lcNombrePrd As String
    DevuelveCodigoYdescripcionSegunId lnIdOtrosAdm, lcCodigoPrd, lcNombrePrd
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        'debb-14022011
        .Fields!idProducto = lnIdOtrosAdm     'lnIdPagosACuenta
        'debb-14022011
        .Fields!Codigo = lcCodigoPrd
        .Fields!NombreProducto = lcNombrePrd
        .Fields!Cantidad = 1
        .Fields!PrecioUnitario = 1
        .Fields!TotalPorPagar = 0
        .Fields!idTipoFinanciamiento = 1
        .Fields!idPuntoCarga = ml_IdPuntoCarga
        On Error Resume Next
        .Fields!idAtencion = mo_DoAtencion.idAtencion
        .Fields!EstadoLocal = "A"   'Agregar
        .Fields!FechaAutorizaPendiente = 0
        .Fields!IdUsuarioAutorizaPendiente = 0
        .Fields!idestadofacturacion = 1
        .Fields!FechaAutorizaSeguro = Now
        .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
        .Fields!IdFuenteFinanciamiento = 0
        .Fields!IdServicioInternamiento = 0
        .Fields!IdUsuarioAuditoria = ml_idUsuario
        .Fields!IdComprobantePago = 0
        .Fields!IdComprobantePagoDevolucion = 0
        .Fields!PermiteEditarPrecio = True
    End With
    mb_CargandoProductos = False

    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode

End Sub

Public Function EsUnPagoOtrosAdm() As Boolean
    EsUnPagoOtrosAdm = False
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
       Do While Not mrs_FacturacionProductos.EOF
          If mrs_FacturacionProductos.Fields!idProducto = lnIdOtrosAdm Then
             EsUnPagoOtrosAdm = True
             Exit Do
          End If
          mrs_FacturacionProductos.MoveNext
       Loop
    End If
End Function

'Actualizado 09102014
Public Sub CargaCptPorAtencion()
'    LimpiarGrilla
    Dim orstemp1 As New Recordset
    Set orstemp1 = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
    If orstemp1.RecordCount > 0 Then
        Do While Not orstemp1.EOF
            With mrs_FacturacionProductos
            .AddNew
            .Fields!IdFacturacionProducto = 0
            .Fields!idProducto = orstemp1.Fields!idProducto
            .Fields!Codigo = orstemp1.Fields!Codigo
            .Fields!NombreProducto = orstemp1.Fields!nombre
            'mgaray201411a
            .Fields!labConfHIS = orstemp1.Fields!labConfHIS
            .Fields!Cantidad = orstemp1.Fields!Cantidad
            .Fields!PrecioUnitario = orstemp1.Fields!Precio
            .Fields!TotalPorPagar = orstemp1.Fields!Total
            .Fields!idTipoFinanciamiento = ml_IdTipoFinanciamiento
            .Fields!idPuntoCarga = orstemp1.Fields!idPuntoCarga
            If Not mo_DoAtencion Is Nothing Then
                .Fields!idAtencion = mo_DoAtencion.idAtencion
            End If
            .Fields!EstadoLocal = "A"   'Agregar
            .Fields!FechaAutorizaPendiente = 0
            .Fields!IdUsuarioAutorizaPendiente = 0
            
            Select Case ml_IdTipoFinanciamiento
            Case 2, 3, 4
                .Fields!idestadofacturacion = 4
                .Fields!FechaAutorizaSeguro = Now
                .Fields!IdUsuarioAutorizaSeguro = ml_idUsuario
            Case Else
                .Fields!idestadofacturacion = 1
                .Fields!FechaAutorizaSeguro = 0
                .Fields!IdUsuarioAutorizaSeguro = 0
            End Select
            .Fields!IdFuenteFinanciamiento = 1
            .Fields!IdServicioInternamiento = 0
            .Fields!IdUsuarioAuditoria = ml_idUsuario
            .Fields!IdComprobantePago = 0
            .Fields!IdComprobantePagoDevolucion = 0
            .Fields!IdOrden = ml_idOrden
            
            End With
            
            orstemp1.MoveNext
         Loop
    End If
End Sub
'mgaray201410e
Public Function ExisteFilaNuevaEnListado() As Boolean
    Dim returnValue As Boolean
    ExisteFilaNuevaEnListado = False
    On Error Resume Next
    Dim oRsProductos As ADODB.Recordset
    If Not (mrs_FacturacionProductos Is Nothing) Then
        Set oRsProductos = mrs_FacturacionProductos.Clone()
        If oRsProductos.RecordCount > 0 Then
            oRsProductos.MoveLast
            Dim sCodigo As String
            sCodigo = IIf(IsNull(oRsProductos.Fields!Codigo), "", oRsProductos.Fields!Codigo)
            If sCodigo = "" Then
                returnValue = True
            End If
        End If
    End If
    ExisteFilaNuevaEnListado = returnValue
    Err = 0
End Function

Private Function DevuelveGenerarRecordsetProductos() As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    
    With oRs
          .Fields.Append "IdFacturacionProducto", adInteger
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
          'mgaray201411a
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
          .Fields.Append "IdTipoFinanciamiento", adInteger
          .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
          .Fields.Append "Poliza", adVarChar, 255, adFldIsNullable
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
          .Fields.Append "SeUsaSinPrecio", adBoolean
          .Fields.Append "PermiteEditarPrecio", adBoolean
          .Fields.Append "PqteIdFactPaquete", adInteger
          .Fields.Append "PqteIdPuntoCarga", adInteger
          .Fields.Append "PqteIdEspecialidadServicio", adInteger
          .Fields.Append "PqteGrupo", adInteger
          .Fields.Append "CantidadSinEditar", adInteger
          .Fields.Append "NumeroDeItem", adInteger                                          'debb-18/05/2016
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set DevuelveGenerarRecordsetProductos = oRs
End Function

'mgaray201411a
Private Function ExisteColumnaLab(oGrilla As SSUltraGrid) As Boolean
    Dim i As Long
    Dim returnValue As Boolean
    returnValue = False
    
    For i = 0 To oGrilla.Bands(0).Columns.Count - 1
        If oGrilla.Bands(0).Columns(i).Key = "labConfHIS" Then
            returnValue = True
            Exit For
        End If
    Next i
    ExisteColumnaLab = returnValue
End Function

Private Function AsignarListaDeLabsEnGridaDiagnosticos(oGrilla As SSUltraGrid, cNombreColumna As String) As Boolean
On Error GoTo miError
    Dim oRsLabHis As ADODB.Recordset

    Set oRsLabHis = mo_reglasComunes.DevuelveHIS_SITUACIOporDescripcion()

    With oGrilla.ValueLists.Add("ListaLab").ValueListItems
           If oRsLabHis.RecordCount > 0 Then
              oRsLabHis.MoveFirst
              Do While Not oRsLabHis.EOF
                 .Add Right(Trim((oRsLabHis.Fields!valores)), 3), Trim(oRsLabHis.Fields!valores) '& "(" & Trim(mo_RsLabHis.Fields!descripcio) & ")"
                 oRsLabHis.MoveNext
              Loop
           End If
    End With
    oRsLabHis.Close
    oGrilla.Bands(0).Columns(cNombreColumna).ValueList = "ListaLab"

    AsignarListaDeLabsEnGridaDiagnosticos = True
miError:
    If Err Then
        MsgBox Err.Description & " : " & Err.Description, vbInformation, "Módulo Perinatal"
    End If
End Function

Function ItemLabYaExiste(lnIdProducto As Long, cLabHIS As String, Optional ByVal oRegistroActual As Variant = 0) As Boolean
On Error Resume Next
    If lnIdProducto > 0 Then
        Dim lo_RsCptFrecuentes As ADODB.Recordset
    
'        Set lo_RsCptFrecuentes = mrs_FacturacionProductos.Clone
'
'        ItemLabYaExiste = False
'        With lo_RsCptFrecuentes
'            If Not (.BOF = True And .EOF = True) Then
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    While .EOF = False
'                        If Not (.CompareBookmarks(oRegistroActual, .Bookmark) = adCompareEqual) Then 'Or EditLab = False Then
'                            If .Fields!idProducto = lnIdProducto And Trim(cLabHIS) = Trim(IIf(IsNull(.Fields!labConfHIS), "", .Fields!labConfHIS)) Then
'                                .Bookmark = oRegistroActual
'                                ItemLabYaExiste = True
'                                MsgBox "Este producto y lab ya está registrado", vbInformation, "Facturación"
'                                Exit Function
'                            End If
'                        End If
'                        .MoveNext
'                    Wend
'                End If
'            End If
'        End With
        
        Dim oRsLabHis As ADODB.Recordset

        Set oRsLabHis = mo_reglasComunes.DevuelveHIS_SITUACIOporDescripcion()
    
        If mo_reglasComunes.existeCodigoLabHis(oRsLabHis, cLabHIS) = False Then
            MsgBox "Código LAB No Valido", vbInformation, "Módulo Perinatal"
            ItemLabYaExiste = True
            Exit Function
        End If
    End If
End Function

Private Function ExisteColumnaLabEnRs(oRs As ADODB.Recordset) As Boolean
On Error Resume Next
    Dim returnValue As Boolean
    returnValue = False
    If Not (oRs Is Nothing) Then
        Dim i As Integer
        'mgaray201411b
        For i = 0 To oRs.Fields.Count - 1
            If oRs.Fields(i).Name = "labConfHIS" Then
                returnValue = True
                Exit For
            End If
        Next i
    End If
    ExisteColumnaLabEnRs = returnValue
Err = 0
End Function

Sub CajaConDescripcionLargaActualizaImporte(lnNuevoImporte As Double)
    If mrs_FacturacionProductos.RecordCount = 0 Then
       AgregaPagoOtrosClinica
    End If
    mrs_FacturacionProductos.MoveFirst
    If mrs_FacturacionProductos!idProducto = 0 Then
        Dim lcCodigoPrd As String, lcNombrePrd As String
        DevuelveCodigoYdescripcionSegunId Val(wxParametro549), lcCodigoPrd, lcNombrePrd
        mrs_FacturacionProductos!idProducto = Val(wxParametro549)
        mrs_FacturacionProductos!Codigo = lcCodigoPrd
        mrs_FacturacionProductos!NombreProducto = lcNombrePrd
    End If
    mrs_FacturacionProductos!PrecioUnitario = lnNuevoImporte
    mrs_FacturacionProductos!TotalPorPagar = lnNuevoImporte
    mrs_FacturacionProductos.Update
    Totalizar
End Sub


VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFactItemsPorCuenta2 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   ScaleHeight     =   5730
   ScaleWidth      =   11685
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   4683
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
      Caption         =   "grillaBusqueda"
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5745
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   10134
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
   Begin VB.Menu mnuProductos 
      Caption         =   "mnuProductos"
      Begin VB.Menu mnuAgregarServicio 
         Caption         =   "Agregar servicio"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAgregarExoneracion 
         Caption         =   "Agregar exoneración"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAgregarPagoACuenta 
         Caption         =   "Agregar pago a cuenta"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutorizaPacienteNormal 
         Caption         =   "Paciente Normal"
         Enabled         =   0   'False
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
Attribute VB_Name = "ucFactItemsPorCuenta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mostrar Procedimientos
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
Dim mo_DoAtencion As DOAtencion
Dim ml_idUsuario As Long
Dim ml_IdPuntoCarga As Long
Dim ms_EstadosFacturacion As String
Dim ms_TiposFinanciamiento As String

'edicion de la grilla
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnTotalExonerado As Double
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnIdPagosACuenta As Long
Dim lnIdDepositoGarantia As Long
Dim lb_inHabilitaEdicionColumnasDelGrid As Boolean

Property Let InHabilitaEdicionColumnasDelGrid(lValue As Boolean)
    lb_inHabilitaEdicionColumnasDelGrid = lValue
End Property

Property Set FacturacionProductos(oValue As Recordset)
    Set mrs_FacturacionProductos = oValue
End Property

Property Get FacturacionProductos() As Recordset
    'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
    'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
    Set FacturacionProductos = mrs_FacturacionProductos.Clone()
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



Sub Inicializar()
    
    Set mrs_FacturacionProductos = New Recordset
    GenerarRecordsetProductos
    
    ms_EstadosFacturacion = ""
    Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
    UserControl.mnuAgregarServicio.Enabled = mo_PermisosFacturacion.AgregarServicios
    UserControl.mnuAgregarExoneracion.Enabled = mo_PermisosFacturacion.AgregarExoneraciones
    
    UserControl.mnuAutorizarSIS.Enabled = False 'mo_PermisosFacturacion.AutorizarSIS
    UserControl.mnuAutorizarSOAT.Enabled = False 'mo_PermisosFacturacion.AutorizarSOAT
    UserControl.mnuAutorizarPendientePago.Enabled = False 'mo_PermisosFacturacion.AutorizarPendientesDePago
    UserControl.mnuAutorizarConvenio.Enabled = False 'mo_PermisosFacturacion.AutorizarConvenios
    UserControl.mnuAutorizarDevolucion.Enabled = False 'mo_PermisosFacturacion.AutorizarDevoluciones
    UserControl.mnuAgregarServicio.Enabled = False
    
    UserControl.grdProductos.DialogStrings(ssDeleteRow) = "Ud. ha seleccionado una fila para borrarla" + Chr(13) + "Seleccione <Si> para borrar la fila o <No> para Cancelar "
    '
    lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
    lnIdDepositoGarantia = Val(lcBuscaParametro.SeleccionaFilaParametro(254))
   
End Sub

Sub AgregaProducto()
        
    grdProductos.SetFocus
    'mrs_FacturacionProductos.Update

    
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
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode

    
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

Sub AgregaDepositoDeGarantia()
        
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!idProducto = lnIdDepositoGarantia
        .Fields!Codigo = "AD045"
        .Fields!NombreProducto = "Depósito de Garantía"
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
        
    End With
    mb_CargandoProductos = False

    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
    
End Sub

Sub AgregaPagoACuenta()
        
    mb_CargandoProductos = True
    With mrs_FacturacionProductos
        .AddNew
        .Fields!IdFacturacionProducto = 0
        .Fields!idProducto = lnIdPagosACuenta
        .Fields!Codigo = "F00001"
        .Fields!NombreProducto = "Pago a cuenta"
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

'***************daniel barrantes**************
'***************aqui tambien se incluye CAMPOS para SIS/SOAT/EXONERACION
'***************
Sub CargaProductosPorIdCuentaAtencion()
    Dim oCargaProductosPorIdCuentaAtencion As New SighFacturacion.dllFactUCGestionCaja
    Dim lnms_TipoProducto As Integer
    lnms_TipoProducto = ms_TipoProducto
    mb_CargandoProductos = True
    Select Case ms_TipoProducto
    Case sghbien
      LimpiarGrilla
      oCargaProductosPorIdCuentaAtencion.CargaDetalleBoletaB mrs_FacturacionProductos, ml_idCuentaAtencion, lnTotalExonerado
    Case sghServicio
      LimpiarGrilla
      oCargaProductosPorIdCuentaAtencion.CargaDetalleBoletaS mrs_FacturacionProductos, ml_idCuentaAtencion, lnTotalExonerado
    End Select
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mrs_FacturacionProductos 'Carga productos de consumo en farmacia
    If mrs_FacturacionProductos.RecordCount > 0 Then
       mrs_FacturacionProductos.MoveFirst
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
    
    rsProductos.MoveFirst
    Do While Not rsProductos.EOF
    
        dSubTotal = rsProductos!TotalPorPagar
        lIdEstadoFacturacion = rsProductos!idestadofacturacion
        lIdProducto = rsProductos!idProducto
        
        Select Case lIdEstadoFacturacion
        Case 1 'Pendiente por pagar
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

    RaiseEvent Totalizado(dTotalIngresado, dTotalPendientePago, dTotalPagoACuenta, dTotalExonerado, dTotalPagado, dTotalPorDevolver, dTotalDevuelto, dTotalAnulado)
        


End Sub

'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
        If mb_CargandoProductos Then
            Exit Sub
        End If
End Sub


Private Sub grdProductos_AfterRowsDeleted()
    If ml_ultimoProductoEliminado > 0 Then
        mo_ProductosEliminados.Add ml_ultimoProductoEliminado
        ml_ultimoProductoEliminado = 0
    End If

End Sub

Private Sub grdProductos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
        
        If mb_FilaEditable Then
            'Si la fila es editable y estamos en la celda de codigo se completa los datos
            'del producto
            Select Case grdProductos.ActiveCell.Column.Key
            Case "Codigo"
                ConfigurarProductoPorCodigo grdProductos
            Case "Cantidad"
                RecalcularSubTotal grdProductos
            Case "TipoFinanciamiento"
            Case "EstadoFacturacion"
            End Select
        End If
        
End Sub

Private Sub grdProductos_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    
    'Si la fila no es editable, cancela cualquier cambio en la fila
    If Not mb_FilaEditable Then
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub grdProductos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    mb_FilaEditable = True
End Sub

Private Sub grdProductos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If lb_inHabilitaEdicionColumnasDelGrid = False Then
        'Si ya esta pagado cancela la eliminacion
        If Rows.Item(0).Cells("EstadoLocal").Value = "M" And Rows.Item(0).Cells("idestadofacturacion").Value = 4 Then
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
    If Button = 2 Then
        PopupMenu mnuProductos
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
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
     
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
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)

        If oGrilla.ActiveCell Is Nothing Then
            Exit Sub
        End If
        
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
                If grillaBusqueda.Visible Then
                   grillaBusqueda.SetFocus
                End If
            Case vbKeyLeft
            End Select
        End Select



End Sub

Sub OnKeyPress(oGrilla As SSUltraGrid, KeyAscii As UltraGrid.SSReturnShort)
                
        'Si la fila no es editable, cancela cualquier cambio en la fila
        If Not mb_FilaEditable Then
            Exit Sub
        End If
        
        If oGrilla.ActiveCell Is Nothing Then
            Exit Sub
        End If

        If oGrilla.ActiveCell.Column.Key = "Cantidad" Then
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
                grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY
                
                grillaBusqueda.Visible = True
                grillaBusqueda.Enabled = True
            End Select
        End If

End Sub


'WILLIAM CASTRO
'***************daniel barrantes**************
'***************tambien se ha incluido CAMPOS para SIS/SOAT/EXONERACIONES
'***************
Sub GenerarRecordsetProductos()
     Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUCGestionCaja
     Set mrs_FacturacionProductos = oGenerarRecordsetProductos.GenerarRecordsetProductos()


    'Set grdProductos.DataSource = mrs_FacturacionProductos
    
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    On Error GoTo ConfigEstilo
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
    oGrilla.Bands(0).Columns("NombreProducto").Width = 8500
    oGrilla.Bands(0).Columns("NombreProducto").Activation = ssActivationAllowEdit
    
    
    oGrilla.Bands(0).Columns("Cantidad").Header.Caption = "Cant.Pag"
    oGrilla.Bands(0).Columns("Cantidad").Format = "###0"
    oGrilla.Bands(0).Columns("Cantidad").Width = 800
    oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
    If lb_inHabilitaEdicionColumnasDelGrid = True Then
       oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    End If
    
    oGrilla.Bands(0).Columns("preciounitario").Header.Caption = "P.U.(S/.)"
    oGrilla.Bands(0).Columns("preciounitario").Format = "#0.00"
    oGrilla.Bands(0).Columns("preciounitario").Width = "900"
    If lb_inHabilitaEdicionColumnasDelGrid = True Then
       oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    End If
    
    oGrilla.Bands(0).Columns("TotalPorPagar").Header.Caption = "Sub Total"
    oGrilla.Bands(0).Columns("TotalPorPagar").Format = "#0.00"
    oGrilla.Bands(0).Columns("TotalPorPagar").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("TotalPorPagar").Width = 1200
    If lb_inHabilitaEdicionColumnasDelGrid = True Then
       oGrilla.Bands(0).Columns("TotalPorPagar").Activation = ssActivationActivateNoEdit
    End If
  
    oGrilla.Bands(0).Columns("IdEstadoFacturacion").Width = 1500
    oGrilla.Bands(0).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
    oGrilla.Bands(0).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList

    oGrilla.Bands(0).Columns("idPuntoCarga").Header.Caption = "Puntos de carga"
    oGrilla.Bands(0).Columns("idPuntoCarga").Width = 1500
    oGrilla.Bands(0).Columns("idPuntoCarga").Style = ssStyleDropDownList

    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Width = 2500
    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Aut. Pend."
    oGrilla.Bands(0).Columns("FechaAutorizaPendiente").Format = sighEntidades.DevuelveFechaSoloFormato_DMY_HM

    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Width = 2500
    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Header.Caption = "Fec. Aut. Seguro."
    oGrilla.Bands(0).Columns("FechaAutorizaSeguro").Format = sighEntidades.DevuelveFechaSoloFormato_DMY_HM
    
    'Configura Values List
    SeteaListaEstado oGrilla, oGrilla.Bands(0).Columns("idEstadoFacturacion")
    SeteaListaTipoFinanciamiento oGrilla, oGrilla.Bands(0).Columns("IdTipoFinanciamiento")
    SeteaPuntosDeCarga oGrilla, oGrilla.Bands(0).Columns("idPuntoCarga")

    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("idPuntoCarga").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("idEstadoFacturacion").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("IdTipoFinanciamiento").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("ImporteSIS").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("ImporteSIS").Header.Caption = "Cant.SIS"
    oGrilla.Bands(0).Columns("ImporteSOAT").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("ImporteSOAT").Header.Caption = "Cant.SOAT"
    oGrilla.Bands(0).Columns("ImporteEXO").Activation = ssActivationActivateNoEdit
    
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

Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, sighEntidades.GrillaConFilasBicolor
End Sub
Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
errInic:
End Sub
Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
Dim lnIdProductoBusqueda As Long
    'debb-hra-ya en version Polsalud
    On Error GoTo ErrGrillaBusqueda
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
           grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
           grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
           grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
           grdProductos.ActiveRow.Cells("preciounitario").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdProductos.ActiveRow.Cells("TotalPorPagar").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdProductos.ActiveRow.Cells("cantidad").Value = 1
           grdProductos.ActiveRow.Cells("idestadofacturacion").Value = 1
        
        End If
        
        Totalizar
        
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
        
        Exit Sub
    End If
ErrGrillaBusqueda:
End Sub

Private Sub grillaBusqueda_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)

    Select Case KeyCode
    Case vbKeyEscape
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
    Case vbKeyReturn
        grillaBusqueda_DblClick
    End Select
    
End Sub

Private Sub mnuAgregarDepositoGarantia_Click()
  AgregaDepositoDeGarantia
End Sub

Private Sub mnuAgregarExoneracion_Click()
    AgregaExoneracion
End Sub

Private Sub mnuAgregarPagoACuenta_Click()
    AgregaPagoACuenta
End Sub

Private Sub mnuAgregarServicio_Click()
    AgregaProducto
End Sub

Private Sub mnuAutorizaPacienteNormal_Click()

    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 1   'Paciente Normal
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 1   'Ingresado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    
End Sub

Private Sub mnuAutorizarConvenio_Click()

    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 4   'Convenio
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado

End Sub

Private Sub mnuAutorizarDevolucion_Click()
    
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 5   'Devuelto
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    
End Sub

Private Sub mnuAutorizarPendientePago_Click()
    
    Select Case grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value
    Case 1
        grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 3   'Pagado
        grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    Case 2, 3, 4
        MsgBox "La autorización de pendientes de pago no aplica a seguros y convenios ", vbInformation, "Facturacion de servicios"
    End Select

End Sub

Private Sub mnuAutorizarSIS_Click()

    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 2   'SIS
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
        
End Sub

Private Sub mnuAutorizarSOAT_Click()

    grdProductos.ActiveRow.Cells("IdTipoFinanciamiento").Value = 3   'SIS
    grdProductos.ActiveRow.Cells("IdEstadoFacturacion").Value = 4   'Pagado
    grdProductos.ActiveRow.Cells("EstadoLocal").Value = "M"   'Modificado
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height
   
End Sub

Sub LimpiarGrilla()
        On Error GoTo ErrLimpiar
        If mrs_FacturacionProductos.RecordCount > 0 Then
            mrs_FacturacionProductos.MoveFirst
            Do While Not mrs_FacturacionProductos.EOF
                mrs_FacturacionProductos.Delete
                mrs_FacturacionProductos.Update
                mrs_FacturacionProductos.MoveNext
            Loop
        End If
        Set grdProductos.DataSource = mrs_FacturacionProductos
ErrLimpiar:
   
End Sub


Function DevuelveTotalImporteExonerado() As Double
    DevuelveTotalImporteExonerado = lnTotalExonerado
End Function


Sub DevuelveTotal()
    Dim rsProductos As New Recordset
    Dim dTotalIngresado As Double

    dTotalIngresado = 0
    Set rsProductos = mrs_FacturacionProductos.Clone
    If rsProductos.RecordCount = 0 Then
        Exit Sub
    End If
    rsProductos.MoveFirst
    Do While Not rsProductos.EOF
       dTotalIngresado = dTotalIngresado + rsProductos!TotalPorPagar
        rsProductos.MoveNext
    Loop
End Sub

'debb-17/02/2011
Function DevuelveTotalFS() As Double
    Dim rsProductos As New Recordset
    Dim dTotalIngresado As Double
    dTotalIngresado = 0
    Set rsProductos = mrs_FacturacionProductos.Clone
    If rsProductos.RecordCount > 0 Then
        rsProductos.MoveFirst
        Do While Not rsProductos.EOF
           dTotalIngresado = dTotalIngresado + rsProductos!TotalPorPagar
            rsProductos.MoveNext
        Loop
    End If
    DevuelveTotalFS = dTotalIngresado
End Function


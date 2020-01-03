VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ucDetalleFacturacion 
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   ScaleHeight     =   9555
   ScaleWidth      =   12045
   Begin VB.CommandButton btnHistoria 
      Caption         =   "..."
      Height          =   315
      Left            =   2610
      TabIndex        =   16
      Top             =   1200
      Width           =   315
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   90
      TabIndex        =   8
      Top             =   540
      Width           =   11895
      Begin VB.ComboBox cmbFechaIngreso 
         Enabled         =   0   'False
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
         Left            =   9210
         TabIndex        =   12
         Top             =   660
         Width           =   2490
      End
      Begin VB.TextBox txtPaciente 
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
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   11
         Top             =   660
         Width           =   4755
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
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   10
         Top             =   660
         Width           =   1065
      End
      Begin VB.ComboBox cmbIdPuntosDeCarga 
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
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Ingreso"
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
         Left            =   7710
         TabIndex        =   15
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Historia"
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
         TabIndex        =   14
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Pto de Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   90
      TabIndex        =   0
      Top             =   8310
      Width           =   11865
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ucDetalleFacturacion.ctx":0000
         DownPicture     =   "ucDetalleFacturacion.ctx":0460
         Height          =   700
         Left            =   4620
         Picture         =   "ucDetalleFacturacion.ctx":08D5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ucDetalleFacturacion.ctx":0D4A
         DownPicture     =   "ucDetalleFacturacion.ctx":120E
         Height          =   700
         Left            =   6120
         Picture         =   "ucDetalleFacturacion.ctx":16FA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab tabCuentas 
      Height          =   6495
      Left            =   60
      TabIndex        =   1
      Top             =   1770
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Servicios"
      TabPicture(0)   =   "ucDetalleFacturacion.ctx":1BE6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdServicios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes e Insumos"
      TabPicture(1)   =   "ucDetalleFacturacion.ctx":1C02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grillaBusqueda"
      Tab(1).Control(1)=   "grdBienes"
      Tab(1).ControlCount=   2
      Begin UltraGrid.SSUltraGrid grillaBusqueda 
         Height          =   2655
         Left            =   -74400
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   10680
         _ExtentX        =   18838
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
      Begin UltraGrid.SSUltraGrid grdBienes 
         Height          =   5895
         Left            =   -74850
         TabIndex        =   3
         Top             =   480
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   10398
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67174420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ValueLists      =   "ucDetalleFacturacion.ctx":1C1E
         Caption         =   "Bienes e Insumos"
      End
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   5895
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   10398
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67174420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ValueLists      =   "ucDetalleFacturacion.ctx":1C89
         Caption         =   "Servicios"
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Catálogo de Servicios"
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
      TabIndex        =   7
      Top             =   0
      Width           =   12015
   End
   Begin VB.Menu mnuBienes 
      Caption         =   "Bienes"
      Begin VB.Menu mnuAgregaBienes 
         Caption         =   "Agregar Bienes"
      End
   End
   Begin VB.Menu mnuServicios 
      Caption         =   "Servicios"
      Begin VB.Menu mnuAgregaServicios 
         Caption         =   "Agregar Servicios"
      End
   End
End
Attribute VB_Name = "ucDetalleFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mb_TransaccionDeNuevoRegistroEnProceso  As Boolean
Dim mb_PresionoEscape As Boolean

Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_Facturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_Comun As New SIGHNegocios.ReglasComunes
'Dim mo_ReglasServicios As New SIGHNegocios.ReglasServiciosHosp


Dim mo_cmbFechaIngreso As New SIGHComun.ListaDespleglable
Dim mo_cmbIdPuntosDeCarga As New SIGHComun.ListaDespleglable
Dim gridInfra As New GridInfragistic

Dim oCuentaAtencion As New DOCuentaAtencion
Dim oComprobantePago As New DOCajaComprobantesPago
Dim oCajaNroDocumento As New DOCajaNroDocumento
Dim oAtencion As DOAtencion

Dim ml_IdTipoComprobante As Integer
Dim ml_IdCaja As Integer
Dim ml_IdPaciente As Long
Dim ml_IdUsuario As Long
Dim ms_NombreUsuario As String
Dim ml_IdCuentaAtencion As Long
Dim ml_NroHistoriaClinica As Long
Dim ms_NombrePaciente As String

Dim mb_Invisible As Boolean
Dim mb_AgregoAtencion As Boolean
Dim me_TipoEmpleado As sghTipoEmpleado

'Dim mi_EstadoAntiguo As Integer
'Dim mi_TipoFinanciamientoAntiguo As Integer
'Dim mi_EstadoActual As Integer
''Dim mi_TipoFinanciamientoActual As Integer
Dim ml_IdPuntosDeCarga As Long

Dim mrs_FacturacionServicios As New ADODB.Recordset
Dim mrs_FacturacionBienes As New ADODB.Recordset

Dim mo_FacturacionServicios As Collection
Dim mo_FacturacionBienes  As Collection

Dim mo_FacturacionServiciosBorrar As Collection
Dim mo_FacturacionBienesBorrar  As Collection
Dim idProductoSelecto() As Long
Dim nombreProductoSelecto() As String
Dim numeroProductosSelectos As Integer


Dim ms_TipoProducto As String

Dim mb_NoEditar As Boolean


Property Let TipoEmpleado(oValue As sghTipoEmpleado)
    me_TipoEmpleado = oValue
End Property

Property Get TipoEmpleado() As sghTipoEmpleado
    TipoEmpleado = me_TipoEmpleado
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

Property Let PuntosDeCargaEnabled(bValue As Boolean)
        cmbIdPuntosDeCarga.Enabled = bValue
End Property

Property Get PuntosDeCargaEnabled() As Boolean
        PuntosDeCargaEnabled = cmbIdPuntosDeCarga.Enabled
End Property

Property Let NombrePaciente(oValue As String)
    ms_NombrePaciente = oValue
End Property
Property Get NombrePaciente() As String
    NombrePaciente = ms_NombrePaciente
End Property

Property Let NroHistoriaClinica(oValue As Long)
    ml_NroHistoriaClinica = oValue
End Property
Property Get NroHistoriaClinica() As Long
    NroHistoriaClinica = ml_NroHistoriaClinica
End Property

Property Let IdTipoComprobante(oValue As Integer)
    ml_IdTipoComprobante = oValue
End Property
Property Get IdTipoComprobante() As Integer
    IdTipoComprobante = ml_IdTipoComprobante
End Property

Property Get IdCaja() As Integer
    IdCaja = ml_IdCaja
End Property

Property Let IdCaja(oValue As Integer)
    ml_IdCaja = oValue
End Property

Property Get IdPaciente() As Long
    IdPaciente = ml_IdPaciente
End Property

Property Let IdPaciente(oValue As Long)
    ml_IdPaciente = oValue
    
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
   ObtenerNombreUsuario
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property

Property Let NombreUsuario(sValue As String)
   ms_NombreUsuario = sValue
End Property
Property Get NombreUsuario() As String
   NombreUsuario = ms_NombreUsuario
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property

Property Let IdPuntosDeCarga(oValue As Long)
    ml_IdPuntosDeCarga = oValue
    mo_cmbIdPuntosDeCarga.BoundText = oValue
End Property
Property Get IdPuntosDeCarga() As Long
    IdPuntosDeCarga = ml_IdPuntosDeCarga
    
End Property

Property Let Invisible(lValue As Boolean)
   mb_Invisible = lValue
End Property
Property Get Invisible() As Boolean
   Invisible = mb_Invisible
End Property

Public Sub Inicializar()
    cmbFechaIngreso.Clear
    Set mo_cmbFechaIngreso.MiComboBox = cmbFechaIngreso
    Set mo_cmbIdPuntosDeCarga.MiComboBox = cmbIdPuntosDeCarga
    'TipoEmpleado = sghtipoempleado.sghSOAT
    txtNroHistoria.Text = ""
    txtPaciente.Text = ""
    txtNroHistoria.Enabled = True
    txtPaciente.Enabled = False
    ConfigurarPuntosDeCarga
    Set grdBienes.DataSource = Nothing
    Set grdServicios.DataSource = Nothing
    
'    If IdPuntosDeCarga <= 0 Then
'        MsgBox "Debe ingresar desde alguna especialidad", vbExclamation, "Por favor seleccione el item adecuado"
'        Invisible = True
'    End If
End Sub
Sub ConfigurarPuntosDeCarga()
    cmbIdPuntosDeCarga.Clear
    
    'Set mo_cmbIdPuntosDeCarga = New ListaDespleglable
    
    mo_cmbIdPuntosDeCarga.ListField = "Descripcion"
    mo_cmbIdPuntosDeCarga.BoundColumn = "IdPuntoCarga"
    
    Set mo_cmbIdPuntosDeCarga.RowSource = mo_Comun.SeleccionarPuntosDeCarga

End Sub

Sub ConfigurarFechaIngreso(lIdPaciente As Long)
    
    Dim rs As New Recordset
    mo_cmbFechaIngreso.ListField = "DescripcionLarga"
    mo_cmbFechaIngreso.BoundColumn = "IdCuentaAtencion"
    
    Set rs = mo_Facturacion.CuentaAtencionSeleccionarUltimaPorIdPaciente(lIdPaciente)
    Set mo_cmbFechaIngreso.RowSource = rs
    
    If rs.RecordCount = 1 Then
        cmbFechaIngreso.ListIndex = 0
    End If

End Sub

Private Sub btnBuscar_Click()
    
    
End Sub

Private Sub btnHistoria_Click()
    Me.BusquedaPaciente
End Sub


Public Sub BusquedaPaciente()
Dim oPaciente As New doPaciente
        Dim rsRespuesta As New Recordset
        
       ' If (UserControl.txtNroHistoria = "") Then
            Dim oFrm As New PacientesBusqueda
            oFrm.TipoFiltro = sghFiltrarTodos
            oFrm.Caption = "Seleccione el empleado"
            oFrm.Show vbModal
            If oFrm.IdRegistroSeleccionado <> 0 Then
                IdPaciente = oFrm.IdRegistroSeleccionado
                 Call ObtenerNombrePaciente(oFrm.IdRegistroSeleccionado)
            End If
            Exit Sub
        'End If
            
'        oPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
'
'        Set rsRespuesta = mo_AdminAdmision.PacientesFiltrar(oPaciente)
'        On Error Resume Next
'        If rsRespuesta.RecordCount = 0 Then
'            MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
'        ElseIf rsRespuesta.RecordCount = 1 Then
'            IdPaciente = rsRespuesta!IdPaciente
'            Call ObtenerNombrePaciente(rsRespuesta!IdPaciente)
'        End If
              
End Sub

Sub ObtenerNombrePaciente(IdPaciente As Long)
    Dim oPaciente As doPaciente
    
    Set oPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(IdPaciente)
    txtPaciente.Text = oPaciente.ApellidoPaterno & " " & oPaciente.ApellidoMaterno & ", " & oPaciente.PrimerNombre & " " & oPaciente.SegundoNombre
    
End Sub
Sub ObtenerNombreUsuario()
    Dim oEmpleado As New dOEmpleado
    
    Set oEmpleado = mo_Comun.EmpleadosSeleccionarPorId(IdUsuario)
    
    NombreUsuario = oEmpleado.ApellidoPaterno & " " & oEmpleado.ApellidoMaterno & ", " & oEmpleado.Nombres
    
    
    Dim mo_seg As New SIGHNegocios.ReglasDeSeguridad
    Dim rs As New Recordset
    Dim i As Integer
    Set rs = mo_seg.UsuariosRolesSeleccionarPorEmpleado(IdUsuario)
    For i = 0 To rs.RecordCount - 1
        If rs.Fields!Nombre = "Cajero" Then
            TipoEmpleado = sghTipoEmpleado.sghCajero
            Exit For
        End If
        If rs.Fields!Nombre = "Cuenta Corriente" Then
            TipoEmpleado = sghTipoEmpleado.sghCuentaCorriente
            Exit For
        End If
        If rs.Fields!Nombre = "Operador SIS" Then
            TipoEmpleado = sghTipoEmpleado.sghSIS
            Exit For
        End If
        If rs.Fields!Nombre = "Asistente Social" Then
            TipoEmpleado = sghTipoEmpleado.sghAsistenta
            Exit For
        End If
        If rs.Fields!Nombre = "Operador Convenio" Then
            TipoEmpleado = sghTipoEmpleado.sghConvenio
            Exit For
        End If
        'If rs.Fields!Nombre = "Cajero" Then
            TipoEmpleado = sghTipoEmpleado.sghOtros
         '   Exit For
        'End If
        rs.MoveNext
    Next i
    'comentar
    'TipoEmpleado = sghtipoempleado.sghCajero
End Sub
Sub CargaDatosServicios()
    Dim rs As New Recordset
    Set mo_FacturacionServiciosBorrar = New Collection
    
    Set rs = mo_Facturacion.FacturacionServicioParaEstadoCuenta(IdCuentaAtencion, IdPuntosDeCarga, "1,2")
    Set grdServicios.DataSource = rs
End Sub
Sub CargaDatosBienes()
    Dim rs As New Recordset
    Set mo_FacturacionBienesBorrar = New Collection
    
    Set rs = mo_Facturacion.FacturacionBienesParaEstadoCuenta(IdCuentaAtencion, IdPuntosDeCarga, "1,2")
    Set grdBienes.DataSource = rs
End Sub
Private Sub cmbFechaIngreso_Change()
    IdCuentaAtencion = Val(mo_cmbFechaIngreso.BoundText)

    SeleccionarUltimaAtencion
    CargaDatosServicios
    CargaDatosBienes
End Sub

Private Sub cmbFechaIngreso_Click()
    ml_IdPuntosDeCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
    IdCuentaAtencion = Val(mo_cmbFechaIngreso.BoundText)
    SeleccionarUltimaAtencion
    CargaDatosServicios
    CargaDatosBienes
End Sub

Private Sub cmbIdPuntosDeCarga_Change()
    ml_IdPuntosDeCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
    If ml_IdPuntosDeCarga = 0 Then
        tabCuentas.TabVisible(0) = True 'servicios
        grdServicios.Visible = True
        tabCuentas.TabVisible(1) = True 'bienes
        grdBienes.Visible = True
        Exit Sub
    End If
End Sub

Private Sub cmbIdPuntosDeCarga_Click()
    Dim texto As String
    ml_IdPuntosDeCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
    If ml_IdPuntosDeCarga = 0 Then
        tabCuentas.TabVisible(0) = True 'servicios
        grdServicios.Visible = True
        tabCuentas.TabVisible(1) = True 'bienes
        grdBienes.Visible = True
        Exit Sub
    End If
    texto = Trim(cmbIdPuntosDeCarga.Text)
    texto = Trim(Mid(texto, InStr(texto, "=") + 1))
    If LCase(texto) = "farmacia" Then
        tabCuentas.TabVisible(0) = False 'servicios
        grdServicios.Visible = False
        tabCuentas.TabVisible(1) = True 'bienes
        grdBienes.Visible = True
        'grdBienes.ZOrder 1
    Else
        tabCuentas.TabVisible(0) = True 'servicios
        grdServicios.Visible = True
        tabCuentas.TabVisible(1) = False 'bienes
        grdBienes.Visible = False
    End If
    
    ml_IdPuntosDeCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
    IdCuentaAtencion = Val(mo_cmbFechaIngreso.BoundText)
    If IdCuentaAtencion > 0 Then
        SeleccionarUltimaAtencion
        CargaDatosServicios
        CargaDatosBienes
    End If

End Sub

Private Sub cmdGrabar_Click()
    
    If IdCuentaAtencion <= 0 Then
        Exit Sub
    End If
    If MsgBox("Por favor confirmar, ¿Realmente desea grabar los cambios que ha realizado?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbNo Then
        Exit Sub
    End If
    CargaDatosAlObjetosDeDatos
    If ValidarReglas() Then
        If ModificarDatos() Then
             MsgBox "Los datos se modificaron correctamente", vbInformation, "Estado de Cuenta"
             cmdNuevo_Click
         Else
             MsgBox "No se pudo agregar los datos" + Chr(13) + mo_Facturacion.MensajeError, vbExclamation, "Estado de Cuenta"
        End If
    End If
    
End Sub
Function ValidarReglas() As Boolean
    ValidarReglas = False
    
    
    ValidarReglas = True
End Function
Function ModificarDatos() As Boolean
    ModificarDatos = mo_Facturacion.ActualizarServiciosYBienes(mo_FacturacionServicios, mo_FacturacionBienes, mo_FacturacionServiciosBorrar, mo_FacturacionBienesBorrar)
End Function

Private Sub cmdNuevo_Click()
    
    txtNroHistoria.Text = ""
    txtPaciente.Text = ""
    cmbFechaIngreso.Clear
    
    IdCuentaAtencion = 0
    Set grdBienes.DataSource = Nothing
    Set grdServicios.DataSource = Nothing
    Me.ConfigurarPuntosDeCarga
    
    If IdPuntosDeCarga > 0 Then
        Call mo_cmbIdPuntosDeCarga.UbicarItemDeComboBoxPorId(cmbIdPuntosDeCarga, IdPuntosDeCarga)
    End If
    txtNroHistoria.SetFocus
    
End Sub

Private Sub cmdSalir_Click()
    IdCuentaAtencion = Val(mo_cmbFechaIngreso.BoundText)
    SeleccionarUltimaAtencion
    CargaDatosServicios
    CargaDatosBienes
    
End Sub

Sub SeleccionarUltimaAtencion()
    Set oAtencion = mo_Facturacion.SeleccionarUltimaAtencion(IdPaciente, IdCuentaAtencion)
End Sub

Sub CargaDatosAlObjetosDeDatos()
On Error GoTo errDescription
    
    grdServicios.Refresh ssRefetchAndFireInitializeRow
    grdBienes.Refresh ssRefetchAndFireInitializeRow
    Set mrs_FacturacionServicios = grdServicios.DataSource
    Set mrs_FacturacionBienes = grdBienes.DataSource
    
    Set mo_FacturacionServicios = New Collection
    Set mo_FacturacionBienes = New Collection
    
    Dim oDOFacturacionServicios As DOFacturacionServicios
    Dim odoFacturacionBienesInsumos As DOFacturacionBienesInsumos
    
    If Not (mrs_FacturacionServicios.EOF And mrs_FacturacionServicios.BOF) Then
        mrs_FacturacionServicios.MoveFirst
        Do While Not mrs_FacturacionServicios.EOF
            'If mrs_FacturacionServicios!EstadoRegistro = "M" Then
            If mrs_FacturacionServicios!Id <= 0 Then
                Set oDOFacturacionServicios = New DOFacturacionServicios
                oDOFacturacionServicios.IdAtencion = oAtencion.IdAtencion
                oDOFacturacionServicios.IdProducto = mrs_FacturacionServicios!IdProducto
                oDOFacturacionServicios.PrecioUnitario = mrs_FacturacionServicios!PrecioUnitario
                oDOFacturacionServicios.cantidad = mrs_FacturacionServicios!cantidad
                oDOFacturacionServicios.TotalPorPagar = oDOFacturacionServicios.cantidad * oDOFacturacionServicios.PrecioUnitario
                If (TipoEmpleado = sghTipoEmpleado.sghConvenio) Then
                    oDOFacturacionServicios.IdFuenteFinanciamiento = 6 'debe haber una querie que saque segun tipo de paciente
                ElseIf (TipoEmpleado = sghTipoEmpleado.sghSIS) Then
                    oDOFacturacionServicios.IdFuenteFinanciamiento = 9
                ElseIf (TipoEmpleado = sghTipoEmpleado.sghSOAT) Then
                    oDOFacturacionServicios.IdFuenteFinanciamiento = 2
                Else
                    oDOFacturacionServicios.IdFuenteFinanciamiento = 1
                End If
            Else
                'Set oDOFacturacionServicios = mo_Facturacion.FacturacionServiciosSeleccionarPorId(Val(mrs_FacturacionServicios!Id))  'New DOFacturacionServicios
            End If
                
                oDOFacturacionServicios.IdTipoFinanciamiento = mrs_FacturacionServicios!IdTipoFinanciamiento
'                If mrs_FacturacionServicios!IdTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal Then
'                    'oDOFacturacionServicios.
'                End If
                If oDOFacturacionServicios.IdAtencion <= 0 Then
                    oDOFacturacionServicios.IdAtencion = oAtencion.IdAtencion
                End If
                oDOFacturacionServicios.IdEstadoFacturacion = mrs_FacturacionServicios!IdEstadoFacturacion
                oDOFacturacionServicios.IdUsuarioAuditoria = Me.IdUsuario
                If oDOFacturacionServicios.IdPuntoCarga <= 0 Then
                    oDOFacturacionServicios.IdPuntoCarga = IdPuntosDeCarga
                End If
                
                mo_FacturacionServicios.Add oDOFacturacionServicios
            
            'End If
            
            mrs_FacturacionServicios.MoveNext
        Loop
        'mrs_FacturacionServicios.MoveFirst
    End If
    
    If Not (mrs_FacturacionBienes.EOF And mrs_FacturacionBienes.BOF) Then
        mrs_FacturacionBienes.MoveFirst
        Do While Not mrs_FacturacionBienes.EOF
            'If mrs_FacturacionBienes!EstadoRegistro = "M" Then
            If mrs_FacturacionBienes!Id <= 0 Then
                Set odoFacturacionBienesInsumos = New DOFacturacionBienesInsumos
                odoFacturacionBienesInsumos.IdAtencion = oAtencion.IdAtencion
                odoFacturacionBienesInsumos.IdProducto = mrs_FacturacionBienes!IdProducto
                odoFacturacionBienesInsumos.PrecioUnitario = mrs_FacturacionBienes!PrecioUnitario
                odoFacturacionBienesInsumos.PrecioUnitario = mrs_FacturacionBienes!PrecioUnitario
                odoFacturacionBienesInsumos.cantidad = mrs_FacturacionBienes!cantidad
                odoFacturacionBienesInsumos.TotalPorPagar = odoFacturacionBienesInsumos.cantidad * odoFacturacionBienesInsumos.PrecioUnitario
                
                If (TipoEmpleado = sghTipoEmpleado.sghConvenio) Then
                    odoFacturacionBienesInsumos.IdFuenteFinanciamiento = 6 'debe haber una querie que saque segun tipo de paciente
                ElseIf (TipoEmpleado = sghTipoEmpleado.sghSIS) Then
                    odoFacturacionBienesInsumos.IdFuenteFinanciamiento = 9
                ElseIf (TipoEmpleado = sghTipoEmpleado.sghSOAT) Then
                    odoFacturacionBienesInsumos.IdFuenteFinanciamiento = 2
                Else
                    odoFacturacionBienesInsumos.IdFuenteFinanciamiento = 1
                End If
            Else
                Set odoFacturacionBienesInsumos = mo_Facturacion.FacturacionBienesInsumosSeleccionarPorId(Val(mrs_FacturacionBienes!Id))
            End If
                'odoFacturacionBienesInsumos.IdFacturacionBienes = mrs_FacturacionBienes!id
                odoFacturacionBienesInsumos.IdTipoFinanciamiento = mrs_FacturacionBienes!IdTipoFinanciamiento
                If mrs_FacturacionBienes!IdTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal Then
                    'oDOFacturacionServicios.
                End If
                If odoFacturacionBienesInsumos.IdAtencion <= 0 Then
                    odoFacturacionBienesInsumos.IdAtencion = oAtencion.IdAtencion
                End If
                odoFacturacionBienesInsumos.IdEstadoFacturacion = mrs_FacturacionBienes!IdEstadoFacturacion
                odoFacturacionBienesInsumos.IdUsuarioAuditoria = Me.IdUsuario
                If odoFacturacionBienesInsumos.IdPuntoCarga <= 0 Then
                    odoFacturacionBienesInsumos.IdPuntoCarga = IdPuntosDeCarga
                End If
                mo_FacturacionBienes.Add odoFacturacionBienesInsumos
            'End If
            mrs_FacturacionBienes.MoveNext
        Loop
        'mrs_FacturacionBienes.MoveFirst
    End If

Exit Sub
errDescription:
    Set mo_FacturacionServicios = New Collection
    Set mo_FacturacionBienes = New Collection
    
End Sub

Private Sub grdBienes_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If (grdBienes.ActiveRow Is Nothing) Then
        Exit Sub
    End If
    If Cell.Column.Key = "Cantidad" Then
        Cell.Row.Cells("subtotal").Value = Cell.Row.Cells("preciounitario").Value * Cell.Value
    End If
    
End Sub

'bienes
'Private Sub grdBienes_AfterRowActivate()
'    If (grdBienes.ActiveRow Is Nothing) Then
'        Exit Sub
'    End If
'    If grdBienes.ActiveRow.Cells("idestadofacturacion").GetText() = "" Or grdBienes.ActiveRow.Cells("idtipofinanciamiento").GetText() = "" Then
'        Exit Sub
'    End If
'        mi_EstadoAntiguo = ObtenerValorEnLista(grdBienes.ActiveRow.Cells("idestadofacturacion").Column, grdBienes.ActiveRow.Cells("idestadofacturacion").Value)
'        mi_TipoFinanciamientoAntiguo = ObtenerValorEnLista(grdBienes.ActiveRow.Cells("idtipofinanciamiento").Column, grdBienes.ActiveRow.Cells("idtipofinanciamiento").Value)
'End Sub

Private Sub grdBienes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    Dim cantidad As Integer
    Dim precio As Double
    Dim subtotal As Double

    If (grdBienes.ActiveRow Is Nothing) Then
        Exit Sub
    End If
    If grdBienes.ActiveCell Is Nothing Then
        Exit Sub
    End If
    If grdBienes.ActiveCell.Column.Key = "Codigo" Then
        ms_TipoProducto = "bienes"
        SeteaProducto grdBienes.ActiveCell.Value
    End If
    If (grdBienes.ActiveCell.Column.Key = "cantidad") Then
        cantidad = grdBienes.ActiveRow.Cells("cantidad").Value
        precio = grdBienes.ActiveRow.Cells("preciounitario").Value
        subtotal = precio * cantidad
        grdBienes.ActiveRow.Cells("totalporpagar").Value = subtotal
    End If
      
End Sub

Private Sub grdBienes_BeforeRowDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)

    Set grillaBusqueda.DataSource = Nothing
    grillaBusqueda.Visible = False
    
    If grdBienes.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    If grdBienes.ActiveRow.Cells("idproducto") = 0 Then
       If Not mb_NoEditar Then
            mb_NoEditar = True
            Set grillaBusqueda.DataSource = Nothing
            grillaBusqueda.Visible = False
            grdBienes.ActiveRow.Delete
            mb_NoEditar = False
       End If
    End If
End Sub

'Private Sub grdBienes_BeforeRowUpdate(ByVal Row As UltraGrid.SSRow, ByVal Cancel As UltraGrid.SSReturnBoolean)
'    Dim tipoFinanc As Integer
'    Dim estado As Integer
'    Dim poliza As String
'    If grdBienes.ActiveRow Is Nothing Then
'        Exit Sub
'    End If
'    If grdBienes.ActiveRow.Cells("idestadofacturacion").GetText() = "" Or grdBienes.ActiveRow.Cells("idtipofinanciamiento").GetText() = "" Then
'        Exit Sub
'    End If
'    tipoFinanc = Me.ObtenerValorEnLista(grdBienes.ActiveRow.Cells("IdTipoFinanciamiento").Column, grdBienes.ActiveRow.Cells("IdTipoFinanciamiento").Value)
'    estado = ObtenerValorEnLista(grdBienes.ActiveRow.Cells("IdEstadoFacturacion").Column, grdBienes.ActiveRow.Cells("IdEstadoFacturacion").Value)
''    If grdBienes.ActiveRow.Cells("Poliza").GetText() <> "" Then
''        poliza = grdBienes.ActiveRow.Cells("Poliza").Value
''    End If
'
'    If Not ValidaAccesos(TipoEmpleado, tipoFinanc, estado) Then
'        grdBienes.ActiveRow.Cells("IdTipoFinanciamiento").Value = mi_TipoFinanciamientoAntiguo
'        grdBienes.ActiveRow.Cells("IdEstadoFacturacion").Value = mi_EstadoAntiguo
'        Cancel = True
'    End If
'
'End Sub

Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrilla grdBienes
End Sub

Private Sub grdBienes_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
 Dim sNombre As String
If grdBienes.ActiveRow Is Nothing Then
        Exit Sub
    End If
    If grdBienes.ActiveCell Is Nothing Then
        Exit Sub
    End If
    If grdBienes.ActiveCell Is Nothing Then
        Exit Sub
    End If
    
    'If grdBienes.ActiveCell.Row.Cells("id").Value = 0 Then
     'If mb_TransaccionDeNuevoRegistroEnProceso Then
        
        mb_PresionoEscape = False
        
        'El cajero presiono ESCAPE
        If KeyAscii = vbKeyEscape Then
            mb_PresionoEscape = True
                grillaBusqueda.Visible = False
                Set grillaBusqueda.DataSource = Nothing
                'If grdBienes.ActiveCell.Row.Cells("id").Value = 0 Then
                mb_NoEditar = True
                grdBienes.ActiveRow.Delete
                mb_NoEditar = False
                'End If
            Exit Sub
        End If
        
        'El cajero esta editando la parte de la DESCRIPCION
        If grdBienes.ActiveCell.Row.Cells("id").Value <> 0 Then
            If grdBienes.ActiveCell.Column.Key <> "Cantidad" Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
        If grdBienes.ActiveCell.Column.Key = "Descripcion" Then
            ms_TipoProducto = "bienes"
            Select Case KeyAscii
            Case 8
                'El cajero ha presionado BACKSPACE
                sNombre = grdBienes.ActiveCell.GetText
                If Len(sNombre) > 1 Then
                    sNombre = Mid(sNombre, 1, Len(sNombre) - 1)
                End If
            Case 13, 9, 10
            Case Else
                sNombre = grdBienes.ActiveCell.GetText + Chr(KeyAscii)
            End Select
            
            Dim lIdTipoFinanciamiento As Long
            Dim lIdPuntoCarga As Long
            
            If TipoEmpleado = sghTipoEmpleado.sghConvenio Then
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghConvenios)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSIS Then
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSIS)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSOAT)
            Else
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghPacienteNormal)
            End If

            lIdPuntoCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
            BuscaProductos sNombre, lIdTipoFinanciamiento, lIdPuntoCarga
        End If
    
    'End If
        
End Sub

Private Sub grdBienes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuBienes
    End If
End Sub


Function SeteaProducto(codigo As String) As Boolean
'EFGL 14/06/2006
    Dim rs As New ADODB.Recordset
    SeteaProducto = False
    Dim IdTipoFinanciamiento  As Long
    
            If TipoEmpleado = sghTipoEmpleado.sghConvenio Then
                IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghConvenios)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSIS Then
                IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSIS)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
                IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSOAT)
            Else
                IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghPacienteNormal)
            End If
    
    If ms_TipoProducto = "bienes" Then
        Set rs = mo_Facturacion.FacturacionBienesPorCodigo(codigo, IdTipoFinanciamiento)
        If rs.RecordCount = 1 Then
           grdBienes.ActiveRow.Cells("codigo").Value = rs.Fields("CODIGO").Value
           grdBienes.ActiveRow.Cells("idproducto").Value = rs.Fields("idproducto").Value
           grdBienes.ActiveRow.Cells("descripcion").Value = rs.Fields("descripcion").Value
           grdBienes.ActiveRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
           grdBienes.ActiveRow.Cells("subtotal").Value = rs.Fields("preciounitario").Value
           grdBienes.ActiveRow.Cells("cantidad").Value = 1
           grdBienes.ActiveRow.Cells("idestadofacturacion").Value = 1
           grdBienes.ActiveRow.Cells("idtipofinanciamiento").Value = IdTipoFinanciamiento
           SeteaProducto = True
        End If
    Else
        Set rs = mo_Facturacion.FacturacionServicioPorCodigo(codigo, IdTipoFinanciamiento)
        If rs.RecordCount = 1 Then
            grdServicios.ActiveRow.Cells("codigo").Value = rs.Fields("CODIGO").Value
           grdServicios.ActiveRow.Cells("idproducto").Value = rs.Fields("idproducto").Value
           grdServicios.ActiveRow.Cells("descripcion").Value = rs.Fields("descripcion").Value
           grdServicios.ActiveRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
           grdServicios.ActiveRow.Cells("subtotal").Value = rs.Fields("preciounitario").Value
           grdServicios.ActiveRow.Cells("cantidad").Value = 1
           grdServicios.ActiveRow.Cells("idestadofacturacion").Value = 1
           grdServicios.ActiveRow.Cells("idtipofinanciamiento").Value = IdTipoFinanciamiento
           SeteaProducto = True
        End If
        
    End If
    grillaBusqueda.Visible = False
    Set grillaBusqueda.DataSource = Nothing
'EFGL 14/06/2006
End Function

Sub BuscaProductos(sNombre As String, lIdTipoFinanciamiento As Long, lIdPuntoCarga As Long)
    Dim rs As New Recordset
    
    If ms_TipoProducto = "servicios" Then
        grillaBusqueda.Left = grdServicios.Left
        Set rs = mo_AdminCaja.ServiciosFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, lIdPuntoCarga)
        grillaBusqueda.Top = grdServicios.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY + 500
    Else
        grillaBusqueda.Left = grdBienes.Left
        Set rs = mo_AdminCaja.BienesFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, lIdPuntoCarga)
        grillaBusqueda.Top = grdBienes.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY + 500
    End If
    
    Set grillaBusqueda.DataSource = rs
    grillaBusqueda.Visible = True
    grillaBusqueda.Enabled = True
    
End Sub
'servicios
Private Sub grdBienes_AfterRowsDeleted()
On Error GoTo errDescription
If numeroProductosSelectos <= 0 Then
    Exit Sub
End If
Dim oFacturacionBienes As New DOFacturacionBienesInsumos
Dim i As Integer
For i = 0 To numeroProductosSelectos - 1
    Set oFacturacionBienes = New DOFacturacionBienesInsumos
    oFacturacionBienes.IdFacturacionBienes = idProductoSelecto(i)
    oFacturacionBienes.IdUsuarioAuditoria = IdUsuario
    mo_FacturacionBienesBorrar.Add oFacturacionBienes
Next
numeroProductosSelectos = 0

errDescription:

End Sub

Private Sub grdServicios_AfterRowActivate()
    
    If (grdServicios.ActiveRow Is Nothing) Then
        Exit Sub
    End If
    
    If grdServicios.ActiveRow.Cells("idestadofacturacion").GetText() = "" Or grdServicios.ActiveRow.Cells("idtipofinanciamiento").GetText() = "" Then
        Exit Sub
    End If
    
End Sub
Private Sub grdBienes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    
    If (Cancel.Value = True) Then
        Exit Sub
    End If
   If (grdBienes.Selected Is Nothing) Then
        Exit Sub
    End If
    If (grdBienes.Selected.Rows Is Nothing) Then
        Exit Sub
    End If
    If (grdBienes.Selected.Rows.Count <= 0) Then
        Exit Sub
    End If
    Dim i As Integer
     
    ReDim idProductoSelecto(grdBienes.Selected.Rows.Count)
    ReDim nombreProductoSelecto(grdBienes.Selected.Rows.Count)
    numeroProductosSelectos = grdBienes.Selected.Rows.Count
    For i = 0 To grdBienes.Selected.Rows.Count - 1
        idProductoSelecto(i) = grdBienes.Selected.Rows(i).Cells("id").Value
        nombreProductoSelecto(i) = grdBienes.Selected.Rows(i).Cells("descripcion").Value
    Next
End Sub
Private Sub grdServicios_AfterRowsDeleted()
On Error GoTo errDescription
If numeroProductosSelectos <= 0 Then
    Exit Sub
End If
Dim oFacturacionServicios As New DOFacturacionServicios
Dim i As Integer
For i = 0 To numeroProductosSelectos - 1
    Set oFacturacionServicios = New DOFacturacionServicios
    oFacturacionServicios.IdFacturacionServicio = idProductoSelecto(i)
    oFacturacionServicios.IdUsuarioAuditoria = IdUsuario
    mo_FacturacionServiciosBorrar.Add oFacturacionServicios
Next
numeroProductosSelectos = 0
errDescription:

End Sub

Private Sub grdServicios_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    Dim precio As Double
    Dim cantidad As Double
    Dim subtotal As Double
    If (grdServicios.ActiveRow Is Nothing) Then
        Exit Sub
    End If
    If grdServicios.ActiveCell Is Nothing Then
        Exit Sub
    End If
    If grdServicios.ActiveCell.Column.Key = "Codigo" Then
        ms_TipoProducto = "servicios"
        SeteaProducto grdServicios.ActiveCell.Value
    End If
    If (grdServicios.ActiveCell.Column.Key = "Cantidad") Then
        cantidad = grdServicios.ActiveRow.Cells("Cantidad").Value
        precio = grdServicios.ActiveRow.Cells("preciounitario").Value
        subtotal = precio * cantidad
        grdServicios.ActiveRow.Cells("subtotal").Value = subtotal
    End If
    
End Sub

Private Sub grdServicios_BeforeRowDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    
    Set grillaBusqueda.DataSource = Nothing
    grillaBusqueda.Visible = False
       
    If grdServicios.ActiveRow Is Nothing Then
        Exit Sub
    End If

    If grdServicios.ActiveRow.Cells("idproducto") = 0 Then
        If Not mb_NoEditar Then
            mb_NoEditar = True
            grillaBusqueda.Visible = False
            Set grillaBusqueda.DataSource = Nothing
            grdServicios.ActiveRow.Delete
            mb_NoEditar = False
       End If
    End If
    mb_TransaccionDeNuevoRegistroEnProceso = False
    
End Sub
'Private Sub grdServicios_BeforeRowUpdate(ByVal Row As UltraGrid.SSRow, ByVal Cancel As UltraGrid.SSReturnBoolean)
'Dim tipoFinanc As Integer
'    Dim estado As Integer
'    Dim poliza As String
'    If grdServicios.ActiveRow Is Nothing Then
'        Exit Sub
'    End If
'
'    If grdServicios.ActiveRow.Cells("idestadofacturacion").GetText() = "" Or grdServicios.ActiveRow.Cells("idtipofinanciamiento").GetText() = "" Then
'        Exit Sub
'    End If
'
'    tipoFinanc = Me.ObtenerValorEnLista(grdServicios.ActiveRow.Cells("IdTipoFinanciamiento").Column, grdServicios.ActiveRow.Cells("IdTipoFinanciamiento").Value)
'
'    estado = ObtenerValorEnLista(grdServicios.ActiveRow.Cells("IdEstadoFacturacion").Column, grdServicios.ActiveRow.Cells("IdEstadoFacturacion").Value)
''    If grdServicios.ActiveRow.Cells("poliza").GetText() <> "" Then
''        poliza = grdServicios.ActiveRow.Cells("Poliza").Value
''    End If
'
'    If Not ValidaAccesos(TipoEmpleado, tipoFinanc, estado) Then
'
'        grdServicios.ActiveRow.Cells("IdTipoFinanciamiento").Value = mi_TipoFinanciamientoAntiguo
'
'        grdServicios.ActiveRow.Cells("IdEstadoFacturacion").Value = mi_EstadoAntiguo
'
'        Cancel = True
'
'    End If
'End Sub

Private Sub grdServicios_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If (Cancel.Value = True) Then
        Exit Sub
    End If
   
   If (grdServicios.Selected Is Nothing) Then
        Exit Sub
    End If
    If (grdServicios.Selected.Rows Is Nothing) Then
        Exit Sub
    End If
    If (grdServicios.Selected.Rows.Count <= 0) Then
        Exit Sub
    End If
    Dim i As Integer
     
    
    ReDim idProductoSelecto(grdServicios.Selected.Rows.Count)
    ReDim nombreProductoSelecto(grdServicios.Selected.Rows.Count)
    numeroProductosSelectos = grdServicios.Selected.Rows.Count
    For i = 0 To grdServicios.Selected.Rows.Count - 1
        idProductoSelecto(i) = grdServicios.Selected.Rows(i).Cells("id").Value
        nombreProductoSelecto(i) = grdServicios.Selected.Rows(i).Cells("descripcion").Value
    Next

End Sub

Private Sub grdServicios_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
Dim sNombre As String

    If grdServicios.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    If grdServicios.ActiveCell Is Nothing Then
        Exit Sub
    End If
   
     'If mb_TransaccionDeNuevoRegistroEnProceso Then
        
        mb_PresionoEscape = False
        
        'El cajero presiono ESCAPE
        If KeyAscii = vbKeyEscape Then
            mb_PresionoEscape = True
                grillaBusqueda.Visible = False
                Set grillaBusqueda.DataSource = Nothing
            'If grdServicios.ActiveCell.Row.Cells("id").Value = 0 Then
                mb_NoEditar = True
                grdServicios.ActiveRow.Delete
                mb_NoEditar = False
            'End If
            Exit Sub
        End If
        
        'El cajero esta editando la parte de la DESCRIPCION
        If grdServicios.ActiveCell.Row.Cells("id").Value <> 0 Then
            If grdServicios.ActiveCell.Column.Key <> "Cantidad" Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
        If grdServicios.ActiveCell.Column.Key = "Descripcion" Then
            ms_TipoProducto = "servicios"
            Select Case KeyAscii
            Case 8
                'El cajero ha presionado BACKSPACE
                sNombre = grdServicios.ActiveCell.GetText
                If Len(sNombre) > 1 Then
                    sNombre = Mid(sNombre, 1, Len(sNombre) - 1)
                End If
            Case 13, 9, 10
            Case Else
                sNombre = grdServicios.ActiveCell.GetText + Chr(KeyAscii)
            End Select
            
            Dim lIdTipoFinanciamiento As Long
            Dim lIdPuntoCarga As Long
            
            If TipoEmpleado = sghTipoEmpleado.sghConvenio Then
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghConvenios)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSIS Then
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSIS)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSOAT)
            Else
                lIdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghPacienteNormal)
            End If
            lIdPuntoCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
            
            BuscaProductos sNombre, lIdTipoFinanciamiento, lIdPuntoCarga
        End If
    
    'End If
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrilla grdServicios
End Sub

Private Sub grdServicios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuServicios
    End If
End Sub

Function ObtenerValorEnLista(oColumna As SSColumn, descripcion As String) As Integer
    Dim i As Integer
    Dim descripcionValue As String
    ObtenerValorEnLista = Val(descripcion)
    If IsNumeric(descripcion) Then
        Exit Function
    End If
    
    For i = 0 To oColumna.ValueList.ValueListItems.Count
        descripcionValue = oColumna.ValueList.ValueListItems(i).DisplayText
        If (descripcionValue = descripcion) Then
            ObtenerValorEnLista = Val(oColumna.ValueList.ValueListItems(i).DataValue)
            Exit For
        End If
    Next i
End Function

Private Sub grillaBusqueda_DblClick()
    Dim fila As New Record
'    If Not grillaBusqueda.ActiveCell Is Nothing Then
'       Set fila.Source = grillaBusqueda.ActiveCell.Row
'       Exit Sub
'    End If
    If Not grillaBusqueda.ActiveRow Is Nothing Then
        If ms_TipoProducto = "bienes" Then
           grdBienes.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
           grdBienes.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
           grdBienes.ActiveRow.Cells("descripcion").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
           grdBienes.ActiveRow.Cells("preciounitario").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdBienes.ActiveRow.Cells("subtotal").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdBienes.ActiveRow.Cells("cantidad").Value = 1
           grdBienes.ActiveRow.Cells("idestadofacturacion").Value = 1
            If TipoEmpleado = sghTipoEmpleado.sghConvenio Then
                grdBienes.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghConvenios)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSIS Then
                grdBienes.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghSIS)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
                grdBienes.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghSOAT)
            Else
                grdBienes.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghPacienteNormal)
            End If
        Else
           grdServicios.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
           grdServicios.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
           grdServicios.ActiveRow.Cells("descripcion").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
           grdServicios.ActiveRow.Cells("preciounitario").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdServicios.ActiveRow.Cells("subtotal").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdServicios.ActiveRow.Cells("cantidad").Value = 1
           grdServicios.ActiveRow.Cells("idestadofacturacion").Value = 1
           If TipoEmpleado = sghTipoEmpleado.sghConvenio Then
                grdServicios.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghConvenios)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSIS Then
                grdServicios.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghSIS)
            ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
                grdServicios.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghSOAT)
            Else
                grdServicios.ActiveRow.Cells("idtipofinanciamiento").Value = Val(sghTipoFinanciamiento.sghPacienteNormal)
            End If
        End If
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
        Exit Sub
    End If
End Sub

Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, SIGHComun.GrillaConFilasBicolor
End Sub
Private Sub FormatoGrillaBusqueda(oGrilla As SSUltraGrid)
   
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub FormatoGrilla(oGrilla As SSUltraGrid)
    
    oGrilla.Override.AllowUpdate = ssAllowUpdateYes
     
    Dim oColumnPoliza As SSColumn
    Dim oColumnTipoFinanciamiento As SSColumn
    Dim oColumnEstado As SSColumn
    
    oGrilla.Bands(0).Columns("id").Hidden = True
    oGrilla.Bands(0).Columns("idProducto").Hidden = True
    oGrilla.Bands(0).Columns("Estado").Hidden = True
    oGrilla.Bands(0).Columns("TipoFinanciamiento").Hidden = True
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    oGrilla.Bands(0).Columns("Descripcion").Header.Caption = "Descripcion"
    oGrilla.Bands(0).Columns("Descripcion").Width = 4000
    
    Set oColumnPoliza = oGrilla.Bands(0).Columns("poliza")
    oColumnPoliza.Header.Caption = "Poliza"
    oColumnPoliza.Width = 1000
    oColumnPoliza.Hidden = True
    
    Set oColumnTipoFinanciamiento = oGrilla.Bands(0).Columns("IdTipoFinanciamiento")
    oColumnTipoFinanciamiento.Width = 3000
    oColumnTipoFinanciamiento.Header.Caption = "Tipo Financiamiento"
    oColumnTipoFinanciamiento.Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    oGrilla.Bands(0).Columns("Cantidad").Format = "#0.00"
    'oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("preciounitario").Header.Caption = "P.U.(s/.)"
    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("preciounitario").Format = "#0.00"
    oGrilla.Bands(0).Columns("subtotal").Header.Caption = "Subtotal"
    oGrilla.Bands(0).Columns("subtotal").Format = "#0.00"
    oGrilla.Bands(0).Columns("subtotal").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("IdPuntoCarga").Hidden = True
  
    Set oColumnEstado = oGrilla.Bands(0).Columns("idEstadoFacturacion")
    oColumnEstado.Width = 2500
    oColumnEstado.Header.Caption = "Estado"
    oColumnEstado.Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHComun.GrillaConFilasBicolor
    
    oColumnPoliza.Style = ssStyleDropDown
    oColumnTipoFinanciamiento.Style = ssStyleDropDownList
    oColumnEstado.Style = ssStyleDropDownList
    
    Set oColumnTipoFinanciamiento.ValueList = oGrilla.ValueLists("listaTipoFinanciamiento")
    Set oColumnEstado.ValueList = oGrilla.ValueLists("listaEstado")
    
    SeteaListaEstado oColumnEstado
    SeteaListaTipoFinanciamiento oColumnTipoFinanciamiento
    
End Sub

'Sub SeteaListaPoliza(oColumn As SSColumn)
'    Dim i As Integer
'     Dim rs As New Recordset
'
'     Set rs = grdPolizas.DataSource
'    If rs Is Nothing Then
'        Exit Sub
'    End If
'    If Not (rs.EOF And rs.BOF) Then
'        rs.MoveFirst
'    End If
'    For i = 0 To rs.RecordCount - 1
'         oColumn.ValueList.ValueListItems.Add Val(rs.Fields!poliza), Trim(rs.Fields!poliza)
'        rs.MoveNext
'    Next i
'
'
'End Sub

Sub SeteaListaTipoFinanciamiento(oColumn As SSColumn)
    Dim rs As New ADODB.Recordset
    Dim i As Integer
     
    Set rs = mo_Facturacion.TiposFinanciamientoSeleccionarTodos

    oColumn.ValueList.ValueListItems.Clear
    For i = 0 To rs.RecordCount - 1
        If rs.Fields!IdTipoFinanciamiento <> 0 Then
            oColumn.ValueList.ValueListItems.Add Val(rs.Fields!IdTipoFinanciamiento), Trim(rs.Fields!descripcion)
        End If
        rs.MoveNext
        
    Next i
    rs.Close
End Sub

Sub SeteaListaEstado(oColumn As SSColumn)
    Dim rs As New ADODB.Recordset
    Dim i As Integer
     
    Set rs = mo_Facturacion.EstadosFacturacionObtenerTodos
    oColumn.ValueList.ValueListItems.Clear
    For i = 0 To rs.RecordCount - 1
        oColumn.ValueList.ValueListItems.Add Val(rs.Fields!IdEstadoFacturacion), Trim(rs.Fields!descripcion)
        rs.MoveNext
    Next i
    rs.Close
End Sub

Private Sub mnuAgregaBienes_Click()

'If TipoEmpleado <> sghtipoempleado.sghSIS And TipoEmpleado <> sghtipoempleado.sghConvenio And TipoEmpleado <> sghtipoempleado.sghSOAT Then
'    MsgBox "Ud no puede agregar un Bien o Insumo", vbInformation, "Estado de Cuenta"
'    Exit Sub
'End If

Me.AgregaBienesInsumos

End Sub

Private Sub mnuAgregaServicios_Click()
'If TipoEmpleado <> sghtipoempleado.sghSIS And TipoEmpleado <> sghtipoempleado.sghConvenio And TipoEmpleado <> sghtipoempleado.sghSOAT Then
'    MsgBox "Ud no puede agregar un Servicio", vbInformation, "Estado de Cuenta"
'    Exit Sub
'End If

Me.AgregaServicios

End Sub

Sub AgregaBienesInsumos()
'    If oCuentaAtencion Is Nothing Then
'        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes e Insumos"
'        Exit Sub
'    End If
    If IdCuentaAtencion <= 0 Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes e Insumos"
        Exit Sub
    End If
    If IdPuntosDeCarga = 0 Then
        MsgBox "Debe seleccionar una Especialidad", vbCritical, "Filtro Bienes e Insumos"
        Exit Sub
    End If
    
'    If (Not AgregarNuevaAtencion) Then
'        Exit Sub
'    End If
    grdBienes.SetFocus
    Set mrs_FacturacionBienes = grdBienes.DataSource
    
    With mrs_FacturacionBienes
        .AddNew
        .Fields!IdProducto = 0
        .Fields!codigo = ""
        .Fields!descripcion = ""
        .Fields!cantidad = 1
        .Fields!PrecioUnitario = 0
        .Fields!subtotal = 0
        .Fields!IdEstadoFacturacion = 1
        '.Fields!IdAtencion = oAtencion.IdAtencion
        If TipoEmpleado = sghTipoEmpleado.sghSIS Then
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSIS)
        ElseIf TipoEmpleado = sghTipoEmpleado.sghConvenio Then
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghConvenios)
        ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSOAT)
        Else
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghPacienteNormal)
        End If
        '.Fields!poliza = ""
    End With
    mb_NoEditar = True
        Set grdBienes.DataSource = mrs_FacturacionBienes
    mb_NoEditar = False
    mb_TransaccionDeNuevoRegistroEnProceso = True
    grdBienes.PerformAction ssKeyActionLastRowInGrid
    grdBienes.ActiveRow.Activation = ssActivationAllowEdit
    grdBienes.ActiveCell = grdBienes.ActiveRow.Cells("codigo")
End Sub


Sub AgregaServicios()
'    If oCuentaAtencion Is Nothing Then
'        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes e Insumos"
'        Exit Sub
'    End If
    If IdCuentaAtencion <= 0 Then
        MsgBox "Por favor ingrese la historia clínica del paciente", vbInformation, "Agregar servicios"
        Exit Sub
    End If
    If IdPuntosDeCarga = 0 Then
        MsgBox "Por favor ingrese el punto de carga", vbInformation, "Agregar servicios"
        Exit Sub
    End If
    
'    If (Not AgregarNuevaAtencion) Then
'        Exit Sub
'    End If
    grdServicios.SetFocus
    Set mrs_FacturacionServicios = grdServicios.DataSource
    
    With mrs_FacturacionServicios
        .AddNew
        .Fields!IdProducto = 0
        '.Fields!id = 0
    
        .Fields!codigo = ""
        .Fields!descripcion = ""
        .Fields!cantidad = 1
        .Fields!PrecioUnitario = 0
        .Fields!subtotal = 0
        .Fields!IdEstadoFacturacion = 1
        If TipoEmpleado = sghTipoEmpleado.sghSIS Then
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSIS)
        ElseIf TipoEmpleado = sghTipoEmpleado.sghConvenio Then
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghConvenios)
        ElseIf TipoEmpleado = sghTipoEmpleado.sghSOAT Then
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghSOAT)
        Else
            .Fields!IdTipoFinanciamiento = Val(sghTipoFinanciamiento.sghPacienteNormal)
        End If
        '.Fields!IdAtencion = oAtencion.IdAtencion
       ' .Fields!poliza = ""
    
    End With
    
    mb_TransaccionDeNuevoRegistroEnProceso = True
    mb_NoEditar = True
    Set grdServicios.DataSource = mrs_FacturacionServicios
    mb_NoEditar = False
    grdServicios.GetRow(ssChildRowLast).Activation = ssActivationAllowEdit
End Sub

Private Sub txtNroHistoria_LostFocus()
Dim oPaciente As New doPaciente
Dim rsRespuesta As New Recordset

    If UserControl.txtNroHistoria = "" Then
        Exit Sub
    End If
    
    oPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
    Set rsRespuesta = mo_AdminAdmision.PacientesFiltrar(oPaciente)
    On Error Resume Next
    
    If rsRespuesta.RecordCount = 0 Then
        MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
    ElseIf rsRespuesta.RecordCount = 1 Then
        IdPaciente = rsRespuesta!IdPaciente
        Call ObtenerNombrePaciente(rsRespuesta!IdPaciente)
        ConfigurarFechaIngreso ml_IdPaciente
    End If
    
    If mo_AdminAdmision.MensajeError <> "" Then
        MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
    End If
        
        
        
End Sub

Function ValidaAccesos(TipoEmpleado As sghTipoEmpleado, tipoFinanciamiento As Integer, estado As Integer) As Boolean
Dim ok As Boolean
ok = False
Select Case TipoEmpleado
 
 Case sghTipoEmpleado.sghCajero
    If tipoFinanciamiento = Val(sghPacienteNormal) And estado = sghPagado Then
        ok = True
    End If
    If tipoFinanciamiento = sghPacienteNormal And estado = sghPendientePago Then
        ok = True
    End If
    If tipoFinanciamiento = sghTipoFinanciamiento.sghPacienteParticular And estado = sghPagado Then
        ok = True
    End If
 Case sghTipoEmpleado.sghAsistenta
 
 Case sghTipoEmpleado.sghConvenio
    If tipoFinanciamiento = sghConvenios And estado = sghPagado Then
        ok = True
    End If
 Case sghTipoEmpleado.sghCuentaCorriente
    If tipoFinanciamiento = sghPacienteNormal And estado = sghPendientePago Then
        ok = True
    End If
    If tipoFinanciamiento = sghConvenios And estado = sghPagado Then
        ok = True
    End If
 Case sghTipoEmpleado.sghOtros
    If estado = sghRegistroManual Then
        ok = True
    End If
 Case sghTipoEmpleado.sghSIS
    If tipoFinanciamiento = sghTipoFinanciamiento.sghSIS And estado = sghPagado Then
        ok = True
    End If
    
 Case sghTipoEmpleado.sghSOAT
    If tipoFinanciamiento = sghTipoFinanciamiento.sghSOAT And estado = sghPagado Then
        ok = True
    End If
End Select
ValidaAccesos = ok
End Function

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FarmDespachoDonaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FarmDespachoDonaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SighFarmacia.ucDespachoDonaciones grdProductos 
      Height          =   5025
      Left            =   0
      TabIndex        =   24
      Top             =   2460
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8864
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
      Height          =   2445
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11805
      Begin VB.TextBox txtNhistoria 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   29
         ToolTipText     =   "Ingrese el Nro de Historia Clínica"
         Top             =   1072
         Width           =   1245
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2850
         TabIndex        =   28
         Top             =   1050
         Width           =   3465
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2820
         TabIndex        =   26
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   25
         Top             =   660
         Width           =   3165
      End
      Begin VB.TextBox txtNcuenta 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   671
         Width           =   1245
      End
      Begin VB.ComboBox cmbCoordinador 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   1473
         Width           =   4770
      End
      Begin VB.TextBox txtDx 
         Height          =   315
         Left            =   7380
         MaxLength       =   30
         TabIndex        =   4
         ToolTipText     =   "Ingrese el Dx (4 dígitos)"
         Top             =   1110
         Width           =   675
      End
      Begin VB.TextBox txtNombreDx 
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
         Height          =   330
         Left            =   8070
         TabIndex        =   14
         Top             =   1110
         Width           =   3645
      End
      Begin VB.ComboBox cmbPrescriptor 
         Height          =   330
         Left            =   7380
         TabIndex        =   2
         Top             =   1500
         Width           =   4350
      End
      Begin VB.TextBox txtHoraRegistro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5580
         MaxLength       =   30
         TabIndex        =   13
         Top             =   270
         Width           =   735
      End
      Begin VB.ComboBox cmbAlmOrigen 
         Height          =   330
         Left            =   7380
         TabIndex        =   12
         Top             =   690
         Width           =   4350
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   315
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1890
         Width           =   4755
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10290
         MaxLength       =   30
         TabIndex        =   11
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtIntervencion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   10
         Top             =   270
         Width           =   1635
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   270
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
      Begin VB.Label lblNcuenta 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prescriptor"
         Height          =   210
         Left            =   6480
         TabIndex        =   23
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Coord.Serv.Social"
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   1530
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dx"
         Height          =   210
         Left            =   7140
         TabIndex        =   20
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Farmacia"
         Height          =   210
         Left            =   6660
         TabIndex        =   19
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   210
         Left            =   9660
         TabIndex        =   18
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Guía Int. Salida"
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         Height          =   210
         Left            =   3360
         TabIndex        =   16
         Top             =   300
         Width           =   810
      End
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
      Left            =   0
      TabIndex        =   7
      Top             =   7560
      Width           =   11820
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmDespachoDonaciones.frx":0CCA
         DownPicture     =   "FarmDespachoDonaciones.frx":112A
         Height          =   700
         Left            =   4470
         Picture         =   "FarmDespachoDonaciones.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmDespachoDonaciones.frx":1A14
         DownPicture     =   "FarmDespachoDonaciones.frx":1ED8
         Height          =   700
         Left            =   6000
         Picture         =   "FarmDespachoDonaciones.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   120
         Picture         =   "FarmDespachoDonaciones.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FarmDespachoDonaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Donaciones
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
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbPrescriptor As New SIGHEntidades.ListaDespleglable
Dim mo_cmbCoordinador As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oRsConceptos As New ADODB.Recordset
Dim oRsAlmacenOrigen As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mRs_Productos As New ADODB.Recordset
Dim mRs_Componente As New ADODB.Recordset
Dim mo_DoFarmMovimiento As New sighComun.DoFarmMovimiento
Dim mo_DoPaciente As New DOPaciente
Dim mo_DoFarmMovimientoDespachoDonaciones As New sighComun.DoFarmMovimientoDonaciones
Const lcConstanteMovimientoSalida As String = "S"
Const lcIdTipoConceptoDonacionesPacientes As Long = 27
Dim lnTotalDocumento As Double
Dim ml_IdCuentaAtencion As Long
Dim ml_IdDiagnostico As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_ReglasLaboratorio As New ReglasLaboratorio
Dim ms_MensajeError As String
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_PermisosFacturacion As New PermisosFacturacion
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_idUsuarioCreo As Long

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
    oRptClase.MovTipo = "S"
    oRptClase.Documento = mo_DoFarmMovimiento.movNumero
    oRptClase.TextoDelFiltro = "DESPACHO DE DONACION"
    oRptClase.Almacen = "Paciente: (" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtNhistoria.Text, False) & ")  " & txtNombrePaciente.Text & IIf(Me.txtNcuenta.Text <> "", " (N° Cuenta: " & Me.txtNcuenta.Text & ")", "")
    oRptClase.AlmacenO = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmOrigen.Text
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = "Guía Interna de Salida" & " - " & txtIntervencion.Text
    oRptClase.Importe = lnTotalDocumento
    oRptClase.EsUnaDonacion = True
    oRptClase.TipoReporte = "NiNs"
    oRptClase.Proveedor = ""
    oRptClase.idUsuario = ml_idUsuarioCreo
    oRptClase.Show vbModal
    Set oRptClase = Nothing
End Sub


Private Sub btnImprimir_Click()
   ImprimeDocumento
End Sub


Private Sub cmbAlmOrigen_Click()
    grdProductos.IdAlmacen = Val(mo_cmbAlmacenOrigen.BoundText)
End Sub

Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen

End Sub



Private Sub cmbCoordinador_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbCoordinador
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbPrescriptor_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPrescriptor
    AdministrarKeyPreview KeyCode

End Sub





Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdCuentaAtencion = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdCuentaAtencion, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    Set oBusqueda = Nothing
    Set oDOPaciente = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub Form_Activate()
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(mo_DoFarmMovimiento.IdAlmacenOrigen) = True Then
        btnCancelar_Click
        Exit Sub
   End If

End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbPrescriptor.MiComboBox = cmbPrescriptor
    Set mo_cmbCoordinador.MiComboBox = cmbCoordinador

End Sub

Private Sub Form_Load()
    ConfigurarGrdProductos
    CargarComboBoxes
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Despacho Donaciones"
    Case sghModificar
        Me.Caption = "Modificar Despacho Donaciones"
    Case sghConsultar
        Me.Caption = "Consultar Despacho Donaciones"
        btnImprimir.Visible = True
    Case sghEliminar
        Me.Caption = "Anular Despacho Donaciones"
    End Select
    CargarDatosAlFormulario
End Sub
Sub ConfigurarGrdProductos()
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.inicializar
    grdProductos.TipoPrecioParaNiNs = sghPrecioDonacion
    
End Sub


Sub CargarComboBoxes()
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='02' and idEstado=1")
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='02' and idEstado=1")
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacenOrigen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
    End If
   '
    mo_cmbPrescriptor.BoundColumn = "idEmpleado"
    mo_cmbPrescriptor.ListField = "ApNom"
    Set mo_cmbPrescriptor.RowSource = mo_ReglasLaboratorio.LabSeleccionaMedicos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
   '
    mo_cmbCoordinador.BoundColumn = "idEmpleado"
    mo_cmbCoordinador.ListField = "ApNom"
    Set mo_cmbCoordinador.RowSource = mo_ReglasFarmacia.EmpleadosDevuelveSoloServicioSocial
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError
    End If
End Sub
Sub CargarDatosAlFormulario()
     mo_Formulario.HabilitarDeshabilitar Me.txtIntervencion, False
     mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
     mo_Formulario.HabilitarDeshabilitar Me.txtHoraRegistro, False
     mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
     mo_Formulario.HabilitarDeshabilitar Me.txtNombrePaciente, False
     mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
     mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
     mo_Formulario.HabilitarDeshabilitar Me.txtNhistoria, False
     mo_Formulario.HabilitarDeshabilitar Me.txtDx, False
     mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
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
   mo_DoFarmMovimiento.movNumero = ml_movNumero
   mo_DoFarmMovimiento.MovTipo = lcConstanteMovimientoSalida
   If Not mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_DoFarmMovimiento) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   
   txtIntervencion.Text = mo_DoFarmMovimiento.DocumentoNumero
   mo_cmbAlmacenOrigen.BoundText = mo_DoFarmMovimiento.IdAlmacenOrigen
   txtObservaciones.Text = mo_DoFarmMovimiento.Observaciones
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDelMovimiento("idEstadoMovimiento=" & mo_DoFarmMovimiento.idEstadoMovimiento)
   txtFregistro.Text = Format(mo_DoFarmMovimiento.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   txtHoraRegistro.Text = Format(mo_DoFarmMovimiento.fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
   ml_idUsuarioCreo = mo_DoFarmMovimiento.idUsuario
   '**************Datos de la tabla FarmMovimientoProgramas *****************
   
   Dim mo_Diagnostico As New DODiagnostico
   With mo_DoFarmMovimientoDespachoDonaciones
       .movNumero = ml_movNumero
       .MovTipo = lcConstanteMovimientoSalida
       If Not mo_ReglasFarmacia.FarmMovimientoDespachoDonacionesSeleccionarPorId(mo_DoFarmMovimientoDespachoDonaciones) Then
            MsgBox mo_ReglasFarmacia.MensajeError
            Exit Sub
       Else
            mo_cmbCoordinador.BoundText = .idCoordinadorServicioSocial
            mo_cmbPrescriptor.BoundText = .idPrescriptorReceta
            'Dx
            ml_IdDiagnostico = .idDiagnostico
            Set mo_Diagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(.idDiagnostico)
            txtDx.Text = mo_Diagnostico.CodigoCIE2004
            txtNombreDx.Text = mo_Diagnostico.Descripcion
            'Cuenta del Paciente
            ml_IdCuentaAtencion = .idCuentaAtencion
            If ml_IdCuentaAtencion > 0 Then
               txtNcuenta.Text = ml_IdCuentaAtencion
               txtNcuenta_LostFocus
            End If
       End If
   End With
   Set mo_Diagnostico = Nothing
   '**************Acceso a Modificar Fecha de Registro *****************
   If mi_Opcion = sghModificar Then
        Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
        If mo_PermisosFacturacion.ActualizaFechaDocumentoES = True Then
           mo_Formulario.HabilitarDeshabilitar txtFregistro, True
        End If
        
   End If
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.movNumero = ml_movNumero
   grdProductos.CargaProductosPorMovNumero
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   If mo_DoFarmMovimiento.idEstadoMovimiento = 0 Then
      btnAceptar.Enabled = False
   End If
End Sub

Sub DeshabilitaCabecera()
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
  
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
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
                ml_idUsuarioCreo = ml_idUsuario
                MsgBox "Se agregó correctamente la Guía Interna de Salida N° " + txtIntervencion.Text, vbExclamation, Me.Caption
                LimpiarDatos
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdProductos.RefrescaSaldos
            End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
                ml_idUsuarioCreo = ml_idUsuario
                MsgBox "Se Modificó correctamente la Guía Interna de Salida N° " + txtIntervencion.Text, vbExclamation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                grdProductos.RefrescaSaldos
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If AnularNS() Then
                MsgBox " Se anuló la Guía Interna de Salida N° " + txtIntervencion.Text, vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub

Function ValidarDatosObligatorios() As Boolean
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If cmbAlmOrigen.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
   ElseIf Me.txtDatosDeCuenta.Text = "" Then
      ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° Cuenta del Paciente" + Chr(13)
      Me.txtNcuenta.SetFocus
   ElseIf cmbCoordinador.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Coordinador de Servicio Social" + Chr(13)
       cmbCoordinador.SetFocus
''   ElseIf cmbPrescriptor.Text = "" Then
''       ms_MensajeError = ms_MensajeError + "Por favor elija el Prescriptor" + Chr(13)
''       cmbPrescriptor.SetFocus
   End If
   lnTotalDocumento = grdProductos.DevuelveTotal
   Set mRs_Productos = grdProductos.DevuelveProductos
   If mRs_Productos.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        mRs_Productos.MoveFirst
        Do While Not mRs_Productos.EOF
           If Trim(mRs_Productos.Fields!codigo) = "" Or Trim(mRs_Productos.Fields!nombreProducto) = "" Then
              mRs_Productos.Delete
              mRs_Productos.Update
           ElseIf mRs_Productos.Fields!Cantidad <= 0 Or mRs_Productos!Cantidad > mRs_Productos!saldo Then
              ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  Tiene problemas de Saldo" + Chr(13)
           End If
           mRs_Productos.MoveNext
        Loop
   End If
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With mo_DoFarmMovimiento
            .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .IdAlmacenDestino = 0   '<<ninguno>>
            .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
            .idEstadoMovimiento = sghEstadoTabla.sghRegistrado    'registrado
            .idTipoConcepto = lcIdTipoConceptoDonacionesPacientes   'Donaciones a pacientes
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoSalida
            .Observaciones = txtObservaciones.Text
            .Total = lnTotalDocumento
        End With
        With mo_DoFarmMovimientoDespachoDonaciones
             .idCoordinadorServicioSocial = Val(mo_cmbCoordinador.BoundText)
             .idCuentaAtencion = ml_IdCuentaAtencion
             .idDiagnostico = ml_IdDiagnostico
             .idPrescriptorReceta = Val(mo_cmbPrescriptor.BoundText)
             .IdUsuarioAuditoria = ml_idUsuario
             .MovTipo = lcConstanteMovimientoSalida
             '.movNumero =
        End With
   Case sghModificar
        With mo_DoFarmMovimiento
            .Observaciones = txtObservaciones.Text
            .IdUsuarioAuditoria = ml_idUsuario
            .Total = lnTotalDocumento
        End With
        With mo_DoFarmMovimientoDespachoDonaciones
             .idCoordinadorServicioSocial = Val(mo_cmbCoordinador.BoundText)
             .idCuentaAtencion = ml_IdCuentaAtencion
             .idDiagnostico = ml_IdDiagnostico
             .idPrescriptorReceta = Val(mo_cmbPrescriptor.BoundText)
             .IdUsuarioAuditoria = ml_idUsuario
             .MovTipo = lcConstanteMovimientoSalida
             '.movNumero =
        End With
   Case sghEliminar
        With mo_DoFarmMovimiento
            .fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .idEstadoMovimiento = sghEstadoTabla.sghAnulado    'Anulado
            .IdUsuarioAuditoria = ml_idUsuario
        End With
   End Select
End Sub
Function AgregarDatos() As Boolean
    AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeDespachoDonaciones(mo_DoFarmMovimiento, mo_DoFarmMovimientoDespachoDonaciones, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    txtIntervencion.Text = mo_DoFarmMovimiento.DocumentoNumero
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasFarmacia.ModificaDatosDeDespachoDonaciones(mo_DoFarmMovimiento, mo_DoFarmMovimientoDespachoDonaciones, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function AnularNS() As Boolean
    AnularNS = mo_ReglasFarmacia.AnulaNotaSalida(mo_DoFarmMovimiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, 0)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function






Private Sub Form_Unload(Cancel As Integer)
     LimpiarVariablesDeMemoria
End Sub

Private Sub grdProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
     End If
End Sub





Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservaciones
    AdministrarKeyPreview KeyCode
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

Sub LimpiarDatos()
   txtNhistoria.Text = ""
   txtNombrePaciente.Text = ""
   ml_IdCuentaAtencion = 0
   cmbPrescriptor.Text = ""
   ml_IdDiagnostico = 0
   txtDx.Text = ""
   txtNombreDx.Text = ""
   cmbCoordinador.Text = ""
   txtObservaciones.Text = ""
   ml_movNumero = ""
   txtIntervencion.Text = ""
   lnTotalDocumento = 0
   Me.txtDatosDeCuenta.Text = ""
   grdProductos.movNumero = 0
   grdProductos.LimpiarGrilla
   grdProductos.AgregaRegistro
   Me.txtNcuenta.SetFocus
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbPrescriptor = Nothing
    Set mo_cmbCoordinador = Nothing
    Set oRsConceptos = Nothing
    Set oRsAlmacenOrigen = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mRs_Componente = Nothing
    Set mo_DoFarmMovimiento = Nothing
    Set mo_DoPaciente = Nothing
    Set mo_DoFarmMovimientoDespachoDonaciones = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_PermisosFacturacion = Nothing
End Sub

Private Sub txtObservaciones_LostFocus()
    grdProductos.TabEnDescripcion
End Sub


Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta

End Sub


Private Sub txtNcuenta_LostFocus()
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
       Dim oConexion As New Connection
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       txtDatosDeCuenta.Text = ""
       ml_IdCuentaAtencion = 0
       txtNombrePaciente.Text = ""
       txtNhistoria.Text = ""
       lbSigue = True
       If oRsTmp.RecordCount > 0 Then
          If oRsTmp.Fields!idEstado <> 1 Then
             If mi_Opcion <> sghConsultar Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          If lbSigue Then
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - " & IIf(oRsTmp.Fields!IdTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!IdTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                ml_IdCuentaAtencion = Val(Me.txtNcuenta.Text)
                txtNombrePaciente.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                txtNhistoria.Text = oRsTmp.Fields!NroHistoriaClinica
                If mi_Opcion = sghAgregar Then
                   Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorNroCuenta(ml_IdCuentaAtencion)
                   If oRsTmp.RecordCount > 0 Then
                        oRsTmp.MoveFirst
                        txtDx.Text = oRsTmp.Fields!CodigoCIE2004
                        txtNombreDx.Text = oRsTmp.Fields!Descripcion
                        ml_IdDiagnostico = oRsTmp.Fields!idDiagnostico
                   Else
                        txtDx.Text = ""
                        txtNombreDx.Text = ""
                        ml_IdDiagnostico = 0
                   End If
                End If
          End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
        oConexion.Close
        Set oConexion = Nothing
   End If
End Sub


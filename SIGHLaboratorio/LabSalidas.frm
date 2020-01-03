VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LabSalidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   Icon            =   "LabSalidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   Begin SIGHLaboratorio.ucLabServicios ucProductos 
      Height          =   6375
      Left            =   60
      TabIndex        =   8
      Top             =   1200
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   11245
   End
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   13725
      Begin VB.ComboBox cmbMotivo 
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
         Left            =   1440
         TabIndex        =   3
         Top             =   630
         Width           =   3240
      End
      Begin VB.TextBox txtNmovimiento 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   0
         Top             =   270
         Width           =   3225
      End
      Begin VB.TextBox txtEstado 
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
         Height          =   315
         Left            =   10800
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmbIdPuntoDeCarga 
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
         Left            =   10800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   630
         Width           =   2805
      End
      Begin VB.ComboBox cmbResponsable 
         Appearance      =   0  'Flat
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
         Left            =   6270
         TabIndex        =   4
         Top             =   630
         Width           =   3240
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   6270
         TabIndex        =   1
         Top             =   240
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
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
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
         Left            =   5250
         TabIndex        =   16
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
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
         TabIndex        =   15
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5370
         TabIndex        =   14
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10170
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pto. Carga"
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
         Left            =   9855
         TabIndex        =   11
         Top             =   690
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   9
      Top             =   7620
      Width           =   13710
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LabSalidas.frx":0CCA
         DownPicture     =   "LabSalidas.frx":118E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6870
         Picture         =   "LabSalidas.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LabSalidas.frx":1B66
         DownPicture     =   "LabSalidas.frx":1FC6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5340
         Picture         =   "LabSalidas.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "LabSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Salidas de Insumos
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim ml_IdMovimiento As Long
Dim mi_Opcion As sghOpciones
Dim ms_MensajeError As String
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdEstado As New sighentidades.ListaDespleglable
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim mo_cmbMotivo As New sighentidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim lbPrimeraVez As Boolean
Dim ml_idPaciente As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdFuenteFinanciamiento  As Long
Dim ml_IdServicioPaciente As Long
Dim oDOPaciente As New doPaciente
Dim oDOLabMovimiento As New DOLabMovimiento
Dim oDOLabMovimientoSalidas As New DOLabMovimientoSalidas
Dim rsProductos As Recordset
Dim ml_IdPuntoCarga As Long
Const ml_IdTipoFinanciamiento As Long = 1
Const lcConstanteMovimientoSalida As String = "S"
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_areaTrabajo As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
  idUsuario = ml_idUsuario
End Property

Property Let IdMovimiento(lValue As Long)
  ml_IdMovimiento = lValue
End Property

Property Get IdMovimiento() As Long
  IdMovimiento = ml_IdMovimiento
End Property

Property Let IdPuntoCarga(lValue As Long)
    ml_IdPuntoCarga = lValue
    If lValue = 32 Or lValue = 31 Then ml_areaTrabajo = 70
    If lValue = 38 Or lValue = 37 Or lValue = 34 Or lValue = 35 Or lValue = 33 Or lValue = 36 Then ml_areaTrabajo = 69
End Property

Property Get IdPuntoCarga() As Long
    IdPuntoCarga = ml_IdPuntoCarga
End Property

Private Sub btnAceptar_Click()
  Select Case mi_Opcion
    Case sghAgregar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If AgregarDatos() Then
            Me.txtNmovimiento = oDOLabMovimiento.IdMovimiento
            MsgBox "Se agregó correctamente el Movimiento N° " & oDOLabMovimiento.IdMovimiento, vbExclamation, Me.Caption
            Me.Visible = False
          Else
            MsgBox "No se pudo agregar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghModificar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If ModificarDatos() Then
            MsgBox "Se Modificó correctamente el Movimiento N° " & oDOLabMovimiento.IdMovimiento, vbExclamation, Me.Caption
            Me.Visible = False
          Else
            MsgBox "No se pudo modificar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghEliminar
      If MsgBox("¿Realmente desea Anular?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
      If ValidarReglas() Then
        CargaDatosAlObjetosDeDatos
        If EliminarDatos() Then
          MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
          Me.Visible = False
        Else
          MsgBox "No se pudo anular los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
  End Select
End Sub

Function ValidarDatosObligatorios() As Boolean
  ValidarDatosObligatorios = False
  ms_MensajeError = ""
  If cmbMotivo.Text = "" Then ms_MensajeError = ms_MensajeError & "- Elija el MOTIVO de la Salida" & Chr(13)
  If cmbResponsable.Text = "" Then ms_MensajeError = ms_MensajeError & "- Elija el Responsable que entrega." & Chr(13)
  If cmbIdPuntoDeCarga.Text = "" Then ms_MensajeError = ms_MensajeError & "- Elija el Punto de Carga." & Chr(13)
  Select Case mi_Opcion
    Case sghAgregar, sghModificar
      Set rsProductos = Me.ucProductos.FacturacionProductos
      If Not (rsProductos.EOF And rsProductos.BOF) Then
        rsProductos.MoveFirst
        Do While Not rsProductos.EOF
          If rsProductos!idProducto = 0 Then
            rsProductos.Delete
            rsProductos.Update
          Else
            If rsProductos!Cantidad <= 0 Then ms_MensajeError = ms_MensajeError & "El producto: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
            If rsProductos!PrecioUnitario <= 0 Then ms_MensajeError = ms_MensajeError & "El producto: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
          End If
          rsProductos.MoveNext
        Loop
      End If
      If Me.ucProductos.DevuelveTotalPagar <= 0 Then ms_MensajeError = ms_MensajeError & "- Agregue insumos a entregar (Importe total es S/. 0.00)" & Chr(13)
  End Select
  If ms_MensajeError = "" Then
    ValidarDatosObligatorios = True
  Else
    MsgBox ms_MensajeError, vbInformation, Me.Caption
  End If
End Function

Sub CargaDatosAlObjetosDeDatos()
  Select Case mi_Opcion
    Case sghAgregar
      With oDOLabMovimiento
        .fecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
        .IdlabEstado = sghEstadoTabla.sghRegistrado
        .IdPuntoCarga = ml_IdPuntoCarga
        .idTipoConcepto = sghTipoConceptoImagen.sghImgTCsalidaDeterioro  'Ingresos
        .idUsuario = ml_idUsuario
        .IdUsuarioAuditoria = ml_idUsuario
        .MovTipo = lcConstanteMovimientoSalida
      End With
      With oDOLabMovimientoSalidas
        .IdResponsable = Val(mo_cmbResponsable.BoundText)
        '.Motivo = txtMotivo.Text
        .idMotivoSalida = Val(mo_cmbMotivo.BoundText)
        .IdUsuarioAuditoria = ml_idUsuario
      End With
    Case sghModificar
      With oDOLabMovimiento
        .IdUsuarioAuditoria = ml_idUsuario
      End With
      With oDOLabMovimientoSalidas
        .IdResponsable = Val(mo_cmbResponsable.BoundText)
        .idMotivoSalida = Val(mo_cmbMotivo.BoundText) '.Motivo = txtMotivo.Text
        .IdUsuarioAuditoria = ml_idUsuario
      End With
    Case sghEliminar
      With oDOLabMovimiento
        .IdUsuarioAuditoria = ml_idUsuario
      End With
  End Select
End Sub

Function ValidarReglas() As Boolean
  ValidarReglas = False
  ValidarReglas = True
End Function

Function AgregarDatos() As Boolean
  AgregarDatos = mo_ReglasLaboratorio.LabMovimientoSalidasAgregar(oDOLabMovimiento, oDOLabMovimientoSalidas, rsProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
  ms_MensajeError = mo_ReglasLaboratorio.MensajeError
  ml_IdMovimiento = oDOLabMovimiento.IdMovimiento
End Function

Function ModificarDatos() As Boolean
  ModificarDatos = mo_ReglasLaboratorio.LabMovimientoSalidasModificar(oDOLabMovimiento, oDOLabMovimientoSalidas, rsProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
  ms_MensajeError = mo_ReglasLaboratorio.MensajeError
End Function

Function EliminarDatos() As Boolean
  EliminarDatos = mo_ReglasLaboratorio.LabMovimientoSalidasAnular(oDOLabMovimiento, oDOLabMovimientoSalidas, rsProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
  ms_MensajeError = mo_ReglasLaboratorio.MensajeError
End Function

Private Sub btnCancelar_Click()
  Me.Visible = False
End Sub

Private Sub cmbIdPuntoDeCarga_Click()
  ml_IdPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
  ucProductos.IdPuntoCarga = ml_IdPuntoCarga
End Sub

Private Sub cmbIdPuntoDeCarga_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbMotivo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbResponsable_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Initialize()
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
  Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
  Set mo_cmbMotivo.MiComboBox = cmbMotivo
End Sub

Private Sub Form_Load()
  txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL
  txtEstado.Text = "Registrado"
  
  CargaDataCombos
  
  Me.ucProductos.HabilitaIngresoDePrecio = False
  Me.ucProductos.idUsuario = ml_idUsuario
  Me.ucProductos.Inicializar
  Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
  Me.ucProductos.TipoProducto = sghServicio
  Me.ucProductos.IdPuntoCarga = ml_IdPuntoCarga
  
  Select Case mi_Opcion
    Case sghAgregar
      Me.Caption = "Agregar Salida de Insumos"
      mo_cmbIdPuntoCarga.BoundText = IdPuntoCarga
      mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
    Case sghModificar
      Me.Caption = "Modificar Salida de Insumos"
    Case sghConsultar
      Me.Caption = "Consultar Salida de Insumos"
    Case sghEliminar
      Me.Caption = "Eliminar Salida de Insumos"
  End Select
  
  CargarDatosAlFormulario
End Sub

Sub CargarDatosAlFormulario()
  mo_Formulario.HabilitarDeshabilitar Me.txtNmovimiento, False
  mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
  mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False

  Select Case mi_Opcion
    Case sghAgregar
      Me.ucProductos.idOrden = -999
      Me.ucProductos.CargaProductosPorIdOrden
      Me.ucProductos.AgregaProducto
    Case sghModificar
      CargarDatosALosControles
    Case sghConsultar
      CargarDatosALosControles
    Case sghEliminar
      CargarDatosALosControles
  End Select
End Sub

Sub CargarDatosALosControles()
  mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
  Set oDOLabMovimiento = mo_ReglasLaboratorio.LabMovimientoSeleccionarPorId(ml_IdMovimiento)
  txtFregistro.Text = Format(oDOLabMovimiento.fecha, sighentidades.DevuelveFechaSoloFormato_DMY)
  txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDeLaboratorio("idLabEstado=" & oDOLabMovimiento.IdlabEstado)
  txtNmovimiento.Text = ml_IdMovimiento
  mo_cmbIdPuntoCarga.BoundText = oDOLabMovimiento.IdPuntoCarga
  '
  Set oDOLabMovimientoSalidas = mo_ReglasLaboratorio.LabMovimientoSalidasSeleccionarPorId(ml_IdMovimiento)
  mo_cmbMotivo.BoundText = oDOLabMovimientoSalidas.idMotivoSalida
  mo_cmbResponsable.BoundText = oDOLabMovimientoSalidas.IdResponsable
  mb_ExistenDatos = True
  'Cargar datos de los servicios
  Me.ucProductos.LimpiarGrilla
  Me.ucProductos.IdMovimiento = ml_IdMovimiento
  Me.ucProductos.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
  Me.ucProductos.CargaProductosPorIdMovimiento
  If oDOLabMovimiento.IdlabEstado = 0 Or mi_Opcion = sghConsultar Then btnAceptar.Enabled = False
  Select Case mi_Opcion
    Case sghModificar
    Case sghEliminar
    Case sghConsultar
  End Select
End Sub

Sub CargaDataCombos()
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  Set mo_cmbResponsable.RowSource = mo_ReglasLaboratorio.EmpleadosDeLab(ml_areaTrabajo)
  mo_cmbIdPuntoCarga.ListField = "Descripcion"
  mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
  Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idUPS=2 or idUPS=3 or idUPS=4")
  Dim rsAlmacen As Recordset
  Set rsAlmacen = mo_reglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghlaboratorio, ml_idUsuario)
  If rsAlmacen.RecordCount > 0 Then
    mo_cmbIdPuntoCarga.BoundText = rsAlmacen.Fields!idLaboraSubArea
    mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
  End If
  mo_cmbMotivo.BoundColumn = "idMotivoSalida"
  mo_cmbMotivo.ListField = "Motivo"
  Set mo_cmbMotivo.RowSource = mo_ReglasLaboratorio.labMotivoSalidasSeleccionarTodos
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, sighentidades.NombreUsuario)
  If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
End Sub

Private Sub txtEstado_GotFocus()
   SeleccionaTexto txtEstado
End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFregistro_GotFocus()
  SeleccionaMask txtFregistro
End Sub

Private Sub txtFregistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNmovimiento_GotFocus()
  SeleccionaTexto txtNmovimiento
End Sub

Private Sub txtNmovimiento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

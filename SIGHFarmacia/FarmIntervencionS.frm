VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FarmIntervencionS 
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
   Icon            =   "FarmIntervencionS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2865
      Left            =   0
      TabIndex        =   12
      Top             =   30
      Width           =   11805
      Begin VB.CommandButton cmdPaquetes 
         Caption         =   "Carga Paquetes"
         Height          =   495
         Left            =   9150
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.ComboBox cmbCoordinador 
         Height          =   330
         Left            =   7590
         TabIndex        =   5
         Top             =   1470
         Width           =   4140
      End
      Begin VB.TextBox txtDx 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Ingrese el Dx (4 dígitos)"
         Top             =   1890
         Width           =   855
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
         Height          =   315
         Left            =   2730
         TabIndex        =   20
         Top             =   1890
         Width           =   3585
      End
      Begin VB.CommandButton cmdBuscaDx 
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
         Left            =   2400
         TabIndex        =   19
         Top             =   1890
         Width           =   315
      End
      Begin VB.TextBox txtNhistoria 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "Ingrese el Nro Historia Clínica"
         Top             =   660
         Width           =   1125
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         TabIndex        =   18
         Top             =   660
         Width           =   3255
      End
      Begin VB.CommandButton btnBuscarPaciente 
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
         Left            =   2730
         TabIndex        =   17
         Top             =   660
         Width           =   315
      End
      Begin VB.ComboBox cmbPrescriptor 
         Height          =   330
         Left            =   7590
         TabIndex        =   4
         Top             =   1080
         Width           =   4140
      End
      Begin VB.TextBox txtHoraRegistro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5580
         MaxLength       =   30
         TabIndex        =   16
         Top             =   270
         Width           =   735
      End
      Begin VB.ComboBox cmbSubComponente 
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   1470
         Width           =   4770
      End
      Begin VB.ComboBox cmbComponente 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   1050
         Width           =   4770
      End
      Begin VB.ComboBox cmbAlmOrigen 
         Height          =   330
         Left            =   7590
         TabIndex        =   15
         Top             =   690
         Width           =   4140
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   315
         Left            =   7590
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1890
         Width           =   4125
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7590
         MaxLength       =   30
         TabIndex        =   14
         Top             =   300
         Width           =   1125
      End
      Begin VB.TextBox txtIntervencion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   13
         Top             =   270
         Width           =   1635
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   4200
         TabIndex        =   21
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
      Begin MSMask.MaskEdBox txtFprescribe 
         Height          =   315
         Left            =   9930
         TabIndex        =   34
         Top             =   300
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "F.Prescripción"
         Height          =   210
         Left            =   8850
         TabIndex        =   35
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Componente"
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Paciente"
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prescriptor"
         Height          =   210
         Left            =   6690
         TabIndex        =   30
         Top             =   1140
         Width           =   870
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   6390
         TabIndex        =   29
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Coordinador"
         Height          =   210
         Left            =   6585
         TabIndex        =   28
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
         Height          =   210
         Left            =   150
         TabIndex        =   27
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SubComponente"
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   1530
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Farmacia"
         Height          =   210
         Left            =   6840
         TabIndex        =   25
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   210
         Left            =   6990
         TabIndex        =   24
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Guía Int. Salida"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         Height          =   210
         Left            =   3360
         TabIndex        =   22
         Top             =   300
         Width           =   810
      End
   End
   Begin SighFarmacia.ucIntervencionS grdProductos 
      Height          =   4515
      Left            =   0
      TabIndex        =   7
      Top             =   2940
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   7964
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
      TabIndex        =   10
      Top             =   7560
      Width           =   11820
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmIntervencionS.frx":0CCA
         DownPicture     =   "FarmIntervencionS.frx":112A
         Height          =   700
         Left            =   4470
         Picture         =   "FarmIntervencionS.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmIntervencionS.frx":1A14
         DownPicture     =   "FarmIntervencionS.frx":1ED8
         Height          =   700
         Left            =   6000
         Picture         =   "FarmIntervencionS.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   120
         Picture         =   "FarmIntervencionS.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FarmIntervencionS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Intervensiones Sanitarias
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
Dim mo_cmbComponente As New SIGHEntidades.ListaDespleglable
Dim mo_cmbSubComponente As New SIGHEntidades.ListaDespleglable
Dim mo_cmbCoordinador As New SIGHEntidades.ListaDespleglable
Dim oRsConceptos As New ADODB.Recordset
Dim oRsAlmacenOrigen As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mRs_Productos As New ADODB.Recordset
Dim mRs_Componente As New ADODB.Recordset
Dim mo_DoFarmMovimiento As New sighComun.DoFarmMovimiento
Dim mo_DoPaciente As New DOPaciente
Dim mo_DoFarmMovimientoProgramas As New sighComun.DOfarmMovimientoProgramas
Const lcConstanteMovimientoSalida As String = "S"
Const lcIdTipoConceptoIntervencionSanitaria As Long = 16
Dim lnTotalDocumento As Double
Dim ml_IdPaciente As Long
Dim ml_IdDiagnostico As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_PermisosFacturacion As New PermisosFacturacion
Dim ms_MensajeError As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim wxParametro347 As String
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
    oRptClase.TextoDelFiltro = "INTERVENCION SANITARIA"
    oRptClase.Almacen = ""
    oRptClase.AlmacenO = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmOrigen.Text
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = Trim("Guía Interna de Salida") & " - " & txtIntervencion.Text
    oRptClase.Importe = lnTotalDocumento
    oRptClase.TipoReporte = "NiNs"
    oRptClase.MuestraTipoSoporteSISMED = True
    oRptClase.Proveedor = cmbComponente.Text & " (" & Label5.Caption & ": " & Trim(cmbSubComponente.Text) & ")"
    oRptClase.idUsuario = ml_idUsuarioCreo
    oRptClase.Show vbModal
    Set oRptClase = Nothing
    Set oDOfarmAlmacen = Nothing
End Sub

Private Sub btnBuscarPaciente_Click()
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
            ml_IdPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
            Me.cmbComponente.SetFocus
        End If
    End If
    Set oBusqueda = Nothing
    Set oDOPaciente = Nothing
    oConexion.Close
    Set oConexion = Nothing
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

Private Sub cmbComponente_Click()
    mRs_Componente.MoveFirst
    mRs_Componente.Find "idComponente=" & mo_cmbComponente.BoundText
    mo_cmbSubComponente.BoundColumn = "IdSubComponente"
    mo_cmbSubComponente.ListField = "Descripcion"
    Set mo_cmbSubComponente.RowSource = mo_ReglasFarmacia.FarmComponenteSubDevuelveTodosSegunComponente(mRs_Componente.Fields!idComponente)
End Sub



Private Sub cmbComponente_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbComponente
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbComponente_LostFocus()
    ml_IdDiagnostico = 0
    txtDx.Text = ""
    txtNombreDx.Text = ""


End Sub

Private Sub cmbCoordinador_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbCoordinador
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbPrescriptor_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPrescriptor
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbSubComponente_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbSubComponente
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbSubComponente_LostFocus()
    Dim lcTipoSOPORTE As String
    ml_IdDiagnostico = 0
    txtDx.Text = ""
    txtNombreDx.Text = ""
    grdProductos.TipoSOPORTE = ""
    If Val(mo_cmbSubComponente.BoundText) > 0 Then
        txtDx.Text = mo_ReglasFarmacia.DevuelveDiagnosticoSegunSubComponente(Val(mo_cmbSubComponente.BoundText), lcTipoSOPORTE)
        grdProductos.TipoSOPORTE = lcTipoSOPORTE
        If txtDx.Text <> "" Then
           txtDx_LostFocus
        End If
        
    End If
End Sub

Private Sub cmdBuscaDx_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.SoloMuestraDxGalenHos = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtDx.Text = oDODiagnostico.CodigoCIE2004
            txtNombreDx.Text = oDODiagnostico.Descripcion
        End If
    End If
    Set oBusqueda = Nothing
    Set oDODiagnostico = Nothing
End Sub

Private Sub cmdPaquetes_Click()
    If cmbAlmOrigen.Text = "" Then
       MsgBox "Debe elegir Almacén", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim oPaquetesBuscar As New SIGHNegocios.BuscaPaquetes
    Dim lnIdFactPaquete As Long
    oPaquetesBuscar.DebeConsiderarPaquete = sghTipoPaqueteSoloFarmacia
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdFactPaquete = oPaquetesBuscar.IdFactPaquete
       grdProductos.cargaPaqueteElegido lnIdFactPaquete
    End If
    Set oPaquetesBuscar = Nothing

End Sub

Private Sub Form_Activate()
    If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenOrigen.BoundText)) = True Then
       btnCancelar_Click
       Exit Sub
    End If
End Sub



Private Sub Form_Initialize()
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbPrescriptor.MiComboBox = cmbPrescriptor
    Set mo_cmbComponente.MiComboBox = cmbComponente
    Set mo_cmbSubComponente.MiComboBox = cmbSubComponente
    Set mo_cmbCoordinador.MiComboBox = cmbCoordinador

End Sub

Private Sub Form_Load()
    txtFprescribe.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    wxParametro347 = lcBuscaParametro.SeleccionaFilaParametro(347)

    
    ConfigurarGrdProductos
    CargarComboBoxes
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Intervención Sanitaria"
        cmdPaquetes.Visible = True
    Case sghModificar
        Me.Caption = "Modificar Intervención Sanitaria"
    Case sghConsultar
        Me.Caption = "Consultar Intervención Sanitaria"
        btnImprimir.Visible = True
    Case sghEliminar
        Me.Caption = "Anular Intervención Sanitaria"
    End Select
    CargarDatosAlFormulario
End Sub
Sub ConfigurarGrdProductos()
    grdProductos.Parametro347 = wxParametro347
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.inicializar
    grdProductos.TipoPrecioParaNiNs = 3    'precio de venta
    
End Sub


Sub CargarComboBoxes()
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1")
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1")
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacenOrigen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
    End If
   '
    mo_cmbPrescriptor.BoundColumn = "idEmpleado"
    mo_cmbPrescriptor.ListField = "ApNom"
    Set mo_cmbPrescriptor.RowSource = mo_ReglasFarmacia.EmpleadosDevuelvePrescriptores
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
   '
    mo_cmbCoordinador.BoundColumn = "idEmpleado"
    mo_cmbCoordinador.ListField = "ApNom"
    Set mo_cmbCoordinador.RowSource = mo_ReglasFarmacia.EmpleadosDevuelveCoordinadores
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
   '
    Set mRs_Componente = mo_ReglasFarmacia.FarmComponenteDevuelveTodos
    mo_cmbComponente.BoundColumn = "idComponente"
    mo_cmbComponente.ListField = "Descripcion"
    Set mo_cmbComponente.RowSource = mo_ReglasFarmacia.FarmComponenteDevuelveTodos
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
   Dim oConexion As New Connection
   oConexion.CursorLocation = adUseClient
   oConexion.CommandTimeout = 300
   oConexion.Open SIGHEntidades.CadenaConexion
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
   With mo_DoFarmMovimientoProgramas
       .movNumero = ml_movNumero
       .MovTipo = lcConstanteMovimientoSalida
       If Not mo_ReglasFarmacia.FarmMovimientoProgramaSeleccionarPorId(mo_DoFarmMovimientoProgramas) Then
            MsgBox mo_ReglasFarmacia.MensajeError
            Exit Sub
       Else
            mo_cmbComponente.BoundText = .idComponente
            mo_cmbCoordinador.BoundText = .idCoordinador
            mo_cmbPrescriptor.BoundText = .idPrescriptor
            mo_cmbSubComponente.BoundText = .idSubComponente
            txtFprescribe.Text = Format(.FechaHoraPrescribe, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            'Dx
            ml_IdDiagnostico = .idDiagnostico
            Set mo_Diagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(.idDiagnostico)
            txtDx.Text = mo_Diagnostico.CodigoCIE2004
            txtNombreDx.Text = mo_Diagnostico.Descripcion
            'Paciente
            ml_IdPaciente = .IdPaciente
            If ml_IdPaciente > 0 Then
                 mo_DoPaciente.IdPaciente = ml_IdPaciente
                 Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
                 txtNhistoria.Text = mo_DoPaciente.NroHistoriaClinica
                 txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
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
   '******permiso a Modificar documento con Fecha Anterior a la actual
   Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
   If mo_PermisosFacturacion.ActualizaFechaDocumentoES = False Then
      If CDate(lcBuscaParametro.RetornaFechaServidorSQL) <> CDate(txtFregistro.Text) Then
         MsgBox "No tiene ACCESO a Modificar/Anular una Intervención Sanitaria" & Chr(13) & " de una Fecha Registro diferente a la actual", vbExclamation, Me.Caption
         btnAceptar.Enabled = False
      End If
   End If
   oConexion.Close
   Set oConexion = Nothing
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
   ElseIf cmbComponente.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Componente" + Chr(13)
       cmbComponente.SetFocus
   ElseIf cmbSubComponente.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el SubComponente" + Chr(13)
       cmbSubComponente.SetFocus
   ElseIf cmbPrescriptor.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Prescriptor" + Chr(13)
       cmbPrescriptor.SetFocus
   ElseIf txtNombreDx.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor ingrese el Dx del Paciente" + Chr(13)
       txtDx.SetFocus
   ElseIf cmbCoordinador.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Coordinador" + Chr(13)
       cmbCoordinador.SetFocus
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
            .idTipoConcepto = lcIdTipoConceptoIntervencionSanitaria    'Intervencion Sanitaria
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoSalida
            .Observaciones = txtObservaciones.Text
            .Total = lnTotalDocumento
        End With
        With mo_DoFarmMovimientoProgramas
             .idComponente = Val(mo_cmbComponente.BoundText)
             .idCoordinador = Val(mo_cmbCoordinador.BoundText)
             .idDiagnostico = ml_IdDiagnostico
             .IdPaciente = ml_IdPaciente
             .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
             .idSubComponente = Val(mo_cmbSubComponente.BoundText)
             .IdUsuarioAuditoria = ml_idUsuario
             .MovTipo = lcConstanteMovimientoSalida
             .FechaHoraPrescribe = txtFprescribe.Text
        End With
   Case sghModificar
        With mo_DoFarmMovimiento
            .Observaciones = txtObservaciones.Text
            .IdUsuarioAuditoria = ml_idUsuario
            .Total = lnTotalDocumento
        End With
        With mo_DoFarmMovimientoProgramas
             .idComponente = Val(mo_cmbComponente.BoundText)
             .idCoordinador = Val(mo_cmbCoordinador.BoundText)
             .idDiagnostico = ml_IdDiagnostico
             .IdPaciente = ml_IdPaciente
             .idPrescriptor = Val(mo_cmbPrescriptor.BoundText)
             .idSubComponente = Val(mo_cmbSubComponente.BoundText)
             .IdUsuarioAuditoria = ml_idUsuario
             .FechaHoraPrescribe = txtFprescribe.Text
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
    AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeIntervencion(mo_DoFarmMovimiento, mo_DoFarmMovimientoProgramas, _
                                                               mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                               wxParametro347)
    txtIntervencion.Text = mo_DoFarmMovimiento.DocumentoNumero
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
End Function
Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasFarmacia.ModificaDatosDeIntervencion(mo_DoFarmMovimiento, mo_DoFarmMovimientoProgramas, _
                                                                   mRs_Productos, mo_lnIdTablaLISTBARITEMS, _
                                                                   mo_lcNombrePc, wxParametro347)
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

Private Sub txtDx_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDx
    AdministrarKeyPreview KeyCode

End Sub


Private Sub txtDx_LostFocus()
        Dim oDODiagnostico As DODiagnostico
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtDx.Text, True)
        If Not oDODiagnostico Is Nothing Then
            ml_IdDiagnostico = oDODiagnostico.idDiagnostico
            txtNombreDx.Text = oDODiagnostico.Descripcion
        Else
            ml_IdDiagnostico = 0
            txtNombreDx.Text = ""
        End If

End Sub

Private Sub txtFprescribe_LostFocus()
If Not IsDate(txtFprescribe.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFprescribe.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtNhistoria_LostFocus()
      If mo_Teclado.TextoEsSoloNumeros(txtNhistoria.Text) Then
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oDOPaciente As New sighComun.DOPaciente
        oDOPaciente.NroHistoriaClinica = txtNhistoria.Text
        Set oRsTmp1 = mo_AdminAdmision.PacientesFiltrar(oDOPaciente, False, False, "")
        If oRsTmp1.RecordCount > 0 Then
           ml_IdPaciente = oRsTmp1.Fields!IdPaciente
           txtNombrePaciente.Text = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & oRsTmp1.Fields!PrimerNombre
        Else
           ml_IdPaciente = 0
           txtNombrePaciente.Text = ""
        End If
        Set oRsTmp1 = Nothing
        Set oDOPaciente = Nothing
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
   ml_IdPaciente = 0
   cmbPrescriptor.Text = ""
   ml_IdDiagnostico = 0
   txtDx.Text = ""
   txtNombreDx.Text = ""
   cmbComponente.Text = ""
   cmbSubComponente.Text = ""
   cmbCoordinador.Text = ""
   txtObservaciones.Text = ""
   ml_movNumero = ""
   txtIntervencion.Text = ""
   lnTotalDocumento = 0
   grdProductos.movNumero = 0
   txtFprescribe.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
   grdProductos.LimpiarGrilla
   grdProductos.AgregaRegistro
   txtNhistoria.SetFocus
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbPrescriptor = Nothing
    Set mo_cmbComponente = Nothing
    Set mo_cmbSubComponente = Nothing
    Set mo_cmbCoordinador = Nothing
    Set oRsConceptos = Nothing
    Set oRsAlmacenOrigen = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mRs_Componente = Nothing
    Set mo_DoFarmMovimiento = Nothing
    Set mo_DoPaciente = Nothing
    Set mo_DoFarmMovimientoProgramas = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_PermisosFacturacion = Nothing
End Sub

Private Sub txtObservaciones_LostFocus()
    grdProductos.TabEnDescripcion
End Sub

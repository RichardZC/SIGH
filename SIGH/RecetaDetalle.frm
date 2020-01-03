VERSION 5.00
Begin VB.Form RecetaDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   Icon            =   "RecetaDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
      Enabled         =   0   'False
      Height          =   8085
      Left            =   30
      TabIndex        =   9
      Top             =   630
      Width           =   11775
      Begin VB.ComboBox cmbIdResponsable 
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
         Left            =   1080
         TabIndex        =   12
         Top             =   120
         Width           =   5550
      End
      Begin SISGalenPlus.ucRecetas ucRecetas1 
         Height          =   7575
         Left            =   15
         TabIndex        =   10
         Top             =   495
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   13361
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Médico"
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
         Left            =   180
         TabIndex        =   11
         Top             =   165
         Width           =   570
      End
   End
   Begin VB.Frame frmCabecera 
      Height          =   615
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   11805
      Begin VB.TextBox txtPlan 
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
         Left            =   6630
         TabIndex        =   7
         Top             =   210
         Width           =   5145
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         Height          =   315
         Left            =   2190
         TabIndex        =   6
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   210
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
         Left            =   2520
         TabIndex        =   5
         Top             =   210
         Width           =   4095
      End
      Begin VB.TextBox txtNcuenta 
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
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
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
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   30
      TabIndex        =   2
      Top             =   8700
      Width           =   11835
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RecetaDetalle.frx":0CCA
         DownPicture     =   "RecetaDetalle.frx":112A
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
         Left            =   4455
         Picture         =   "RecetaDetalle.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RecetaDetalle.frx":1A14
         DownPicture     =   "RecetaDetalle.frx":1ED8
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
         Left            =   5970
         Picture         =   "RecetaDetalle.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1365
      End
   End
End
Attribute VB_Name = "RecetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Registro de Recetas automáticas
'        Programado por: Barrantes D
'        Fecha: Enero 2011
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_idCuentaAtencion As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdServicioPaciente As Long
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica

Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnRecetaRayosX As Long, lnRecetaEcografiaO As Long, lnRecetaEcografiaG As Long
Dim lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long, lnRecetaPatologiaC As Long
Dim lnRecetaBancoS As Long, lnRecetaFarmacia As Long, lnRecetaOtrosCpt As Long
Dim ml_IdTipoServicio  As Long
Dim ml_FechaReceta As Date
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim ml_IdMedicoServicioActual As Long, ml_IdFuenteFinanciamiento As Long

Property Let IdMedicoServicioActual(lValue As Long)
    ml_IdMedicoServicioActual = lValue
       Dim oRsTmp2 As New Recordset
       Dim oConexion As New Connection
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       Set oRsTmp2 = oBuscaMedicos.MedicosXidEmpleado(sighentidades.Usuario, oConexion)
       If oRsTmp2.RecordCount > 0 Then
           ml_IdMedicoServicioActual = oRsTmp2!idMedico
       End If
       oRsTmp2.Close
       Set oRsTmp2 = Nothing
       oConexion.Close
       Set oConexion = Nothing
    
End Property


Property Let FechaReceta(lValue As Date)
    ml_FechaReceta = lValue
End Property


Property Let idTipoServicio(lValue As Long)
    ml_IdTipoServicio = lValue
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
   If mi_Opcion = sghAgregar Then
      txtNcuenta.Text = ml_idCuentaAtencion
      'txtNcuenta_LostFocus
   End If
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property



'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
    mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
    Select Case mi_Opcion
     Case sghAgregar
         Me.UcRecetas1.Opcion = sghAgregar    'debb-09/07/2015
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
    End Select
 
    Select Case mi_Opcion
         Case sghAgregar
         Case sghModificar
         Case sghConsultar
            Me.btnAceptar.Enabled = False
         Case sghEliminar
     End Select
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_ReglasAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(oDOPaciente.idPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Receta"
       Case sghModificar
           Me.Caption = "Modificar Receta"
       Case sghConsultar
           Me.Caption = "Consultar Receta"
       Case sghEliminar
           Me.Caption = "Eliminar Receta"
       End Select
       '
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oConexion As New Connection
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       mo_cmbIdResponsable.BoundColumn = "IdMedico"
       mo_cmbIdResponsable.ListField = "Dmedico"
       Set oRsTmp1 = oBuscaMedicos.MedicosSeleccionarTodosOrdenadoAlfabeticamente
       oRsTmp1.Filter = "esActivo=true and idcolegioHIS<>'00' and idColegioHIS<>'02' and idColegioHIS<>'06' and idColegioHIS<>'07'  and idColegioHIS<>'10' and idColegioHIS<>'11'"
       Set mo_cmbIdResponsable.RowSource = oRsTmp1
       Set oRsTmp2 = oBuscaMedicos.MedicosXidEmpleado(sighentidades.Usuario, oConexion)
       If oRsTmp2.RecordCount > 0 Then
           mo_cmbIdResponsable.BoundText = oRsTmp2!idMedico
       End If
       oRsTmp2.Close
       Set oRsTmp2 = Nothing
       oConexion.Close
       Set oConexion = Nothing
       
       '
       UcRecetas1.Inicializar
       UcRecetas1.lnWnd = Me.hwnd
       '
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       
       wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
       wxParametro545 = lcBuscaParametro.SeleccionaFilaParametro(545)
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Triaje
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       Else
            On Error Resume Next
            'Me.txtPresion.SetFocus
       End If
   Else
       If txtNcuenta.Text <> "" And mi_Opcion = sghAgregar And ml_IdMedicoServicioActual > 0 Then
            txtNcuenta_LostFocus
            mo_cmbIdResponsable.BoundText = ml_IdMedicoServicioActual
       Else
            On Error Resume Next
            txtNcuenta.SetFocus
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
            btnAceptar_Click
        End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente" & Chr(13) & DevuelveNroRecetasGeneradas, vbInformation, Me.Caption
                    LimpiarFormulario
                    Me.Visible = False
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente" & Chr(13) & DevuelveNroRecetasGeneradas, vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   ValidarDatosObligatorios = False
   If Me.txtNcuenta.Text = "" Or Me.txtPlan.Text = "" Then
       MsgBox "Ingrese el N° Cuenta", vbInformation, Me.Caption
       Exit Function
   End If
   If Me.UcRecetas1.AlMenosHayUnaReceta() = False Then
       MsgBox "Se debe registrar al menos 1 Receta", vbInformation, Me.Caption
       Exit Function
   End If
   If mo_cmbIdResponsable.BoundText = "" Then
       MsgBox "Por favor elija al Médico que prescribe la receta", vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim lcMensaje11 As String
   'debb-03/09/2015
   If mo_ReglasFarmacia.RecetaChequeaSiFechaVigenciaEsCorrecta(Me.UcRecetas1.DevuelveFarmacia) = False Then
      Exit Function
   End If
   If Me.UcRecetas1.ValidaReglas = False Then
      Exit Function
   End If
   '
   lcMensaje11 = mo_ReglasSISgalenhos.ReglasDeConsistenciaSISsoloFarmaciaXmonto(Val(txtNcuenta.Text), _
                             ml_IdFuenteFinanciamiento, Format(IIf(mi_Opcion = sghAgregar, Date, ml_FechaReceta), "dd/mm/yyyy"), _
                             mi_Opcion, Me.UcRecetas1.DevuelveFarmacia, True)
   If lcMensaje11 <> "" Then
      MsgBox lcMensaje11, vbInformation, ""
      Exit Function
   End If
   '
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Triaje
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    AgregarDatos = mo_ReglasAdmision.RecetaAgregar(ml_idCuentaAtencion, ml_IdServicioPaciente, ml_idUsuario, _
                                     lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                     lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                     Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                     Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                     Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                     Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, _
                                     mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & txtDatosDeCuenta.Text, _
                                     Val(mo_cmbIdResponsable.BoundText), Me.UcRecetas1.DevuelveOtrosCpt, lnRecetaOtrosCpt)
    If lnRecetaRayosX > 0 Or lnRecetaEcografiaO > 0 Or lnRecetaEcografiaG > 0 Or lnRecetaTomografia > 0 Or _
             lnRecetaAnatomiaP > 0 Or lnRecetaPatologiaC > 0 Or lnRecetaBancoS > 0 Or lnRecetaFarmacia > 0 Or lnRecetaOtrosCpt > 0 Then
       Me.UcRecetas1.idCuentaAtencion = ml_idCuentaAtencion 'Actualizado 30092014
       Me.UcRecetas1.CargaNumeroDeRecetaEimprime lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                                 lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, True, _
                                                 lnRecetaOtrosCpt
    End If
End Function


'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_ReglasAdmision.RecetaModificar(ml_idCuentaAtencion, ml_IdServicioPaciente, ml_idUsuario, _
                                     lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                     lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                     Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                     Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                     Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                     Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, ml_FechaReceta, _
                                     mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & txtDatosDeCuenta.Text, _
                                     Val(mo_cmbIdResponsable.BoundText), True, Me.UcRecetas1.DevuelveOtrosCpt, lnRecetaOtrosCpt)
    If lnRecetaRayosX > 0 Or lnRecetaEcografiaO > 0 Or lnRecetaEcografiaG > 0 Or lnRecetaTomografia > 0 Or _
             lnRecetaAnatomiaP > 0 Or lnRecetaPatologiaC > 0 Or lnRecetaBancoS > 0 Or lnRecetaFarmacia > 0 Or lnRecetaOtrosCpt > 0 Then
       Me.UcRecetas1.CargaNumeroDeRecetaEimprime lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                                 lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, True, _
                                                 lnRecetaOtrosCpt
    End If
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_ReglasAdmision.RecetaEliminar(ml_idCuentaAtencion, ml_IdServicioPaciente, ml_idUsuario, _
                                     lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                     lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                     Me.UcRecetas1.DevuelveRayosX, Me.UcRecetas1.DevuelveEcografiaO, _
                                     Me.UcRecetas1.DevuelveEcografiaG, Me.UcRecetas1.DevuelveTomografia, _
                                     Me.UcRecetas1.DevuelveAnatomia, Me.UcRecetas1.DevuelvePatologia, _
                                     Me.UcRecetas1.DevuelveBancoSangre, Me.UcRecetas1.DevuelveFarmacia, _
                                     mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & txtDatosDeCuenta.Text)
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
       Dim oRecetaCabecera As New RecetaCabecera
       Dim oRsCabeceraReceta As New Recordset
       Dim oConexion As New Connection
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       Set oRecetaCabecera.Conexion = oConexion
       Set oRsCabeceraReceta = oRecetaCabecera.SeleccionarPorIdCuentaAtencion(ml_idCuentaAtencion)
       oRsCabeceraReceta.Filter = "fechaReceta='" & ml_FechaReceta & "'"
       If oRsCabeceraReceta.RecordCount = 0 Then
           MsgBox "No se pudo obtener los datos" + Chr(13) + oRecetaCabecera.MensajeError, vbInformation, Me.Caption
           mb_ExistenDatos = False
           Exit Sub
       End If
       If Not IsNull(oRsCabeceraReceta.Fields!idMedicoREceta) Then
          mo_cmbIdResponsable.BoundText = oRsCabeceraReceta.Fields!idMedicoREceta
       End If
       Me.txtNcuenta.Text = ml_idCuentaAtencion
       txtNcuenta_LostFocus
       '
       Me.UcRecetas1.idTipoFinanciamiento = ml_IdTipoFinanciamiento
       Me.UcRecetas1.idCuentaAtencion = ml_idCuentaAtencion 'actualizado 22092014
       Me.UcRecetas1.CargaDatosAcontroles oRsCabeceraReceta, lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, _
                     lnRecetaTomografia, lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                     lnRecetaOtrosCpt
                     
       Me.UcRecetas1.Opcion = mi_Opcion     'debb-09/07/2015
       '
       mb_ExistenDatos = True
End Sub


Sub LimpiarFormulario()
End Sub











Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub

Private Sub txtNcuenta_LostFocus()
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
       Dim oRsDx As New Recordset
       Dim oRsTmp As New Recordset
       Dim oRsTmp1 As New Recordset   'debb-24/06/2015
       Dim lbSigue As Boolean
       Dim oConexion As New Connection
       oConexion.Open sighentidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       lbSigue = True
       txtDatosDeCuenta.Text = ""
       txtPlan.Text = ""
       ml_idCuentaAtencion = 0
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
            ml_IdFuenteFinanciamiento = oRsTmp!IdFuenteFinanciamiento
            If mi_Opcion = sghAgregar And _
                   mo_ReglasAdmision.AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(Val(txtNcuenta.Text), wxParametro302, _
                                                                                oRsTmp!IdFuenteFinanciamiento) = False Then
                                                                             
                   lbSigue = False
            End If

            If lbSigue Then
                If oRsTmp.Fields!idTipoServicio <> ml_IdTipoServicio Then
                    Select Case oRsTmp.Fields!idTipoServicio
                        Case sghConsultaExterna
                            MsgBox "El Nro de Cuenta corresponde al servicio de Consultorios Externos", vbInformation, Me.Caption
                        Case sghHospitalizacion
                            MsgBox "El Nro de Cuenta corresponde al servicio de Hospitalización", vbInformation, Me.Caption
                        Case sghEmergenciaObservacion
                            MsgBox "El Nro de Cuenta corresponde al servicio de Observación Emergencia ", vbInformation, Me.Caption
                        Case sghEmergenciaConsultorios
                            MsgBox "El Nro de Cuenta corresponde al servicio de Emergencias", vbInformation, Me.Caption
                        Case Else
                            MsgBox "El Nro de Cuenta no corresponde al servicio seleccionado", vbInformation, Me.Caption
                    End Select
                    lbSigue = False
                End If
                'debb-24/06/2015
                If lbSigue = True And oRsTmp.Fields!idTipoServicio = sghEmergenciaConsultorios Then
                   Set oRsTmp1 = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarPorNroCuenta(Val(Me.txtNcuenta.Text))
                   If oRsTmp1.RecordCount = 0 Then
                      MsgBox "No han registrado DIAGNOSTICO", vbInformation, Me.Caption
                      lbSigue = False
                   End If
                   oRsTmp1.Close
                End If
                '
                If lbSigue Then
                    txtDatosDeCuenta.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp.Fields!NroHistoriaClinica)), False) & " - " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre) & " (Edad: " & Trim(Str(oRsTmp.Fields!Edad)) & " " & Trim(oRsTmp.Fields!tedad) & ")"
                    txtPlan.Text = "IAFA Act.: " & Trim(oRsTmp.Fields!dFuenteFinanciamiento) & " (" & Trim(mo_ReglasFacturacion.BuscaServicioActualDelPaciente(oRsTmp.Fields!IdServicioIngreso)) & ")"
                    ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(Val(txtNcuenta.Text), CDate(lcBuscaParametro.RetornaFechaServidorSQL), lcBuscaParametro.RetornaHoraServidorSQL)
                    ml_idCuentaAtencion = Val(Me.txtNcuenta.Text)
                    ml_IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
                    Me.UcRecetas1.idTipoSexo = oRsTmp!idTipoSexo
                    Me.UcRecetas1.idTipoFinanciamiento = ml_IdTipoFinanciamiento
                    Me.UcRecetas1.DatoCabeceraReceta = "(N° Cuenta=" & Trim(txtNcuenta.Text) & ") Paciente=" & txtDatosDeCuenta.Text
                    Me.UcRecetas1.idCuentaAtencion = Val(Me.txtNcuenta.Text)
                    
                    Set oRsDx = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(oRsTmp!idAtencion, IIf(oRsTmp!idTipoServicio = 1, 1, sghHospitalizacionIngreso), oConexion)
                    Me.UcRecetas1.ActualizaDxEnGrilla oRsDx
                    
                    FraDetalle.Enabled = True
                    frmCabecera.Enabled = False
                End If
            End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       'oConexion.Close
      ' Set oConexion = Nothing
       Set oRsTmp1 = Nothing
   End If
End Sub

Function DevuelveNroRecetasGeneradas() As String
    DevuelveNroRecetasGeneradas = ""
    If lnRecetaRayosX > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Rayos X: " & Trim(Str(lnRecetaRayosX))
    End If
    If lnRecetaEcografiaO > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Ecografía Obstétrica: " & Trim(Str(lnRecetaEcografiaO))
    End If
    If lnRecetaEcografiaG > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Ecografía General: " & Trim(Str(lnRecetaEcografiaG))
    End If
    If lnRecetaTomografia > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Tomografía: " & Trim(Str(lnRecetaTomografia))
    End If
    If lnRecetaAnatomiaP > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Anatomía Patológica: " & Trim(Str(lnRecetaAnatomiaP))
    End If
    If lnRecetaPatologiaC > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Patológia Clínica: " & Trim(Str(lnRecetaPatologiaC))
    End If
    If lnRecetaBancoS > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Banco de Sangre: " & Trim(Str(lnRecetaBancoS))
    End If
    If lnRecetaFarmacia > 0 Then
       DevuelveNroRecetasGeneradas = DevuelveNroRecetasGeneradas & Chr(13) & "N° Receta para: Farmacia: " & Trim(Str(lnRecetaFarmacia))
    End If
End Function

Sub CargaDxParaFarmacia(oRsDx As Recordset)
    UcRecetas1.ActualizaDxEnGrilla oRsDx
End Sub


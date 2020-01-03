VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form SolicitudHistoriaDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBuscarPaciente 
      Caption         =   "..."
      Height          =   315
      Left            =   2805
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   570
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarServicio 
      Caption         =   "..."
      Height          =   315
      Left            =   2790
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   960
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarRespSolicita 
      Caption         =   "..."
      Height          =   315
      Left            =   2580
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2085
      Width           =   315
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   75
      TabIndex        =   13
      Top             =   3060
      Width           =   8310
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SolicitudHistoriaDetalle.frx":0000
         DownPicture     =   "SolicitudHistoriaDetalle.frx":0460
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
         Left            =   2565
         Picture         =   "SolicitudHistoriaDetalle.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   255
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SolicitudHistoriaDetalle.frx":0D4A
         DownPicture     =   "SolicitudHistoriaDetalle.frx":120E
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
         Left            =   4110
         Picture         =   "SolicitudHistoriaDetalle.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   255
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3060
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   8310
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
         Left            =   1470
         TabIndex        =   0
         Top             =   195
         Width           =   3645
      End
      Begin VB.TextBox txtIdEmpleadoSolicita 
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
         Left            =   1470
         TabIndex        =   8
         Top             =   2070
         Width           =   975
      End
      Begin VB.TextBox txtNombreEmpleadoSolicita 
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
         Left            =   2910
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2085
         Width           =   3645
      End
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1470
         TabIndex        =   10
         Top             =   2445
         Width           =   6705
      End
      Begin VB.TextBox txtIdServicio 
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
         Left            =   1470
         TabIndex        =   3
         Top             =   960
         Width           =   1170
      End
      Begin VB.TextBox txtNombreServicio 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   5055
      End
      Begin VB.ComboBox cmbIdMotivo 
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
         Left            =   5025
         TabIndex        =   9
         Top             =   1335
         Width           =   3150
      End
      Begin MSMask.MaskEdBox txtHoraSolicitud 
         Height          =   315
         Left            =   2910
         TabIndex        =   5
         Top             =   1335
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaSolicitud 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   1335
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox txtNombrePaciente 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   5055
      End
      Begin VB.TextBox txtIdHistoriaClinica 
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
         Left            =   1470
         TabIndex        =   1
         Top             =   585
         Width           =   1185
      End
      Begin MSMask.MaskEdBox txtFechaRequerida 
         Height          =   315
         Left            =   1470
         TabIndex        =   6
         Top             =   1710
         Width           =   1380
         _ExtentX        =   2434
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
      Begin MSMask.MaskEdBox txtHoraRequerida 
         Height          =   315
         Left            =   2895
         TabIndex        =   7
         Top             =   1710
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo historia"
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
         Left            =   150
         TabIndex        =   24
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Solicitante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   23
         Top             =   2085
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Servicio destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Nº Historia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label lblIdMotivo 
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
         Height          =   315
         Left            =   4410
         TabIndex        =   16
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label lblFechaSolicitud 
         Caption         =   "Fecha solicitud"
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
         Left            =   150
         TabIndex        =   15
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Label lblFechaRequerida 
         Caption         =   "Fecha requerida"
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
         Left            =   150
         TabIndex        =   14
         Top             =   1725
         Width           =   1335
      End
   End
End
Attribute VB_Name = "SolicitudHistoriaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Solicitud de Historia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_HistoriasSolicitadas As New DOHistoriaSolicitada
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdHistoriaSolicitada As Long
Dim mo_cmbIdMotivo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHEntidades.ListaDespleglable
Dim ml_IdMovimiento As Long
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       
    mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
    mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
    If mi_Opcion = sghAgregar Then
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriaSeleccionarDefinitivos(0)
    Else
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    End If
    
    mo_cmbIdMotivo.BoundColumn = "IdMotivo"
    mo_cmbIdMotivo.ListField = "DescripcionLarga"
    Set mo_cmbIdMotivo.RowSource = mo_AdminArchivoClinico.MotivosMovimientoHistoriaSeleccionarTodos()
    
    sMensaje = mo_AdminArchivoClinico.MensajeError
    
    If sMensaje <> "" Then
        MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption
    End If

End Sub
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdHistoriaSolicitada(lValue As Long)
   ml_IdHistoriaSolicitada = lValue
End Property
Property Get IdHistoriaSolicitada() As Long
   IdHistoriaSolicitada = ml_IdHistoriaSolicitada
End Property

Private Sub btnBuscarRespSolicita_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoSolicita, Me.txtNombreEmpleadoSolicita
End Sub

Private Sub btnBuscarServicio_Click()
    CompletarDatosDeServicio Me.txtIdServicio, Me.txtNombreServicio
End Sub

Private Sub cmbIdMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdMotivo
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdMotivo_LostFocus()
   If cmbIdMotivo.Text <> "" Then
       mo_cmbIdMotivo.BoundText = Val(Split(cmbIdMotivo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdMotivo
End Sub

Private Sub cmbIdMotivo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub btnBuscarPaciente_Click()
Dim oBusqueda As New SIGHNegocios.BuscaPacientes
Dim oDOPaciente As New doPaciente
Dim oConexion As New Connection
oConexion.Open SIGHEntidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarConHistoriasDefinitivas
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            Me.txtIdHistoriaClinica.Tag = oDOPaciente.idPaciente
            Me.txtIdHistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
            mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.IdTipoNumeracion
            Me.txtNombrePaciente = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdMotivo.MiComboBox = cmbIdMotivo
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
End Sub

Private Sub txtHoraRequerida_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraRequerida
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraRequerida_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraRequerida
End Sub

Private Sub txtHoraRequerida_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaRequerida_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaRequerida
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaRequerida_LostFocus()
   mo_Formulario.MarcarComoVacio txtFechaRequerida
End Sub

Private Sub txtFechaRequerida_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraSolicitud
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraSolicitud_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraSolicitud
End Sub

Private Sub txtHoraSolicitud_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitud
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaSolicitud_LostFocus()
   mo_Formulario.MarcarComoVacio txtFechaSolicitud
End Sub

Private Sub txtFechaSolicitud_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
 End Select

    mo_Formulario.HabilitarDeshabilitar Me.txtNombrePaciente, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreServicio, False

 Select Case mi_Opcion
     Case sghAgregar
        Me.txtFechaSolicitud.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        Me.txtHoraSolicitud.Text = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
     Case sghModificar
     Case sghConsultar
         
     Case sghEliminar
         
 End Select

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar solicitud de historia clínica"
       Case sghModificar
           Me.Caption = "Modificar solicitud de historia clínica"
       Case sghConsultar
           Me.Caption = "Consultar solicitud de historia clínica"
       Case sghEliminar
           Me.Caption = "Eliminar solicitud de historia clínica"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
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
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   If Val(mo_cmbIdMotivo.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdMotivo" + Chr(13)
   End If
   If Me.txtHoraRequerida.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de HoraRequerida" + Chr(13)
   End If
   If Me.txtFechaRequerida.Text = SIGHEntidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese el valor de FechaRequerida" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_HistoriasSolicitadas
           .IdHistoriaSolicitada = ml_IdHistoriaSolicitada
           .IdMovimiento = ml_IdMovimiento
           .idMotivo = mo_cmbIdMotivo.BoundText
           .HoraRequerida = Me.txtHoraRequerida.Text
           .FechaRequerida = Me.txtFechaRequerida.Text
           .HoraSolicitud = Me.txtHoraSolicitud.Text
           .FechaSolicitud = Me.txtFechaSolicitud.Text
           .idPaciente = Val(Me.txtIdHistoriaClinica.Tag)
           .idMotivo = Val(mo_cmbIdMotivo.BoundText)
           .Observacion = Me.txtObservacion.Text
           .IdServicio = Me.txtIdServicio.Tag
           .IdEmpleadoSolicita = Me.txtIdEmpleadoSolicita.Tag
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminArchivoClinico.HistoriasSolicitadasAgregar(mo_HistoriasSolicitadas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtIdHistoriaClinica.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminArchivoClinico.HistoriasSolicitadasModificar(mo_HistoriasSolicitadas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtIdHistoriaClinica.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminArchivoClinico.HistoriasSolicitadasEliminar(mo_HistoriasSolicitadas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtIdHistoriaClinica.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
        Dim oConexion As New Connection
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set mo_HistoriasSolicitadas = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarPorId(Me.IdHistoriaSolicitada)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
                MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption
                mb_ExistenDatos = False
                Exit Sub
        End If
        
        If Not mo_HistoriasSolicitadas Is Nothing Then
           With mo_HistoriasSolicitadas
                'Me.IdHistoriaSolicitada = .IdHistoriaSolicitada
                ml_IdMovimiento = .IdMovimiento
                mo_cmbIdMotivo.BoundText = .idMotivo
                Me.txtHoraRequerida.Text = .HoraRequerida
                Me.txtFechaRequerida.Text = .FechaRequerida
                Me.txtHoraSolicitud.Text = .HoraSolicitud
                Me.txtFechaSolicitud.Text = Format(.FechaSolicitud, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                mo_cmbIdMotivo.BoundText = .idMotivo
                Me.txtObservacion = .Observacion
                
                Dim oDOPaciente As New doPaciente
                Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(.idPaciente, oConexion)
                If Not oDOPaciente Is Nothing Then
                    Me.txtIdHistoriaClinica.Tag = .idPaciente
                    Me.txtIdHistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
                    mo_cmbIdTipoGenHistoriaClinica.BoundText = oDOPaciente.IdTipoNumeracion
                    Me.txtNombrePaciente = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
                End If
                
                Dim oDoServicio As New doServicio
                Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicio, oConexion)
                If Not oDoServicio Is Nothing Then
                    Me.txtIdServicio.Tag = oDoServicio.IdServicio
                    Me.txtIdServicio.Text = oDoServicio.Codigo
                    Me.txtNombreServicio = oDoServicio.Nombre
                Else
                    Me.txtIdServicio.Tag = ""
                    Me.txtIdServicio.Text = ""
                    Me.txtNombreServicio = ""
                End If
                
                Dim oDOEmpleado As New dOEmpleado
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoSolicita)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoSolicita.Tag = oDOEmpleado.IdEmpleado
                    Me.txtIdEmpleadoSolicita.Text = oDOEmpleado.CodigoPlanilla
                    Me.txtNombreEmpleadoSolicita = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                
                
                mb_ExistenDatos = True
                
                If ml_IdMovimiento <> 0 Then
                    MsgBox "La solicitud ya fue procesada, sólo podra consultar los datos", vbInformation, Me.Caption
                    Me.btnAceptar.Enabled = False
                End If
                
            End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       oConexion.Close
       Set oConexion = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasSolicitadas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdHistoriaSolicitada = 0
           mo_cmbIdMotivo.BoundText = ""
           Me.txtHoraRequerida.Text = SIGHEntidades.HORA_VACIA_HM
           Me.txtFechaRequerida.Text = SIGHEntidades.FECHA_VACIA_DMY
           Me.txtHoraSolicitud.Text = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
           Me.txtFechaSolicitud.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
           
           Me.txtIdHistoriaClinica.Text = ""
           Me.txtIdHistoriaClinica.Tag = ""
           
           Me.txtObservacion.Text = ""
           Me.txtNombrePaciente.Text = ""
           Me.txtNombreServicio.Text = ""
           Me.txtIdServicio.Text = ""
           Me.txtIdEmpleadoSolicita.Tag = ""
           Me.txtIdEmpleadoSolicita.Text = ""
           Me.txtNombreEmpleadoSolicita.Text = ""
   
End Sub


Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.HabilitarTipoServicio = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDoServicio Is Nothing Then
            txtIdServicio.Text = oDoServicio.Codigo
            txtIdServicio.Tag = oDoServicio.IdServicio
            lblDescripcionServicio.Text = oDoServicio.Nombre
        Else
            txtIdServicio.Text = ""
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing

End Sub



Private Sub txtIdEmpleadoSolicita_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoSolicita
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoSolicita_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoSolicita, Me.txtNombreEmpleadoSolicita
    mo_Formulario.MarcarComoVacio txtIdEmpleadoSolicita
End Sub

Private Sub txtIdEmpleadoSolicita_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtIdHistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdHistoriaClinica
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdHistoriaClinica_LostFocus()
Dim oDOPaciente As New doPaciente



    Set oDOPaciente = mo_AdminAdmision.PacientesObtenerPacientePorHistoriaClinica(Val(HCigualDNI_AgregaNUEVEaLaHistoria(Me.txtIdHistoriaClinica.Text)), Val(mo_cmbIdTipoGenHistoriaClinica.BoundText))
    Me.txtIdHistoriaClinica.Tag = oDOPaciente.idPaciente
    Me.txtNombrePaciente = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
    mo_Formulario.MarcarComoVacio txtIdHistoriaClinica
    
End Sub

Private Sub txtIdHistoriaClinica_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdServicio_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicio, Me.txtNombreServicio
    mo_Formulario.MarcarComoVacio txtIdServicio
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CompletarDatosResponsable(txtIdResponsable As TextBox, txtNombreResponsable As TextBox)
'Dim oBusqueda As New EmpleadosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    'oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtIdResponsable.Tag = oDOEmpleado.IdEmpleado
            txtIdResponsable.Text = oDOEmpleado.CodigoPlanilla
            txtNombreResponsable = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub

Sub CompletarDatosDeEmpleadoEnElLostFocus(txtCodigoPlanilla As TextBox, txtNombre As TextBox)
Dim oDOEmpleado As New dOEmpleado

        If mo_AdminComun.EmpleadosSeleccionarPorCodigo(txtCodigoPlanilla.Text, oDOEmpleado) Then
            txtCodigoPlanilla.Tag = oDOEmpleado.IdEmpleado
            txtCodigoPlanilla.Text = oDOEmpleado.CodigoPlanilla
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub


Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDoServicio As doServicio
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDoServicio Is Nothing Then
            txtIdServicio.Tag = oDoServicio.IdServicio
            lblDescripcionServicio.Text = oDoServicio.Nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio.Text = ""
        End If
   End If

End Sub



VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HistoriaClinicaDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "HistoriaClinicaDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   18
      Top             =   15
      Width           =   7440
      Begin VB.CommandButton btnBuscarPacientes 
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
         Left            =   6900
         Picture         =   "HistoriaClinicaDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   270
         Width           =   330
      End
      Begin VB.TextBox txtIdHistoriaClinicaAnt 
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
         Left            =   1635
         TabIndex        =   1
         Top             =   645
         Width           =   1065
      End
      Begin VB.ComboBox cmbIdTipoNumeracionHistoriaAnt 
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
         Left            =   2790
         TabIndex        =   2
         Top             =   645
         Width           =   4470
      End
      Begin VB.TextBox lblNombre 
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
         Left            =   1635
         TabIndex        =   0
         Top             =   270
         Width           =   5220
      End
      Begin VB.Label Label3 
         Caption         =   "Nº historia anterior"
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
         Left            =   90
         TabIndex        =   20
         Top             =   690
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Nombres"
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
         Left            =   120
         TabIndex        =   19
         Top             =   285
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   17
      Top             =   2925
      Width           =   7440
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HistoriaClinicaDetalle.frx":1254
         DownPicture     =   "HistoriaClinicaDetalle.frx":1718
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
         Left            =   3780
         Picture         =   "HistoriaClinicaDetalle.frx":1C04
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HistoriaClinicaDetalle.frx":20F0
         DownPicture     =   "HistoriaClinicaDetalle.frx":2550
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
         Left            =   2235
         Picture         =   "HistoriaClinicaDetalle.frx":29C5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   60
      TabIndex        =   11
      Top             =   1080
      Width           =   7440
      Begin VB.ComboBox cmbIdTipoHistoria 
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
         Left            =   4935
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   630
         Width           =   2340
      End
      Begin VB.ComboBox cmbIdEstadoHistoria 
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
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   975
         Width           =   2115
      End
      Begin VB.ComboBox cmbIdTipoNumeracionHistoria 
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
         Left            =   2805
         TabIndex        =   4
         Top             =   240
         Width           =   4485
      End
      Begin MSMask.MaskEdBox txtFechaPasoAPasivo 
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Top             =   1335
         Width           =   1410
         _ExtentX        =   2487
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
      Begin MSMask.MaskEdBox txtFechaCreacion 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   615
         Width           =   1440
         _ExtentX        =   2540
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
         Left            =   1650
         TabIndex        =   3
         Top             =   240
         Width           =   1065
      End
      Begin MSMask.MaskEdBox txtFultMovimiento 
         Height          =   315
         Left            =   5835
         TabIndex        =   21
         Top             =   1005
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   17
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo Movimiento"
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
         Left            =   4290
         TabIndex        =   22
         Top             =   1020
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Nº historia actual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   315
         Width           =   1395
      End
      Begin VB.Label lblFechaCreacion 
         Caption         =   "Fecha de creación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   655
         Width           =   1545
      End
      Begin VB.Label lblFechaPasoAPasivo 
         AutoSize        =   -1  'True
         Caption         =   "Paso a pasivo"
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
         TabIndex        =   14
         Top             =   1365
         Width           =   1080
      End
      Begin VB.Label lblIdTipoHistoria 
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
         Height          =   315
         Left            =   3855
         TabIndex        =   13
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label lblIdEstadoHistoria 
         Caption         =   "Estado de historia"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1005
         Width           =   1545
      End
   End
End
Attribute VB_Name = "HistoriaClinicaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca temporales creados y genera Historia Clinica final
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_IdUsuario As Long
Dim ml_idPaciente As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mb_ExistenDatos As Boolean
Dim mo_HistoriasClinicas As New DOHistoriaClinica
Dim ml_IdHistoriaClinica As Long
Dim mo_cmbIdEstadoHistoria As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoHistoria As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoNumeracionHistoria As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoNumeracionHistoriaAnt As New sighentidades.ListaDespleglable
Dim mo_lnIdTablaLISTBARITEMS As Long, lnIdHistoriaCl As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
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
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let IdHistoriaClinica(lValue As Long)
   ml_IdHistoriaClinica = lValue
End Property
Property Get IdHistoriaClinica() As Long
   IdHistoriaClinica = ml_IdHistoriaClinica
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       mo_cmbIdEstadoHistoria.BoundColumn = "IdEstadoHistoria"
       mo_cmbIdEstadoHistoria.ListField = "DescripcionLarga"
       Set mo_cmbIdEstadoHistoria.RowSource = mo_AdminArchivoClinico.EstadosHistoriaClinicaSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
       
       mo_cmbIdTipoHistoria.BoundColumn = "IdTipoHistoria"
       mo_cmbIdTipoHistoria.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoHistoria.RowSource = mo_AdminArchivoClinico.TiposHistoriaClinicaSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
       
        mo_cmbIdTipoNumeracionHistoria.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoNumeracionHistoria.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoNumeracionHistoria.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriaSeleccionarDefinitivos(0)
        sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
        
        mo_cmbIdTipoNumeracionHistoriaAnt.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoNumeracionHistoriaAnt.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoNumeracionHistoriaAnt.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
        
       If sMensaje <> "" Then
           MsgBox sMensaje, vbInformation, Me.Caption
       End If


End Sub

Private Sub btnBuscarPacientes_Click()
Dim oBusqueda As New SIGHNegocios.BuscaPacientes
Dim oDOPaciente As New doPaciente
Dim oConexion As New Connection
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarConHistoriasTemporales
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            Me.txtIdHistoriaClinica.Tag = oDOPaciente.idPaciente
            Me.txtIdHistoriaClinicaAnt = oDOPaciente.NroHistoriaClinica
            mo_cmbIdTipoNumeracionHistoriaAnt.BoundText = oDOPaciente.idTipoNumeracion
            Me.lblNombre = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub cmbIdEstadoHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEstadoHistoria
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdEstadoHistoria_LostFocus()
   If cmbIdEstadoHistoria.Text <> "" Then
       mo_cmbIdEstadoHistoria.BoundText = Val(Split(cmbIdEstadoHistoria.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdEstadoHistoria
End Sub

Private Sub cmbIdEstadoHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoHistoria
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoHistoria_LostFocus()
   If cmbIdTipoHistoria.Text <> "" Then
       mo_cmbIdTipoHistoria.BoundText = Val(Split(cmbIdTipoHistoria.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoHistoria
End Sub

Private Sub cmbIdTipoHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmbIdTipoNumeracionHistoria_Click()
    
    Me.txtIdHistoriaClinica = ""
    
    Select Case mo_cmbIdTipoNumeracionHistoria.BoundText
    Case sghHistoriaDefinitivaManual
        mo_Formulario.HabilitarDeshabilitar Me.txtIdHistoriaClinica, True
    Case Else
        mo_Formulario.HabilitarDeshabilitar Me.txtIdHistoriaClinica, False
    End Select
End Sub
Private Sub cmbIdTipoNumeracionHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoNumeracionHistoria
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoNumeracionHistoria_LostFocus()
   If cmbIdTipoNumeracionHistoria.Text <> "" Then
       mo_cmbIdTipoNumeracionHistoria.BoundText = Val(Split(cmbIdTipoNumeracionHistoria.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoNumeracionHistoria
End Sub

Private Sub cmbIdTipoNumeracionHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub Form_Initialize()

    Set mo_cmbIdEstadoHistoria.MiComboBox = cmbIdEstadoHistoria
    Set mo_cmbIdTipoHistoria.MiComboBox = cmbIdTipoHistoria
    Set mo_cmbIdTipoNumeracionHistoria.MiComboBox = cmbIdTipoNumeracionHistoria
    Set mo_cmbIdTipoNumeracionHistoriaAnt.MiComboBox = cmbIdTipoNumeracionHistoriaAnt

End Sub

Private Sub txtFechaPasoAPasivo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaPasoAPasivo
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaPasoAPasivo_LostFocus()

       If txtFechaPasoAPasivo <> sighentidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaPasoAPasivo, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaPasoAPasivo = sighentidades.FECHA_VACIA_DMY
            End If
        End If

   mo_Formulario.MarcarComoVacio txtFechaPasoAPasivo
End Sub

Private Sub txtFechaPasoAPasivo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaCreacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaCreacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaCreacion_LostFocus()
       
       If txtFechaCreacion <> sighentidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaCreacion, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaCreacion = sighentidades.FECHA_VACIA_DMY
            End If
        End If
        
   mo_Formulario.MarcarComoVacio txtFechaCreacion
End Sub

Private Sub txtFechaCreacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdHistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdHistoriaClinica
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdHistoriaClinica_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
        'Valores por defecto
        mo_cmbIdEstadoHistoria.BoundText = 1
        mo_cmbIdTipoHistoria.BoundText = 1
        Me.txtFechaCreacion = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
     
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
         Frame3.Enabled = False
         Frame1.Enabled = False
     Case sghEliminar
         CargarDatosAlosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
            Me.Caption = "Agregar historias clínicas"
       Case sghModificar
            Me.Caption = "Modificar historias clínicas"
            Me.txtIdHistoriaClinica.Enabled = False
            Me.cmbIdTipoNumeracionHistoria.Enabled = False
            Me.btnBuscarPacientes.Enabled = False
            Me.Frame3.Enabled = False
       Case sghConsultar
            Me.Caption = "Consultar historias clínicas"
            Me.Frame1.Enabled = False
            Me.btnAceptar.Visible = False
       Case sghEliminar
            Me.Frame1.Enabled = False
            Me.Caption = "Eliminar historias clínicas"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       mo_Formulario.HabilitarDeshabilitar txtFultMovimiento, False
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasClinicas
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Se agregó correctamente el N° Historia Clínica " & mo_HistoriasClinicas.NroHistoriaClinica, vbInformation, Me.Caption
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
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
       If mo_AdminFacturacion.PacienteSePuedeEliminar(Me.idPaciente) Then
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       Else
                MsgBox "El paciente no se puede eliminar porque tiene Atenciones registradas", vbInformation, Me.Caption
       End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String
   
   ValidarDatosObligatorios = False
   
    If Val(Me.txtIdHistoriaClinica.Tag) = 0 Then
        sMensaje = sMensaje + "Por favor ingrese el paciente" + Chr(13)
    End If
    
   If mo_cmbIdEstadoHistoria.BoundText = 0 Then
       sMensaje = sMensaje + "Por favor ingrese el estado de la historia clínica" + Chr(13)
   End If
   If mo_cmbIdTipoHistoria.BoundText = 0 Then
       sMensaje = sMensaje + "Por favor ingrese el tipo de historia clínica" + Chr(13)
   End If
   If Me.txtFechaCreacion.Text = sighentidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Por favor ingrese la fecha de creacion" + Chr(13)
   End If
   If mo_cmbIdTipoNumeracionHistoria.BoundText = "" Then
        sMensaje = sMensaje + "Por favor ingrese la numeración de la historia clínica" + Chr(13)
   End If
   
    If Val(mo_cmbIdTipoNumeracionHistoria.BoundText) = 2 Then
        If Val(Me.txtIdHistoriaClinica.Tag) = 0 Then
            sMensaje = sMensaje + "Por favor Ingrese el valor de nro de historia clinica" + Chr(13)
        End If
    End If
    
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
Dim sMensaje As String
Dim oRsTmp As New Recordset

   ValidarReglas = False
   
    If Val(mo_cmbIdTipoNumeracionHistoria.BoundText) = sghHistoriaDefinitivaManual Then
        If Me.txtIdHistoriaClinica = "" Then
            sMensaje = sMensaje + "Ingrese el número de historia clínica" + Chr(13)
        End If
    End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
    If CDate(Me.txtFechaCreacion) > Date Then
        MsgBox "La fecha de creación no pueder ser mayor que la fecha de hoy", vbExclamation, Me.Caption
        Exit Function
    End If
    If Val(mo_cmbIdEstadoHistoria.BoundText) <> sghEstadosHistoria.sghActiva And Me.txtFechaPasoAPasivo = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "Tiene que registrar la fecha de paso a pasivo", vbExclamation, Me.Caption
        Exit Function
    End If
    If Val(mo_cmbIdEstadoHistoria.BoundText) = sghEstadosHistoria.sghActiva Then
       Me.txtFechaPasoAPasivo.Text = sighentidades.FECHA_VACIA_DMY
    End If
    If Me.txtFechaPasoAPasivo <> sighentidades.FECHA_VACIA_DMY Then
        If CDate(Me.txtFechaCreacion) > CDate(Me.txtFechaPasoAPasivo) Then
            MsgBox "La fecha de creación no pueder ser mayor que la fecha de paso a pasivo", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
   ValidarReglas = True
   Set oRsTmp = Nothing
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_HistoriasClinicas
           .idPaciente = Me.txtIdHistoriaClinica.Tag
           .IdEstadoHistoria = mo_cmbIdEstadoHistoria.BoundText
           .idTipoHistoria = mo_cmbIdTipoHistoria.BoundText
           .FechaPasoAPasivo = IIf(Me.txtFechaPasoAPasivo.Text = sighentidades.FECHA_VACIA_DMY, 0, Me.txtFechaPasoAPasivo.Text)
           .fechacreacion = Me.txtFechaCreacion.Text
           If mi_Opcion = sghAgregar Then
              .NroHistoriaClinica = Val(Me.txtIdHistoriaClinica)
           Else
              .NroHistoriaClinica = lnIdHistoriaCl
           End If
           .IdUsuarioAuditoria = Me.IdUsuario
           .idTipoNumeracion = Val(mo_cmbIdTipoNumeracionHistoria.BoundText)
           .IdTipoNumeracionAnterior = Val(mo_cmbIdTipoNumeracionHistoriaAnt.BoundText)
           .NroHistoriaClinicaAnterior = Me.txtIdHistoriaClinicaAnt
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    Dim oRsTmp As New Recordset
    Dim lbCreaHistoria As Boolean
    Dim lcSql As String
    Dim lnIdPacienteEncontrado As Long
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    lbCreaHistoria = True
    Set oRsTmp = mo_AdminArchivoClinico.HistoriaClinicaSeleccionarXhistoriaYtipoNumeracion(Val(txtIdHistoriaClinica.Text), Val(mo_cmbIdTipoNumeracionHistoria.BoundText))
    If oRsTmp.RecordCount > 0 Then
       lbCreaHistoria = False
       If MsgBox("El 'N° Historia actual' existe para el Paciente: " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre) & Chr(3) & Chr(13) & "el 'N° Historia anterior' desea cambiarla por el 'N° historia actual' ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
          lnIdPacienteEncontrado = oRsTmp.Fields!idPaciente
          oRsTmp.Close
          mo_AdminArchivoClinico.ActualizaIdPacienteEnTodasLasTablas lnIdPacienteEncontrado, Me.txtIdHistoriaClinica.Tag, Val(txtIdHistoriaClinica.Text), mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Val(txtIdHistoriaClinicaAnt), ml_IdUsuario
          AgregarDatos = True
       Else
          AgregarDatos = False
          oRsTmp.Close
       End If
    Else
        oRsTmp.Close
    End If
    
    If lbCreaHistoria = True Then
        CargaDatosAlObjetosDeDatos
        AgregarDatos = mo_AdminArchivoClinico.HistoriaClinicaAgregar(mo_HistoriasClinicas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lblNombre.Text)
    End If
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_AdminArchivoClinico.HistoriaClinicaModificar(mo_HistoriasClinicas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "HC.Anterior: " & Trim(txtIdHistoriaClinicaAnt.Text) & " " & lblNombre.Text)
   
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
   
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_AdminArchivoClinico.HistoriaClinicaEliminar(mo_HistoriasClinicas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "HC.Anterior: " & Trim(txtIdHistoriaClinicaAnt.Text) & " " & lblNombre.Text)
   
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
Dim oDOPaciente As New doPaciente
Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set mo_HistoriasClinicas = mo_AdminArchivoClinico.HistoriaClinicaSeleccionarPorId(Me.IdHistoriaClinica)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
       If Not mo_HistoriasClinicas Is Nothing Then
           With mo_HistoriasClinicas
                Me.idPaciente = .idPaciente
                mo_cmbIdEstadoHistoria.BoundText = .IdEstadoHistoria
                mo_cmbIdTipoHistoria.BoundText = .idTipoHistoria
                Me.txtFechaPasoAPasivo.Text = IIf(.FechaPasoAPasivo = 0, sighentidades.FECHA_VACIA_DMY, .FechaPasoAPasivo)
                Me.txtFechaCreacion.Text = IIf(.fechacreacion = 0, sighentidades.FECHA_VACIA_DMY, Format(.fechacreacion, sighentidades.DevuelveFechaSoloFormato_DMY))
                mo_cmbIdTipoNumeracionHistoria.BoundText = .idTipoNumeracion
                Me.txtIdHistoriaClinica.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(.NroHistoriaClinica)), False)
                lnIdHistoriaCl = .NroHistoriaClinica
                
                Me.txtIdHistoriaClinicaAnt.Text = .NroHistoriaClinicaAnterior
                mo_cmbIdTipoNumeracionHistoriaAnt.BoundText = .IdTipoNumeracionAnterior
                Me.txtFultMovimiento.Text = IIf(.FechaUltimoMovimiento = 0, sighentidades.FECHA_VACIA_DMY, Format(.FechaUltimoMovimiento, sighentidades.DevuelveFechaSoloFormato_DMY))
                mb_ExistenDatos = True
           
                Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(Me.idPaciente, oConexion)
                If Not oDOPaciente Is Nothing Then
                    Me.txtIdHistoriaClinica.Tag = oDOPaciente.idPaciente
                    'Me.txtIdHistoriaClinica = oDOPaciente.NroHistoriaClinica
                    'Me.cmbIdTipoNumeracionHistoria.BoundText = oDOPaciente.IdTipoNumeracion
                    Me.lblNombre = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
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
'   Descripción:    Seleccionar un registro unico de la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.idPaciente = 0
           mo_cmbIdTipoNumeracionHistoria.BoundText = ""
           mo_cmbIdEstadoHistoria.BoundText = 1
           mo_cmbIdTipoHistoria.BoundText = 1
           Me.txtFechaPasoAPasivo.Text = sighentidades.FECHA_VACIA_DMY
           Me.txtFechaCreacion.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
           Me.txtIdHistoriaClinica.Text = ""
           Me.txtIdHistoriaClinica.Tag = ""
           Me.lblNombre = ""
   
End Sub



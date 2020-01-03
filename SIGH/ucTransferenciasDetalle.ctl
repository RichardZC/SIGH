VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucTransferenciasDetalle 
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   ScaleHeight     =   3780
   ScaleWidth      =   9255
   Begin VB.Frame fraTransferencia 
      Height          =   2070
      Left            =   30
      TabIndex        =   11
      Top             =   -60
      Width           =   9180
      Begin VB.CommandButton btnBuscarMedicoTransf 
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
         Left            =   2400
         Picture         =   "ucTransferenciasDetalle.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   945
         Width           =   405
      End
      Begin VB.CommandButton btnBuscarServicioTransf 
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
         Left            =   2400
         Picture         =   "ucTransferenciasDetalle.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   585
         Width           =   405
      End
      Begin VB.CommandButton btnBuscarMedicoTransfOrigen 
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
         Left            =   2400
         Picture         =   "ucTransferenciasDetalle.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   405
      End
      Begin VB.CommandButton btnVerDisponibilidadCamaTransf 
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
         Left            =   5565
         Picture         =   "ucTransferenciasDetalle.ctx":109E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1275
         Width           =   315
      End
      Begin VB.TextBox txtIdMedicoOrdenaOrigen 
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
         Left            =   1485
         TabIndex        =   0
         Top             =   210
         Width           =   885
      End
      Begin VB.TextBox lblNombreMedicoOrigen 
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
         Left            =   2805
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   5415
      End
      Begin VB.TextBox lblNombreMedico 
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
         Left            =   2805
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   945
         Width           =   5415
      End
      Begin VB.TextBox lblNombreServicio 
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
         Left            =   2805
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   570
         Width           =   5415
      End
      Begin VB.TextBox txtNroCamaTransf 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4665
         TabIndex        =   8
         Top             =   1290
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtIdMedicoOrdenaTransf 
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
         Left            =   1485
         TabIndex        =   4
         Top             =   930
         Width           =   885
      End
      Begin VB.TextBox txtIdServicioTransferencia 
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
         Left            =   1485
         TabIndex        =   2
         Top             =   570
         Width           =   885
      End
      Begin MSMask.MaskEdBox txtHoraTransf 
         Height          =   315
         Left            =   2970
         TabIndex        =   7
         Top             =   1290
         Width           =   780
         _ExtentX        =   1376
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
      Begin MSMask.MaskEdBox txtFechaTransf 
         Height          =   315
         Left            =   1485
         TabIndex        =   6
         Top             =   1290
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
      Begin Threed.SSCommand btnAgregar 
         Height          =   465
         Left            =   6240
         TabIndex        =   9
         Top             =   1290
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "ucTransferenciasDetalle.ctx":1628
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand btnQuitar 
         Height          =   465
         Left            =   7560
         TabIndex        =   10
         Top             =   1290
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "ucTransferenciasDetalle.ctx":45B4
         Caption         =   "Quitar"
         PictureAlignment=   9
         ShapeSize       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Médico Ordena"
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
         TabIndex        =   17
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label lblNroCama 
         Caption         =   "Nro cama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3825
         TabIndex        =   15
         Top             =   1350
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label58 
         Caption         =   "Médico recibe"
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
         Left            =   180
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label59 
         Caption         =   "Servicio recibe"
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
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label62 
         Caption         =   "Fecha recep."
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
         Left            =   180
         TabIndex        =   12
         Top             =   1320
         Width           =   1485
      End
   End
   Begin UltraGrid.SSUltraGrid grdTransferencias 
      Height          =   1650
      Left            =   30
      TabIndex        =   16
      Top             =   2070
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   2910
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
      Caption         =   "Lista de transferencias"
   End
End
Attribute VB_Name = "ucTransferenciasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar Transferencias
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idUsuario As Long
Dim ml_idCuentaAtencion As Long
Dim ml_TipoServicio As sghTipoServicio

Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia

Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ms_MensajeError As String
Dim mrs_OcupacionCamas As New ADODB.Recordset
Dim mda_FechaIngreso As Date
Dim lcBuscaParametro As New SIGHDatos.Parametros

Dim mb_UsuarioEsMedico As Boolean
Dim ml_IdMedico As Long
Dim lbElServicioExigeCama As Boolean
Public Event UltimoServicioTransferido(lcUltimoCodigoServicio As String)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)

Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property
Property Get IdServicioUltimaTransferencia() As Long
'    On Error Resume Next
'    If Not (mrs_OcupacionCamas.EOF And mrs_OcupacionCamas.BOF) Then
'        mrs_OcupacionCamas.MoveLast
'        IdServicioUltimaTransferencia = mrs_OcupacionCamas!idServicio
'        mrs_OcupacionCamas.MoveFirst
'    Else
'        IdServicioUltimaTransferencia = 0
'    End If
    IdServicioUltimaTransferencia = getIdServicioUltimaTransferencia()
End Property
Property Get IdCamaUltimaTransferencia() As Long
    
    If Not (mrs_OcupacionCamas.EOF And mrs_OcupacionCamas.BOF) Then
        mrs_OcupacionCamas.MoveLast
        IdCamaUltimaTransferencia = IIf(IsNull(mrs_OcupacionCamas!idCama), 0, mrs_OcupacionCamas!idCama)
        mrs_OcupacionCamas.MoveFirst
    Else
        IdCamaUltimaTransferencia = 0
    End If
    
End Property

Private Sub btnBuscarMedicoTransfOrigen_Click()
    CompletarDatosDeMedico txtIdMedicoOrdenaOrigen, lblNombreMedicoOrigen, Val(getIdEspecialidadUltimaTransferencia())
    txtIdServicioTransferencia.SetFocus
End Sub

Private Sub btnBuscarMedicoTransf_Click()
    CompletarDatosDeMedico txtIdMedicoOrdenaTransf, lblNombreMedico, Val(lblNombreServicio.Tag)
    If lblNombreMedico.Text <> "" And lblNombreServicio.Text <> "" Then
       txtFechaTransf.Text = lcBuscaParametro.RetornaFechaServidorSQL
'       txtHoraTransf.Text = lcBuscaParametro.RetornaHoraServidorSQL
       txtHoraTransf.Text = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
    End If
    txtFechaTransf.SetFocus
End Sub

Private Sub btnVerDisponibilidadCamaTransf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        btnBuscarMedicoTransf_Click
    End If
End Sub

Private Sub grdTransferencias_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    RaiseEvent SePresionoTeclaEspecial(lnKeyCode)
End Sub

Private Sub txtFechaTransf_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaTransf
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaTransf_LostFocus()
       
       If txtFechaTransf <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaTransf, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Procedimientos"
                 txtFechaTransf = SIGHEntidades.FECHA_VACIA_DMY
            End If
        End If
        
   mo_Formulario.MarcarComoVacio txtFechaTransf
End Sub

Private Sub txtFechaTransf_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraTransf_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraTransf
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtHoraTransf_LostFocus()
    If txtHoraTransf <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.ValidaHora(txtHoraTransf) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, "Registro de transferencias"
             txtHoraTransf = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
    mo_Formulario.MarcarComoVacio txtHoraTransf
End Sub

Private Sub txtHoraTransf_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdMedicoOrdenaOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoOrdenaOrigen
    If KeyCode = vbKeyF1 Then
        btnBuscarMedicoTransfOrigen_Click
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtIdMedicoOrdenaOrigen_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdMedicoOrdenaOrigen_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoOrdenaOrigen, lblNombreMedicoOrigen
    mo_Formulario.MarcarComoVacio txtIdMedicoOrdenaOrigen
End Sub

Private Sub txtIdMedicoOrdenaTransf_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoOrdenaTransf
    If KeyCode = vbKeyF1 Then
        btnBuscarMedicoTransf_Click
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtIdMedicoOrdenaTransf_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoOrdenaTransf, lblNombreMedico
    mo_Formulario.MarcarComoVacio txtIdMedicoOrdenaTransf
End Sub

Private Sub txtIdMedicoOrdenaTransf_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub btnBuscarServicioTransf_Click()
    CompletarDatosDeServicio txtIdServicioTransferencia, lblNombreServicio
    txtIdMedicoOrdenaTransf.Tag = ""
    lblNombreMedico.Text = ""
    If lblNombreServicio.Text <> "" Then
       txtFechaTransf.Text = lcBuscaParametro.RetornaFechaServidorSQL
       txtHoraTransf.Text = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
    End If
    btnBuscarMedicoTransf.SetFocus
End Sub

Private Sub btnVerDisponibilidadCamaTransf_Click()
Dim oBusqueda As New CamasBusqueda
Dim oDOCama As New DOCama
Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.idTipoServicio = ml_TipoServicio
    oBusqueda.IdServicio = Val(txtIdServicioTransferencia.Tag)

    oBusqueda.Show 1
    
    
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOCama Is Nothing Then
            If Val(txtIdServicioTransferencia.Tag) = oDOCama.IdServicioUbicacionActual Then
                txtNroCamaTransf.Text = oDOCama.Codigo
                txtNroCamaTransf.Tag = oDOCama.idCama
            Else
                MsgBox "La cama seleccionada no pertenece a un servicio de emergencia", vbInformation, "Transferencia de paciente"
                txtNroCamaTransf.Text = ""
                txtNroCamaTransf.Tag = ""
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDOCama = Nothing
End Sub

Private Sub txtIdServicioTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioTransferencia
    If KeyCode = vbKeyF1 Then
        btnBuscarServicioTransf_Click
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtIdServicioTransferencia_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioTransferencia, lblNombreServicio
    mo_Formulario.MarcarComoVacio txtIdServicioTransferencia
End Sub

Private Sub txtIdServicioTransferencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub ValidarSiEsMedico()
On Error GoTo miError
    Dim oRsTmp As ADODB.Recordset
    Dim oReglasDeProgMedica As New ReglasDeProgMedica
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    mb_UsuarioEsMedico = False
    ml_IdMedico = 0
    
    Set oRsTmp = oReglasDeProgMedica.MedicosXidEmpleado(ml_idUsuario, oConexion)
    If oRsTmp.RecordCount > 0 Then
        mb_UsuarioEsMedico = True
        ml_IdMedico = oRsTmp!idMedico
    End If
    oConexion.Close
    Set oConexion = Nothing
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Mensaje"
    End If
End Sub

Public Function Inicializar()
    GenerarRecordsetTemporal
    'mgaray
    mo_Formulario.HabilitarDeshabilitar UserControl.lblNombreMedicoOrigen, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtIdMedicoOrdenaOrigen, False
    
    mo_Formulario.HabilitarDeshabilitar UserControl.lblNombreMedico, False
    mo_Formulario.HabilitarDeshabilitar UserControl.lblNombreServicio, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtIdServicioTransferencia, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtIdMedicoOrdenaTransf, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtNroCamaTransf, False
    ValidarSiEsMedico
End Function
Private Sub UserControl_Resize()
    
    lblNombreMedico.Width = UserControl.Width - 3200
    lblNombreMedicoOrigen.Width = UserControl.Width - 3200
    lblNombreServicio.Width = UserControl.Width - 3200
    fraTransferencia.Width = UserControl.Width - 20
    grdTransferencias.Width = UserControl.Width - 30
    grdTransferencias.Height = UserControl.Height - 2000 '1850
    
End Sub

Private Sub grdTransferencias_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdTransferencias.Bands(0).Columns("IdServicio").Hidden = True
    grdTransferencias.Bands(0).Columns("IdMedicoOrdena").Hidden = True
    grdTransferencias.Bands(0).Columns("IdMedicoOrdenaOrigen").Hidden = True
    grdTransferencias.Bands(0).Columns("FechaDesocupacion").Hidden = True
    grdTransferencias.Bands(0).Columns("HoraDesocupacion").Hidden = True
    grdTransferencias.Bands(0).Columns("IdCama").Hidden = True
    grdTransferencias.Bands(0).Columns("LlegoAlServicio").Hidden = True
    
    If ml_TipoServicio = sghEmergenciaConsultorios Then
        grdTransferencias.Bands(0).Columns("NroCama").Hidden = True
    End If
    
    grdTransferencias.Bands(0).Columns("FechaOcupacion").Header.Caption = "Fecha Transf"
    grdTransferencias.Bands(0).Columns("FechaOcupacion").Width = 1200
    
    grdTransferencias.Bands(0).Columns("HoraOcupacion").Header.Caption = "Hora Transf."
    grdTransferencias.Bands(0).Columns("HoraOcupacion").Width = 1000
    
    grdTransferencias.Bands(0).Columns("NroCama").Header.Caption = "Nro Cama"
    
    grdTransferencias.Bands(0).Columns("CodigoServicio").Header.Caption = "Servicio"
    grdTransferencias.Bands(0).Columns("CodigoServicio").Width = 1000

    grdTransferencias.Bands(0).Columns("NombreServicio").Header.Caption = ""
    grdTransferencias.Bands(0).Columns("NombreServicio").Width = 3500

    grdTransferencias.Bands(0).Columns("NombreMedicoOrigen").Header.Caption = "Médico Ordena"
    grdTransferencias.Bands(0).Columns("NombreMedicoOrigen").Width = 3500
    
    grdTransferencias.Bands(0).Columns("NombreMedico").Header.Caption = "Médico Recibe"
    grdTransferencias.Bands(0).Columns("NombreMedico").Width = 3500
    'mgaray
    Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdServicio", _
                                        "IdMedicoOrdena", "FechaDesocupacion", _
                                        "HoraDesocupacion", "IdCama", "FechaOcupacion", _
                                        "HoraOcupacion", "NroCama", "CodigoServicio", _
                                        "NombreServicio", "NombreMedico", "LlegoAlServicio", _
                                        "IdMedicoOrdenaOrigen", "NombreMedicoOrigen")
End Sub

Public Sub CargarDatosDeTransferencias(oConexion As Connection)
Dim rsOcupacion As New Recordset
    GenerarRecordsetTemporal
    Set rsOcupacion = mo_AdminAdmision.EstanciaHospitalariaSeleccionarPorAtencion(idAtencion, 1, oConexion)
    Do While Not rsOcupacion.EOF
        With mrs_OcupacionCamas
            .AddNew
            .Fields!IdServicio = rsOcupacion!IdServicio
            .Fields!CodigoServicio = rsOcupacion!CodigoServicio
            .Fields!NombreServicio = rsOcupacion!NombreServicio
            .Fields!IdMedicoOrdena = rsOcupacion!IdMedicoOrdena
            .Fields!NombreMedico = rsOcupacion!NombreMedico
            .Fields!FechaOcupacion = rsOcupacion!FechaOcupacion
            .Fields!HoraOcupacion = rsOcupacion!HoraOcupacion
            .Fields!FechaDesocupacion = 0
            .Fields!HoraDesocupacion = ""
            .Fields!idCama = rsOcupacion!idCama
            .Fields!NroCama = rsOcupacion!CodigoCama
            .Fields!LlegoAlServicio = rsOcupacion!LlegoAlServicio
            'mgaray
            .Fields!idMedicoOrdenaOrigen = IIf(IsNull(rsOcupacion!idMedicoOrdenaOrigen), 0, IsNull(rsOcupacion!idMedicoOrdenaOrigen))
            .Fields!NombreMedicoOrigen = IIf(IsNull(rsOcupacion!NombreMedicoOrigen), "", rsOcupacion!NombreMedicoOrigen)
        End With
        rsOcupacion.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores grdTransferencias, SIGHEntidades.GrillaConFilasBicolor
    
End Sub

Sub GenerarRecordsetTemporal()
    If mrs_OcupacionCamas.State = 1 Then Set mrs_OcupacionCamas = Nothing
    With mrs_OcupacionCamas
          .Fields.Append "FechaOcupacion", adDate
          .Fields.Append "HoraOcupacion", adChar, 5
          .Fields.Append "IdCama", adInteger, 4, adFldIsNullable
          .Fields.Append "NroCama", adChar, 5, adFldIsNullable
          .Fields.Append "IdServicio", adInteger
          .Fields.Append "CodigoServicio", adVarChar, 10
          .Fields.Append "NombreServicio", adVarChar, 100
          'mgaray
          .Fields.Append "IdMedicoOrdenaOrigen", adInteger, , adFldIsNullable
          .Fields.Append "IdMedicoOrdena", adInteger
          .Fields.Append "NombreMedicoOrigen", adVarChar, 150, adFldIsNullable
          .Fields.Append "NombreMedico", adVarChar, 150
          .Fields.Append "FechaDesocupacion", adDate, , adFldIsNullable
          .Fields.Append "HoraDesocupacion", adChar, 5, adFldIsNullable
          .Fields.Append "LlegoAlServicio", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
          .Sort = "FechaOcupacion,HoraOcupacion"
    End With
    Set grdTransferencias.DataSource = mrs_OcupacionCamas
    
End Sub
Sub LimpiarDatos()
    On Error GoTo errLimp
    txtIdServicioTransferencia.Text = ""
    lblNombreServicio.Text = ""
    txtIdMedicoOrdenaTransf.Text = ""
    lblNombreMedico.Text = ""
    txtFechaTransf.Text = SIGHEntidades.FECHA_VACIA_DMY
    txtHoraTransf.Text = SIGHEntidades.HORA_VACIA_HM
    txtNroCamaTransf.Text = ""
    'mgaray
    txtIdMedicoOrdenaOrigen.Text = ""
    lblNombreMedicoOrigen.Text = ""
    
    With mrs_OcupacionCamas
       If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
             .Delete
             .Update
             .MoveNext
          Loop
       End If
    End With
errLimp:
End Sub

Function ElServicioTieneCama() As Boolean
    Dim oConexion As New Connection
    Dim oRsTmp986 As New Recordset
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oRsTmp986 = mo_AdminHoteleria.CamasSeleccionarPorIdServicio(Val(txtIdServicioTransferencia.Tag), oConexion)
    If oRsTmp986.RecordCount > 0 Then
       ElServicioTieneCama = True
    Else
       ElServicioTieneCama = False
    End If
    oRsTmp986.Close
    oConexion.Close
    Set oRsTmp986 = Nothing
    Set oConexion = Nothing
End Function

Private Sub btnAgregar_Click()
    
    If txtIdMedicoOrdenaOrigen.Tag = "" Then
        MsgBox "Por favor ingrese el médico que ordena la transferencia", vbInformation, "Trasnferencias"
        Exit Sub
    End If
    
    If txtIdMedicoOrdenaTransf.Tag = "" Then
        MsgBox "Por favor ingrese el médico que Recibe la transferencia", vbInformation, "Trasnferencias"
        Exit Sub
    End If
    
    If txtIdServicioTransferencia.Tag = "" Then
        MsgBox "Por favor ingrese el servicio hacia donde se realiza la transferencia", vbInformation, "Trasnferencias"
        Exit Sub
    End If
    
    If txtFechaTransf = SIGHEntidades.FECHA_VACIA_DMY Or txtHoraTransf = SIGHEntidades.HORA_VACIA_HM Then
        MsgBox "Por favor ingrese la fecha y hora de la transferencia", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If
    
    If mda_FechaIngreso > CDate(UserControl.txtFechaTransf & " " & UserControl.txtHoraTransf.Text) Then
        MsgBox "La fecha de transferencia no puede ser menor que la fecha de ingreso", vbExclamation, "Ingreso de procedimientos"
        Exit Sub
    End If

    If CDate(UserControl.txtFechaTransf.Text & " " & UserControl.txtHoraTransf.Text) > lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos Then
        MsgBox "La fecha de transferencia no puede ser mayor que la fecha de hoy", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If
    If ElServicioTieneCama = False And lbElServicioExigeCama = True Then
        MsgBox "El Servicio que recibe no tiene CAMAS", vbExclamation, "Ingreso de exámenes"
        Exit Sub
    End If
    
    If mrs_OcupacionCamas.RecordCount > 0 Then
        mrs_OcupacionCamas.MoveFirst
    
        Dim mda_MaxFecha As Date
        mda_MaxFecha = mda_FechaIngreso    '"01/01/1900"
        Do While Not mrs_OcupacionCamas.EOF
            If mda_MaxFecha < CDate(mrs_OcupacionCamas!FechaOcupacion & " " & mrs_OcupacionCamas!HoraOcupacion) Then
                mda_MaxFecha = CDate(mrs_OcupacionCamas!FechaOcupacion & " " & mrs_OcupacionCamas!HoraOcupacion)
            End If
            mrs_OcupacionCamas.MoveNext
        Loop
    
    End If
    
    'mgaray
    If EsMismoServicioDelPaciente(txtIdServicioTransferencia.Tag) = True Then
        MsgBox "Paciente ya se encuentra en el servicio", vbExclamation, "Registro de transferencias"
        Exit Sub
    End If
    
    If CDate(txtFechaTransf + " " + txtHoraTransf) < mda_MaxFecha Then
        MsgBox "La fecha de transferencia actual no puede ser menor que la ultima fecha de transferencia registrada", vbExclamation, "Registro de transferencias"
        Exit Sub
    End If
    
    With mrs_OcupacionCamas
        .AddNew
        .Fields!IdServicio = txtIdServicioTransferencia.Tag
        .Fields!CodigoServicio = txtIdServicioTransferencia.Text
        .Fields!NombreServicio = lblNombreServicio
        
        .Fields!IdMedicoOrdena = txtIdMedicoOrdenaTransf.Tag
        .Fields!NombreMedico = lblNombreMedico
        
        .Fields!FechaOcupacion = txtFechaTransf
        .Fields!HoraOcupacion = txtHoraTransf
        .Fields!FechaDesocupacion = 0
        .Fields!HoraDesocupacion = ""
        .Fields!idCama = Val(txtNroCamaTransf.Tag)
        .Fields!NroCama = txtNroCamaTransf.Text
        'mgaray09
        .Fields!idMedicoOrdenaOrigen = IIf(txtIdMedicoOrdenaOrigen.Tag = "", Null, txtIdMedicoOrdenaOrigen.Tag)
        .Fields!NombreMedicoOrigen = lblNombreMedicoOrigen.Text
        
        .Update
        .Sort = "FechaOcupacion,HoraOcupacion"
        .MoveLast
    End With
    SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "Transf: " & lblNombreServicio.Text
    RaiseEvent UltimoServicioTransferido(txtIdServicioTransferencia.Text)
    On Error Resume Next
    grdTransferencias.SetFocus
End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_OcupacionCamas
        If Not .EOF And Not .BOF Then
           'chequea si tiene consumos en SERVICIOS
           Dim oRsTmp9 As New Recordset
           Dim lbTieneConsumo As Boolean
           lbTieneConsumo = False
           Set oRsTmp9 = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdCuenta(ml_idCuentaAtencion)
           oRsTmp9.Filter = "idEstadoFacturacion=1 or idEstadoFacturacion=4"
           If oRsTmp9.RecordCount > 0 Then
              oRsTmp9.MoveFirst
              Do While Not oRsTmp9.EOF
                 If mrs_OcupacionCamas.Fields!IdServicio = oRsTmp9.Fields!idServicioPaciente And oRsTmp9.Fields!fechacreacion >= CDate(mrs_OcupacionCamas.Fields!FechaOcupacion & " " & mrs_OcupacionCamas.Fields!HoraOcupacion) Then
                    lbTieneConsumo = True
                    Exit Do
                 End If
                 oRsTmp9.MoveNext
              Loop
           End If
           If lbTieneConsumo = True Then
              MsgBox "No se puede ELIMINAR porque tiene consumos registrados en SERVICIOS", vbInformation, "Mensaje"
              Exit Sub
           Else
              'chequea si tiene consumos en FARMACIA
               Set oRsTmp9 = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarPorIdCuentaAtencion(ml_idCuentaAtencion)
               oRsTmp9.Filter = "idEstadoMovimiento=1"
               If oRsTmp9.RecordCount > 0 Then
                    oRsTmp9.MoveFirst
                    Do While Not oRsTmp9.EOF
                       If mrs_OcupacionCamas.Fields!IdServicio = oRsTmp9.Fields!idServicioPaciente And oRsTmp9.Fields!fechacreacion >= CDate(mrs_OcupacionCamas.Fields!FechaOcupacion & " " & mrs_OcupacionCamas.Fields!HoraOcupacion) Then
                          lbTieneConsumo = True
                          Exit Do
                       End If
                       oRsTmp9.MoveNext
                    Loop
               End If
               'mgaray
               If lbTieneConsumo = True Then
                    MsgBox "No se puede ELIMINAR porque tiene consumos registrados en FARMACIA", vbInformation, "Mensaje"
                    Exit Sub
                End If
           End If
           Set oRsTmp9 = Nothing
           Dim lIdMedicoOrdenaOrigen  As Long
           lIdMedicoOrdenaOrigen = 0
           If Not (IsNull(mrs_OcupacionCamas!idMedicoOrdenaOrigen)) Then
                lIdMedicoOrdenaOrigen = mrs_OcupacionCamas!idMedicoOrdenaOrigen
           End If
           If lIdMedicoOrdenaOrigen <> 0 Then
                If mb_UsuarioEsMedico = True And ml_IdMedico <> lIdMedicoOrdenaOrigen Then
                     MsgBox "No se puede ELIMINAR Transferencia ordenada por otro médico", vbInformation, "Mensaje"
                     Exit Sub
                End If
           End If
           
           Dim oRsUltimaTransferecia As ADODB.Recordset
          
'            Set oRsUltimaTransferecia = mrs_OcupacionCamas
'            If oRsUltimaTransferecia.RecordCount > 1 Then
'                  oRsUltimaTransferecia.MoveLast
'                  If Not (oRsUltimaTransferecia!FechaOcupacion = mrs_OcupacionCamas!FechaOcupacion _
'                          And oRsUltimaTransferecia!HoraOcupacion = mrs_OcupacionCamas!HoraOcupacion _
'                          And oRsUltimaTransferecia!IdServicio = mrs_OcupacionCamas!IdServicio) Then
'                        MsgBox "Solo puede eliminar la ùltima Transferencia Registrada", vbInformation, "Mensaje"
'                        Exit Sub
'                  End If
'            End If
           
           '
           .Delete
           .Update
           .Sort = "FechaOcupacion,HoraOcupacion"
           If .RecordCount > 0 Then
              .MoveLast
              RaiseEvent UltimoServicioTransferido(.Fields!CodigoServicio)
           Else
              RaiseEvent UltimoServicioTransferido("")
           End If
        Else
            If .BOF And .EOF Then
                MsgBox "No existen transferencias para eliminar", vbInformation, "Mensaje"
            Else
                MsgBox "Seleccione la transferencia a eliminar", vbInformation, "Mensaje"
            End If
        End If
    End With
End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
Dim oFechaHOra As New FechaHora
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.IdEspecialidad = lIdEspecialidad
    'mgaray
    oBusqueda.idTipoServicio = ml_TipoServicio
    oBusqueda.FechaProgramada = IIf(IsDate(UserControl.txtFechaTransf.Text), UserControl.txtFechaTransf.Text, 0)
    oBusqueda.HoraProgramada = IIf(UserControl.txtHoraTransf.Text = oFechaHOra.HORA_VACIA_HM, "00:00", UserControl.txtHoraTransf.Text)
    
    oBusqueda.NoMuestraInactivos = True
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing

End Sub
Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection

    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New dOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo1(CStr(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.idMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
    End If
    
End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    lbElServicioExigeCama = True
    oBusqueda.idTipoServicio = ml_TipoServicio
    oBusqueda.HabilitarTipoServicio = False
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDoServicio Is Nothing Then
            If ml_TipoServicio = oDoServicio.idTipoServicio Then
                txtIdServicio.Text = oDoServicio.Codigo
                txtIdServicio.Tag = oDoServicio.IdServicio
                lblDescripcionServicio.Text = oDoServicio.nombre
                lblDescripcionServicio.Tag = oDoServicio.IdEspecialidad
                If ml_TipoServicio = sghEmergenciaConsultorios Then   '09/08/2011
                    
                    If oDoServicio.EsObservacionEmergencia = False Then
                      lbElServicioExigeCama = False
                    End If
                    
                    lblNroCama.Visible = False
                    txtNroCamaTransf.Visible = False
                    btnVerDisponibilidadCamaTransf.Visible = False
                    If oDoServicio.EsObservacionEmergencia = True Then
                       lblNroCama.Visible = True
                       txtNroCamaTransf.Visible = True
                       btnVerDisponibilidadCamaTransf.Visible = True
                    End If
                End If
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, "Transferencias"
                txtIdServicio.Text = ""
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
                lblDescripcionServicio.Tag = ""
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub
Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDoServicio As doServicio
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDoServicio Is Nothing Then
            If ml_TipoServicio = oDoServicio.idTipoServicio Then
                txtIdServicio.Tag = oDoServicio.IdServicio
                lblDescripcionServicio.Text = oDoServicio.nombre
                lblDescripcionServicio.Tag = oDoServicio.IdEspecialidad
            Else
                MsgBox "El servicio ingresado no pertenece es de emergencia", vbInformation, "Transferencias"
                txtIdServicio.Tag = ""
                lblDescripcionServicio.Text = ""
                lblDescripcionServicio.Tag = ""
            End If
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio.Text = ""
        End If
   End If

End Sub

Sub CargaTransferenciasAlObjetosDatos(oOcupacionCamas As Collection, oDOOcupacionIngreso As DOEstanciaHospitalaria, _
    sFechaEgreso As String, sHoraEgreso As String, lnLlegoServicioIngreso As Long, lnLlegoAlServicioTransferido As Long, _
    lnSecuenciaTransferido As Long, lnCamaAlServicioTransferido As Long, lnIdCamaEgreso As Long)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LAS OCUPACIONES
    '---------------------------------------------------------------------------------
    Dim oDOOcupacion As DOEstanciaHospitalaria
    Dim oDOOcupacionAnterior As DOEstanciaHospitalaria
    Dim iSecuencia As Integer
    Dim iFor As Integer
    Do While oOcupacionCamas.Count > 0
        iFor = oOcupacionCamas.Count
        If iFor > 0 Then
           For iSecuencia = 0 To iFor + 1
               On Error Resume Next
               oOcupacionCamas.Remove (iSecuencia)
               
           Next
        End If
    Loop
    iSecuencia = 1
    
    Set oDOOcupacion = New DOEstanciaHospitalaria
    Set oDOOcupacion = oDOOcupacionIngreso
    oDOOcupacion.IdEstanciaHospitalaria = 0
    oDOOcupacion.Secuencia = iSecuencia
    oDOOcupacion.FechaDesocupacion = 0
    oDOOcupacion.HoraDesocupacion = ""
    oDOOcupacion.IdUsuarioAuditoria = Me.idUsuario
    oDOOcupacion.DiasEstancia = 0
    oDOOcupacion.IdFacturacionServicio = 0
    oDOOcupacion.LlegoAlServicio = lnLlegoServicioIngreso
    
    oOcupacionCamas.Add oDOOcupacion
    Set oDOOcupacionAnterior = oDOOcupacion
    If mrs_OcupacionCamas.RecordCount > 0 Then
        mrs_OcupacionCamas.MoveFirst
        
        Do While Not mrs_OcupacionCamas.EOF
            iSecuencia = iSecuencia + 1
            
            oDOOcupacionAnterior.FechaDesocupacion = mrs_OcupacionCamas!FechaOcupacion
            oDOOcupacionAnterior.HoraDesocupacion = mrs_OcupacionCamas!HoraOcupacion
            oDOOcupacionAnterior.DiasEstancia = Format(DateDiff("h", CDate(Format(oDOOcupacionAnterior.FechaOcupacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY) + " " + Format(oDOOcupacionAnterior.HoraOcupacion, "hh:nn")), CDate(Format(oDOOcupacionAnterior.FechaDesocupacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY) + " " + Format(oDOOcupacionAnterior.HoraDesocupacion, "hh:nn"))) / 24, "#0.000")
        
            Set oDOOcupacion = New DOEstanciaHospitalaria
            oDOOcupacion.IdEstanciaHospitalaria = 0
            oDOOcupacion.idAtencion = 0 'Me.IdAtencion
            oDOOcupacion.IdServicio = mrs_OcupacionCamas!IdServicio
            oDOOcupacion.IdMedicoOrdena = mrs_OcupacionCamas!IdMedicoOrdena
            oDOOcupacion.idMedicoOrdenaOrigen = IIf(IsNull(mrs_OcupacionCamas!idMedicoOrdenaOrigen), 0, mrs_OcupacionCamas!idMedicoOrdenaOrigen)
            oDOOcupacion.Secuencia = iSecuencia
            oDOOcupacion.FechaOcupacion = mrs_OcupacionCamas!FechaOcupacion
            oDOOcupacion.HoraOcupacion = mrs_OcupacionCamas!HoraOcupacion
            oDOOcupacion.FechaDesocupacion = 0
            oDOOcupacion.HoraDesocupacion = ""
            oDOOcupacion.idCama = IIf(IsNull(mrs_OcupacionCamas!idCama), 0, mrs_OcupacionCamas!idCama)
            oDOOcupacion.IdFacturacionServicio = 0
            oDOOcupacion.IdUsuarioAuditoria = Me.idUsuario
            oDOOcupacion.DiasEstancia = 0
            If lnSecuenciaTransferido = iSecuencia Then
               oDOOcupacion.LlegoAlServicio = lnLlegoAlServicioTransferido
               oDOOcupacion.idCama = lnCamaAlServicioTransferido
            Else
               oDOOcupacion.LlegoAlServicio = IIf(IsNull(mrs_OcupacionCamas!LlegoAlServicio), 0, mrs_OcupacionCamas!LlegoAlServicio)
            End If
            oOcupacionCamas.Add oDOOcupacion
            
            Set oDOOcupacionAnterior = oDOOcupacion
            
            mrs_OcupacionCamas.MoveNext
        Loop
    
    End If
    If lnIdCamaEgreso > 0 Then
       oDOOcupacionAnterior.idCama = lnIdCamaEgreso
    End If
    oDOOcupacionAnterior.FechaDesocupacion = IIf(sFechaEgreso = SIGHEntidades.FECHA_VACIA_DMY, 0, Format(sFechaEgreso, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
    oDOOcupacionAnterior.HoraDesocupacion = IIf(sHoraEgreso = SIGHEntidades.HORA_VACIA_HM, "", sHoraEgreso)
    If oDOOcupacionAnterior.FechaDesocupacion <> 0 Then
        oDOOcupacionAnterior.DiasEstancia = Format(DateDiff("h", CDate(Format(oDOOcupacionAnterior.FechaOcupacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY) + " " + Format(oDOOcupacionAnterior.HoraOcupacion, "hh:nn")), CDate(Format(oDOOcupacionAnterior.FechaDesocupacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY) + " " + Format(oDOOcupacionAnterior.HoraDesocupacion, "hh:nn"))) / 24, "#0.000")
    End If

End Sub

Public Sub OcultarDatosDeCama()
    
    UserControl.lblNroCama.Visible = False
    UserControl.txtNroCamaTransf.Visible = False
    UserControl.btnVerDisponibilidadCamaTransf.Visible = False

End Sub

Private Sub txtNroCamaTransf_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroCamaTransf
    If KeyCode = vbKeyF1 Then
        btnVerDisponibilidadCamaTransf_Click
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtNroCamaTransf_LostFocus()
    CompletarDatosDeCamasEnElLostFocus txtNroCamaTransf
    mo_Formulario.MarcarComoVacio txtNroCamaTransf
End Sub

Private Sub txtNroCamaTransf_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CompletarDatosDeCamasEnElLostFocus(txtNroCama As TextBox)
    
    txtNroCama.Tag = ""
    txtNroCama.Text = UCase(txtNroCama.Text)
    
    If txtNroCama.Text <> "" Then
        Dim oDOCama As DOCama
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorCodigo(txtNroCama.Text)
            If oDOCama Is Nothing Then
                MsgBox "El código ingresado no es válido", vbInformation, "Busqueda de camas"
            Else
                If Val(txtIdServicioTransferencia.Tag) = oDOCama.IdServicioUbicacionActual Then
                    txtNroCama.Tag = oDOCama.idCama
                Else
                    MsgBox "La cama seleccionada no pertenece al mismo servicio de ingreso", vbInformation, "Búsqueda de camas"
                    txtNroCama.Tag = ""
                End If
            End If
    End If

End Sub

'mgaray
Private Function EsMismoServicioDelPaciente(lIdServicioTransferido As Long) As Boolean
            On Error GoTo miError
    Dim EstaEnElServicio As Boolean
    Dim lIdUltimoServicioTransferido As Long
    
    lIdUltimoServicioTransferido = getIdServicioUltimaTransferencia()
    
    EstaEnElServicio = False
    
    If lIdUltimoServicioTransferido = 0 Then
        '
        Dim oDOAtencion  As New DOAtencion
        Dim oConexion As New Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.ConnectionTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion



        Set oDOAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_idAtencion, oConexion)
        If Not (oDOAtencion Is Nothing) Then
            If oDOAtencion.IdServicioIngreso = lIdServicioTransferido Then
                EstaEnElServicio = True
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
    Else
        If lIdUltimoServicioTransferido = lIdServicioTransferido Then
            EstaEnElServicio = True
        End If
    End If
'    lIdUltimoServicioTransferido = mrs_OcupacionCamas!IdServicio
    
    EsMismoServicioDelPaciente = EstaEnElServicio
miError:
    Err = 0
End Function

Public Function getIdServicioUltimaTransferencia() As Long
    Dim IdServicio As Long
    On Error Resume Next
    If Not (mrs_OcupacionCamas.EOF And mrs_OcupacionCamas.BOF) Then
        mrs_OcupacionCamas.MoveLast
        IdServicio = mrs_OcupacionCamas!IdServicio
        mrs_OcupacionCamas.MoveFirst
    Else
        IdServicio = 0
    End If
    getIdServicioUltimaTransferencia = IdServicio
End Function

Private Function getIdEspecialidadUltimaTransferencia() As Long
    Dim IdServicio As Long, lIdEspecialidad As Long
    On Error Resume Next
    
    lIdEspecialidad = 0
    
    If Not (mrs_OcupacionCamas.EOF And mrs_OcupacionCamas.BOF) Then
        mrs_OcupacionCamas.MoveLast
        IdServicio = mrs_OcupacionCamas!IdServicio
        mrs_OcupacionCamas.MoveFirst
    Else
        IdServicio = getIdServicioAtencion()
    End If
    
    If IdServicio > 0 Then
        Dim oDoServicio As New doServicio
        Set oDoServicio = getDatosDeServicio(IdServicio)
        If Not oDoServicio Is Nothing Then
            lIdEspecialidad = oDoServicio.IdEspecialidad
        End If
    End If
    getIdEspecialidadUltimaTransferencia = lIdEspecialidad
End Function

Private Function getDatosDeServicio(lIdServicio As Long) As doServicio
    Dim oDoServicio As New doServicio
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(lIdServicio, oConexion)
    
    Set getDatosDeServicio = oDoServicio
    oConexion.Close
    Set oConexion = Nothing
End Function

Private Function getIdServicioAtencion() As Long
On Error GoTo miError
    Dim oDOAtencion  As New DOAtencion
    Dim oConexion As New Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.ConnectionTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion

    getIdServicioAtencion = 0
    
    Set oDOAtencion = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_idAtencion, oConexion)
    If Not (oDOAtencion Is Nothing) Then
        getIdServicioAtencion = oDOAtencion.IdServicioIngreso
    End If
    oConexion.Close
miError:
    If Err Then
        MsgBox Err.Description, vbInformation, "Datos Atencion"
    End If
    Set oConexion = Nothing
End Function


Function getIdMedicoUltimaTransferencia() As Long
    Dim idMedico As Long
    On Error Resume Next
    If Not (mrs_OcupacionCamas.EOF And mrs_OcupacionCamas.BOF) Then
        mrs_OcupacionCamas.MoveLast
        idMedico = mrs_OcupacionCamas!IdMedicoOrdena
        mrs_OcupacionCamas.MoveFirst
    Else
        idMedico = 0
    End If
    getIdMedicoUltimaTransferencia = idMedico
End Function


Public Sub ColocarsePrimerRegistroTransferencia()
    If mrs_OcupacionCamas.RecordCount > 0 Then
        mrs_OcupacionCamas.MoveLast
    End If
End Sub

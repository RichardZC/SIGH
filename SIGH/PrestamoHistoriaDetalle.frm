VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form PrestamoHistoriaDetalle 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   Icon            =   "PrestamoHistoriaDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBuscarServicios 
      Caption         =   "..."
      Height          =   315
      Left            =   2940
      TabIndex        =   11
      Top             =   2430
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarPaciente 
      Caption         =   "..."
      Height          =   315
      Left            =   2940
      TabIndex        =   31
      Top             =   240
      Width           =   315
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   26
      Top             =   2805
      Width           =   7935
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "PrestamoHistoriaDetalle.frx":08CA
         DownPicture     =   "PrestamoHistoriaDetalle.frx":0D8E
         Height          =   700
         Left            =   4020
         Picture         =   "PrestamoHistoriaDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "PrestamoHistoriaDetalle.frx":1766
         DownPicture     =   "PrestamoHistoriaDetalle.frx":1BC6
         Height          =   700
         Left            =   2475
         Picture         =   "PrestamoHistoriaDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   7935
      Begin VB.TextBox lblNombreServicio 
         Height          =   315
         Left            =   3240
         TabIndex        =   30
         Top             =   2400
         Width           =   4470
      End
      Begin VB.ComboBox cmbIdTipoServicio 
         Height          =   315
         Left            =   1710
         TabIndex        =   9
         Top             =   2040
         Width           =   3165
      End
      Begin VB.ComboBox cmbIdEstadoPrestamo 
         Height          =   315
         Left            =   4920
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmbIdMotivo 
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox lblNombrePaciente 
         Height          =   315
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   4440
      End
      Begin MSMask.MaskEdBox txtHoraSolicitud 
         Height          =   315
         Left            =   2850
         TabIndex        =   2
         Top             =   600
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtObservacion 
         Height          =   1035
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   2805
      End
      Begin VB.TextBox txtIdServicios 
         Height          =   315
         Left            =   1710
         TabIndex        =   10
         Top             =   2400
         Width           =   1125
      End
      Begin VB.TextBox txtIdHistoriaClinica 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtFechaSolicitud 
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaPrestamoRequerida 
         Height          =   315
         Left            =   1710
         TabIndex        =   3
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaDevolucion 
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraPrestamoRequerida 
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Top             =   960
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraDevolucion 
         Height          =   315
         Left            =   2850
         TabIndex        =   8
         Top             =   1680
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaPrestamoReal 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraPrestamoReal 
         Height          =   315
         Left            =   2850
         TabIndex        =   6
         Top             =   1320
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha efectiva prestamo"
         Height          =   405
         Left            =   180
         TabIndex        =   28
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label15 
         Caption         =   "Servicio"
         Height          =   285
         Left            =   180
         TabIndex        =   27
         Top             =   2415
         Width           =   1185
      End
      Begin VB.Label Label14 
         Caption         =   "Tipo servicio"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   2070
         Width           =   1425
      End
      Begin VB.Label Label13 
         Caption         =   "Estado actual"
         Height          =   225
         Left            =   3750
         TabIndex        =   24
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Motivo"
         Height          =   315
         Left            =   3750
         TabIndex        =   23
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label Label11 
         Caption         =   "Observación"
         Height          =   195
         Left            =   3750
         TabIndex        =   22
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha devolución"
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   1770
         Width           =   1545
      End
      Begin VB.Label lblFechaCreacion 
         Caption         =   "Fecha solicitud"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   615
         Width           =   1485
      End
      Begin VB.Label lblFechaPasoAPasivo 
         Caption         =   "Fecha requerida prestamo"
         Height          =   405
         Left            =   180
         TabIndex        =   19
         Top             =   900
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Historia"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   330
         Width           =   945
      End
   End
End
Attribute VB_Name = "PrestamoHistoriaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POPrestamosHistoriaClinica
'        Autor: William Castro Grijalva
'        Fecha: 18/08/2004 12:26:58 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim mo_AdminServHosp As New ReglasServiciosHosp
Dim mb_ExistenDatos As Boolean
Dim mo_PrestamosHistoriaClinica As New DOPrestamoHistoriaClinica
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_AdminServiciosHosp As New ReglasServiciosHosp
Dim ml_IdEnvio As Long
Dim ml_IdPrestamo As Long
Dim ml_EtapaPrestamoHistoria As sghEtapaPrestamoHistoriaClinica
Dim ml_IdPaciente As Long
Dim mo_cmbIdMotivo As New SIGHComun.ListaDespleglable
Dim mo_cmbIdEstadoPrestamo As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoServicio As New SIGHComun.ListaDespleglable

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

        mo_cmbIdMotivo.BoundColumn = "IdMotivo"
        mo_cmbIdMotivo.ListField = "DescripcionLarga"
        Set mo_cmbIdMotivo.RowSource = mo_AdminArchivoClinico.MotivosPrestamoHistoriaSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
        
        mo_cmbIdEstadoPrestamo.BoundColumn = "IdEstadoPrestamo"
        mo_cmbIdEstadoPrestamo.ListField = "DescripcionLarga"
        Set mo_cmbIdEstadoPrestamo.RowSource = mo_AdminArchivoClinico.EstadosPrestamoHistoriaSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
       
        mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
        mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
       
        If sMensaje <> "" Then
            MsgBox sMensaje, vbCritical, Me.Caption
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
Property Let IdEnvio(lValue As Long)
   ml_IdEnvio = lValue
End Property
Property Get IdEnvio() As Long
   IdEnvio = ml_IdEnvio
End Property
Property Let IdPrestamo(lValue As Long)
   ml_IdPrestamo = lValue
End Property
Property Get IdPrestamo() As Long
   IdPrestamo = ml_IdPrestamo
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property

Property Let EtapaPrestamoHistoria(lValue As sghEtapaPrestamoHistoriaClinica)
   ml_EtapaPrestamoHistoria = lValue
End Property
Property Get EtapaPrestamoHistoria() As sghEtapaPrestamoHistoriaClinica
   EtapaPrestamoHistoria = ml_EtapaPrestamoHistoria
End Property


Private Sub btnBuscarPaciente_Click()
Dim oBusqueda As New PacientesBusqueda
Dim oDOPaciente As New doPaciente

    oBusqueda.TipoFiltro = sghFiltrarConHistoriasDefinitivas
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOPaciente Is Nothing Then
            Me.txtIdHistoriaClinica.Tag = oDOPaciente.IdPaciente
            Me.txtIdHistoriaClinica = oDOPaciente.NroHistoriaClinica
            Me.lblNombrePaciente = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
        End If
    End If
    
End Sub

Private Sub btnBuscarServicios_Click()
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicios.Text = oDOServicio.Codigo
            Me.txtIdServicios.Tag = oDOServicio.IdServicio
            Me.lblNombreServicio = oDOServicio.Nombre
        End If
    End If
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmbIdEstadoPrestamo_Click()

    If mi_Opcion = sghModificar Then
    If Me.cmbIdEstadoPrestamo.Tag <> "" Then
    If (Val(mo_cmbIdEstadoPrestamo.BoundText) = 1) Or (Val(mo_cmbIdEstadoPrestamo.BoundText) = 2 And Val(Me.cmbIdEstadoPrestamo.Tag) = 1) Then
        Me.txtFechaDevolucion.Enabled = False
        Me.txtHoraDevolucion.Enabled = False
        MsgBox "Sólo puede modificar al estado en prestamo (2), devuelto (3) o cancelado(4)" + Chr(13) + "Si desea volver al estado En Solicitud, debe anular el envío de la historia", vbInformation, Me.Caption
        mo_cmbIdEstadoPrestamo.BoundText = Me.cmbIdEstadoPrestamo.Tag
    End If
    
    If Val(mo_cmbIdEstadoPrestamo.BoundText) = 3 Then
        Me.txtFechaDevolucion.Enabled = False
        Me.txtHoraDevolucion.Enabled = False
        
        On Error Resume Next
        Me.txtFechaDevolucion.Text = Date 'Me.txtFechaDevolucion.Tag
        Me.txtHoraDevolucion.Text = Format(Now, "hh:mm") 'Me.txtHoraDevolucion.Tag
    
    End If
    
    If Val(mo_cmbIdEstadoPrestamo.BoundText) = 2 And Val(Me.cmbIdEstadoPrestamo.Tag) = 3 Then
        Me.txtFechaDevolucion.Enabled = True
        Me.txtHoraDevolucion.Enabled = True
        
        Me.txtFechaDevolucion.Text = "__/__/____"
        Me.txtHoraDevolucion.Text = "__:__"
    
    End If
    End If
    End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdMotivo.MiComboBox = cmbIdMotivo
    Set mo_cmbIdEstadoPrestamo.MiComboBox = cmbIdEstadoPrestamo
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
End Sub

Private Sub txtIdServicios_LostFocus()

    Me.txtIdServicios.Text = UCase(Me.txtIdServicios.Text)

   If Me.txtIdServicios.Text <> "" Then
    Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(Me.txtIdServicios.Text)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicios.Tag = oDOServicio.IdServicio
            Me.lblNombreServicio = oDOServicio.Nombre
            mo_cmbIdTipoServicio.BoundText = oDOServicio.IdTipoServicio
        Else
            Me.txtIdServicios.Tag = ""
            Me.lblNombreServicio = ""
            mo_cmbIdTipoServicio.BoundText = ""
        End If
   End If
   
   mo_Formulario.MarcarComoVacio txtIdServicios
End Sub
Private Sub txtIdServicios_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdServicios
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdServicios_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtObservacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtObservacion_LostFocus()
   mo_Formulario.MarcarComoVacio txtObservacion
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdHistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdHistoriaClinica
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdHistoriaClinica_LostFocus()
    
    If Val(txtIdHistoriaClinica.Text) <> 0 Then
        Dim oDOPaciente As doPaciente
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorHistoriaClinicaDefinitiva(Val(txtIdHistoriaClinica.Text))
        If Not oDOPaciente Is Nothing Then
            Me.txtIdHistoriaClinica.Tag = oDOPaciente.IdPaciente
            Me.txtIdHistoriaClinica = oDOPaciente.NroHistoriaClinica
            Me.lblNombrePaciente = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
            'Me.cmbIdTipoGeneracionHistoria.BoundText = oDOPAciente.IdTipoGeneracion
        Else
            MsgBox "El Nº de historia clínica no existe", vbExclamation, Me.Caption
            Me.txtIdHistoriaClinica.Tag = ""
            Me.lblNombrePaciente = ""
            'Me.cmbIdTipoGeneracionHistoria.BoundText = ""
        End If
    Else
        Me.txtIdHistoriaClinica.Tag = ""
        Me.lblNombrePaciente = ""
    End If
    
   mo_Formulario.MarcarComoVacio txtIdHistoriaClinica
End Sub

Private Sub txtIdHistoriaClinica_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
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


Private Sub cmbIdEstadoPrestamo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEstadoPrestamo
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdEstadoPrestamo_LostFocus()
   If cmbIdEstadoPrestamo.Text <> "" Then
       mo_cmbIdEstadoPrestamo.BoundText = Val(Split(cmbIdEstadoPrestamo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdEstadoPrestamo
End Sub

Private Sub cmbIdEstadoPrestamo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaDevolucion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaDevolucion
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaDevolucion_LostFocus()

    If txtFechaDevolucion <> SIGHComun.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaDevolucion, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
             txtFechaDevolucion = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
   mo_Formulario.MarcarComoVacio txtFechaDevolucion
End Sub

Private Sub txtFechaDevolucion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtHoraDevolucion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraDevolucion
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraDevolucion_LostFocus()
    
    If txtHoraDevolucion <> SIGHComun.HORA_VACIA_HM Then
        If Not SIGHComun.ValidaHora(txtHoraDevolucion) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraDevolucion = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
   mo_Formulario.MarcarComoVacio txtHoraDevolucion
End Sub

Private Sub txtHoraDevolucion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaPrestamoReal_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaPrestamoReal
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaPrestamoReal_LostFocus()

    If txtFechaPrestamoReal <> SIGHComun.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaPrestamoReal, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
             txtFechaPrestamoReal = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
   
   mo_Formulario.MarcarComoVacio txtFechaPrestamoReal
End Sub

Private Sub txtFechaPrestamoReal_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtHoraPrestamoReal_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraPrestamoReal
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraPrestamoReal_LostFocus()

    If txtHoraPrestamoReal <> SIGHComun.HORA_VACIA_HM Then
        If Not SIGHComun.ValidaHora(txtHoraPrestamoReal) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraPrestamoReal = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
   mo_Formulario.MarcarComoVacio txtHoraPrestamoReal
End Sub

Private Sub txtHoraPrestamoReal_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtFechaPrestamoRequerida_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaPrestamoRequerida
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaPrestamoRequerida_LostFocus()

    If txtFechaPrestamoRequerida <> SIGHComun.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaPrestamoRequerida, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
             txtFechaPrestamoRequerida = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
   mo_Formulario.MarcarComoVacio txtFechaPrestamoRequerida
End Sub

Private Sub txtFechaPrestamoRequerida_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtHoraPrestamoRequerida_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraPrestamoRequerida
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraPrestamoRequerida_LostFocus()

    If txtHoraPrestamoRequerida <> SIGHComun.HORA_VACIA_HM Then
        If Not SIGHComun.ValidaHora(txtHoraPrestamoRequerida) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraPrestamoRequerida = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
    
   mo_Formulario.MarcarComoVacio txtHoraPrestamoRequerida
End Sub

Private Sub txtHoraPrestamoRequerida_KeyPress(KeyAscii As Integer)
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

        If txtFechaSolicitud <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaSolicitud, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaSolicitud = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
   mo_Formulario.MarcarComoVacio txtFechaSolicitud
End Sub

Private Sub txtFechaSolicitud_KeyPress(KeyAscii As Integer)
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

    If txtHoraSolicitud <> SIGHComun.HORA_VACIA_HM Then
        If Not SIGHComun.ValidaHora(txtHoraSolicitud) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraSolicitud = SIGHComun.FECHA_VACIA_DMY
        End If
    End If
        
    mo_Formulario.MarcarComoVacio txtHoraSolicitud
    
End Sub

Private Sub txtHoraSolicitud_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select

 Select Case mi_Opcion
     Case sghAgregar
           Me.txtFechaSolicitud = Format(Now, "dd/mm/yyyy")
           Me.txtHoraSolicitud = Format(Now, "hh:mm")
           
           Me.txtFechaPrestamoRequerida = Format(Now, "dd/mm/yyyy")
           Me.txtHoraPrestamoRequerida = Format(Now, "hh:mm")
           
           Me.txtFechaSolicitud.Enabled = False
           Me.txtHoraSolicitud.Enabled = False
           
           Me.txtFechaPrestamoRequerida.Enabled = True
           Me.txtHoraPrestamoRequerida.Enabled = True
           
           Me.txtFechaPrestamoReal.Enabled = False
           Me.txtHoraPrestamoReal.Enabled = False
           
           Me.txtFechaDevolucion.Enabled = False
           Me.txtHoraDevolucion.Enabled = False
           
           Me.cmbIdEstadoPrestamo.Enabled = False
           mo_cmbIdEstadoPrestamo.BoundText = 1
     
     Case sghModificar
           Me.txtFechaSolicitud.Enabled = False
           Me.txtHoraSolicitud.Enabled = False
           
            Me.txtFechaPrestamoReal.Enabled = False
            Me.txtHoraPrestamoReal.Enabled = False
           
            Select Case mo_cmbIdEstadoPrestamo.BoundText
            Case 1
                Me.txtFechaPrestamoRequerida.Enabled = True
                Me.txtHoraPrestamoRequerida.Enabled = True
           
                Me.txtFechaDevolucion.Enabled = False
                Me.txtHoraDevolucion.Enabled = False
            Case 2
                Me.txtFechaPrestamoRequerida.Enabled = False
                Me.txtHoraPrestamoRequerida.Enabled = False
                
                Me.txtFechaDevolucion.Enabled = False
                Me.txtHoraDevolucion.Enabled = False
            
            End Select
           
     Case sghConsultar
            Me.Frame1.Enabled = False
            Me.btnAceptar.Enabled = False
            Me.btnBuscarPaciente.Enabled = False
            Me.btnBuscarServicios.Enabled = False
     Case sghEliminar
            Me.Frame1.Enabled = False
            Me.btnBuscarPaciente.Enabled = False
            Me.btnBuscarServicios.Enabled = False
    End Select

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar solicitud de prestamo de historia clinica"
       Case sghModificar
           Me.Caption = "Modificar prestamo historia clinica"
           Me.btnBuscarPaciente.Enabled = False
       Case sghConsultar
           Me.Caption = "Consultar prestamos historia clinica"
       Case sghEliminar
           Me.Caption = "Eliminar prestamos historia clinica"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla PrestamosHistoriaClinica
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
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
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
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
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
   
   If Me.txtIdHistoriaClinica.Tag = "" Then
       sMensaje = sMensaje + "Ingrese el nro de historia clínica" + Chr(13)
   End If
   If Val(mo_cmbIdMotivo.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el motivo" + Chr(13)
   End If
   If mo_cmbIdEstadoPrestamo.BoundText = 0 Then
       sMensaje = sMensaje + "Ingrese el estado del prestamo" + Chr(13)
   End If
   If Val(mo_cmbIdTipoServicio.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el tipo del servicio" + Chr(13)
   End If
   If Me.txtIdServicios.Text = "" Then
       sMensaje = sMensaje + "Ingrese el servicio. (Destino de la historia clínica)" + Chr(13)
   End If
   If Me.txtFechaPrestamoRequerida.Text = SIGHComun.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la fecha de prestamo requerida" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
   If CDate(Me.txtFechaSolicitud) > Date Then
        MsgBox "La fecha de solicitud no puede ser mayor que la fecha de hoy", vbExclamation, Me.Caption
        Exit Function
   End If
   
   If Me.txtFechaPrestamoReal <> SIGHComun.FECHA_VACIA_DMY Then
        If CDate(Me.txtHoraPrestamoReal) > Date Then
             MsgBox "La fecha de prestamo real no puede ser mayor que la fecha de hoy", vbExclamation, Me.Caption
             Exit Function
        End If
   End If
   
   If Me.txtFechaDevolucion <> SIGHComun.FECHA_VACIA_DMY Then
        If CDate(Me.txtFechaDevolucion) > Date Then
             MsgBox "La fecha de devolución real no puede ser mayor que la fecha de hoy", vbExclamation, Me.Caption
             Exit Function
        End If
   End If

    If Me.txtFechaDevolucion <> SIGHComun.FECHA_VACIA_DMY Then
        If Me.txtFechaPrestamoReal = SIGHComun.FECHA_VACIA_DMY Then
            MsgBox "No puede devolver si la fecha de prestamo real esta vacia!", vbInformation, Me.Caption
            Exit Function
        Else
        If CDate(Me.txtFechaDevolucion + " " + Me.txtHoraDevolucion) < CDate(Me.txtFechaPrestamoReal + " " + Me.txtHoraPrestamoReal) Then
            MsgBox "La fecha de devolución no puede ser menor que la fecha de prestamo real.", vbExclamation, Me.Caption
            Exit Function
        End If
        End If
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_PrestamosHistoriaClinica
   
           .Observacion = Me.txtObservacion.Text
           .IdEnvio = Me.IdEnvio
           .IdMotivo = mo_cmbIdMotivo.BoundText
           .IdEstadoPrestamo = mo_cmbIdEstadoPrestamo.BoundText
           
           .FechaSolicitud = IIf(Me.txtFechaSolicitud.Text = SIGHComun.FECHA_VACIA_DMY, 0, Me.txtFechaSolicitud.Text)
           .HoraSolicitud = IIf(Me.txtHoraSolicitud.Text = SIGHComun.HORA_VACIA_HM, 0, Me.txtHoraSolicitud.Text)
           
           .FechaPrestamoReal = IIf(Me.txtFechaPrestamoReal.Text = SIGHComun.FECHA_VACIA_DMY, 0, Me.txtFechaPrestamoReal.Text)
           .HoraPrestamoReal = IIf(Me.txtHoraPrestamoReal.Text = SIGHComun.HORA_VACIA_HM, 0, Me.txtHoraPrestamoReal.Text)
           
           .FechaDevolucion = IIf(Me.txtFechaDevolucion.Text = SIGHComun.FECHA_VACIA_DMY, 0, Me.txtFechaDevolucion.Text)
           .HoraDevolucion = IIf(Me.txtHoraDevolucion.Text = SIGHComun.HORA_VACIA_HM, 0, Me.txtHoraDevolucion.Text)
           
           .FechaPrestamoRequerida = IIf(Me.txtFechaPrestamoRequerida.Text = SIGHComun.FECHA_VACIA_DMY, 0, Me.txtFechaPrestamoRequerida.Text)
           .HoraPrestamoRequerida = IIf(Me.txtHoraPrestamoRequerida.Text = SIGHComun.HORA_VACIA_HM, 0, Me.txtHoraPrestamoRequerida.Text)

           .IdServicio = Val(Me.txtIdServicios.Tag)
           .IdPrestamo = Me.IdPrestamo
           
           .IdPaciente = Val(Me.txtIdHistoriaClinica.Tag)
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    AgregarDatos = mo_AdminArchivoClinico.PrestamosHistoriaClinicaAgregar(mo_PrestamosHistoriaClinica)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_AdminArchivoClinico.PrestamosHistoriaClinicaModificar(mo_PrestamosHistoriaClinica)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_AdminArchivoClinico.PrestamosHistoriaClinicaEliminar(mo_PrestamosHistoriaClinica)
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

        Set mo_PrestamosHistoriaClinica = mo_AdminArchivoClinico.PrestamosHistoriaClinicaSeleccionarPorId(Me.IdPrestamo)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbCritical, Me.Caption
             mb_ExistenDatos = False
             Exit Sub
        End If
        
        If Not mo_PrestamosHistoriaClinica Is Nothing Then
           With mo_PrestamosHistoriaClinica
           
                Me.txtObservacion.Text = .Observacion
                Me.IdEnvio = .IdEnvio
                mo_cmbIdMotivo.BoundText = .IdMotivo
                mo_cmbIdEstadoPrestamo.BoundText = .IdEstadoPrestamo
                Me.cmbIdEstadoPrestamo.Tag = .IdEstadoPrestamo
                
                Me.txtFechaSolicitud.Text = IIf(.FechaSolicitud <> 0, .FechaSolicitud, SIGHComun.FECHA_VACIA_DMY)
                Me.txtHoraSolicitud.Text = IIf(Val(.HoraSolicitud) <> 0, .HoraSolicitud, SIGHComun.HORA_VACIA_HM)
                
                Me.txtFechaDevolucion.Text = IIf(.FechaDevolucion <> 0, .FechaDevolucion, SIGHComun.FECHA_VACIA_DMY)
                Me.txtHoraDevolucion.Text = IIf(Val(.HoraDevolucion) <> 0, .HoraDevolucion, SIGHComun.HORA_VACIA_HM)
                
                Me.txtFechaDevolucion.Tag = Me.txtFechaDevolucion.Text
                Me.txtHoraDevolucion.Tag = Me.txtHoraDevolucion.Text
                
                
                Me.txtFechaPrestamoRequerida.Text = IIf(.FechaPrestamoRequerida <> 0, .FechaPrestamoRequerida, SIGHComun.FECHA_VACIA_DMY)
                Me.txtHoraPrestamoRequerida.Text = IIf(Val(.HoraPrestamoRequerida) <> 0, .HoraPrestamoRequerida, SIGHComun.HORA_VACIA_HM)
                
                Me.txtFechaPrestamoReal.Text = IIf(.FechaPrestamoReal <> 0, .FechaPrestamoReal, SIGHComun.FECHA_VACIA_DMY)
                Me.txtHoraPrestamoReal.Text = IIf(Val(.HoraPrestamoReal) <> 0, .HoraPrestamoReal, SIGHComun.HORA_VACIA_HM)
                
                
                Me.IdPrestamo = .IdPrestamo
                Me.IdPaciente = .IdPaciente
                
                Dim oDOPaciente As New doPaciente
                Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(.IdPaciente)
                If Not oDOPaciente Is Nothing Then
                    Me.lblNombrePaciente = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
                    Me.txtIdHistoriaClinica.Tag = oDOPaciente.IdPaciente
                    Me.txtIdHistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
                End If
                
                Me.txtIdServicios.Tag = .IdServicio
                Dim oDOServicio As New DOServicio
                Set oDOServicio = mo_AdminServHosp.ServiciosSeleccionarPorId(.IdServicio)
                If Not oDOServicio Is Nothing Then
                    Me.txtIdServicios = oDOServicio.Codigo
                    Me.lblNombreServicio = oDOServicio.Nombre
                    mo_cmbIdTipoServicio.BoundText = oDOServicio.IdTipoServicio
                End If
                
                mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla PrestamosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.txtObservacion.Text = ""
           Me.IdEnvio = 0
           Me.txtIdHistoriaClinica.Text = ""
           Me.lblNombrePaciente = ""
           Me.lblNombreServicio = ""
           mo_cmbIdMotivo.BoundText = ""
           mo_cmbIdEstadoPrestamo.BoundText = 1
           
           Me.txtFechaSolicitud.Text = Format(Now, "dd/mm/yyyy")
           Me.txtHoraSolicitud.Text = Format(Now, "hh:mm")
           
           Me.txtFechaDevolucion.Text = SIGHComun.FECHA_VACIA_DMY
           Me.txtHoraDevolucion.Text = SIGHComun.HORA_VACIA_HM
           
           Me.txtFechaPrestamoRequerida.Text = SIGHComun.FECHA_VACIA_DMY
           Me.txtHoraPrestamoRequerida.Text = SIGHComun.HORA_VACIA_HM
           
           Me.txtFechaPrestamoReal.Text = SIGHComun.FECHA_VACIA_DMY
           Me.txtHoraPrestamoReal.Text = SIGHComun.HORA_VACIA_HM
           
           Me.IdPrestamo = 0
   
End Sub



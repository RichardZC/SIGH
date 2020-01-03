VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmProEstablecimientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control en otro establecimiento"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmProEstablecimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Establecimiento donde se realizó el control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   7335
      Begin VB.TextBox txtCodigoEstablecimiento 
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
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscarEstablecimiento 
         Height          =   315
         Left            =   5880
         Picture         =   "frmProEstablecimientos.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox txtNombreEstablecimiento 
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
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
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
         TabIndex        =   9
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label 
         Caption         =   "Nombre"
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
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7335
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmProEstablecimientos.frx":2C55
         DownPicture     =   "frmProEstablecimientos.frx":30B5
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
         Left            =   2258
         Picture         =   "frmProEstablecimientos.frx":352A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmProEstablecimientos.frx":399F
         DownPicture     =   "frmProEstablecimientos.frx":3E63
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
         Left            =   3788
         Picture         =   "frmProEstablecimientos.frx":434F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin MSMask.MaskEdBox txtFechaControl 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "__/__/____"
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
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
   Begin VB.Label Label1 
      Caption         =   "Fecha de control"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmProEstablecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de los establecimientos de Salud de la Micro Red
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_lnIdTablaLISTBARITEMS As Long
Dim ms_lcNombrePc As String
Dim ml_IdEstablecimiento As Long
Dim ms_mesajeError As String

Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos           'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mr_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbEspecialidad As New sighentidades.ListaDespleglable

Dim oRcs_Servicios As New ADODB.Recordset                         'Representan los servicios de un establecimiento dado
Dim oRcs_ServiciosEstablecimiento As New ADODB.Recordset    'Representa los servicios regsitrados en la MR

Dim ms_LoginPC As String
Dim ml_idUsuario As Long                                'Indica el ID del Usuario que esta en session activa.
Dim mi_Opcion As sghOpciones
Dim mo_Establecimiento As DOEstablecimiento         'Representa el establecimiento actual
Dim ms_NombreEstablcimiento As String
Dim ms_CodigoEstablecimiento As String
Dim mc_FechaControl As String
Dim ml_IdPrograma As Long
Dim ml_IdProCabecera As Long
Dim ml_IdControl As Long
Dim mc_FechaUltimoControl As String
Dim mb_Aceptar As Boolean
Dim mi_BotonPresionado As sghBotonDetallePresionado
'========================== PROPIEDADES ===============================
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   ml_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get lnIdTablaLISTBARITEMS() As Long
   lnIdTablaLISTBARITEMS = ml_lnIdTablaLISTBARITEMS
End Property
Property Let lcNombrePc(sValue As String)
   ms_lcNombrePc = sValue
End Property
Property Get lcNombrePc() As String
   lcNombrePc = ms_lcNombrePc
End Property

Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let NombreEstablecimiento(lValue As String)
   ms_NombreEstablcimiento = lValue
End Property
Property Get NombreEstablecimiento() As String
   NombreEstablecimiento = ms_NombreEstablcimiento
End Property

Property Let CodigoEstablecimiento(lValue As String)
   ms_CodigoEstablecimiento = lValue
End Property
Property Get CodigoEstablecimiento() As String
   CodigoEstablecimiento = ms_CodigoEstablecimiento
End Property

Property Let FechaControl(lValue As String)
   mc_FechaControl = lValue
End Property
Property Get FechaControl() As String
   FechaControl = mc_FechaControl
End Property

Property Let IdPrograma(lValue As Long)
   ml_IdPrograma = lValue
End Property
Property Get IdPrograma() As Long
   IdPrograma = ml_IdPrograma
End Property

Property Let IdProCabecera(lValue As Long)
   ml_IdProCabecera = lValue
End Property
Property Get IdProCabecera() As Long
   IdProCabecera = ml_IdProCabecera
End Property

Property Let IdControl(lValue As Long)
   ml_IdControl = lValue
End Property
Property Get IdControl() As Long
   IdControl = ml_IdControl
End Property

Property Let FechaUltimoControl(lValue As String)
   mc_FechaUltimoControl = lValue
End Property
Property Get FechaUltimoControl() As String
   FechaUltimoControl = mc_FechaUltimoControl
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub btnCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

'========================== EVENTOS ===================================
Private Sub Form_Load()
    Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agregar Control Nº " & ml_IdControl
        Case sghModificar
            Me.Caption = "Modificar Control Nº " & ml_IdControl
        Case sghEliminar
            Me.Caption = "Eliminar Control Nº " & ml_IdControl
    End Select
    If mi_Opcion = sghAgregar Then limpiarIngresoControl
    If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
        CargarDatosAlFormulario
    End If
    'Me.txtFechaControl.SetFocus
End Sub

Sub limpiarIngresoControl()
    ml_IdEstablecimiento = 0
    Me.txtNombreEstablecimiento.Text = ""
    Me.txtCodigoEstablecimiento.Text = ""
    Me.txtFechaControl.Text = sighentidades.FECHA_VACIA_DMY
End Sub

Sub CargarDatosAlFormulario()
    Dim oDoEstablecimiento As New DOEstablecimiento
    Set oDoEstablecimiento = mr_ReglasComunes.EstablecimientosSeleccionarPorId(ml_IdEstablecimiento)
    If Not oDoEstablecimiento Is Nothing Then
        Set mo_Establecimiento = oDoEstablecimiento
        Me.txtNombreEstablecimiento.Text = mo_Establecimiento.Nombre
        Me.txtCodigoEstablecimiento.Text = mo_Establecimiento.Codigo
    End If
    Me.txtFechaControl.Text = mc_FechaControl
End Sub

Private Sub cmdBuscarEstablecimiento_Click()
Dim oForm As New SIGHNegocios.BuscaEstablecimientos
Dim oDoEstablecimiento As New DOEstablecimiento
Dim mo_RcsListaEstablecimientos  As New Recordset

oForm.DescripcionEstablecimiento = Me.txtNombreEstablecimiento.Text
oForm.NivelMaximoEstablecimiento = 0
oForm.MostrarFormulario

Me.btnAceptar.Enabled = True

If oForm.idRegistroSeleccionado = 0 Then
    Call MsgBox("No ha seleccionado ningún registro de la Lista.", vbExclamation, Me.Caption)
Else
    'Ingresando los valores del Establecimiento Elegido
    If oForm.BotonPresionado = sghAceptar Then
        Set oDoEstablecimiento = mr_ReglasComunes.EstablecimientosSeleccionarPorId(oForm.idRegistroSeleccionado)
        If Not oDoEstablecimiento Is Nothing Then
            Set mo_Establecimiento = oDoEstablecimiento
            ml_IdEstablecimiento = oDoEstablecimiento.IdEstablecimiento
            Me.txtNombreEstablecimiento.Text = mo_Establecimiento.Nombre
            Me.txtCodigoEstablecimiento.Text = mo_Establecimiento.Codigo
        End If
    End If
End If
End Sub

Private Sub btnAceptar_Click()
    If btnAceptar.Enabled = False Then
       Exit Sub
    End If
    Select Case mi_Opcion
    Case sghAgregar
        If ValidarDatosObligatorios() Then
            If ValidarReglas() Then
                If AgregarControl() Then
                    'Call MsgBox("Los datos fuerón agregados satisfactoriamente.", vbInformation, Me.Caption)
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    Call MsgBox("No se pudo agregar los datos.", vbCritical Or vbSystemModal, Me.Caption)
                End If
            End If
        End If
    Case sghModificar
        If ValidarDatosObligatorios() Then
            If ValidarReglas() Then
                If ModificarControl() Then
                    'Call MsgBox("Los datos fuerón modificados satisfactoriamente.", vbInformation, Me.Caption)
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    Call MsgBox("No se pudo modificar los datos.", vbCritical Or vbSystemModal, Me.Caption)
                End If
            End If
        End If
    Case sghEliminar
            If EliminarControl() Then
                Call MsgBox("Los datos fuerón eliminados satisfactoriamente.", vbInformation, Me.Caption)
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                Call MsgBox("No se pudo eliminar los datos.", vbCritical Or vbSystemModal, Me.Caption)
            End If
    End Select
    mi_BotonPresionado = sghAceptar
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

'========================== METODOS ===================================

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim mb_resultado As Boolean
'mb_resultado = True
ms_mesajeError = ""
ValidarDatosObligatorios = True
If ml_IdEstablecimiento = 0 Then
    MsgBox "No ha elegido ningún establecimiento.", vbInformation, Me.Caption
    ValidarDatosObligatorios = False
    Exit Function
End If
If txtFechaControl.Text = txtFechaControl.Tag Then
    MsgBox "No se ingreso la fecha de control.", vbInformation, Me.Caption
    ValidarDatosObligatorios = False
    Exit Function
End If
If Not IsDate(txtFechaControl.Text) Then
    MsgBox "La fecha de control no tiene el formato correcto.", vbInformation, Me.Caption
    ValidarDatosObligatorios = False
    Exit Function
End If
End Function

Function ValidarReglas()
    Dim mb_ValidacionReglas As Boolean
    ValidarReglas = True
    If mc_FechaUltimoControl <> "" Then
        If CDate(mc_FechaUltimoControl) < CDate(txtFechaControl.Text) Then
            MsgBox "La fecha del control ingresado no debe ser mayor al del ultimo control", vbExclamation, Me.Caption
            ValidarReglas = False
        End If
    End If
End Function

Function ActualizarDatos() As Boolean
If mi_Opcion = sghEliminar Then
    oRcs_ServiciosEstablecimiento.MoveFirst
    While Not oRcs_ServiciosEstablecimiento.EOF
        oRcs_ServiciosEstablecimiento!IdEstado = 3
        oRcs_ServiciosEstablecimiento.MoveNext
    Wend
End If
ActualizarDatos = mr_ReglasHIS.ActualizarServiciosPorEstablecimientos(ml_IdEstablecimiento, mi_Opcion, oRcs_ServiciosEstablecimiento)
End Function

Function AgregarControl() As Boolean
    mc_FechaControl = txtFechaControl.Text
    AgregarControl = True
End Function

Function ModificarControl() As Boolean
    mc_FechaControl = txtFechaControl.Text
    ModificarControl = True
End Function

Function EliminarControl() As Boolean
    ml_IdEstablecimiento = 0
    mc_FechaControl = sighentidades.FECHA_VACIA_DMY
    EliminarControl = True
End Function

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            cmdBuscarEstablecimiento_Click
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtCodigoEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoEstablecimiento
End Sub

Private Sub txtFechaControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaControl
End Sub

Private Sub txtNombreEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
    mo_Teclado.RealizarNavegacion KeyCode, txtNombreEstablecimiento
End Sub

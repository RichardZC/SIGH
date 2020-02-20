VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CamaDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "CamaDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   4125
      Left            =   4620
      TabIndex        =   25
      Top             =   30
      Width           =   7230
      Begin VB.CommandButton btnBuscarServicioUbicacionActual 
         Caption         =   "..."
         Height          =   315
         Left            =   2715
         TabIndex        =   11
         Top             =   1335
         Width           =   315
      End
      Begin VB.CommandButton btnBuscarServicios 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   2715
         TabIndex        =   8
         Top             =   975
         Width           =   315
      End
      Begin VB.TextBox txtIdServicioPropietario 
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
         Left            =   1770
         TabIndex        =   7
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtIdServicioUbicacionActual 
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
         Left            =   1770
         TabIndex        =   10
         Top             =   1350
         Width           =   885
      End
      Begin VB.TextBox lblNombreServicioPropietario 
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   990
         Width           =   2655
      End
      Begin VB.TextBox lblNombreServicioUbicacionActual 
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2655
      End
      Begin VB.CommandButton btnQuitar 
         DisabledPicture =   "CamaDetalle.frx":08CA
         DownPicture     =   "CamaDetalle.frx":0C55
         Height          =   315
         Left            =   2685
         Picture         =   "CamaDetalle.frx":0FE8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1830
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregar 
         DisabledPicture =   "CamaDetalle.frx":1379
         DownPicture     =   "CamaDetalle.frx":1762
         Height          =   315
         Left            =   1620
         Picture         =   "CamaDetalle.frx":1B6E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1830
         Width           =   1005
      End
      Begin UltraGrid.SSUltraGrid grdMovimientos 
         Height          =   1740
         Left            =   30
         TabIndex        =   15
         Top             =   2250
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   3069
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Movimientos de la cama"
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1770
         TabIndex        =   5
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   1770
         TabIndex        =   6
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label lblIdServicioUbicacionActual 
         Caption         =   "Ubicación actual"
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
         Left            =   105
         TabIndex        =   29
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Label lblIdServicioPropietario 
         Caption         =   "Servicio propietario"
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
         Left            =   105
         TabIndex        =   28
         Top             =   1020
         Width           =   1605
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Salida"
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
         Left            =   105
         TabIndex        =   27
         Top             =   660
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Ingreso "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   105
         TabIndex        =   26
         Top             =   285
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   24
      Top             =   4230
      Width           =   11805
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
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
         Left            =   6128
         Picture         =   "CamaDetalle.frx":1F7A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
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
         Left            =   4568
         Picture         =   "CamaDetalle.frx":2466
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   60
      TabIndex        =   18
      Top             =   30
      Width           =   4515
      Begin VB.TextBox txtCodigo 
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
         Left            =   2010
         MaxLength       =   5
         TabIndex        =   1
         Top             =   630
         Width           =   1000
      End
      Begin MSDataListLib.DataCombo cmbIdEstadoCama 
         Height          =   330
         Left            =   2010
         TabIndex        =   3
         Top             =   1350
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbIdCondicionOcupacion 
         Height          =   330
         Left            =   2010
         TabIndex        =   4
         Top             =   1710
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbIdTiposCama 
         Height          =   330
         Left            =   2010
         TabIndex        =   2
         Top             =   990
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbIdTipoServicio 
         Height          =   330
         Left            =   2010
         TabIndex        =   0
         Top             =   270
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo servicio"
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
         Left            =   210
         TabIndex        =   23
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblIdTiposCama 
         Caption         =   "Tipo cama"
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
         Left            =   210
         TabIndex        =   22
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label lblIdCondicionOcupacion 
         Caption         =   "Condición ocupación"
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
         Left            =   210
         TabIndex        =   21
         Top             =   1770
         Width           =   1665
      End
      Begin VB.Label lblIdEstadoCama 
         Caption         =   "Estado cama"
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
         Left            =   210
         TabIndex        =   20
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código cama"
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
         Left            =   210
         TabIndex        =   19
         Top             =   660
         Width           =   1185
      End
   End
End
Attribute VB_Name = "CamaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de CAMAS para Hospitalización y Emergencia
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_Camas As New DOCama
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_idTipoServicio  As Long
Dim ms_X As Long
Dim ms_Y As Long
Dim ml_IdCama As Long
Dim ml_ConfirmoOperacion As Long
Dim mo_CamasMovimiento As New sighComun.DOCamasMovimientos
Dim mrs_Movimientos As New ADODB.Recordset
Dim ml_IdServicioActual As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
'mgaray20141014
Const ESTADO_CAMA_OCUPADA = 3

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let IdCama(lValue As Long)
   ml_IdCama = lValue
End Property
Property Get IdCama() As Long
   IdCama = ml_IdCama
End Property

Property Let idTipoServicio(lValue As Long)
   ml_idTipoServicio = lValue
End Property
Property Get idTipoServicio() As Long
   idTipoServicio = ml_idTipoServicio
End Property

Property Let ConfirmoOperacion(lValue As Long)
   ml_ConfirmoOperacion = lValue
End Property
Property Get ConfirmoOperacion() As Long
   ConfirmoOperacion = ml_ConfirmoOperacion
End Property

Property Let IdTipoServicioActual(lValue As Long)
   ml_idTipoServicio = lValue
   Me.cmbIdTipoServicio.BoundText = ml_idTipoServicio
End Property



Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

        cmbIdTipoServicio.ListField = "DescripcionLarga"
        cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
        Set cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()

       cmbIdEstadoCama.BoundColumn = "IdEstadoCama"
       cmbIdEstadoCama.ListField = "DescripcionLarga"
       Set cmbIdEstadoCama.RowSource = mo_AdminHoteleria.EstadosCamaSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminHoteleria.MensajeError
       
       cmbIdTiposCama.BoundColumn = "IdTipoCama"
       cmbIdTiposCama.ListField = "DescripcionLarga"
       Set cmbIdTiposCama.RowSource = mo_AdminHoteleria.TiposCamaSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminHoteleria.MensajeError
       
       If sMensaje <> "" Then
           MsgBox mo_AdminHoteleria.MensajeError, vbInformation, Me.Caption
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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let Xini(sValue As Long)
   ms_X = sValue
End Property
Property Get Xini() As Long
   Xini = ms_X
End Property
Property Let Yini(sValue As Long)
   ms_Y = sValue
End Property
Property Get Yini() As Long
   Yini = ms_Y
End Property

Private Sub cmbIdEstadoCama_Change()
       'mgaray20141014
       If Val(cmbIdEstadoCama.BoundText) = ESTADO_CAMA_OCUPADA Then
            cmbIdCondicionOcupacion.Enabled = True
            cmbIdCondicionOcupacion.BoundColumn = "IdCondicionOcupacion"
            cmbIdCondicionOcupacion.ListField = "DescripcionLarga"
            Set cmbIdCondicionOcupacion.RowSource = mo_AdminHoteleria.TiposCondicionOcupacionSeleccionarTodos()
        Else
            cmbIdCondicionOcupacion.BoundText = ""
            cmbIdCondicionOcupacion.Enabled = False
        End If

End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
AdministrarKeyPreview KeyCode
End Sub

Private Sub btnBuscarServicios_Click()
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDOServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.idTipoServicio = Val(Me.cmbIdTipoServicio.BoundText)
    oBusqueda.HabilitarTipoServicio = False

    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOServicio Is Nothing Then
            If Val(Me.cmbIdTipoServicio.BoundText) = oDOServicio.idTipoServicio Then
                Me.txtIdServicioPropietario.Text = oDOServicio.codigo
                Me.txtIdServicioPropietario.Tag = oDOServicio.idServicio
                Me.lblNombreServicioPropietario = oDOServicio.nombre
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, Me.Caption
                Me.txtIdServicioPropietario.Text = ""
                Me.txtIdServicioPropietario.Tag = ""
                Me.lblNombreServicioPropietario = ""
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing

End Sub

Private Sub btnBuscarServicioUbicacionActual_Click()
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDOServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.idTipoServicio = Val(Me.cmbIdTipoServicio.BoundText)
    oBusqueda.HabilitarTipoServicio = False

    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOServicio Is Nothing Then
            If Val(Me.cmbIdTipoServicio.BoundText) = oDOServicio.idTipoServicio Then
                Me.txtIdServicioUbicacionActual.Text = oDOServicio.codigo
                Me.txtIdServicioUbicacionActual.Tag = oDOServicio.idServicio
                Me.lblNombreServicioUbicacionActual = oDOServicio.nombre
                ml_IdServicioActual = oDOServicio.idServicio
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, Me.Caption
                Me.txtIdServicioUbicacionActual.Text = ""
                Me.txtIdServicioUbicacionActual.Tag = ""
                Me.lblNombreServicioUbicacionActual = ""
                ml_IdServicioActual = 0
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub grdMovimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     grdMovimientos.Bands(0).Columns("IdServicio").Hidden = True
     grdMovimientos.Bands(0).Columns("Dservicio").Width = 4000
     grdMovimientos.Bands(0).Columns("dServicio").Header.Caption = "Servicio"
End Sub

Private Sub lblNombreServicioPropietario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, lblNombreServicioPropietario
    AdministrarKeyPreview KeyCode
End Sub

Private Sub lblNombreServicioUbicacionActual_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, lblNombreServicioUbicacionActual
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(Trim(txtCodigo))
    mo_Formulario.MarcarComoVacio txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdEstadoCama_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEstadoCama
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdEstadoCama_LostFocus()
   If cmbIdEstadoCama.Text <> "" Then
       cmbIdEstadoCama.BoundText = Val(Split(cmbIdEstadoCama.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdEstadoCama
End Sub

Private Sub cmbIdEstadoCama_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdCondicionOcupacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCondicionOcupacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdCondicionOcupacion_LostFocus()
   If cmbIdCondicionOcupacion.Text <> "" Then
       cmbIdCondicionOcupacion.BoundText = Val(Split(cmbIdCondicionOcupacion.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdCondicionOcupacion
End Sub

Private Sub cmbIdCondicionOcupacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTiposCama_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTiposCama
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTiposCama_LostFocus()
   If cmbIdTiposCama.Text <> "" Then
       cmbIdTiposCama.BoundText = Val(Split(cmbIdTiposCama.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTiposCama
End Sub

Private Sub cmbIdTiposCama_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
mo_Apariencia.ConfigurarFilasBiColores grdMovimientos, sighentidades.GrillaConFilasBicolor

mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoServicio, False
mo_Formulario.HabilitarDeshabilitar Me.txtIdServicioPropietario, False
mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicioPropietario, False
 Select Case mi_Opcion
     Case sghAgregar
        Me.cmbIdTiposCama.BoundText = 1
        Me.cmbIdEstadoCama.BoundText = 1
        Me.cmbIdEstadoCama.Enabled = False
        Me.cmbIdCondicionOcupacion.BoundText = ""
        Me.cmbIdCondicionOcupacion.Enabled = False
        'ml_IdServicioActual = Val(Me.txtIdServicioUbicacionActual.Tag)
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
 'mgaray20141012
 Call BloquearCamasOcupadas
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       GenerarRecordsetTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Camas"
       Case sghModificar
           Me.Caption = "Modificar Camas"
       Case sghConsultar
           Me.Caption = "Consultar Camas"
       Case sghEliminar
           Me.Caption = "Eliminar Camas"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
    btnAceptar.Enabled = True
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   Else
       ml_IdServicioActual = Val(Me.txtIdServicioUbicacionActual.Tag)
   End If
   If mi_Opcion = sghConsultar Then
    btnAceptar.Enabled = False
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
   'AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   Me.IdCama = mo_Camas.IdCama
                   LimpiarFormularioDespuesAgregar
                   Me.txtCodigo.SetFocus
'                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminHoteleria.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminHoteleria.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
        ml_ConfirmoOperacion = 0
               CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
                   ml_ConfirmoOperacion = 1
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminHoteleria.MensajeError, vbExclamation, Me.Caption
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
   
   If Me.txtIdServicioPropietario.Text = "" Then
       sMensaje = sMensaje + "Ingrese el servicio propietario de la cama" + Chr(13)
   End If
   If Me.txtIdServicioUbicacionActual.Text = "" Then
       sMensaje = sMensaje + "Ingrese el servicio donde esta ubicado la cama actualmente" + Chr(13)
   End If
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de la cama" + Chr(13)
   End If
   
    If mi_Opcion <> sghAgregar Then
         If Val(Me.cmbIdEstadoCama.BoundText) = 0 Then
             sMensaje = sMensaje + "Ingrese el estado de la cama" + Chr(13)
         Else
            'mgaray20141014
             If Val(Me.cmbIdEstadoCama.BoundText) = ESTADO_CAMA_OCUPADA Then
                 If Val(Me.cmbIdCondicionOcupacion.BoundText) = 0 Then
                     sMensaje = sMensaje + "Ingrese el valor de IdCondicionOcupacion" + Chr(13)
                 End If
             End If
        End If
    End If
   If Val(Me.cmbIdTiposCama.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el tipo de cama" + Chr(13)
   End If
   If mrs_Movimientos.RecordCount = 0 Then
            sMensaje = sMensaje + "Tiene que Registrar al menos un Movimiento de la Cama (Fecha de Ingreso y Servicio)"
   End If

   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
Dim rsCamas As New ADODB.Recordset
Dim lcSql As String
   ValidarReglas = False
   
   Select Case mi_Opcion
   Case sghAgregar
        Set rsCamas = mo_AdminHoteleria.CamasBuscarCodigoDeCama(txtCodigo.Text, Me.txtIdServicioPropietario.Tag)
        If rsCamas.RecordCount > 0 Then
             MsgBox "Ya existe una cama con el mismo /Código/Servicio Propietario/" + Chr(13) + "Servicio: " & rsCamas!ServicioActual, vbExclamation, Me.Caption
             rsCamas.Close
             Exit Function
        End If
        Set rsCamas = Nothing
    Case sghModificar
        If Val(cmbIdEstadoCama.BoundText) = ESTADO_CAMA_OCUPADA Then
        
        End If
    Case sghEliminar
        'mgaray10141014
        If Val(cmbIdEstadoCama.BoundText) = ESTADO_CAMA_OCUPADA Then
            MsgBox "No se puede eliminar una cama que esta siendo ocupada", vbExclamation, Me.Caption
            Exit Function
        End If
        
    End Select
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Camas
           .x = Me.Xini
           .Y = Me.Yini
           .IdServicioUbicacionActual = Val(Me.txtIdServicioUbicacionActual.Tag)
           .codigo = Trim(Me.txtCodigo.Text)
           .IdEstadoCama = Val(Me.cmbIdEstadoCama.BoundText)
           .IdCondicionOcupacion = Val(Me.cmbIdCondicionOcupacion.BoundText)
           .IdTiposCama = Val(Me.cmbIdTiposCama.BoundText)
           .IdServicioPropietario = Val(Me.txtIdServicioPropietario.Tag)
           .IdCama = Me.IdCama
           .IdUsuarioAuditoria = ml_idUsuario
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   'CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminHoteleria.CamasAgregar(mo_Camas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtCodigo.Text)
   Call AgregaCamasMovimientos(mo_Camas.IdCama, mo_Camas.IdUsuarioAuditoria)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

    '   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminHoteleria.CamasModificar(mo_Camas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtCodigo.Text)
   Call EliminaCamasMovimientosPorCama(mo_CamasMovimiento)
   Call AgregaCamasMovimientos(mo_CamasMovimiento.IdCama, mo_Camas.IdUsuarioAuditoria)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   'CargaDatosAlObjetosDeDatos
   Call EliminaCamasMovimientosPorCama(mo_CamasMovimiento)
   EliminarDatos = mo_AdminHoteleria.CamasEliminar(mo_Camas, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtCodigo.Text)
   
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
        Dim oDOServicio As New doServicio
        Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set mo_Camas = mo_AdminHoteleria.CamasSeleccionarPorId(Me.IdCama, oConexion)
        If mo_AdminHoteleria.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminHoteleria.MensajeError, vbInformation, Me.Caption
             mb_ExistenDatos = False
             Exit Sub
        End If
        
       If Not mo_Camas Is Nothing Then
           With mo_Camas
                Me.Xini = .x
                Me.Yini = .Y
                Me.txtCodigo.Text = .codigo
                
                Me.cmbIdEstadoCama.BoundText = .IdEstadoCama
                Me.cmbIdCondicionOcupacion.BoundText = .IdCondicionOcupacion
                Me.cmbIdTiposCama.BoundText = .IdTiposCama
                
                Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioPropietario, oConexion)
                If Not oDOServicio Is Nothing Then
                    Me.txtIdServicioPropietario.Tag = oDOServicio.idServicio
                    Me.txtIdServicioPropietario.Text = oDOServicio.codigo
                    Me.lblNombreServicioPropietario = oDOServicio.nombre
                Else
                    Me.txtIdServicioPropietario.Tag = ""
                    Me.lblNombreServicioPropietario = ""
                End If
                
                Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioUbicacionActual, oConexion)
                If Not oDOServicio Is Nothing Then
                    Me.txtIdServicioUbicacionActual.Tag = oDOServicio.idServicio
                    Me.txtIdServicioUbicacionActual.Text = oDOServicio.codigo
                    Me.lblNombreServicioUbicacionActual = oDOServicio.nombre
                    ml_IdServicioActual = oDOServicio.idServicio
                Else
                    Me.txtIdServicioUbicacionActual.Tag = ""
                    Me.lblNombreServicioUbicacionActual = ""
                End If
               End With
           If BusquedaCamasMovimientos(mo_Camas.IdCama) Then
               Call CargaUltimoMovimiento
               mb_ExistenDatos = True
           End If
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       oConexion.Close
       Set oConexion = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
        Me.txtIdServicioUbicacionActual.Text = ""
        Me.txtCodigo.Text = ""
        Me.cmbIdEstadoCama.BoundText = ""
        Me.cmbIdCondicionOcupacion.BoundText = ""
        Me.cmbIdTiposCama.BoundText = ""
        Me.txtIdServicioPropietario.Text = ""

         txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY
         txtFechaFin.Text = sighentidades.FECHA_VACIA_DMY
         
         With mrs_Movimientos
             If .RecordCount > 0 Then
                 If Not (.BOF = True And .EOF = True) Then
                     .MoveFirst
                     While .EOF = False
                         .Delete
                         .Update
                         .MoveNext
                     Wend
                 End If
             End If
         End With

End Sub






Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaFin
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin.Text <> sighentidades.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaFin.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin.Text = sighentidades.FECHA_VACIA_DMY
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaInicio
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio.Text <> sighentidades.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaInicio.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY
            Exit Sub
        End If
    End If
End Sub

Private Sub txtIdServicioPropietario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioPropietario
    If KeyCode = vbKeyF1 Then
        btnBuscarServicios_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicioPropietario_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioPropietario, lblNombreServicioPropietario
    mo_Formulario.MarcarComoVacio txtIdServicioPropietario
End Sub

Private Sub txtIdServicioPropietario_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioUbicacionActual_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioUbicacionActual
    If KeyCode = vbKeyF1 Then
        btnBuscarServicios_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicioUbicacionActual_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioUbicacionActual, lblNombreServicioUbicacionActual
    mo_Formulario.MarcarComoVacio txtIdServicioUbicacionActual
End Sub

Private Sub txtIdServicioUbicacionActual_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDOServicio As doServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            If cmbIdTipoServicio.BoundText = oDOServicio.idTipoServicio Then
                txtIdServicio.Tag = oDOServicio.idServicio
                lblDescripcionServicio = oDOServicio.nombre
            Else
                MsgBox "El servicio ingresado no pertenece es de emergencia", vbInformation, Me.Caption
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
            End If
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
   End If

End Sub

'***************daniel barrantes**************
'***************Temporal para Movimientos de cama
'***************
Sub GenerarRecordsetTemporal()
    With mrs_Movimientos
          .Fields.Append "idServicio", adInteger
          .Fields.Append "dServicio", adVarChar, 150
          .Fields.Append "FechaIngreso", adDate, 8, adFldIsNullable
          .Fields.Append "FechaSalida", adDate, 8, adFldIsNullable
          .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdMovimientos.DataSource = mrs_Movimientos
    
End Sub
Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_Movimientos
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
    Set Me.grdMovimientos.DataSource = mrs_Movimientos
End Sub
Private Sub btnAgregar_Click()
    On Error Resume Next
    If txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY Then
       MsgBox "Debe ingresar la Fecha que llegó la cama", vbInformation, Me.Caption
       Exit Sub
    ElseIf txtIdServicioPropietario.Text = "" Then
       MsgBox "Debe elegir el Servicio propietario de la cama", vbInformation, Me.Caption
       Exit Sub
    ElseIf txtIdServicioUbicacionActual.Text = "" Then
       MsgBox "Debe elegir la Ubicacion actual de la cama", vbInformation, Me.Caption
       Exit Sub
    End If
    If IsDate(Me.txtFechaFin.Text) Then
        If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
           Exit Sub
        End If
    End If
    
    mrs_Movimientos.MoveFirst
    Do While Not mrs_Movimientos.EOF
        If mrs_Movimientos.Fields("fechaIngreso").Value = CDate(txtFechaInicio.Text) Then
            Exit Do
        End If
        mrs_Movimientos.MoveNext
    Loop
    With mrs_Movimientos
        If .EOF Then
           .AddNew
           .Fields!FechaIngreso = txtFechaInicio.Text
        End If
        .Fields!idServicio = ml_IdServicioActual
        .Fields!DServicio = lblNombreServicioUbicacionActual.Text
        If txtFechaFin.Text <> sighentidades.FECHA_VACIA_DMY Then
           .Fields!fechaSalida = txtFechaFin.Text
        Else
           .Fields!fechaSalida = Null
        End If
        .Update
        .Sort = "fechaIngreso desc"
        .MoveFirst
    End With
    Set Me.grdMovimientos.DataSource = mrs_Movimientos
End Sub

'***************daniel barrantes**************
'***************carga los movimientos de las camas
'***************
Function BusquedaCamasMovimientos(lnIdCama As Long) As Boolean
    Dim oBuscaMov As New SIGHDatos.CamasMovimientos
    Dim rsTmp As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    On Error GoTo ErrorBusquedaCamasMovimientos
    BusquedaCamasMovimientos = False
    oConexion.Open sighentidades.CadenaConexion
    Set oBuscaMov.Conexion = oConexion
    Set rsTmp = oBuscaMov.SeleccionarPorCama(lnIdCama)
    If rsTmp.RecordCount > 0 Then
       rsTmp.MoveFirst
       Do While Not rsTmp.EOF
            mrs_Movimientos.AddNew
            mrs_Movimientos.Fields!idServicio = rsTmp.Fields("idServicio").Value
            mrs_Movimientos.Fields!DServicio = rsTmp.Fields("Dservicio").Value
            mrs_Movimientos.Fields!FechaIngreso = rsTmp.Fields("fechaIngreso").Value
            mrs_Movimientos.Fields!fechaSalida = rsTmp.Fields("fechaSalida").Value
            mrs_Movimientos.Update
            rsTmp.MoveNext
       Loop
       Set Me.grdMovimientos.DataSource = mrs_Movimientos
    End If
    BusquedaCamasMovimientos = True
    Exit Function
ErrorBusquedaCamasMovimientos:
    MsgBox Err.Description
End Function

Sub EliminaCamasMovimientosPorCama(oCamaMovimiento As sighComun.DOCamasMovimientos)
    Dim oBuscaMov As New SIGHDatos.CamasMovimientos
    Dim oConexion As New ADODB.Connection
    On Error GoTo ErrorEliminaCamasMovimientosPorCama
    oConexion.Open sighentidades.CadenaConexion
    Set oBuscaMov.Conexion = oConexion
    If oBuscaMov.EliminarPorCama(oCamaMovimiento) Then
    End If
    Exit Sub
ErrorEliminaCamasMovimientosPorCama:
    MsgBox Err.Description
End Sub

Sub AgregaCamasMovimientos(lnIdCama As Long, lnIdUsuarioAuditoria As Long)
         Dim oAgregarMov As New SIGHDatos.CamasMovimientos
         Dim oConexion As New ADODB.Connection
         On Error GoTo ErrorAgregaCamasMovimiento
         oConexion.Open sighentidades.CadenaConexion
         Set oAgregarMov.Conexion = oConexion
         mrs_Movimientos.MoveFirst
         Do While Not mrs_Movimientos.EOF
            mo_CamasMovimiento.IdCama = lnIdCama
            mo_CamasMovimiento.IdFechaIngreso = mrs_Movimientos.Fields("FechaIngreso").Value
            If Not IsNull(mrs_Movimientos.Fields("FechaSalida").Value) Then
               mo_CamasMovimiento.IdFechaSalida = mrs_Movimientos.Fields("FechaSalida").Value
            End If
            mo_CamasMovimiento.IdMovimiento = 0
            mo_CamasMovimiento.idServicio = mrs_Movimientos.Fields("idServicio").Value
            mo_CamasMovimiento.IdUsuarioAuditoria = lnIdUsuarioAuditoria
            If oAgregarMov.Insertar(mo_CamasMovimiento) Then
               
            End If
            mrs_Movimientos.MoveNext
         Loop
         Exit Sub
ErrorAgregaCamasMovimiento:
   MsgBox Err.Description
End Sub
'***************daniel barrantes**************
'***************carga el Ultimo movimiento de la cama
'***************
Sub CargaUltimoMovimiento()
    If mrs_Movimientos.RecordCount > 0 Then
        mrs_Movimientos.MoveFirst
        mo_CamasMovimiento.IdFechaIngreso = mrs_Movimientos.Fields("fechaIngreso").Value
        If Not IsNull(mrs_Movimientos.Fields("fechaSalida").Value) Then
           mo_CamasMovimiento.IdFechaSalida = mrs_Movimientos.Fields("fechaSalida").Value
           txtFechaFin.Text = mo_CamasMovimiento.IdFechaSalida
        End If
        mo_CamasMovimiento.idServicio = mrs_Movimientos.Fields("idServicio").Value
        txtFechaInicio.Text = Format(mo_CamasMovimiento.IdFechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
        ml_IdServicioActual = mrs_Movimientos.Fields("idServicio").Value
    End If
    mo_CamasMovimiento.IdCama = mo_Camas.IdCama
    mo_CamasMovimiento.IdUsuarioAuditoria = mo_Camas.IdUsuarioAuditoria
End Sub
'mgaray20141014
Public Function BloquearCamasOcupadas() As Boolean
    mo_Formulario.HabilitarDeshabilitar Me.cmbIdEstadoCama, True
    If mi_Opcion = sghModificar Then
        If Val(cmbIdEstadoCama.BoundText) = ESTADO_CAMA_OCUPADA Then
            mo_Formulario.HabilitarDeshabilitar Me.cmbIdEstadoCama, False
        End If
    End If
End Function

Sub LimpiarFormularioDespuesAgregar()
On Error GoTo miError
        Me.txtCodigo.Text = ""
        Me.cmbIdCondicionOcupacion.BoundText = ""
        Me.cmbIdTiposCama.BoundText = ""

         txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY
         txtFechaFin.Text = sighentidades.FECHA_VACIA_DMY
         
         With mrs_Movimientos
             If .RecordCount > 0 Then
                 If Not (.BOF = True And .EOF = True) Then
                     .MoveFirst
                     While .EOF = False
                         .Delete
                         .Update
                         .MoveNext
                     Wend
                 End If
             End If
         End With
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation, "Error"
    End If
End Sub

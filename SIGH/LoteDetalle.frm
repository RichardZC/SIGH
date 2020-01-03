VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form LoteDetalle 
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "LoteDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   0
      TabIndex        =   15
      Top             =   2700
      Width           =   5115
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LoteDetalle.frx":0CCA
         DownPicture     =   "LoteDetalle.frx":118E
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
         Left            =   2520
         Picture         =   "LoteDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LoteDetalle.frx":1B66
         DownPicture     =   "LoteDetalle.frx":1FC6
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
         Left            =   975
         Picture         =   "LoteDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5130
      Begin VB.ComboBox cmbTurno 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   3825
      End
      Begin VB.TextBox txtSaldoInicialDolares 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   2160
         Width           =   1395
      End
      Begin VB.TextBox txtSaldoInicialSoles 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   1395
      End
      Begin VB.ComboBox cmbCajero 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   3825
      End
      Begin VB.ComboBox cmbCaja 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3825
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
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
      Begin VB.Label Label5 
         Caption         =   "Turno"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial                            ($)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   2220
         Width           =   2865
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial                            (S/.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1860
         Width           =   3000
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1500
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Cajero"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Caja"
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "LoteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD 22/06/2005 [Todo el Archivo]
'MZD02 Ini 04/07/2005 Cambios varios

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: PODiagnosticos
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 12:17:18 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_CajaLote As New DOCajaLote
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdLote As Long
Dim mo_AdminCaja As New ReglasCaja
Dim mo_AdminComun As New ReglasComunes
Dim mo_cmbCaja  As New SIGHComun.ListaDespleglable
Dim mo_cmbCajero  As New SIGHComun.ListaDespleglable
Dim mo_cmbTurno  As New SIGHComun.ListaDespleglable

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
Property Let IdLote(lValue As Long)
   ml_IdLote = lValue
End Property
Property Get IdLote() As Long
   IdLote = ml_IdLote
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
         Me.txtFecha = Format(Now, SIGHComun.FormatoFechaCorta)
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         Frame1.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         Frame1.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Private Sub cmbCaja_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbCaja
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbCajero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbCajero
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbCaja.MiComboBox = cmbCaja
    Set mo_cmbCajero.MiComboBox = cmbCajero
    Set mo_cmbTurno.MiComboBox = cmbTurno
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Lote"
       Case sghModificar
           Me.Caption = "Modificar Lote"
       Case sghConsultar
           Me.Caption = "Consultar Lote"
       Case sghEliminar
           Me.Caption = "Eliminar Lote"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
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
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, Me.Caption
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
   
   If mo_cmbCaja.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese la Caja" + Chr(13)
   End If
   If mo_cmbCajero.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el Cajero" + Chr(13)
   End If
   If mo_cmbTurno.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el Turno" + Chr(13)
   End If
   If Me.txtFecha = "" Or Me.txtFecha = SIGHComun.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la Fecha" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   CargaDatosAlObjetosDeDatos
   If mo_AdminCaja.LoteExisteAsignacionCaja(mo_CajaLote) Then
        MsgBox "Ya existe la asignación del Cajero a la Caja Seleccionada, para la fecha " & mo_CajaLote.Fecha, vbExclamation, Me.Caption
        Exit Function
   End If
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   Me.txtSaldoInicialSoles = Replace(Me.txtSaldoInicialSoles, ".", ",")
   Me.txtSaldoInicialDolares = Replace(Me.txtSaldoInicialDolares, ".", ",")
   With mo_CajaLote
        .IdCaja = Val(mo_cmbCaja.BoundText)
        .IdCajero = Val(mo_cmbCajero.BoundText)
        .IdTurno = Val(mo_cmbTurno.BoundText)
        .Fecha = IIf(Me.txtFecha = SIGHComun.FECHA_VACIA_DMY, 0, Me.txtFecha)
        .IdCajero = Val(mo_cmbCajero.BoundText)
        .SaldoInicialSoles = Val(Me.txtSaldoInicialSoles)
        .SaldoInicialDolares = Val(Me.txtSaldoInicialDolares)
        .IdUsuarioAuditoria = Me.IdUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   mo_CajaLote.EstadoLote = "A"
   AgregarDatos = mo_AdminCaja.LoteAgregar(mo_CajaLote)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminCaja.LoteModificar(mo_CajaLote)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminCaja.LoteEliminar(mo_CajaLote)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CajaLote = mo_AdminCaja.LoteSeleccionarPorId(Me.IdLote)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminCaja.MensajeError, vbCritical, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CajaLote Is Nothing Then
        With mo_CajaLote
            mo_cmbCaja.BoundText = .IdCaja
            mo_cmbCajero.BoundText = .IdCajero
            mo_cmbTurno.BoundText = .IdTurno
            Me.txtSaldoInicialSoles = .SaldoInicialSoles
            Me.txtSaldoInicialDolares = .SaldoInicialDolares
            Me.txtFecha = IIf(.Fecha = 0, SIGHComun.FECHA_VACIA_DMY, .Fecha)
            
            mb_ExistenDatos = True
        End With
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdLote = 0
    Me.txtSaldoInicialSoles = ""
    Me.txtSaldoInicialDolares = ""
    mo_cmbCaja.BoundText = ""
    mo_cmbCajero.BoundText = ""
    mo_cmbTurno.BoundText = ""
    
End Sub

Sub CargarComboBoxes()
       
    mo_cmbCaja.BoundColumn = "IdCaja"
    mo_cmbCaja.ListField = "Descripcion"
    Set mo_cmbCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()

    mo_cmbCajero.BoundColumn = "IdCajero"
    mo_cmbCajero.ListField = "NombreCompleto"
    Set mo_cmbCajero.RowSource = mo_AdminCaja.CajerosSeleccionarTodosParaLista()
    
    mo_cmbTurno.BoundColumn = "IdTurno"
    mo_cmbTurno.ListField = "Descripcion"
    Set mo_cmbTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
    
End Sub

Private Sub txtSaldoInicialDolares_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSaldoInicialDolares
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtSaldoInicialDolares_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtSaldoInicialSoles_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSaldoInicialSoles
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtSaldoInicialSoles_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

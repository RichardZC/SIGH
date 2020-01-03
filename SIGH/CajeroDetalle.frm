VERSION 5.00
Begin VB.Form CajeroDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "CajeroDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1995
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   6750
      Begin VB.CheckBox chkEstadoCajero 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton btnEmpleado 
         Caption         =   "..."
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   315
      End
      Begin VB.ComboBox cmbIdCaja 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   5625
      End
      Begin VB.TextBox txtIdEmpleado 
         BackColor       =   &H00FFEBD9&
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
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   0
         Top             =   300
         Width           =   1000
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label lblCodigoCIE2004 
         Caption         =   "Empleado"
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
         TabIndex        =   10
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblNombreEmpleado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   5595
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   7
      Top             =   2100
      Width           =   6735
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CajeroDetalle.frx":0CCA
         DownPicture     =   "CajeroDetalle.frx":112A
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
         Left            =   1830
         Picture         =   "CajeroDetalle.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CajeroDetalle.frx":1A14
         DownPicture     =   "CajeroDetalle.frx":1ED8
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
         Left            =   3360
         Picture         =   "CajeroDetalle.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CajeroDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD 19/06/2005 [Todo el Archivo]
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

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_CajaCajero As New DOCajaCajero
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdCajero As Long
Dim mo_AdminCaja As New ReglasCaja
Dim mo_AdminComun As New ReglasComunes
Dim mrs_Supervisores As New ADODB.Recordset
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
'Dim mo_Supervisores As New Collection
Dim mo_cmbIdCaja As New SIGHEntidades.ListaDespleglable

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
Property Let IdCajero(lValue As Long)
   ml_IdCajero = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Inicializar()

End Function
Sub CargarDatosAlFormulario()

    Select Case mi_Opcion
        Case sghAgregar
            'CargarSupervisores
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

Private Sub btnEmpleado_Click()
    Dim oFrm As New EmpleadosBusqueda
    Dim dOEmpleado As dOEmpleado
    oFrm.Caption = "Seleccione el empleado"
    oFrm.Show vbModal
    If oFrm.idRegistroSeleccionado <> 0 Then
        Me.txtIdEmpleado = CStr(oFrm.idRegistroSeleccionado)
        Call ObtenerNombreEmpleado(oFrm.idRegistroSeleccionado)
    End If
End Sub

Private Sub cmbIdCaja_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCaja
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
    'GenerarRecordsetTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Cajero"
       Case sghModificar
           Me.Caption = "Modificar Cajero"
       Case sghConsultar
           Me.Caption = "Consultar Cajero"
       Case sghEliminar
           Me.Caption = "Eliminar Cajero"
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
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
   
   If Me.txtIdEmpleado = "" Then
       sMensaje = sMensaje + "Ingrese el código del Empleado " + Chr(13)
   End If
   
   If mo_cmbIdCaja.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese la Caja" + Chr(13)
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
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   With mo_CajaCajero
        .IdEmpleado = Val(Me.txtIdEmpleado.Text)
        .EstadoCajero = CStr(Me.chkEstadoCajero.Value)
        .IdUsuarioAuditoria = Me.idUsuario
        .IdCaja = Val(mo_cmbIdCaja.BoundText)
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminCaja.CajeroAgregar(mo_CajaCajero)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminCaja.CajeroModificar(mo_CajaCajero)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminCaja.CajeroEliminar(mo_CajaCajero)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CajaCajero = mo_AdminCaja.CajeroSeleccionarPorId(Me.IdCajero)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminCaja.MensajeError, vbCritical, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CajaCajero Is Nothing Then
        With mo_CajaCajero
            Me.txtIdEmpleado = .IdEmpleado
            Me.chkEstadoCajero.Value = Val(.EstadoCajero)
            mo_cmbIdCaja.BoundText = .IdCaja
            ObtenerNombreEmpleado .IdEmpleado
            mb_ExistenDatos = True
            
        End With
        'CargarSupervisores
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

    Me.IdCajero = 0
    Me.txtIdEmpleado = ""
    Me.lblNombreEmpleado = ""
    Me.chkEstadoCajero.Value = 0
    mo_cmbIdCaja.BoundText = ""
End Sub

Sub ObtenerNombreEmpleado(IdEmpleado As Long)
    Dim dOEmp As dOEmpleado
    Set dOEmp = mo_AdminComun.EmpleadosSeleccionarPorId(IdEmpleado)
    Me.lblNombreEmpleado.Caption = dOEmp.ApellidoPaterno & " " & dOEmp.ApellidoMaterno & " " & dOEmp.Nombres
End Sub

Sub CargarComboBoxes()
       
    mo_cmbIdCaja.BoundColumn = "IdCaja"
    mo_cmbIdCaja.ListField = "Descripcion"
    Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()

       
End Sub






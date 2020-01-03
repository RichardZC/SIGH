VERSION 5.00
Begin VB.Form SupervisorDetalle 
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "SupervisorDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnEmpleado 
      Caption         =   "..."
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   8
      Top             =   1470
      Width           =   6735
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SupervisorDetalle.frx":0CCA
         DownPicture     =   "SupervisorDetalle.frx":118E
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
         Picture         =   "SupervisorDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SupervisorDetalle.frx":1B66
         DownPicture     =   "SupervisorDetalle.frx":1FC6
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
         Left            =   1815
         Picture         =   "SupervisorDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   2
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
      Height          =   1455
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   6750
      Begin VB.CheckBox chkEstadoSupervisor 
         Caption         =   "Activo"
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         Top             =   1080
         Width           =   1755
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
         Left            =   1050
         MaxLength       =   7
         TabIndex        =   5
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label lblNombreEmpleado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1050
         TabIndex        =   9
         Top             =   660
         Width           =   5595
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
         Left            =   60
         TabIndex        =   7
         Top             =   660
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
         Left            =   75
         TabIndex        =   6
         Top             =   285
         Width           =   1335
      End
   End
End
Attribute VB_Name = "SupervisorDetalle"
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

Dim mo_Teclado As New SIGHCOmun.Teclado
Dim mo_Formulario As New SIGHCOmun.Formulario
Dim mo_CajaSupervisor As New DOCajaSupervisor
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdSupervisor As Long
Dim mo_AdminCaja As New ReglasCaja
Dim mo_AdminComun As New ReglasComunes

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
Property Let IdSupervisor(lValue As Long)
   ml_IdSupervisor = lValue
End Property
Property Get IdSupervisor() As Long
   IdSupervisor = ml_IdSupervisor
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
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
    If oFrm.IdRegistroSeleccionado <> 0 Then
        Me.txtIdEmpleado = CStr(oFrm.IdRegistroSeleccionado)
        ObtenerNombreEmpleado oFrm.IdRegistroSeleccionado
    End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Supervisor"
       Case sghModificar
           Me.Caption = "Modificar Supervisor"
       Case sghConsultar
           Me.Caption = "Consultar Supervisor"
       Case sghEliminar
           Me.Caption = "Eliminar Supervisor"
       End Select
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
   If Me.txtIdEmpleado = "" Then
       sMensaje = sMensaje + "Ingrese el código del Empleado " + Chr(13)
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
   With mo_CajaSupervisor
        .IdEmpleado = Me.txtIdEmpleado.Text
        .EstadoSupervisor = CStr(Me.chkEstadoSupervisor.Value)
        .IdUsuarioAuditoria = Me.IdUsuario
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminCaja.SupervisorAgregar(mo_CajaSupervisor)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminCaja.SupervisorModificar(mo_CajaSupervisor)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminCaja.SupervisorEliminar(mo_CajaSupervisor)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CajaSupervisor = mo_AdminCaja.SupervisorSeleccionarPorId(Me.IdSupervisor)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminCaja.MensajeError, vbCritical, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CajaSupervisor Is Nothing Then
        With mo_CajaSupervisor
            Me.txtIdEmpleado = .IdEmpleado
            Me.chkEstadoSupervisor.Value = Val(.EstadoSupervisor)
            ObtenerNombreEmpleado .IdEmpleado
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

    Me.IdSupervisor = 0
    Me.txtIdEmpleado = ""
    Me.lblNombreEmpleado = ""
    Me.chkEstadoSupervisor.Value = 0
End Sub
Sub ObtenerNombreEmpleado(IdEmpleado As Long)
    Dim dOEmp  As dOEmpleado
    Set dOEmp = mo_AdminComun.EmpleadosSeleccionarPorId(IdEmpleado)
    Me.lblNombreEmpleado.Caption = dOEmp.ApellidoPaterno & " " & dOEmp.ApellidoMaterno & " " & dOEmp.Nombres
End Sub




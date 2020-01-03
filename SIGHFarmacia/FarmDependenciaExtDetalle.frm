VERSION 5.00
Begin VB.Form FarmDependenciaExtDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FarmDependenciaExtDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   5790
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FarmDependenciaExtDetalle.frx":0CCA
         DownPicture     =   "FarmDependenciaExtDetalle.frx":118E
         Height          =   700
         Left            =   2940
         Picture         =   "FarmDependenciaExtDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FarmDependenciaExtDetalle.frx":1B66
         DownPicture     =   "FarmDependenciaExtDetalle.frx":1FC6
         Height          =   700
         Left            =   1395
         Picture         =   "FarmDependenciaExtDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
      Height          =   1995
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5790
      Begin VB.TextBox txtCodigoDigemid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1050
         Width           =   1410
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   645
         Width           =   4755
      End
      Begin VB.Label Label5 
         Caption         =   "Código DIGEMID"
         ForeColor       =   &H00000000&
         Height          =   660
         Left            =   165
         TabIndex        =   9
         Top             =   1020
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   645
      End
   End
End
Attribute VB_Name = "FarmDependenciaExtDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Manimiento de Dependencias externas
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_DoFarmAlmacen As New DoFarmAlmacen
Dim ms_MensajeError As String
Dim mb_ExistenDatos As Boolean
Dim mo_ReglasFarmacia As New ReglasFarmacia
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdAlmacen As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdDependenciaExt(lValue As Long)
   ml_IdAlmacen = lValue
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
     mo_Formulario.HabilitarDeshabilitar Me.txtCodigo, False
     Select Case mi_Opcion
     Case sghAgregar
         CargaUltimoCorrelativoIdAlmacen
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Sub CargaUltimoCorrelativoIdAlmacen()
    txtCodigo.Text = Trim(Str(mo_ReglasFarmacia.CargaUltimoCorrelativoIdAlmacen))
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Dependencia Externa"
       Case sghModificar
           Me.Caption = "Modificar Dependencia Externa"
       Case sghConsultar
           Me.Caption = "Consultar Dependencia Externa"
       Case sghEliminar
           Me.Caption = "Anular Dependencia Externa"
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
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   'LimpiarFormulario
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
       If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbExclamation, Me.Caption
               End If
           End If
        End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "No hay el Id" + Chr(13)
   End If
   If Trim(Me.txtDescripcion) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre de la Dependencia Externa" + Chr(13)
       txtDescripcion.SetFocus
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
   
   With mo_DoFarmAlmacen
        .IdAlmacen = Val(Me.txtCodigo.Text)
        .descripcion = UCase(Me.txtDescripcion.Text)
        .idEstado = sghEstadoTabla.sghRegistrado
        .idTipoLocales = "X"
        .IdUsuarioAuditoria = ml_idUsuario
        .CodigoSismed = txtCodigoDigemid.Text
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_ReglasFarmacia.FarmAlmacenAgregar(mo_DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
   With mo_DoFarmAlmacen
       .descripcion = UCase(txtDescripcion.Text)
       .IdUsuarioAuditoria = ml_idUsuario
       .CodigoSismed = txtCodigoDigemid.Text
   End With
   ModificarDatos = mo_ReglasFarmacia.farmalmacenmodificar(mo_DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
   With mo_DoFarmAlmacen
       .idEstado = sghEstadoTabla.sghAnulado
       .IdUsuarioAuditoria = ml_idUsuario
   End With
   EliminarDatos = mo_ReglasFarmacia.farmalmacenmodificar(mo_DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_DoFarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(ml_IdAlmacen)
    If Not mo_DoFarmAlmacen Is Nothing Then
        With mo_DoFarmAlmacen
            Me.txtDescripcion = .descripcion
            Me.txtCodigo = .IdAlmacen
            txtCodigoDigemid.Text = .CodigoSismed
            mb_ExistenDatos = True
            If .idEstado = 0 Then
               btnAceptar.Enabled = False
            End If
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

    ml_IdAlmacen = 0
    
    Me.txtDescripcion = ""
    Me.txtCodigo = ""
    
End Sub

Sub CargarComboBoxes()
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub Label6_Click()

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
   End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_DoFarmAlmacen = Nothing
    Set mo_ReglasFarmacia = Nothing

End Sub

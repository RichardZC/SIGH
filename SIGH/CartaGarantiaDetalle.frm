VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CartaGarantiaDetalle 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "CartaGarantiaDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   60
      TabIndex        =   22
      Top             =   0
      Width           =   5970
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   1740
         TabIndex        =   0
         Top             =   330
         Width           =   1530
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   3420
         Picture         =   "CartaGarantiaDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label Label50 
         Caption         =   "Nro Historia Clinica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del paciente"
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
      Height          =   1725
      Left            =   60
      TabIndex        =   13
      Top             =   885
      Width           =   5970
      Begin VB.TextBox lblServicioIngreso 
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
         Left            =   1695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4020
      End
      Begin VB.TextBox lblFechaIngreso 
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
         Left            =   1695
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   615
         Width           =   1680
      End
      Begin VB.TextBox lblPaciente 
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
         Left            =   1695
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   4020
      End
      Begin VB.TextBox lblNroCuenta 
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
         Left            =   1695
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   255
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Servicio Ingreso"
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
         TabIndex        =   21
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Ingreso"
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
         TabIndex        =   20
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Paciente"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Cuenta"
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
         Left            =   150
         TabIndex        =   18
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   4920
      Width           =   5970
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CartaGarantiaDetalle.frx":3913
         DownPicture     =   "CartaGarantiaDetalle.frx":3DD7
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
         Left            =   3120
         Picture         =   "CartaGarantiaDetalle.frx":42C3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CartaGarantiaDetalle.frx":47AF
         DownPicture     =   "CartaGarantiaDetalle.frx":4C0F
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
         Left            =   1575
         Picture         =   "CartaGarantiaDetalle.frx":5084
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosGenerales 
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
      Height          =   2175
      Left            =   60
      TabIndex        =   7
      Top             =   2700
      Width           =   5970
      Begin VB.TextBox txtValorCobertura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   3
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox txtNroCarta 
         Height          =   315
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtObservacion 
         Height          =   315
         Left            =   120
         MaxLength       =   250
         TabIndex        =   4
         Top             =   1680
         Width           =   5775
      End
      Begin MSMask.MaskEdBox txtFechaVigencia 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Cobertura"
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
         Left            =   180
         TabIndex        =   12
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Vigencia"
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
         Left            =   180
         TabIndex        =   11
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Carta"
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
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   1380
         Width           =   990
      End
   End
End
Attribute VB_Name = "CartaGarantiaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim mo_CartaGarantia As New DOCartaGarantia
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdCartaGarantia As Long
Dim mo_AdminComun As New ReglasComunes
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_AdminCaja As New ReglasCaja



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
Property Let IdCartaGarantia(lValue As Long)
   ml_IdCartaGarantia = lValue
End Property
Property Get IdCartaGarantia() As Long
   IdCartaGarantia = ml_IdCartaGarantia
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
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Private Sub btnBuscar_Click()
Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
Dim lIdCuentaAtencionActual As Long
    
    LimpiarDatosDeAtencion
    If (Me.txtNroHistoria) = "" Then
        MsgBox "Ingrese la Historia Clínica a buscar", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Dim rsCuentasAtencion As New ADODB.Recordset
    Dim iCount As Integer

    lIdCuentaAtencionActual = 0
    Set rsCuentasAtencion = mo_AdminCaja.ObtenerCuentasAtencionPorHistoriaClinica(Val(Me.txtNroHistoria))
    iCount = 0
    Do While Not rsCuentasAtencion.EOF
        iCount = iCount + 1
        lIdCuentaAtencionActual = rsCuentasAtencion!IdCuentaAtencion
        rsCuentasAtencion.MoveNext
    Loop
    If iCount > 1 Then
        'Levantamos el formulario para seleccionar la cuenta de atención
        Dim oFrmCuentasAtencion As New CuentasAtencionSeleccionar
        Set oFrmCuentasAtencion.DataSource = rsCuentasAtencion
        oFrmCuentasAtencion.Show vbModal
        If oFrmCuentasAtencion.BotonPresionado = sghCancelar Then
            lIdCuentaAtencionActual = 0
        Else
            lIdCuentaAtencionActual = oFrmCuentasAtencion.IdRegistroSeleccionado
        End If
    End If
    ObtenerDatosCuentaAtencion lIdCuentaAtencionActual
End Sub
Private Sub ObtenerDatosCuentaAtencion(lIdCuentaAtencionActual As Long)
Dim rsPaciente As New Recordset
Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion

    oDOCuentaAtencion.IdCuentaAtencion = lIdCuentaAtencionActual
    
    Screen.MousePointer = vbHourglass
    Set rsPaciente = mo_AdminAdmision.AtencionesFiltrarPacientesParaIngresarProcedimientos(oDOPaciente, oDOCuentaAtencion)
    Screen.MousePointer = vbDefault
    
    'Si hay una sola coincidencia
    If rsPaciente.RecordCount = 1 Then
        rsPaciente.MoveFirst
        LimpiarDatosDeAtencion
        
        Me.lblFechaIngreso = rsPaciente!FechaIngreso
        Me.lblServicioIngreso = rsPaciente!ServicioIngreso
        Me.lblPaciente = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
        Me.lblNroCuenta = rsPaciente!IdCuentaAtencion
    
    ElseIf rsPaciente.RecordCount > 1 Then
        'cmbNroHistoriaBusqueda.ShowDropDown
        
    ElseIf rsPaciente.RecordCount = 0 Then
        MsgBox "No se encontraron atenciones para el nro de historia o nro de cuenta ingresado", vbInformation, Me.Caption
        LimpiarDatosDeAtencion
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
           Me.Caption = "Agregar Carta de  Garantia"
       Case sghModificar
           Me.Caption = "Modificar Carta de  Garantia"
       Case sghConsultar
           Me.Caption = "Consultar Carta de  Garantia"
       Case sghEliminar
           Me.Caption = "Eliminar Carta de  Garantia"
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
       Case vbKeyF6
           btnBuscar_Click
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
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
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
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
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
   
   If Me.lblNroCuenta = "" Then
       sMensaje = sMensaje + "Ingrese el Nro de Cuenta" + Chr(13)
   End If
   If Me.txtFechaVigencia = "" Or txtFechaVigencia.Text = SIGHComun.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Debe ingresar la fecha de vigencia" + Chr(13)
   End If
   If Trim(Me.txtNroCarta) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If CCur(Me.txtValorCobertura) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Valor de la Cobertura" + Chr(13)
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
   
   With mo_CartaGarantia
        .IdCuentaAtencion = Me.lblNroCuenta
        .NroCarta = Me.txtNroCarta
        .Observacion = Me.txtObservacion
        .FechaVigencia = IIf(Me.txtFechaVigencia = "" Or Me.txtValorCobertura = SIGHComun.FECHA_VACIA_DMY, 0, CDate(Me.txtFechaVigencia))
        .ValorCobertura = CCur(Me.txtValorCobertura)
        
        .IdUsuarioAuditoria = Me.IdUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminComun.CartaGarantiaAgregar(mo_CartaGarantia)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminComun.CartaGarantiaModificar(mo_CartaGarantia)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminComun.CartaGarantiaEliminar(mo_CartaGarantia)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CartaGarantia = mo_AdminComun.CartaGarantiaSeleccionarPorId(Me.IdCartaGarantia)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminComun.MensajeError, vbCritical, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CartaGarantia Is Nothing Then
        With mo_CartaGarantia
            Me.txtNroCarta = .NroCarta
            Me.lblNroCuenta = .IdCuentaAtencion
            Me.txtFechaVigencia = IIf(.FechaVigencia = 0, "", .FechaVigencia)
            Me.txtObservacion = .Observacion
            Me.txtValorCobertura = .ValorCobertura
            ObtenerDatosCuentaAtencion Val(Me.lblNroCuenta)
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

    Me.IdCartaGarantia = 0
    
    Me.txtNroCarta = ""
    Me.lblNroCuenta = ""
    Me.txtFechaVigencia = SIGHComun.FECHA_VACIA_DMY
    Me.txtObservacion = ""
    Me.txtValorCobertura = ""
    LimpiarDatosDeAtencion
    Me.txtNroHistoria = ""
End Sub

Sub CargarComboBoxes()
    
End Sub
Private Sub LimpiarDatosDeAtencion()
    
    Me.lblFechaIngreso = ""
    Me.lblNroCuenta = ""
    Me.lblPaciente = ""
    Me.lblServicioIngreso = ""
End Sub

Private Sub txtFechaVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub
Private Sub txtNroCarta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroCarta
End Sub

Private Sub txtNroCarta_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservacion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txtValorCobertura_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtValorCobertura
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtValorCobertura_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

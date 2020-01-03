VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form TipoFinanciamientoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "TipoFinanciamientoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
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
      TabIndex        =   17
      Top             =   1200
      Width           =   6615
      Begin VB.ComboBox cmbGeneraPago 
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
         ItemData        =   "TipoFinanciamientoDetalle.frx":08CA
         Left            =   3030
         List            =   "TipoFinanciamientoDetalle.frx":08DD
         TabIndex        =   2
         Top             =   240
         Width           =   3510
      End
      Begin Threed.SSCheck chkEsOficina 
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   600
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Es Oficina del Hospital ?"
      End
      Begin Threed.SSCheck chkSeIngresPrecios 
         Height          =   315
         Left            =   150
         TabIndex        =   19
         Top             =   960
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Se ingresa Precios (Tarifario) ?"
      End
      Begin VB.Label Label4 
         Caption         =   "Columna ROJA en 'Estado Cuenta'"
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
         TabIndex        =   20
         Top             =   270
         Width           =   2865
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos para Farmacia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   60
      TabIndex        =   11
      Top             =   2670
      Width           =   6615
      Begin VB.CheckBox chkEsFuenteFinanciamiento 
         Caption         =   "Tiene Fuente Financiamiento (IAFA) ?"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   6195
      End
      Begin VB.CheckBox chkEsSalida 
         Caption         =   "Se usa en la opción VENTAS ?"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2610
         Width           =   5955
      End
      Begin VB.CheckBox chkSeImprimeComprobante 
         Caption         =   "Se imprime Comprobante (Fisicamente) ?"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1650
         Width           =   5565
      End
      Begin VB.ComboBox cmbTipoComprobante 
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
         Left            =   2070
         TabIndex        =   4
         Top             =   1950
         Width           =   4470
      End
      Begin VB.ComboBox cmbTipoConcepto 
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
         Left            =   2280
         TabIndex        =   3
         Top             =   1050
         Width           =   4260
      End
      Begin VB.Frame fraTipoVenta 
         Caption         =   "Tipo de Venta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   90
         TabIndex        =   12
         Top             =   2910
         Width           =   6435
         Begin Threed.SSOption optVentaD 
            Height          =   345
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   609
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Venta Directa (con IAFA)"
         End
         Begin Threed.SSOption optPreVenta 
            Height          =   345
            Left            =   5250
            TabIndex        =   14
            Top             =   300
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   609
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PreVenta"
         End
         Begin Threed.SSOption optVtaSinPlan 
            Height          =   345
            Left            =   2580
            TabIndex        =   24
            Top             =   300
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   609
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Venta Directa (sin IAFA)"
         End
      End
      Begin Threed.SSCheck chkEsFarmacia 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   690
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   556
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Es una Forma Pago (igual que SISMEDV2) ?"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   6600
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Comprobante"
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
         TabIndex        =   21
         Top             =   1980
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   30
         X2              =   6555
         Y1              =   1470
         Y2              =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Concepto (Form.ICI)"
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
         TabIndex        =   16
         Top             =   1110
         Width           =   2115
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   6570
      Width           =   6615
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "TipoFinanciamientoDetalle.frx":0958
         DownPicture     =   "TipoFinanciamientoDetalle.frx":0E1C
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
         Left            =   3472
         Picture         =   "TipoFinanciamientoDetalle.frx":1308
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "TipoFinanciamientoDetalle.frx":17F4
         DownPicture     =   "TipoFinanciamientoDetalle.frx":1C54
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
         Left            =   1927
         Picture         =   "TipoFinanciamientoDetalle.frx":20C9
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1125
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   6645
      Begin VB.TextBox txtIdTipoFinanciamiento 
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
         Left            =   1680
         TabIndex        =   0
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   630
         Width           =   4845
      End
      Begin VB.Label Label1 
         Caption         =   "Id"
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
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Producto/Plan"
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
         TabIndex        =   8
         Top             =   690
         Width           =   1425
      End
   End
End
Attribute VB_Name = "TipoFinanciamientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Tipos Financiamientos del Establecimiento
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_TiposFinanciamiento As New DOTiposFinanciamiento
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdTipoFinanciamiento As Long
Dim mo_cmbTipoConceptoF As New SIGHEntidades.ListaDespleglable
Dim mo_cmbCajaTiposComprobante As New SIGHEntidades.ListaDespleglable
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim oRsTmpF As New Recordset
Dim oRsTmpR As New Recordset
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub CargarComboBoxes()
       Set oRsTmpF = mo_ReglasFarmacia.FarmTipoConceptosDevuelveTodos
       mo_cmbTipoConceptoF.BoundColumn = "idTipoConcepto"
       mo_cmbTipoConceptoF.ListField = "Concepto"
       Set mo_cmbTipoConceptoF.RowSource = oRsTmpF
       '
       Set oRsTmpR = mo_ReglasCaja.CajaTiposComprobanteSeleccionarTodos(False, False)
       mo_cmbCajaTiposComprobante.BoundColumn = "idTipoComprobante"
       mo_cmbCajaTiposComprobante.ListField = "Descripcion"
       Set mo_cmbCajaTiposComprobante.RowSource = oRsTmpR
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
Property Let idTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get idTipoFinanciamiento() As Long
   idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

















Private Sub chkEsSalida_Click()
    If chkEsSalida.Value = 1 Then
       fraTipoVenta.Enabled = True
    Else
       optVentaD.Value = False
       optVtaSinPlan.Value = False
       optPreVenta.Value = False
       fraTipoVenta.Enabled = False
    End If

End Sub


Private Sub cmbGeneraPago_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbGeneraPago
    AdministrarKeyPreview KeyCode

End Sub




Private Sub cmbTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoComprobante
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbTipoConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoConcepto
    AdministrarKeyPreview KeyCode

End Sub

Private Sub Form_Initialize()
    Set mo_cmbTipoConceptoF.MiComboBox = cmbTipoConcepto
    Set mo_cmbCajaTiposComprobante.MiComboBox = cmbTipoComprobante

End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcion_LostFocus()
   mo_Formulario.MarcarComoVacio txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
        mo_Formulario.HabilitarDeshabilitar Me.txtIdTipoFinanciamiento, False
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Producto/Plan"
       Case sghModificar
           Me.Caption = "Modificar Producto/Plan"
       Case sghConsultar
           Me.Caption = "Consultar Producto/Plan"
           Me.fraDatos.Enabled = False
       Case sghEliminar
           Me.Caption = "Eliminar Producto/Plan"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
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
                   'LimpiarFormulario
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
   If Me.txtIdTipoFinanciamiento = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de IdTipoFinanciamiento" + Chr(13)
   End If
   If Me.txtDescripcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Descripcion" + Chr(13)
   End If
   If cmbGeneraPago.Text = "" Then
       sMensaje = sMensaje + "Elija 'Columna ROJA en Estado de Cuenta'" + Chr(13)
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
'   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_TiposFinanciamiento
           .idTipoFinanciamiento = Me.txtIdTipoFinanciamiento
           .Descripcion = Me.txtDescripcion.Text
           .esOficina = Me.chkEsOficina.Value
           .esSalida = Me.chkEsSalida.Value
           .SeIngresPrecios = Me.chkSeIngresPrecios.Value
           .EsFarmacia = Me.chkEsFarmacia.Value
           .idCajaTiposComprobante = Val(mo_cmbCajaTiposComprobante.BoundText)
           .tipoVenta = IIf(Me.optPreVenta = True, "P", IIf(Me.optVentaD.Value = True, "D", IIf(Me.optVtaSinPlan.Value = True, "N", "")))
           .SeImprimeComprobante = Me.chkSeImprimeComprobante.Value
           .esFuenteFinanciamiento = IIf(Me.chkEsFuenteFinanciamiento.Value = ssCBChecked, 1, 0)
           .GeneraPago = cmbGeneraPago.ListIndex
           .idTipoConcepto = Val(mo_cmbTipoConceptoF.BoundText)
           .IdUsuarioAuditoria = ml_idUsuario
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.TiposFinanciamientoAgregar(mo_TiposFinanciamiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.TiposFinanciamientoModificar(mo_TiposFinanciamiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
   Dim oRsTmp As New Recordset
   Set oRsTmp = mo_ReglasAdmision.AtencionesSeleccionarPorIdTipoFinanciamiento(Val(Me.txtIdTipoFinanciamiento.Text))
   If oRsTmp.RecordCount > 0 Then
      MsgBox "Ya existen Atenciones, no podrá Eliminar", vbInformation, Me.Caption
      Exit Function
   Else
      oRsTmp.Close
      Set oRsTmp = mo_ReglasCaja.CajaComprobantesPagosSeleccionarPorIdTipoFinanciamiento(Val(Me.txtIdTipoFinanciamiento.Text))
      If oRsTmp.RecordCount > 0 Then
         MsgBox "Ya existen Comprobantes Emitidos, no podrá Eliminar", vbInformation, Me.Caption
         Exit Function
      End If
   End If
   Set oRsTmp = Nothing
   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.TiposFinanciamientoEliminar(mo_TiposFinanciamiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

        Set mo_TiposFinanciamiento = mo_AdminFacturacion.TiposFinanciamientoSeleccionarPorId(Me.idTipoFinanciamiento)
        If mo_AdminFacturacion.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminFacturacion.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
        If Not mo_TiposFinanciamiento Is Nothing Then
            With mo_TiposFinanciamiento
                Me.idTipoFinanciamiento = .idTipoFinanciamiento
                txtIdTipoFinanciamiento = .idTipoFinanciamiento
                Me.txtDescripcion.Text = .Descripcion
                Me.chkEsOficina.Value = .esOficina
                Me.chkEsSalida.Value = IIf(.esSalida = True, 1, 0)
                Me.chkSeIngresPrecios.Value = .SeIngresPrecios
                Me.chkEsFarmacia.Value = .EsFarmacia
                mo_cmbCajaTiposComprobante.BoundText = .idCajaTiposComprobante
                If Not IsNull(.tipoVenta) Then
                   Select Case .tipoVenta
                   Case "D"
                      Me.optVentaD.Value = True
                   Case "P"
                      Me.optPreVenta.Value = True
                   Case Else
                      Me.optVtaSinPlan.Value = True
                   End Select
                End If
                chkSeImprimeComprobante.Value = IIf(.SeImprimeComprobante = True, 1, 0)
                chkEsFuenteFinanciamiento.Value = IIf(.esFuenteFinanciamiento = True, 1, 0)
                cmbGeneraPago.ListIndex = .GeneraPago
                mo_cmbTipoConceptoF.BoundText = .idTipoConcepto
                mb_ExistenDatos = True
            End With
        Else
            mb_ExistenDatos = False
            Exit Sub
        End If
       If Val(txtIdTipoFinanciamiento.Text) < 11 Or Val(txtIdTipoFinanciamiento.Text) = 1000 Then
          btnAceptar.Enabled = False
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.idTipoFinanciamiento = 0
           Me.txtIdTipoFinanciamiento = ""
           Me.txtDescripcion.Text = ""
   
End Sub

Private Sub txtIdTipoFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdTipoFinanciamiento
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdTipoFinanciamiento_LostFocus()
   mo_Formulario.MarcarComoVacio txtIdTipoFinanciamiento
End Sub

Private Sub txtIdTipoFinanciamiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


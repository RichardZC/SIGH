VERSION 5.00
Begin VB.Form CatalogoBienesSoloFarmacia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8880
   Icon            =   "CatalogoBienesSoloFarmacia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFarmacia 
      Caption         =   "TipoProducto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   60
      TabIndex        =   14
      Top             =   1200
      Width           =   8760
      Begin VB.Frame frmPrecios 
         Height          =   1635
         Left            =   45
         TabIndex        =   22
         Top             =   570
         Width           =   4320
         Begin VB.TextBox txtPrCompra 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   27
            Top             =   150
            Width           =   1395
         End
         Begin VB.TextBox txtPrDistribucion 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   26
            Top             =   540
            Width           =   1395
         End
         Begin VB.TextBox txtPrDonaciones 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   25
            Top             =   1290
            Width           =   1395
         End
         Begin VB.TextBox txtPrVenta 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   24
            Top             =   900
            Width           =   1395
         End
         Begin VB.CommandButton cmdCalculaPrec 
            Caption         =   "..."
            Height          =   345
            Left            =   3570
            TabIndex        =   23
            ToolTipText     =   "Calcula Precio de Distribución y Precio de Venta en base al precio de Compra"
            Top             =   150
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio de compra"
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
            TabIndex        =   31
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio de Distribución"
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
            TabIndex        =   30
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio de Venta"
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
            TabIndex        =   29
            Top             =   990
            Width           =   1320
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio para Donaciones"
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
            TabIndex        =   28
            Top             =   1350
            Width           =   1890
         End
      End
      Begin VB.ComboBox cmbTpSISMED 
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
         ItemData        =   "CatalogoBienesSoloFarmacia.frx":0CCA
         Left            =   7260
         List            =   "CatalogoBienesSoloFarmacia.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2100
         Width           =   1425
      End
      Begin VB.CheckBox chkActualizaPV 
         Alignment       =   1  'Right Justify
         Caption         =   "Actualiza el Precio de Venta en los demás IAFA"
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
         Left            =   4365
         TabIndex        =   6
         Top             =   1365
         Value           =   1  'Checked
         Width           =   4245
      End
      Begin VB.ComboBox cmbTipoProducto 
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
         ItemData        =   "CatalogoBienesSoloFarmacia.frx":0CED
         Left            =   7260
         List            =   "CatalogoBienesSoloFarmacia.frx":0CF7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1770
         Width           =   1425
      End
      Begin VB.TextBox txtStockM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7230
         MaxLength       =   20
         TabIndex        =   5
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox txtPrUltCompra 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   7230
         MaxLength       =   20
         TabIndex        =   4
         Top             =   570
         Width           =   1395
      End
      Begin VB.ComboBox cmbEstado 
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
         ItemData        =   "CatalogoBienesSoloFarmacia.frx":0D10
         Left            =   2205
         List            =   "CatalogoBienesSoloFarmacia.frx":0D1A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2325
         Width           =   1425
      End
      Begin VB.ComboBox cmbIdTipoSalida 
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
         ItemData        =   "CatalogoBienesSoloFarmacia.frx":0D30
         Left            =   2220
         List            =   "CatalogoBienesSoloFarmacia.frx":0D32
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   6420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Producto SISMED"
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
         Left            =   5340
         TabIndex        =   21
         Top             =   2175
         Width           =   1860
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Producto"
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
         Left            =   6060
         TabIndex        =   19
         Top             =   1830
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Mínimo"
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
         Left            =   6105
         TabIndex        =   18
         Top             =   1020
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de última compra"
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
         Left            =   5220
         TabIndex        =   17
         Top             =   630
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Left            =   165
         TabIndex        =   16
         Top             =   2385
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Salida"
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
         Left            =   195
         TabIndex        =   15
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
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
      Height          =   1170
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   8760
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         MaxLength       =   250
         TabIndex        =   1
         Top             =   660
         Width           =   6435
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   0
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   4080
      Width           =   8760
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoBienesSoloFarmacia.frx":0D34
         DownPicture     =   "CatalogoBienesSoloFarmacia.frx":1194
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
         Left            =   2985
         Picture         =   "CatalogoBienesSoloFarmacia.frx":1609
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoBienesSoloFarmacia.frx":1A7E
         DownPicture     =   "CatalogoBienesSoloFarmacia.frx":1F42
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
         Left            =   4530
         Picture         =   "CatalogoBienesSoloFarmacia.frx":242E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoBienesSoloFarmacia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Muestra tipo salida de un Medicamento e Insumo
'        Programado por: Castro W
'        Fecha: Agosto 2004
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_CatalogoBienesInsumos As New DOCatalogoBienesInsumos
Dim oCatalogoBienesPrecios As New FinanciamientoCatalogoBien
Dim oDoCatalogoBienesPrecios As New DoFinanciamientoCatalogoBien
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminComun As New ReglasComunes
Dim mo_cmbIdClasificacionBienInsumo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdTipoSalida As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTpSISMED As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdPlanCatalogo As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
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
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdPlanCatalogo(lValue As Long)
   ml_IdPlanCatalogo = lValue
End Property
Property Get IdPlanCatalogo() As Long
   IdPlanCatalogo = ml_IdPlanCatalogo
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






Private Sub cmbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbEstado
    AdministrarKeyPreview KeyCode

End Sub


Private Sub cmbIdTipoSalida_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSalida
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbTipoProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoProducto
    AdministrarKeyPreview KeyCode

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdCalculaPrec_Click()
    If Val(txtPrCompra.Text) > 0 Then
        txtPrDistribucion.Text = Round(CDbl(txtPrCompra.Text) + (CDbl(lcBuscaParametro.SeleccionaFilaParametro(307)) * CDbl(txtPrCompra.Text) / 100), 2)
        txtPrVenta.Text = Round(CDbl(txtPrCompra.Text) + ((CDbl(lcBuscaParametro.SeleccionaFilaParametro(307)) + CDbl(lcBuscaParametro.SeleccionaFilaParametro(307))) * CDbl(txtPrCompra.Text) / 100), 2)
    End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoSalida.MiComboBox = cmbIdTipoSalida
    Set mo_cmbTpSISMED.MiComboBox = cmbTpSISMED
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       mo_Formulario.HabilitarDeshabilitar cmbTpSISMED, False
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Bien e Insumo"
       Case sghModificar
           Me.Caption = "Modificar Bien e Insumo"
       Case sghConsultar
           Me.Caption = "Consultar Bien e Insumo"
       Case sghEliminar
           Me.Caption = "Eliminar Bien e Insumo"
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
   
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtNombre) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre" + Chr(13)
   End If
   If cmbIdTipoSalida.Text = "" Then
       sMensaje = sMensaje + "Elija el Tipo de Salida" + Chr(13)
   End If
   If CDbl(txtPrCompra.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio de la última Compra" + Chr(13)
   End If
   If CDbl(txtPrDistribucion.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio de Distribución" + Chr(13)
   End If
   If CDbl(txtPrVenta.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el Precio de Venta" + Chr(13)
   End If
   If Val(txtPrCompra.Text) > 0 And Val(txtPrDistribucion.Text) > 0 And Val(txtPrVenta.Text) > 0 Then
       If Not (CDbl(txtPrVenta.Text) > CDbl(txtPrDistribucion.Text) And CDbl(txtPrDistribucion.Text) > CDbl(txtPrCompra.Text)) Then
          If MsgBox("Se tiene que seguir el orden: Pr.Venta>Pr.Distribución>Pr.Compra" + Chr(13) + "¿Desea grabar?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
             sMensaje = sMensaje + "Se tiene que seguir el orden: Pr.Venta>Pr.Distribución>Pr.Compra" + Chr(13)
          End If
       End If
   End If
   If cmbTipoProducto.Text = "" Then
      sMensaje = sMensaje + "Elija el Tipo de Producto" + Chr(13)
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
   'Me.txtPrecioUnitario = Replace(Me.txtPrecioUnitario, ".", ",")
   With mo_CatalogoBienesInsumos
        .Codigo = Me.txtCodigo.Text
        .Nombre = Me.txtNombre.Text
        .IdUsuarioAuditoria = Me.idUsuario
        .idTipoSalidaBienInsumo = Val(mo_cmbIdTipoSalida.BoundText)
        .PrecioCompra = CDbl(txtPrCompra.Text)
        .PrecioDistribucion = CDbl(txtPrDistribucion.Text)
        .PrecioDonacion = CDbl(txtPrDonaciones.Text)
        .StockMinimo = Val(txtStockM.Text)
        .TipoProducto = cmbTipoProducto.ListIndex
        .TipoProductoSismed = Chr(mo_cmbTpSISMED.BoundText)
   End With
   With oDoCatalogoBienesPrecios
        .PrecioUnitario = CDbl(txtPrVenta.Text)
        .Activo = cmbEstado.ListIndex
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminComun.CatalogoBienesInsumosAgregar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_AdminComun.CatalogoBienesInsumosModificar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
    If ModificarDatos Then
        Dim oConexion As New ADODB.Connection
        Dim lnPrecio As Double
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        
        Set oCatalogoBienesPrecios.Conexion = oConexion
        If oDoCatalogoBienesPrecios.Activo = False Then
            ModificarDatos = oCatalogoBienesPrecios.Eliminar(oDoCatalogoBienesPrecios)
        Else
            If chkActualizaPV.Value = 0 Then
               ModificarDatos = oCatalogoBienesPrecios.Modificar(oDoCatalogoBienesPrecios)
            Else
               Dim oRsTmp1 As New Recordset
               Dim oRsTmp As New Recordset
               Dim lcSql As String
               Dim lbNuevo As Boolean
               Dim oDoFinanciamientoCatalogoBien As New DoFinanciamientoCatalogoBien, oFinanciamientoCatalogoBien As New FinanciamientoCatalogoBien
               Set oFinanciamientoCatalogoBien.Conexion = oConexion
               'Agrega Insumo en otros Tipo de Financiamiento y actualiza Precio
               Set oRsTmp = mo_ReglasFacturacion.CatalogoBienesInsumosHospSeleccionarXIdProducto(oDoCatalogoBienesPrecios.IdProducto)
               Set oRsTmp1 = mo_AdminComun.TiposFinanciamientoSegunFiltro("SeIngresPrecios=1")
               If oRsTmp1.RecordCount > 0 Then
                  oRsTmp1.MoveFirst
                  Do While Not oRsTmp1.EOF
                     lbNuevo = True
                     If oRsTmp.RecordCount > 0 Then
                        oRsTmp.MoveFirst
                        oRsTmp.Find "idTipoFinanciamiento=" & oRsTmp1.Fields!idTipoFinanciamiento
                        If Not oRsTmp.EOF Then
                           lbNuevo = False
                        End If
                     End If
                     
                     lnPrecio = oDoCatalogoBienesPrecios.PrecioUnitario
                     If Not IsNull(oRsTmp1!porcPrecio) Then
                        If oRsTmp1!porcPrecio > 0 Then
                           lnPrecio = oDoCatalogoBienesPrecios.PrecioUnitario + Round(oDoCatalogoBienesPrecios.PrecioUnitario * oRsTmp1!porcPrecio / 100, 2)
                        End If
                     End If

                     oDoFinanciamientoCatalogoBien.IdProducto = oDoCatalogoBienesPrecios.IdProducto
                     oDoFinanciamientoCatalogoBien.idTipoFinanciamiento = oRsTmp1.Fields!idTipoFinanciamiento
                     oDoFinanciamientoCatalogoBien.Activo = 1
                     oDoFinanciamientoCatalogoBien.PrecioUnitario = lnPrecio
                     oDoFinanciamientoCatalogoBien.IdUsuarioAuditoria = ml_idUsuario
                     If lbNuevo = True Then
                        If oFinanciamientoCatalogoBien.Insertar(oDoFinanciamientoCatalogoBien) = False Then
                        End If
                     Else
                         oDoFinanciamientoCatalogoBien.IdPlanCatalogo = oRsTmp.Fields!IdPlanCatalogo
                         If oFinanciamientoCatalogoBien.Modificar(oDoFinanciamientoCatalogoBien) = False Then
                         End If
                     End If
                     oRsTmp1.MoveNext
                  Loop
               End If
               Set oRsTmp = Nothing
               Set oRsTmp1 = Nothing
               Set oDoFinanciamientoCatalogoBien = Nothing
               Set oFinanciamientoCatalogoBien = Nothing
               ModificarDatos = True
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
    End If
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminComun.CatalogoBienesInsumosEliminar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CatalogoBienesInsumos = mo_AdminComun.CatalogoBienesInsumosSeleccionarPorId(Me.IdProducto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos " + Chr(13) + mo_AdminComun.MensajeError, vbInformation, Me.Caption
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CatalogoBienesInsumos Is Nothing Then
        With mo_CatalogoBienesInsumos
            mo_cmbIdTipoSalida.BoundText = .idTipoSalidaBienInsumo
            Me.txtNombre = .Nombre
            Me.txtCodigo = .Codigo
            txtPrCompra.Text = .PrecioCompra
            txtPrDistribucion.Text = .PrecioDistribucion
            txtPrDonaciones.Text = .PrecioDonacion
            txtPrUltCompra.Text = .PrecioUltCompra
            txtStockM.Text = .StockMinimo
            mb_ExistenDatos = True
            cmbTipoProducto.ListIndex = .TipoProducto
            mo_cmbTpSISMED.BoundText = Asc(.TipoProductoSismed)
            
        End With
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oCatalogoBienesPrecios.Conexion = oConexion
    oDoCatalogoBienesPrecios.IdPlanCatalogo = ml_IdPlanCatalogo
    If oCatalogoBienesPrecios.SeleccionarPorId(oDoCatalogoBienesPrecios) Then
       txtPrVenta.Text = oDoCatalogoBienesPrecios.PrecioUnitario
       cmbEstado.ListIndex = IIf(oDoCatalogoBienesPrecios.Activo, 1, 0)
    End If
    oConexion.Close
    Set oConexion = Nothing
    'debb-08/11/2016
    If mo_ReglasFarmacia.CatalogoDIGEMIDesCodigoPaquete(txtCodigo.Text) = True Then
       MsgBox "No podrá MODIFICAR PRECIOS, porque es un CODIGO DE PAQUETE", vbInformation, Me.Caption
       frmPrecios.Enabled = False
    End If
    '
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdProducto = 0
    mo_cmbIdTipoSalida.BoundText = ""
    mo_cmbTpSISMED.BoundText = ""
    
    Me.txtNombre = ""
    Me.txtCodigo = ""
End Sub

Sub CargarComboBoxes()
       
    mo_cmbIdTipoSalida.BoundColumn = "idTipoSalidaBienInsumo"
    mo_cmbIdTipoSalida.ListField = "Tipo"
    Set mo_cmbIdTipoSalida.RowSource = mo_ReglasFarmacia.farmTipoSalidaBienInsumoDevuelveTodos

    mo_cmbTpSISMED.BoundColumn = "identificador"
    mo_cmbTpSISMED.ListField = "Descripcion"
    Set mo_cmbTpSISMED.RowSource = mo_ReglasFarmacia.farmTipoProductosSismedDevuelveTodos
    
End Sub





Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtPrCompra_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrCompra
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtPrCompra_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtPrDistribucion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrDistribucion
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtPrDistribucion_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtPrDonaciones_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrDonaciones
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtPrDonaciones_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub





Private Sub txtPrUltCompra_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrUltCompra
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtPrVenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrVenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtPrVenta_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub



Private Sub txtStockM_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtStockM
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtStockM_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

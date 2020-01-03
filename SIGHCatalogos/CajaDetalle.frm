VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form CajaDetalle 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "CajaDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Generación de Comprobantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   120
      TabIndex        =   10
      Top             =   4755
      Width           =   10560
      Begin UltraGrid.SSUltraGrid grdItems 
         Height          =   2130
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   3757
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nro Documento"
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
      Height          =   4665
      Left            =   120
      TabIndex        =   9
      Top             =   90
      Width           =   10560
      Begin VB.Frame Frame 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   4725
         TabIndex        =   28
         Top             =   195
         Width           =   5745
         Begin VB.ComboBox cmbCentroCostos 
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
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1125
            Width           =   4260
         End
         Begin VB.ComboBox cmbPartidas 
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
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   750
            Width           =   4260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Centro Costo"
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
            Left            =   180
            TabIndex        =   32
            Top             =   1170
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Partida"
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
            Left            =   180
            TabIndex        =   31
            Top             =   765
            Width           =   555
         End
         Begin VB.Label Label 
            Caption         =   "* solo si la caja usará item con descripciones grandes, no son códigos CPT, usado en FACTURAS que no son a pacientes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   195
            TabIndex        =   30
            Top             =   210
            Width           =   5145
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "(Impresora Windows NO DEFAULT, que se usará en Boletas de FARMACIA)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1230
         Left            =   90
         TabIndex        =   19
         Top             =   3315
         Width           =   10335
         Begin VB.CheckBox chkCinta2 
            Caption         =   "Formato Cinta"
            Height          =   375
            Left            =   8760
            TabIndex        =   27
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cmbTipoComprobante2 
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
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   720
            Width           =   2415
         End
         Begin VB.ComboBox cboImpresora2 
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
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtSerieImpresora2 
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
            Left            =   3840
            MaxLength       =   250
            TabIndex        =   20
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Comprobante"
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
            Left            =   6240
            TabIndex        =   25
            Top             =   480
            Width           =   1530
         End
         Begin VB.Label Label5 
            Caption         =   "Impresora 2"
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
            Left            =   240
            TabIndex        =   23
            Top             =   420
            Width           =   1035
         End
         Begin VB.Label Label12 
            Caption         =   "Nº Serie :"
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
            Left            =   3840
            TabIndex        =   22
            Top             =   480
            Width           =   1035
         End
      End
      Begin VB.Frame fraServicios 
         Caption         =   "(Impresora Windows DEFAULT, que se usará en Boletas de SERVICIOS)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   105
         TabIndex        =   12
         Top             =   1890
         Width           =   10335
         Begin VB.CheckBox chkCinta1 
            Caption         =   "Formato Cinta"
            Height          =   375
            Left            =   8760
            TabIndex        =   26
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cboImpresoraDefault 
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
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   3615
         End
         Begin VB.ComboBox cmbIdTipoComprobante 
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
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtSerieImpresoraDefault 
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
            Left            =   3840
            MaxLength       =   250
            TabIndex        =   13
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Impresora 1"
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
            Left            =   240
            TabIndex        =   18
            Top             =   420
            Width           =   1035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Comprobante"
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
            Left            =   6240
            TabIndex        =   17
            Top             =   480
            Width           =   1530
         End
         Begin VB.Label Label10 
            Caption         =   "Nº Serie :"
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
            Left            =   3840
            TabIndex        =   16
            Top             =   480
            Width           =   1035
         End
      End
      Begin VB.TextBox txtNombrePc 
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
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   2
         Top             =   975
         Width           =   1830
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
         Height          =   330
         Left            =   1170
         MaxLength       =   250
         TabIndex        =   1
         Top             =   585
         Width           =   3420
      End
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
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   0
         Top             =   225
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Pc"
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
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblCodigoCIE2004 
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
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
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
         TabIndex        =   7
         Top             =   690
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   10560
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CajaDetalle.frx":0CCA
         DownPicture     =   "CajaDetalle.frx":112A
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
         Left            =   3892
         Picture         =   "CajaDetalle.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CajaDetalle.frx":1A14
         DownPicture     =   "CajaDetalle.frx":1ED8
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
         Left            =   5437
         Picture         =   "CajaDetalle.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CajaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD 19/06/2005 [Todo el Archivo]
'MZD02
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
Dim mo_CajaCaja As New DOCajaCaja
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdCaja As Long
Dim mo_AdminCaja As New ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mrs_NroDocumentos As New ADODB.Recordset
Dim mo_NroDocumentos As New Collection
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_cmbIdTipoComprobante As New ListaDespleglable
Dim mo_cmbIdTipoComprobante2 As New ListaDespleglable
Dim mo_cmbPartidas As New ListaDespleglable
Dim mo_cmbCentroCostos As New ListaDespleglable
'mgaray20141003
Dim mo_cmbImpresoraDefault As New ListaDespleglable
Dim mo_cmbImpresora2 As New ListaDespleglable

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
Property Let IdCaja(lValue As Long)
   ml_IdCaja = lValue
End Property
Property Get IdCaja() As Long
   IdCaja = ml_IdCaja
End Property
'mgaray20141003
Private Sub cboImpresora2_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cboImpresora2
   AdministrarKeyPreview KeyCode
End Sub
'mgaray20141003
Private Sub cboImpresoraDefault_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cboImpresoraDefault
   AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoComprobante
   AdministrarKeyPreview KeyCode
End Sub

Private Sub grdItems_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (Cell.Column.Key = "NroSerie" Or Cell.Column.Key = "NroDocumento" Or Cell.Column.Key = "NroDocumentoInicial" Or Cell.Column.Key = "NroDocumentoFinal" Or Cell.Column.Key = "FacturaSinIGV") Then
        Cancel = True
    End If
End Sub

Private Sub grdItems_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdItems.Bands(0).Columns("IdCaja").Hidden = True
    grdItems.Bands(0).Columns("IdTipoComprobante").Hidden = True
    
    grdItems.Bands(0).Columns("TipoComprobante").Header.Caption = "Tipo Comprobante"
    grdItems.Bands(0).Columns("TipoComprobante").Width = 2000
    
    grdItems.Bands(0).Columns("NroSerie").Header.Caption = "Nº Serie"
    grdItems.Bands(0).Columns("NroSerie").Width = 900

    grdItems.Bands(0).Columns("NroDocumentoInicial").Header.Caption = "Nro.Doc.Inicial"
    grdItems.Bands(0).Columns("NroDocumentoInicial").Width = 1500
    
    grdItems.Bands(0).Columns("NroDocumentoFinal").Header.Caption = "Nro.Doc.Final"
    grdItems.Bands(0).Columns("NroDocumentoFinal").Width = 1500

    grdItems.Bands(0).Columns("NroDocumento").Header.Caption = "Ult.Doc.Generado"
    grdItems.Bands(0).Columns("NroDocumento").Width = 1800

End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
   AdministrarKeyPreview KeyCode
End Sub
Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(txtCodigo)
   mo_Formulario.MarcarComoVacio txtCodigo
End Sub
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

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
         Frame2.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         Frame1.Enabled = False
         Frame2.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
       Set mo_cmbIdTipoComprobante2.MiComboBox = Me.cmbTipoComprobante2
       
       Set mo_cmbPartidas.MiComboBox = cmbPartidas
       mo_cmbPartidas.BoundColumn = "IdPartidaPresupuestal"
       mo_cmbPartidas.ListField = "Descripcion"
       Set mo_cmbPartidas.RowSource = mo_ReglasComunes.PartidasPresupuestalesSeleccionarTodos
       
       Set mo_cmbCentroCostos.MiComboBox = cmbCentroCostos
       mo_cmbCentroCostos.BoundColumn = "IdCentroCosto"
       mo_cmbCentroCostos.ListField = "Descripcion"
       Set mo_cmbCentroCostos.RowSource = mo_ReglasComunes.CentrosCostoSeleccionarTodos
       
       
       'mgaray20141003
       Set mo_cmbImpresoraDefault.MiComboBox = cboImpresoraDefault
       Set mo_cmbImpresora2.MiComboBox = cboImpresora2
       
       GenerarRecordsetTemporal
       CargaComboBoxes
       Select Case mi_Opcion
       Case sghAgregar
           CargarDocumentosCaja
           Me.Caption = "Agregar Caja"
       Case sghModificar
           Me.Caption = "Modificar Caja"
       Case sghConsultar
           Me.Caption = "Consultar Caja"
       Case sghEliminar
           Me.Caption = "Eliminar Caja"
       End Select
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

Sub CargaComboBoxes()
    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
    
    mo_cmbIdTipoComprobante2.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante2.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante2.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
    
    'mgaray20141003
    Dim oImpresoraUtil As New ImpresoraUtil
    
    mo_cmbImpresoraDefault.BoundColumn = "printerName"
    mo_cmbImpresoraDefault.ListField = "printerName"
    Set mo_cmbImpresoraDefault.RowSource = oImpresoraUtil.listaImpresorasInstaladas()
    
    mo_cmbImpresora2.BoundColumn = "printerName"
    mo_cmbImpresora2.ListField = "printerName"
    Set mo_cmbImpresora2.RowSource = oImpresoraUtil.listaImpresorasInstaladas()
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
                   LimpiarFormulario
                   Me.txtCodigo.SetFocus
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
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código " + Chr(13)
   End If
   If Me.txtDescripcion = "" Then
       sMensaje = sMensaje + "Ingrese la descripción" + Chr(13)
   End If
   If Me.cmbIdTipoComprobante.Text = "" Then
       sMensaje = sMensaje + "Elija el TIPO DE COMPROBANTE" + Chr(13)
   End If
   
   'Validamos que se hayan puesto los números de documentos en el orden debido
   If Not (mrs_NroDocumentos.EOF = True And mrs_NroDocumentos.BOF = True) Then
        mrs_NroDocumentos.MoveFirst
   End If
   Do While Not mrs_NroDocumentos.EOF
        Dim lNumeroInicial As Long
        Dim lNumeroFinal As Long
        Dim lNumeroGenerado As Long
        
        lNumeroInicial = Val(IIf(IsNull(mrs_NroDocumentos!NroDocumentoInicial), "0", mrs_NroDocumentos!NroDocumentoInicial))
        lNumeroFinal = Val(IIf(IsNull(mrs_NroDocumentos!NroDocumentoFinal), "0", mrs_NroDocumentos!NroDocumentoFinal))
        lNumeroGenerado = Val(mrs_NroDocumentos!NroDocumento)
        
        If IIf(IsNull(mrs_NroDocumentos!NroSerie), "0", mrs_NroDocumentos!NroSerie) = "" Then
            MsgBox "Debe Ingresar el número de Serie", vbInformation, Me.Caption
            Exit Function
        End If
        
        If lNumeroInicial > lNumeroFinal Then
            MsgBox "El número inicial debe ser menor que el número de documento final", vbInformation, Me.Caption
            Exit Function
        End If
        If lNumeroInicial > lNumeroGenerado Then
            MsgBox "El número inicial debe ser menor que el ultimo número generado", vbInformation, Me.Caption
            Exit Function
        End If
        If lNumeroGenerado > lNumeroFinal Then
            MsgBox "El número final no puede ser menor que el último número generado", vbInformation, Me.Caption
            Exit Function
        End If
        mrs_NroDocumentos.MoveNext
   Loop
   
   
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
   With mo_CajaCaja
        .Codigo = Me.txtCodigo.Text
        .Descripcion = Me.txtDescripcion.Text
        .IdUsuarioAuditoria = Me.idUsuario
        .LoginPc = Me.txtNombrePc.Text
'        .ImpresoraDefault = Me.txtImpresoraDefault.Text
'        .Impresora2 = Me.txtImpresora2.Text
        '**** Programa: se arreglo la cantidad de caracteres  Right("00000" + lcCodigoMed, 5))
        '**** Programado por:Eder Yamill Palomino Espinoza
        '**** Fecha: 06102014
        .SerieImpresoraDefault = Me.txtSerieImpresoraDefault.Text
        .SerieImpresora2 = Me.txtSerieImpresora2.Text
        .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
        'mgaray20141003
        .ImpresoraDefault = cboImpresoraDefault.Text ' .BoundText
        .Impresora2 = cboImpresora2.Text
        .IdTipoComprobante2 = Val(mo_cmbIdTipoComprobante2.BoundText)
        .FormatoImp2Cinta = Me.chkCinta2.Value
        .FormatoImpDefaultCinta = Me.chkCinta1.Value
        .IdPartida = Val(mo_cmbPartidas.BoundText)
        .IdCentroCosto = Val(mo_cmbCentroCostos.BoundText)
   End With
   CargarNroDocumentosAlObjetoDatos mo_NroDocumentos
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminCaja.CajaAgregar(mo_CajaCaja, mo_NroDocumentos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminCaja.CajaModificar(mo_CajaCaja, mo_NroDocumentos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminCaja.CajaEliminar(mo_CajaCaja, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    Set mo_CajaCaja = mo_AdminCaja.CajaSeleccionarPorId(Me.IdCaja)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminCaja.MensajeError, vbInformation, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CajaCaja Is Nothing Then
        With mo_CajaCaja
            Me.txtCodigo = .Codigo
            Me.txtDescripcion = .Descripcion
            Me.txtNombrePc.Text = .LoginPc
'            Me.txtImpresoraDefault.Text = .ImpresoraDefault
'            Me.txtImpresora2.Text = .Impresora2
            '**** Programa: se arreglo la cantidad de caracteres  Right("00000" + lcCodigoMed, 5))
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
            Me.txtSerieImpresoraDefault.Text = .SerieImpresoraDefault
            Me.txtSerieImpresora2.Text = .SerieImpresora2
            mo_cmbIdTipoComprobante.BoundText = .IdTipoComprobante
            mo_cmbIdTipoComprobante2.BoundText = .IdTipoComprobante2
'            Me.chkCinta1.Value = .FormatoImpDefaultCinta
            If .FormatoImpDefaultCinta Then
                Me.chkCinta1.Value = 1
            Else
                Me.chkCinta1.Value = 0
            End If
'            Me.chkCinta2.Value = .FormatoImp2Cinta
            If .FormatoImp2Cinta Then
                Me.chkCinta2.Value = 1
            Else
                Me.chkCinta2.Value = 0
            End If
            'mgaray20141003
            On Error Resume Next
            cboImpresoraDefault.Text = .ImpresoraDefault
            cboImpresora2.Text = .Impresora2
            mo_cmbPartidas.BoundText = .IdPartida
            mo_cmbCentroCostos.BoundText = .IdCentroCosto
            Err = 0
            mb_ExistenDatos = True
        End With
        CargarDocumentosCaja
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

Sub GenerarRecordsetTemporal()
    Set mrs_NroDocumentos = New ADODB.Recordset
    With mrs_NroDocumentos
          .Fields.Append "IdCaja", adInteger, 4, adFldIsNullable
          .Fields.Append "IdTipoComprobante", adInteger, 4, adFldIsNullable
          .Fields.Append "TipoComprobante", adVarChar, 100, adFldIsNullable
          'sunat facturador nov2016 fcv
          .Fields.Append "NroSerie", adVarChar, 4, adFldIsNullable
          .Fields.Append "NroDocumentoInicial", adVarChar, 8, adFldIsNullable
          .Fields.Append "NroDocumentoFinal", adVarChar, 8, adFldIsNullable
          .Fields.Append "NroDocumento", adVarChar, 8, adFldIsNullable
          .Fields.Append "FacturaSinIGV", adBoolean
          'sunat facturador nov2016 fcv
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdItems.DataSource = mrs_NroDocumentos
    
End Sub

Sub CargarDocumentosCaja()
    Dim rsDocumentos As New Recordset
    GenerarRecordsetTemporal
    
    Set rsDocumentos = mo_AdminCaja.NroDocumentoSeleccionarPorIdCaja(ml_IdCaja)
    Do While Not rsDocumentos.EOF
        With mrs_NroDocumentos
            .AddNew
            .Fields!IdCaja = rsDocumentos!IdCaja
            .Fields!IdTipoComprobante = rsDocumentos!IdTipoComprobante
            .Fields!TipoComprobante = rsDocumentos!TipoComprobante
            .Fields!NroSerie = rsDocumentos!NroSerie
            'sunat facturador nov2016 fcv
            .Fields!NroDocumentoInicial = Right("00000000" & Trim(rsDocumentos!NroDocumentoInicial), 8)
            .Fields!NroDocumentoFinal = Right("00000000" & Trim(rsDocumentos!NroDocumentoFinal), 8)
            .Fields!NroDocumento = Right("00000000" & Trim(rsDocumentos!NroDocumento), 8)
            .Fields!FacturaSinIGV = rsDocumentos!FacturaSinIGV
            'sunat facturador nov2016 fcv
        End With
        rsDocumentos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores Me.grdItems, SIGHEntidades.GrillaConFilasBicolor

End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdCaja = 0
    Me.txtCodigo = ""
    Me.txtDescripcion = ""
    Me.txtNombrePc.Text = ""
    Me.txtNombrePc.Text = ""
    'ypalomino01102014
    Me.txtSerieImpresoraDefault.Text = ""
    Me.txtSerieImpresora2.Text = ""
'    Me.txtImpresoraDefault.Text = ""
'    Me.txtImpresora2.Text = ""
    'mgaray20141003
    On Error Resume Next
    cboImpresoraDefault.Text = ""
    cboImpresora2.Text = ""
    CargarDocumentosCaja
    Err = 0
End Sub

Sub CargarNroDocumentosAlObjetoDatos(oNroDocumentos As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS DOCUMENTOS
    '---------------------------------------------------------------------------------
    Dim oNroDocumento As DOCajaNroDocumento
    
    Set oNroDocumentos = New Collection
    
    If Not (mrs_NroDocumentos.BOF And mrs_NroDocumentos.EOF) Then
        mrs_NroDocumentos.MoveFirst
        Do While Not mrs_NroDocumentos.EOF
            Set oNroDocumento = New DOCajaNroDocumento
            oNroDocumento.IdTipoComprobante = mrs_NroDocumentos!IdTipoComprobante
'            oNroDocumento.nroSerie = Right("000" & Trim(mrs_NroDocumentos!nroSerie), 3)
'            oNroDocumento.NroDocumento = Right("000000" & Trim(mrs_NroDocumentos!NroDocumento), 6)
'            oNroDocumento.NroDocumentoInicial = Right("000000" & Trim(mrs_NroDocumentos!NroDocumentoInicial), 6)
'            oNroDocumento.NroDocumentoFinal = Right("000000" & Trim(mrs_NroDocumentos!NroDocumentoFinal), 6)
            oNroDocumento.NroSerie = Trim(IIf(IsNull(mrs_NroDocumentos!NroSerie) = True, "", mrs_NroDocumentos!NroSerie))
            oNroDocumento.NroDocumento = Trim(mrs_NroDocumentos!NroDocumento)
            oNroDocumento.NroDocumentoInicial = Trim(mrs_NroDocumentos!NroDocumentoInicial)
            oNroDocumento.NroDocumentoFinal = Trim(mrs_NroDocumentos!NroDocumentoFinal)
            oNroDocumento.IdUsuarioAuditoria = ml_idUsuario
            oNroDocumento.FacturaSinIGV = mrs_NroDocumentos!FacturaSinIGV
            oNroDocumentos.Add oNroDocumento
            mrs_NroDocumentos.MoveNext
        Loop
    End If
End Sub






'Private Sub txtImpresora2_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtImpresora2
'   AdministrarKeyPreview KeyCode
'
'End Sub

'Private Sub txtImpresoraDefault_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtImpresoraDefault
'   AdministrarKeyPreview KeyCode
'
'End Sub

Private Sub txtNombrePc_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNombrePc

End Sub

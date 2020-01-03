VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form rProductosIngresados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..."
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rProductosIngresados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabReportes 
      Height          =   7800
      Left            =   -30
      TabIndex        =   0
      Top             =   -15
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   13758
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Productos Ingresados por Proveedor"
      TabPicture(0)   =   "rProductosIngresados.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(1)=   "grdProductos"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "fraCabecera"
      Tab(0).Control(4)=   "TxtBusca"
      Tab(0).Control(5)=   "CmbFiltro"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Productos vendidos según forma de pago"
      TabPicture(1)   =   "rProductosIngresados.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdVentas"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "rProductosIngresados.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.ComboBox CmbFiltro 
         Height          =   330
         ItemData        =   "rProductosIngresados.frx":0D1E
         Left            =   -74250
         List            =   "rProductosIngresados.frx":0D28
         TabIndex        =   47
         Top             =   6270
         Width           =   1815
      End
      Begin VB.TextBox TxtBusca 
         Height          =   330
         Left            =   -72450
         TabIndex        =   46
         Top             =   6270
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   105
         TabIndex        =   24
         Top             =   450
         Width           =   13665
         Begin VB.CheckBox chkExcel 
            Caption         =   "En excel"
            Height          =   210
            Left            =   12225
            TabIndex        =   50
            Top             =   750
            Width           =   1125
         End
         Begin VB.TextBox txtCajero 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6045
            MaxLength       =   30
            TabIndex        =   43
            Top             =   600
            Width           =   2415
         End
         Begin VB.CommandButton cmdBuscaCajero 
            Caption         =   "..."
            Height          =   315
            Left            =   8475
            TabIndex        =   42
            Top             =   600
            Width           =   315
         End
         Begin VB.ComboBox cmbTproducto 
            Height          =   330
            ItemData        =   "rProductosIngresados.frx":0D44
            Left            =   6045
            List            =   "rProductosIngresados.frx":0D51
            TabIndex        =   40
            Top             =   225
            Width           =   2760
         End
         Begin VB.CommandButton cmbRealizarBusqueda 
            Height          =   315
            Left            =   12240
            Picture         =   "rProductosIngresados.frx":0D75
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   315
            Width           =   1305
         End
         Begin VB.TextBox txtPaciente 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1095
            MaxLength       =   30
            TabIndex        =   27
            Top             =   645
            Width           =   3300
         End
         Begin VB.ComboBox cmbFarmacia 
            Height          =   330
            Left            =   1095
            TabIndex        =   26
            Top             =   240
            Width           =   3660
         End
         Begin VB.CommandButton cmdBuscaPaciente 
            Caption         =   "..."
            Height          =   315
            Left            =   4380
            TabIndex        =   25
            Top             =   645
            Width           =   360
         End
         Begin MSMask.MaskEdBox txtFinicio 
            Height          =   330
            Left            =   9840
            TabIndex        =   28
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFFinal 
            Height          =   330
            Left            =   9840
            TabIndex        =   29
            Top             =   615
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin MSMask.MaskEdBox txtHoraInicio1 
            Height          =   315
            Left            =   11295
            TabIndex        =   51
            Top             =   225
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtHoraFinal1 
            Height          =   315
            Left            =   11310
            TabIndex        =   52
            Top             =   615
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cajero"
            Height          =   210
            Left            =   5535
            TabIndex        =   44
            Top             =   660
            Width           =   510
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "T.Producto"
            Height          =   210
            Left            =   5130
            TabIndex        =   41
            Top             =   255
            Width           =   930
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Height          =   210
            Left            =   150
            TabIndex        =   35
            Top             =   2010
            Width           =   60
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Height          =   210
            Left            =   150
            TabIndex        =   34
            Top             =   1590
            Width           =   60
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Farmacia"
            Height          =   210
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "F.Final"
            Height          =   210
            Left            =   9270
            TabIndex        =   32
            Top             =   675
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "F.Inicio"
            Height          =   210
            Left            =   9225
            TabIndex        =   31
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Paciente"
            Height          =   210
            Left            =   120
            TabIndex        =   30
            Top             =   705
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   105
         TabIndex        =   21
         Top             =   6615
         Width           =   13680
         Begin VB.CommandButton cmdSalir 
            Caption         =   "Cancelar(ESC)"
            DisabledPicture =   "rProductosIngresados.frx":39BE
            DownPicture     =   "rProductosIngresados.frx":3E82
            Height          =   700
            Left            =   6998
            Picture         =   "rProductosIngresados.frx":436E
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton cmdImprime 
            Caption         =   "Imprimir"
            Height          =   700
            Left            =   5588
            Picture         =   "rProductosIngresados.frx":485A
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   10680
            TabIndex        =   23
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraCabecera 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -74910
         TabIndex        =   6
         Top             =   405
         Width           =   13665
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   12270
            Picture         =   "rProductosIngresados.frx":4D33
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   330
            Width           =   1305
         End
         Begin VB.TextBox txtRuc 
            Height          =   315
            Left            =   1530
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1065
            Width           =   1635
         End
         Begin VB.TextBox txtProveedor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3615
            MaxLength       =   30
            TabIndex        =   12
            Top             =   1080
            Width           =   3945
         End
         Begin VB.ComboBox cmbAlmDestino 
            Height          =   330
            Left            =   1530
            TabIndex        =   10
            Top             =   345
            Width           =   6045
         End
         Begin VB.ComboBox cmbAlmOrigen 
            Height          =   330
            Left            =   1530
            TabIndex        =   9
            Top             =   705
            Width           =   6045
         End
         Begin VB.CommandButton btnBuscaProv 
            Caption         =   "..."
            Height          =   315
            Left            =   3195
            TabIndex        =   7
            Top             =   1080
            Width           =   390
         End
         Begin MSMask.MaskEdBox txtFechaInicio 
            Height          =   375
            Left            =   9780
            TabIndex        =   8
            Top             =   375
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaFinal 
            Height          =   375
            Left            =   9780
            TabIndex        =   11
            Top             =   825
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Height          =   210
            Left            =   150
            TabIndex        =   20
            Top             =   2010
            Width           =   60
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Height          =   210
            Left            =   150
            TabIndex        =   19
            Top             =   1590
            Width           =   60
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Almacén Destino"
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
            Height          =   210
            Left            =   8820
            TabIndex        =   17
            Top             =   900
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
            Height          =   210
            Left            =   8775
            TabIndex        =   16
            Top             =   450
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   1095
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Almacén origen"
            Height          =   210
            Left            =   120
            TabIndex        =   14
            Top             =   727
            Width           =   1260
         End
      End
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
         Height          =   1095
         Left            =   -74910
         TabIndex        =   2
         Top             =   6645
         Width           =   13680
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar(ESC)"
            DisabledPicture =   "rProductosIngresados.frx":797C
            DownPicture     =   "rProductosIngresados.frx":7E40
            Height          =   700
            Left            =   7718
            Picture         =   "rProductosIngresados.frx":832C
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton btnImprimirDetallado 
            Caption         =   "Detallado"
            Height          =   700
            Left            =   6293
            Picture         =   "rProductosIngresados.frx":8818
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton btnImprimirConsolidado 
            Caption         =   "Consolidado"
            Height          =   700
            Left            =   4868
            Picture         =   "rProductosIngresados.frx":8CF1
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label lbtotalregistro 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   10680
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
      End
      Begin UltraGrid.SSUltraGrid grdProductos 
         Height          =   4125
         Left            =   -74910
         TabIndex        =   1
         Top             =   2085
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   7276
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "grdProductos"
      End
      Begin UltraGrid.SSUltraGrid grdVentas 
         Height          =   4665
         Left            =   105
         TabIndex        =   45
         Top             =   1680
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   8229
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "..."
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Para ORDENAR pulse clic en la CABECERA DE COLUMNA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   105
         TabIndex        =   49
         Top             =   6405
         Width           =   5025
      End
      Begin VB.Label Label16 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74895
         TabIndex        =   48
         Top             =   6330
         Width           =   615
      End
   End
End
Attribute VB_Name = "rProductosIngresados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Notas de Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim mo_cmbConceptos As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenDestino As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoDocum As New SIGHEntidades.ListaDespleglable
Dim mo_cmbFarmacia As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasComunes
Dim oRsAlmacenOrigen As New ADODB.Recordset
Dim oRsAlmacenDestino As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim rsTmp As New Recordset
Dim mrs_Tmp As New Recordset
Dim oRsProveedor As New ADODB.Recordset
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ms_MensajeError As String
Dim idAlmacenD As Long
Dim idAlmacenO As Long
Dim IdProveedor As Long
Dim lnVentaTotal As Double

Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mRs_Productos As New ADODB.Recordset
Dim mo_farmMovimiento As New sighComun.DoFarmMovimiento
Dim lnTotalDocumento As Double
Dim mo_farmMovimientoNotaIngreso As New sighComun.DOfarmMovimientoNotaIngreso
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim oDoProveedores As New DoProveedores
Dim lcTipoLocalesAlmOrigen As String
Dim lbDocumentoEsAutomatico As Boolean
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcTipoLocalesAlmDestino As String
Dim mo_lbElEstablecimentoEsCS As Boolean
Dim ml_idUsuarioCreo As Long
Dim ml_NroReporte As Long
Property Let NroReporte(lValue As Long)
   ml_NroReporte = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property







Private Sub ImprimeDocumento()
    'Dim oRptClase As New rCrystal
    Dim oDOfarmAlmacen As New DoFarmAlmacen
    
    Set oDOfarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(Val(mo_cmbAlmacenOrigen.BoundText))
    
'    MsgBox oDOfarmAlmacen
    'oRptClase.TextoDelFiltro = "NOTA DE SALIDA"
    'oRptClase.Almacen = cmbAlmDestino.Text
    'oRptClase.AlmacenO = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmOrigen.Text
    'oRptClase.HoraInicio = txtFechaInicio.Text
    'oRptClase.HoraFin = Trim(cmbTipoDocum.Text) & " - " & txtRuc.Text
    
    'oRptClase.Show vbModal
    'Set oRptClase = Nothing
    Set oDOfarmAlmacen = Nothing
End Sub

Private Sub btnBuscaProv_Click()
    BuscaProveedores.Show 1
    txtRuc.Text = BuscaProveedores.ruc
    txtRuc_KeyPress 13
End Sub

Private Sub btnBuscar_Click()


If ValidarDatosObligatorios Then
 IdRepProveedor
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    If IdProveedor = "0" Then 'consulta sin proveedor
        txtRuc.Text = ""
        txtProveedor.Text = ""
        'Set rsTmp = mo_ReglasFarmacia.FarmAlmacenProductosIngresados(txtFechaInicio.Text, txtFechaFinal.Text, idAlmacenD, idAlmacenO, oConexion)
        Set rsTmp = mo_ReglasFarmacia.FarmMovimientoSeleccionPorAlmacenProductosIngresados(txtFechaInicio.Text, txtFechaFinal.Text, idAlmacenD, idAlmacenO, oConexion)
        Set Me.grdProductos.DataSource = rsTmp
        lbtotalregistro.Caption = "Total Registros: " & rsTmp.RecordCount
    'Do While rsTmp.EOF
     '   grdProductos.Bands(0).Columns("Concepto").Header.Caption = "Concepto" & rsTmp!Concepto
      '  grdProductos.Bands(0).Columns("Concepto").Width = 2700
       ' rsTmp.MoveNext
    'Loop
    
    Else 'consulta con proveedor
        If txtProveedor.Text = "" Then MsgBox "Proveedor Incorrecto...!", vbInformation, Me.Caption: Exit Sub
        Set rsTmp = mo_ReglasFarmacia.FarmMovimientoSeleccionPorProveedorProductosIngresados(txtFechaInicio.Text, txtFechaFinal.Text, idAlmacenD, idAlmacenO, IdProveedor, oConexion)
        'Set rsReporte = mo_ReglasFarmacia.FarmMovimientoSeleccionPorProveedorProductosIngresados(mda_FechaInicio, mda_FechaFin, lnIdAlmacenDestino, lnIdAlmacenOrigen, ml_Proveedor, oConexion)
        Set Me.grdProductos.DataSource = rsTmp
        lbtotalregistro.Caption = "Total Registros: " & rsTmp.RecordCount
    End If
        'Set Me.grdProductos.DataSource = rsTmp
        'Set rsTmp = Nothing
        Set oConexion = Nothing
    
        mo_Apariencia.ConfigurarFilasBiColores Me.grdProductos, SIGHEntidades.GrillaConFilasBicolor
End If
End Sub

Function ValidarDatosObligatorios() As Boolean
    ValidarDatosObligatorios = False
        If cmbAlmDestino.ListIndex < 0 Then MsgBox "Por Favor Elija el Almacen Destino" + Chr(13), vbInformation, Me.Caption: cmbAlmDestino.SetFocus: Exit Function
        If cmbAlmOrigen.ListIndex < 0 Then MsgBox "Por Favor Elija el Almacen Origen" + Chr(13), vbInformation, Me.Caption: cmbAlmOrigen.SetFocus: Exit Function
        If cmbAlmDestino.Text = cmbAlmOrigen.Text Then MsgBox "El Almacén Origen y Destino deben ser DIFERENTES" + Chr(13), vbInformation, Me.Caption: cmbAlmOrigen.SetFocus: Exit Function
        
        If Not IsDate(Me.txtFechaInicio.Text) Then MsgBox "Fecha Inicio Incorrecta", vbInformation, Me.Caption: txtFechaInicio.SetFocus: Exit Function
        If Not IsDate(Me.txtFechaFinal.Text) Then MsgBox "Fecha Final Incorrecta", vbInformation, Me.Caption: txtFechaFinal.SetFocus: Exit Function
        
        'If Me.txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY Then MsgBox "Por favor Ingrese la Fecha Inicio", vbInformation, Me.Caption: txtFechaInicio.SetFocus: Exit Sub
        If Me.txtFechaInicio.Text = "__/__/____" Then MsgBox "Por favor Ingrese la Fecha Inicio", vbInformation, Me.Caption: txtFechaInicio.SetFocus: Exit Function
        'If Me.txtFechaFinal.Text = sighentidades.FECHA_VACIA_DMY Then MsgBox "Por favor ingrese la Fecha hasta", vbInformation, Me.Caption: txtFechaFinal.SetFocus: Exit Sub
        If Me.txtFechaFinal.Text = "__/__/____" Then MsgBox "Por favor ingrese la Fecha hasta", vbInformation, Me.Caption: txtFechaFinal.SetFocus: Exit Function
        If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFinal.Text) Then MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "": txtFechaFinal.SetFocus: Exit Function
    ValidarDatosObligatorios = True
End Function

Sub IdRepProveedor()
IdProveedor = "0"
Set oRsProveedor = mo_ReglasFarmacia.FarmProveedorSeleccionarSegunFiltro("Ruc='" & Me.txtRuc.Text & "'")
       Do While Not oRsProveedor.EOF
            IdProveedor = oRsProveedor!IdProveedor
        oRsProveedor.MoveNext
    Loop
       oRsProveedor.Close: Set oRsProveedor = Nothing
    'MsgBox IdProveedor
End Sub
Private Sub btnImprimirConsolidado_Click()
If ValidarDatosObligatorios = True Then
    'id proveedor
    IdRepProveedor
    
    Dim oRptClase As New rCrytalInventario ' rCrystal
    oRptClase.TextoDelFiltro = "Reporte Consolidado"
    oRptClase.IdAlmacenDestino = idAlmacenD
    oRptClase.IdAlmacenOrigen = idAlmacenO
    oRptClase.FechaInicio = txtFechaInicio.Text
    oRptClase.FechaFin = txtFechaFinal.Text
    oRptClase.IdProveedores = IdProveedor
    'oRptClase.HoraInicio = txtFechaInicio.Text
    'oRptClase.HoraFin = Trim(cmbTipoDocum.Text) & " - " & txtRuc.Text
    oRptClase.AlmacenO = cmbAlmOrigen.Text 'descripcion origen
    oRptClase.Almacen = cmbAlmDestino 'descripcion destino
    oRptClase.TipoReporte = Me.Name
    oRptClase.Rreportes = "ProductosIngConso"
    oRptClase.Show vbModal
    Set oRptClase = Nothing
 '  ImprimeDocumento
End If
End Sub



Private Sub btnImprimirDetallado_Click()
If ValidarDatosObligatorios = True Then
   'id proveedor
    IdRepProveedor
    Dim oRptClase As New rCrytalInventario ' rCrystal
    
    oRptClase.TextoDelFiltro = "Reporte Consolidado"
    oRptClase.IdAlmacenDestino = idAlmacenD
    oRptClase.IdAlmacenOrigen = idAlmacenO
    oRptClase.FechaInicio = txtFechaInicio.Text
    oRptClase.FechaFin = txtFechaFinal.Text
    oRptClase.IdProveedores = IdProveedor
    oRptClase.AlmacenO = cmbAlmOrigen.Text 'descripcion origen
    oRptClase.Almacen = cmbAlmDestino 'descripcion destino
    'oRptClase.HoraInicio = txtFechaInicio.Text
    'oRptClase.HoraFin = Trim(cmbTipoDocum.Text) & " - " & txtRuc.Text
    oRptClase.TipoReporte = Me.Name
    oRptClase.Rreportes = "ProductosIngDet"
    oRptClase.Show vbModal
    Set oRptClase = Nothing
 '  ImprimeDocumento
End If
End Sub

Private Sub cmbAlmDestino_Click()
       Set oRsAlmacenDestino = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("Descripcion='" & cmbAlmDestino.Text & "'")
       Do While Not oRsAlmacenDestino.EOF
            idAlmacenD = oRsAlmacenDestino!IdAlmacen
        oRsAlmacenDestino.MoveNext
       Loop
       oRsAlmacenDestino.Close: Set oRsAlmacenDestino = Nothing
       'MsgBox idAlmacenD
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
End Sub

Private Sub cmbAlmDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmDestino
End Sub

Private Sub cmbAlmOrigen_Click()
       Set oRsAlmacenOrigen = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("Descripcion='" & cmbAlmOrigen.Text & "'")
       
       Do While Not oRsAlmacenOrigen.EOF
            idAlmacenO = oRsAlmacenOrigen!IdAlmacen
        oRsAlmacenOrigen.MoveNext
       Loop
       oRsAlmacenOrigen.Close: Set oRsAlmacenOrigen = Nothing
       
      ' MsgBox idAlmacenO
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
End Sub


Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen
End Sub









Private Sub CmbFiltro_Click()
If CmbFiltro.ListIndex = 0 Then
            rsTmp.Sort = "Codigo"
            TxtBusca.Text = ""
        Else
            rsTmp.Sort = "Nombre"
            TxtBusca.Text = ""
End If
grdProductos.Refresh
End Sub






Private Sub cmbRealizarBusqueda_Click()


    If cmbFarmacia.Text = "" Then
       MsgBox "Elija la Farmacia", vbInformation, ""
       Exit Sub
    End If
    If CDate(txtFinicio.Text) > CDate(txtFFinal.Text) Then
       MsgBox "La Fecha Final no puede ser Menor a la Fecha Inicial", vbInformation, ""
       Exit Sub
    End If
    
    ParteDiario CDate(txtFinicio.Text & " " & txtHoraInicio1.Text & ":01"), _
                CDate(txtFFinal.Text & " " & txtHoraFinal1.Text & ":59"), "/" & mo_cmbFarmacia.BoundText & "/", False, True, False
    If mrs_Tmp.RecordCount > 0 Then
       mrs_Tmp.MoveFirst
    End If
    Set grdVentas.DataSource = mrs_Tmp
    mo_Apariencia.ConfigurarFilasBiColores Me.grdVentas, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub cmdBuscaCajero_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
    Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    txtCajero.Tag = ""
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_ReglasAdmision.EmpleadosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtCajero.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            txtCajero.Tag = oDOEmpleado.IdEmpleado
        End If
    End If
    Set oBusqueda = Nothing
    Set oDOEmpleado = Nothing

End Sub

Private Sub cmdBuscaPaciente_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    Dim mo_AdminAdmision As New ReglasAdmision
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    txtPaciente.Tag = ""
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
           txtPaciente.Text = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno & " " & oDOPaciente.PrimerNombre & "(" & Trim(Str(oDOPaciente.NroHistoriaClinica)) & ")"
           txtPaciente.Tag = oDOPaciente.IdPaciente
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOPaciente = Nothing
    Set oBusqueda = Nothing
    Set mo_AdminAdmision = Nothing
End Sub



Private Sub cmdImprime_Click()
    Dim lcFiltro As String
    lcFiltro = "Farmacia: " & cmbFarmacia.Text & _
                            "   Desde: " & txtFinicio.Text & " " & Me.txtHoraInicio1.Text & " al " & txtFFinal.Text & " " & txtHoraFinal1.Text & _
                            " (T.Producto: " & cmbTproducto.Text & _
                            ") " & IIf(txtPaciente.Text <> "", " (Paciente: " & txtPaciente.Text & ") ", "") & _
                            IIf(txtCajero.Text <> "", " (Cajero: " & txtCajero.Text & ")", "")
                            
    If chkExcel.Value = 1 Then
           mo_AdminReportes.ExportarRecordSetAexcel mrs_Tmp, tabReportes.Caption, lcFiltro, "", Me.hwnd, False, _
                                                    True
        
    Else
        mrs_Tmp.Sort = "nombre"
        Set RpConsumoItems.DataSource = mrs_Tmp
        RpConsumoItems.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
        RpConsumoItems.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
        RpConsumoItems.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
        RpConsumoItems.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
        RpConsumoItems.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
        RpConsumoItems.Sections("cabecera").Controls("lblTitulo").Caption = tabReportes.Caption
        RpConsumoItems.Sections("cabecera").Controls("lblSubTitulo").Caption = lcFiltro
        Set RpConsumoItems.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
        RpConsumoItems.Sections("piePag").Controls("lblCantidad").Caption = "N° " & Trim(Str(mrs_Tmp.RecordCount))
        RpConsumoItems.Sections("piePag").Controls("lblImporteT").Caption = Trim(Str(lnVentaTotal))
        RpConsumoItems.Orientation = rptOrientLandscape
        RpConsumoItems.Show 1
   End If
End Sub

Private Sub cmdSalir_Click()
Me.Visible = False
End Sub



Private Sub Form_Initialize()
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbAlmacenDestino.MiComboBox = cmbAlmDestino
    Set mo_cmbFarmacia.MiComboBox = cmbFarmacia
End Sub

Private Sub Form_Load()
    If ml_NroReporte = 1 Then
       tabReportes.TabVisible(1) = False
       tabReportes.TabVisible(2) = False
    ElseIf ml_NroReporte = 2 Then
       tabReportes.TabVisible(0) = False
       tabReportes.TabVisible(2) = False
       txtHoraInicio1.Text = "00:00"
       txtHoraFinal1.Text = "23:59"
    End If
    
    mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, False
    '
    mo_Formulario.HabilitarDeshabilitar Me.txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCajero, False
    

    CargarComboBoxes

    txtFechaInicio.Text = DateAdd("yyyy", -5, Date)
    txtFechaFinal.Text = Date
    
    txtFinicio.Text = Date
    txtFFinal.Text = Date
    cmbTproducto.ListIndex = 0
    
End Sub



Sub CargarComboBoxes()
    Dim rsIdAlmacen As Recordset
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("")
    '
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    mo_cmbAlmacenDestino.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenDestino.ListField = "Descripcion"
    Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' or idTipoLocales='A'")
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If cmbAlmDestino.ListCount = 1 Then
       cmbAlmDestino.ListIndex = 0
    End If
    
    mo_cmbFarmacia.BoundColumn = "idAlmacen"
    mo_cmbFarmacia.ListField = "Descripcion"
    Set mo_cmbFarmacia.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F'")
    
    cmbTproducto.ListIndex = 1
    
    
End Sub


Private Sub btnCancelar_Click()
     Me.Visible = False
     LimpiarVariablesDeMemoria
End Sub


Private Sub Form_Terminate()
LimpiarVariablesDeMemoria
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub







Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
 Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    'grdProductos.Bands(0).Columns("idProveedor").Hidden = True
    '
    grdProductos.Bands(0).Columns("fechaCreacion").Header.Caption = "Fecha Creación"
    grdProductos.Bands(0).Columns("fechaCreacion").Width = 1200
    '
    grdProductos.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdProductos.Bands(0).Columns("Codigo").Width = 1000
    
    grdProductos.Bands(0).Columns("Nombre").Header.Caption = "Nombres"
    grdProductos.Bands(0).Columns("Nombre").Width = 7000

    grdProductos.Bands(0).Columns("PrecioUnitario").Header.Caption = "Pre. Oper."
    grdProductos.Bands(0).Columns("PrecioUnitario").Width = 950

    grdProductos.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    grdProductos.Bands(0).Columns("Cantidad").Width = 900

    grdProductos.Bands(0).Columns("Concepto").Header.Caption = "Concepto"
    grdProductos.Bands(0).Columns("Concepto").Width = 2700
    
    grdProductos.Bands(0).Columns("Concepto").Hidden = True
    grdProductos.Bands(0).Columns("cantidad").Hidden = True
    grdProductos.Bands(0).Columns("monto").Hidden = True
    grdProductos.Bands(0).Columns("lote").Hidden = True
    grdProductos.Bands(0).Columns("FechaVencimiento").Hidden = True
    grdProductos.Bands(0).Columns("IdAlmacenDestino").Hidden = True
    grdProductos.Bands(0).Columns("idAlmacenOrigen").Hidden = True
    grdProductos.Bands(0).Columns("idProveedor").Hidden = True
End Sub



Private Sub grdVentas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdVentas.Bands(0).Columns("codigo").Header.Caption = "Código"
    grdVentas.Bands(0).Columns("codigo").Width = 700
    grdVentas.Bands(0).Columns("Nombre").Header.Caption = "Producto"
    grdVentas.Bands(0).Columns("Nombre").Width = 5300
    grdVentas.Bands(0).Columns("precio").Hidden = True
    grdVentas.Bands(0).Columns("saldoI").Hidden = True
    grdVentas.Bands(0).Columns("Ingresos").Hidden = True
    grdVentas.Bands(0).Columns("DevolucionesP").Hidden = True
    grdVentas.Bands(0).Columns("totIngresos").Hidden = True
    grdVentas.Bands(0).Columns("TotSalidas").Header.Caption = "Cantidad"
    grdVentas.Bands(0).Columns("TotSalidas").Width = 900
    grdVentas.Bands(0).Columns("ventas").Header.Caption = "CONT"
    grdVentas.Bands(0).Columns("Ventas").Width = 700
    grdVentas.Bands(0).Columns("creditoH").Header.Caption = "HOSP"
    grdVentas.Bands(0).Columns("creditoH").Width = 700
    grdVentas.Bands(0).Columns("sis").Header.Caption = "SIS"
    grdVentas.Bands(0).Columns("sis").Width = 700
    grdVentas.Bands(0).Columns("soat").Header.Caption = "SOAT"
    grdVentas.Bands(0).Columns("soat").Width = 700
    grdVentas.Bands(0).Columns("exonerac").Header.Caption = "EXON"
    grdVentas.Bands(0).Columns("exonerac").Width = 700
    grdVentas.Bands(0).Columns("convenio").Header.Caption = "CONV"
    grdVentas.Bands(0).Columns("convenio").Width = 700
    grdVentas.Bands(0).Columns("creditoP").Header.Caption = "CRPE"
    grdVentas.Bands(0).Columns("creditoP").Width = 700
    grdVentas.Bands(0).Columns("ImporteSal").Header.Caption = "Total"
    grdVentas.Bands(0).Columns("ImporteSal").Width = 1200
    grdVentas.Bands(0).Columns("defensaN").Hidden = True
    grdVentas.Bands(0).Columns("OsDevol").Hidden = True
    grdVentas.Bands(0).Columns("OsVencim").Hidden = True
    grdVentas.Bands(0).Columns("OsMerma").Hidden = True
    grdVentas.Bands(0).Columns("intervencionS").Hidden = True
    grdVentas.Bands(0).Columns("otrasS").Hidden = True
    grdVentas.Bands(0).Columns("FechaVencimiento").Hidden = True
    grdVentas.Bands(0).Columns("tipo").Hidden = True
    grdVentas.Bands(0).Columns("saldoF").Hidden = True

End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If TxtBusca.Text <> "" Then
            TxtBusca.Text = Trim(TxtBusca.Text)
            rsTmp.MoveFirst
            If CmbFiltro.ListIndex = 0 Then
               rsTmp.Find "Codigo='" & TxtBusca.Text & "'"
            Else
               Do While Not rsTmp.EOF
                  If Left(rsTmp!Nombre, Len(TxtBusca.Text)) = UCase(TxtBusca.Text) Then
                  Exit Do
                  End If
                  rsTmp.MoveNext
               Loop
            End If
            grdProductos.Refresh
      End If
   End If
End Sub



Private Sub txtRuc_GotFocus()
If txtRuc.Text = "" Then
    txtProveedor.Text = ""
End If
End Sub

Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRuc
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
    'If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
     '      KeyAscii = 0
    'End If
    If KeyAscii = 13 Then
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set rsTmp = oConexion.Execute("select*from proveedores where ruc = '" & Me.txtRuc.Text & "'")
        txtProveedor.Text = ""
        Do While Not rsTmp.EOF
            txtProveedor.Text = rsTmp!razonSocial
        rsTmp.MoveNext
        Loop
        rsTmp.Close
        Set rsTmp = Nothing
        Set oConexion = Nothing

    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbConceptos = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbAlmacenDestino = Nothing
    Set mo_cmbTipoDocum = Nothing
    Set mo_ReglasFarmacia = Nothing
    'Set oRsConceptos = Nothing
    Set oRsAlmacenOrigen = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mo_farmMovimiento = Nothing
    Set mo_farmMovimientoNotaIngreso = Nothing
    Set oDoProveedores = Nothing
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            btnBuscar_Click
        Case vbKeyEscape
            btnCancelar_Click
       End Select
End Sub



Sub ParteDiario(mda_FechaInicio As Date, mda_FechaFin As Date, lc_AlmacenesParaICI As String, mb_ConsiderarRecalculo As Boolean, _
                mb_ConsideraOSH As Boolean, mb_ConsiderarSinMovimientos As Boolean)
        Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
        Dim oBuscaMovimientos As New farmMovimientoDetalle
        Dim mo_ReglasFacturacion As New ReglasFacturacion
        Dim mo_ReglasCaja As New ReglasCaja
        Dim rsReporte As New ADODB.Recordset
        Dim rsTmp11 As New Recordset
        Dim rsTmp12 As New Recordset
        Dim rsTmp13 As New Recordset
        Dim rsTmp14 As New Recordset
        Dim rsTmp15 As New Recordset
        Dim rsErrores As New Recordset
        Dim oRsTmp984 As New Recordset
        Dim ldFechaInicioMovim As Date, ldFechaHistoricoXmes As Date, lcUltDiaMes As String
        Dim lnIdProducto As Long, lnSaldoInicial As Long, lnIdAlmacen As Long, lbContinua As Boolean
        Dim lnRegistro As Long, lbPrimeraVez As Boolean, ldFechaVencimiento As Date
        Dim lnRegTope As Long, lnPrecio As Double, lbAgregaPorPaciente As Boolean
        Dim lnidTipoConceptoFarmacia As Long
        Dim lcSql As String, lcTexto2 As String
        Dim lnIngresos As Long, LnDevolucionesP As Long, TotIngresos As Long
        Dim LnVentas As Long, lnSis As Long, lnSoat As Long, LnConvenio As Long, lnCreditoH As Long, lnDefensaN As Long
        Dim LnOsDevol As Long, LnOsVencim As Long, LnOsMerma As Long, LnExonerac As Long, LnIntervencionS As Long
        Dim LnOtrasS As Long, TotSalidas As Long, LnCreditoP As Long, lnFor As Integer
        Dim lnTotalRegistros As Long, lnIdAlmacenRep As Long, lcTexto1 As String
        Dim lcCodigo As String, lcNombre As String, lnTotSalidas As Long
        Dim LnVentas1 As Long, lnCreditoH1 As Long, lnSis1 As Long, lnSoat1 As Long, LnExonerac1 As Long
        Dim LnConvenio1 As Long, LnCreditoP1 As Long
        '
        On Error GoTo ErrParteDia
        '
        oConexion.CursorLocation = adUseClient
        
        oConexion.Open SIGHEntidades.CadenaConexion
        Set oFarmMovimientoDetalle.Conexion = oConexion
        
        'Proceso
        lcUltDiaMes = Trim(Str(SIGHEntidades.DevuelveUltimoDiaDelMes(Month(mda_FechaInicio), Year(mda_FechaInicio))))
        ldFechaHistoricoXmes = CDate("01" & Format(mda_FechaInicio, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263) & ":59") - 1
        ldFechaHistoricoXmes = SIGHEntidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaHistoricoXmes)
        ldFechaInicioMovim = DateAdd("n", 1, ldFechaHistoricoXmes)
        'Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(CDate("01/01/1990"), mda_FechaFin, 0, "")
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(ldFechaInicioMovim, mda_FechaFin, Val(mo_cmbFarmacia.BoundText), "")
        If Me.cmbTproducto.ListIndex > 0 Then
           rsReporte.Filter = "tipoProducto=" & IIf(Me.cmbTproducto.ListIndex = 1, "0", "1")
        End If
        lnTotalRegistros = rsReporte.RecordCount
        lnVentaTotal = 0
        
        If lnTotalRegistros > 0 Then
            GenerarRecordsetTemporalICI
            
            '
            lnRegistro = 1
            lnRegTope = 28320
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF

            
                lnIdProducto = rsReporte.Fields!idProducto
                lcCodigo = rsReporte.Fields!codigo
                lcNombre = rsReporte.Fields!Nombre
                '*******Saldo Inicial****************************************
                lnSaldoInicial = 0
                'saldos-barre historico mensual
                For lnFor = 1 To Len(lc_AlmacenesParaICI)
                    If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                       lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                       lnFor = Len(lc_AlmacenesParaICI)
                    Else
                        lcTexto1 = ""
                        Do While True
                           If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                              Exit Do
                           Else
                              lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                              lnFor = lnFor + 1
                           End If
                        Loop
                        lnIdAlmacenRep = Val(lcTexto1)
                    End If
                    If lnIdAlmacenRep > 1 Then
                        Set rsErrores = mo_ReglasFarmacia.FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes(lnIdProducto, lnIdAlmacenRep, ldFechaHistoricoXmes)
                        Do While Not rsErrores.EOF
                            lnSaldoInicial = lnSaldoInicial + rsErrores.Fields!saldo
                            rsErrores.MoveNext
                        Loop
                        rsErrores.Close
                    End If
                Next
                'saldos-barre movimiento
                Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaInicio
                   If rsReporte.Fields!MovTipo = "S" Then
                      If InStr(lc_AlmacenesParaICI, "/" & Trim(Str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                        lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                      End If
                   Else
                      If InStr(lc_AlmacenesParaICI, "/" & Trim(Str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                         lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                      End If
                   End If
                   rsReporte.MoveNext
                   lnRegistro = lnRegistro + 1

                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
                '****** Movimientos en el Rango de Fechas***********************************
                lnIngresos = 0: LnDevolucionesP = 0: TotIngresos = 0
                LnVentas = 0: lnSis = 0: lnSoat = 0: LnConvenio = 0: lnCreditoH = 0: lnDefensaN = 0
                LnOsDevol = 0: LnOsVencim = 0: LnOsMerma = 0: LnExonerac = 0: LnIntervencionS = 0
                LnOtrasS = 0: TotSalidas = 0: LnCreditoP = 0
                LnVentas1 = 0: lnCreditoH1 = 0: lnSis1 = 0: lnSoat1 = 0: LnExonerac1 = 0: LnConvenio1 = 0: LnCreditoP1 = 0
                If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                       lbPrimeraVez = True
                       If rsReporte.Fields!MovTipo = "S" Then
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(Str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                                If mb_ConsiderarRecalculo = True Then
                                    '********* con recalculo
                                    lcTexto1 = rsReporte.Fields!MovTipo
                                    lcTexto2 = rsReporte.Fields!movNumero
                                    'Busca si tiene Pagos
                                    'debb 02/02/2011
                                    Set rsTmp11 = mo_ReglasFarmacia.FacturacionBienesPagoSeleccionarPorMovNumeroProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                    If rsTmp11.RecordCount > 0 Then
                                       If rsTmp11.Fields!IdComprobantePago > 0 And rsTmp11.Fields!idEstadoFacturacion = 4 Then
                                          lbPrimeraVez = False
                                          rsTmp11.MoveFirst
                                          Do While Not rsTmp11.EOF
                                                lbAgregaPorPaciente = False
                                                If Val(txtPaciente.Tag) > 0 Or Val(txtCajero.Tag) > 0 Then
                                                     If oRsTmp984.State = 1 Then oRsTmp984.Close
                                                     Set oRsTmp984 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(rsTmp11!IdComprobantePago, oConexion)
                                                     If oRsTmp984.RecordCount > 0 Then
                                                        If Val(txtPaciente.Tag) > 0 Then
                                                            If Not IsNull(oRsTmp984!IdPaciente) Then
                                                               If oRsTmp984!IdPaciente = Val(txtPaciente.Tag) Then
                                                                  lbAgregaPorPaciente = True
                                                               End If
                                                            End If
                                                        End If
                                                        If Val(txtCajero.Tag) > 0 Then
                                                            If Not IsNull(oRsTmp984!idCajero) Then
                                                               If oRsTmp984!idCajero = Val(txtCajero.Tag) Then
                                                                  lbAgregaPorPaciente = True
                                                               End If
                                                            End If
                                                        End If
                                                     End If
                                                End If
                                                Select Case rsReporte.Fields!idTipoConcepto
                                                Case 10       'Ventas
                                                     LnVentas = LnVentas + rsTmp11.Fields!CantidadPagar
                                                     If lbAgregaPorPaciente = True Then
                                                           LnVentas1 = LnVentas1 + rsTmp11.Fields!CantidadPagar
                                                     End If
                                                Case 13       'Sis
                                                     lnSis = lnSis + rsTmp11.Fields!CantidadPagar
                                                     If lbAgregaPorPaciente = True Then
                                                           lnSis1 = lnSis1 + rsTmp11.Fields!CantidadPagar
                                                     End If
                                                Case 14       'Soat
                                                     lnSoat = lnSoat + rsTmp11.Fields!CantidadPagar
                                                     If lbAgregaPorPaciente = True Then
                                                           lnSoat1 = lnSoat1 + rsTmp11.Fields!CantidadPagar
                                                     End If
                                                Case 23      'Convenios
                                                     LnConvenio = LnConvenio + rsTmp11.Fields!CantidadPagar
                                                     If lbAgregaPorPaciente = True Then
                                                           LnConvenio1 = LnConvenio1 + rsTmp11.Fields!CantidadPagar
                                                     End If
                                                Case 26      'Credito Personal
                                                     LnCreditoP = LnCreditoP + rsTmp11.Fields!CantidadPagar
                                                     If lbAgregaPorPaciente = True Then
                                                           LnCreditoP1 = LnCreditoP1 + rsTmp11.Fields!CantidadPagar
                                                     End If
                                                Case 17       'Credito Hospitalario
                                                     lnCreditoH = lnCreditoH + rsTmp11.Fields!CantidadPagar
                                                     If lbAgregaPorPaciente = True Then
                                                           lnCreditoH1 = lnCreditoH1 + rsTmp11.Fields!CantidadPagar
                                                     End If
                                               ' Case 22       'Defensa nacional
                                               '      lnDefensaN = lnDefensaN + rsTmp11.Fields!cantidadPagar
                                               ' Case 7       'Otras salidas Devolucion
                                               '      LnOsDevol = LnOsDevol + rsTmp11.Fields!cantidadPagar
                                               ' Case 5       'Otras salidas Vencimiento
                                               '      LnOsVencim = LnOsVencim + rsTmp11.Fields!cantidadPagar
                                               ' Case 6       'Otras salidas Merma
                                               '      LnOsMerma = LnOsMerma + rsTmp11.Fields!cantidadPagar
                                               ' Case 15       'Exoneraciones
                                               '      LnExonerac = LnExonerac + rsTmp11.Fields!cantidadPagar
                                                Case 16       'Intervencion Sanitaria
                                                     LnIntervencionS = LnIntervencionS + rsTmp11.Fields!CantidadPagar
                                                Case Else
                                                     LnOtrasS = LnOtrasS + rsTmp11.Fields!CantidadPagar
                                                End Select
                                                TotSalidas = TotSalidas + rsTmp11.Fields!CantidadPagar
                                                rsTmp11.MoveNext
                                          Loop
                                       End If
                                    End If
                                    rsTmp11.Close
                                    'Busca si tiene algun seguro o exoneracion (Plan)
                                    'debb 02/02/2011
                                    Set rsTmp12 = mo_ReglasFarmacia.FacturacionBienesFinancSeleccionarPorProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                    If rsTmp12.RecordCount > 0 Then
                                       lbPrimeraVez = False
                                        rsTmp12.MoveFirst
                                        Do While Not rsTmp12.EOF
                                            lbAgregaPorPaciente = False
                                            If Val(txtPaciente.Tag) > 0 Or Val(txtCajero.Tag) > 0 Then
                                                 If Val(txtCajero.Tag) > 0 Then
                                                        lbAgregaPorPaciente = False
                                                 Else
                                                        If oRsTmp984.State = 1 Then oRsTmp984.Close
                                                        Set oRsTmp984 = mo_ReglasFacturacion.FarmMovimientoVentasDetalleSeleccionarPorMovNumero(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                                                        oRsTmp984.Filter = "idProducto=" & rsReporte!idProducto
                                                        If oRsTmp984.RecordCount > 0 Then
                                                           If Val(txtPaciente.Tag) > 0 Then
                                                               If Not IsNull(oRsTmp984!IdPaciente) Then
                                                                  If oRsTmp984!IdPaciente = Val(txtPaciente.Tag) Then
                                                                     lbAgregaPorPaciente = True
                                                                  End If
                                                               End If
                                                           End If
                                                        End If
                                                 End If
                                            End If
                                        
                                            Select Case mo_ReglasFacturacion.FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia(oConexion, rsTmp12.Fields!idFuenteFinanciamiento)
                                            Case 10       'Ventas
                                                 LnVentas = LnVentas + rsTmp12.Fields!CantidadFinanciada
                                                 If lbAgregaPorPaciente = True Then
                                                    LnVentas1 = LnVentas1 + rsTmp12.Fields!CantidadFinanciada
                                                 End If
                                            Case 13       'Sis
                                                 lnSis = lnSis + rsTmp12.Fields!CantidadFinanciada
                                                 If lbAgregaPorPaciente = True Then
                                                    lnSis1 = lnSis1 + rsTmp12.Fields!CantidadFinanciada
                                                 End If
                                            Case 14      'Soat
                                                 lnSoat = lnSoat + rsTmp12.Fields!CantidadFinanciada
                                                 If lbAgregaPorPaciente = True Then
                                                    lnSoat1 = lnSoat1 + rsTmp12.Fields!CantidadFinanciada
                                                 End If
                                            Case 16       'Intervencion Sanitaria
                                                 LnIntervencionS = LnIntervencionS + rsTmp12.Fields!CantidadFinanciada
                                            Case 17       'Credito Hospitalario
                                                 lnCreditoH = lnCreditoH + rsTmp12.Fields!CantidadFinanciada
                                                 If lbAgregaPorPaciente = True Then
                                                    lnCreditoH1 = lnCreditoH1 + rsTmp12.Fields!CantidadFinanciada
                                                 End If
                                            Case 23       'Convenios
                                                 LnConvenio = LnConvenio + rsTmp12.Fields!CantidadFinanciada
                                                 If lbAgregaPorPaciente = True Then
                                                    LnConvenio1 = LnConvenio1 + rsTmp12.Fields!CantidadFinanciada
                                                 End If
                                            Case 26      'Credito Personal
                                                 LnCreditoP = LnCreditoP + rsTmp11.Fields!CantidadFinanciada
                                                 If lbAgregaPorPaciente = True Then
                                                    LnCreditoP1 = LnCreditoP1 + rsTmp11.Fields!CantidadFinanciada
                                                 End If
'                                            Case 0      'Exoneraciones
'                                                 If rsTmp12.Fields!cantidadFinanciada > 0 Then
'                                                    LnExonerac = LnExonerac + rsTmp12.Fields!cantidadFinanciada
'                                                 Else
'                                                    '**no se sabe la CANTIDAD EXONERADA solo el IMPORTE EXONERADO
'                                                    LnExonerac = LnExonerac + rsTmp12.Fields!cantidadFinanciada
'                                                    LnExonerac = LnExonerac - rsTmp12.Fields!cantidadFinanciada
'                                                    TotSalidas = TotSalidas - rsTmp12.Fields!cantidadFinanciada
'                                                 End If
                                            Case Else
                                                 LnOtrasS = LnOtrasS + rsTmp12.Fields!CantidadFinanciada
                                            End Select
                                            TotSalidas = TotSalidas + rsTmp12.Fields!CantidadFinanciada
                                            rsTmp12.MoveNext
                                        Loop
                                    End If
                                    rsTmp12.Close
                                    '
                                    If lbPrimeraVez = False Then
                                        Do While Not rsReporte.EOF And lcTexto1 = rsReporte.Fields!MovTipo And lcTexto2 = rsReporte.Fields!movNumero And lnIdProducto = rsReporte.Fields!idProducto
                                           rsReporte.MoveNext
                                           lnRegistro = lnRegistro + 1

                                           If rsReporte.EOF Then
                                              Exit Do
                                           End If
                                        Loop
                                    End If
                                End If
                                If lbPrimeraVez = True Then
                                    '******** sin recalculo
                                    lbAgregaPorPaciente = False
                                    If Val(txtCajero.Tag) > 0 And (rsReporte.Fields!idTipoConcepto = 10 Or rsReporte.Fields!idTipoConcepto = 17) Then
                                            If oRsTmp984.State = 1 Then oRsTmp984.Close
                                            Set oRsTmp984 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorMovnumero(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                                            If oRsTmp984.RecordCount > 0 Then
                                                   If Not IsNull(oRsTmp984!idCajero) Then
                                                      If oRsTmp984!idCajero = Val(txtCajero.Tag) Then
                                                         lbAgregaPorPaciente = True
                                                      End If
                                                   End If
                                            End If
                                    End If
                                    If Val(txtPaciente.Tag) > 0 Then
                                       If Val(txtCajero.Tag) = 0 Or (Val(txtCajero.Tag) > 0 And lbAgregaPorPaciente = True) Then
                                            lbAgregaPorPaciente = False
                                            If oRsTmp984.State = 1 Then oRsTmp984.Close
                                            Set oRsTmp984 = mo_ReglasFacturacion.FarmMovimientoVentasDetalleSeleccionarPorMovNumero(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                                            oRsTmp984.Filter = "idProducto=" & rsReporte!idProducto
                                            If oRsTmp984.RecordCount > 0 Then
                                                If Not IsNull(oRsTmp984!IdPaciente) Then
                                                   If oRsTmp984!IdPaciente = Val(txtPaciente.Tag) Then
                                                      lbAgregaPorPaciente = True
                                                   End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                    Select Case rsReporte.Fields!idTipoConcepto
                                    Case 10       'Ventas
                                         LnVentas = LnVentas + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                            LnVentas1 = LnVentas1 + rsReporte.Fields!Cantidad
                                         End If
                                    Case 13       'Sis
                                         lnSis = lnSis + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                            lnSis1 = lnSis1 + rsReporte.Fields!Cantidad
                                         End If
                                    Case 14       'Soat
                                         lnSoat = lnSoat + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                            lnSoat1 = lnSoat1 + rsReporte.Fields!Cantidad
                                         End If
                                    Case 23      'Convenios
                                         LnConvenio = LnConvenio + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                            LnConvenio1 = LnConvenio1 + rsReporte.Fields!Cantidad
                                         End If
                                    Case 26      'Credito Personal
                                          LnCreditoP = LnCreditoP + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                            LnCreditoP1 = LnCreditoP1 + rsReporte.Fields!Cantidad
                                         End If
                                    Case 17       'Credito Hospitalario
                                         lnCreditoH = lnCreditoH + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                            lnCreditoH1 = lnCreditoH1 + rsReporte.Fields!Cantidad
                                         End If
'                                    Case 22       'Defensa nacional
'                                         lnDefensaN = lnDefensaN + rsReporte.Fields!cantidad
'                                    Case 7       'Otras salidas Devolucion
'                                         LnOsDevol = LnOsDevol + rsReporte.Fields!cantidad
'                                    Case 5       'Otras salidas Vencimiento
'                                         LnOsVencim = LnOsVencim + rsReporte.Fields!cantidad
'                                    Case 6       'Otras salidas Merma
'                                         LnOsMerma = LnOsMerma + rsReporte.Fields!cantidad
'                                    Case 15       'Exoneraciones
'                                         LnExonerac = LnExonerac + rsReporte.Fields!cantidad
                                    Case 16       'Intervencion Sanitaria
                                         LnIntervencionS = LnIntervencionS + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                         End If
                                    Case Else
                                         LnOtrasS = LnOtrasS + rsReporte.Fields!Cantidad
                                         If lbAgregaPorPaciente = True Then
                                         End If
                                    End Select
                                    If rsReporte.Fields!IdAlmacenDestino = 10 And rsReporte.Fields!idTipoConcepto = 4 And mb_ConsideraOSH = False Then
                                       'destino=otros servicios hospital, tipoConcepto=distribucion
                                    Else
                                        TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                    End If
                                End If
                          End If
                       Else
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(Str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                                Select Case rsReporte.Fields!idTipoConcepto
                                Case 19        'Inventario
                                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                                Case 21        'Devolucion de Pacientes
                                     LnDevolucionesP = LnDevolucionesP + rsReporte.Fields!Cantidad
                                     TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                Case Else        'Ingresos
                                     lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                                     TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                End Select
                                
                           End If
                       End If
                       If lbPrimeraVez = True Then
                           rsReporte.MoveNext
                           lnRegistro = lnRegistro + 1

                       End If
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                       rsReporte.MoveNext
                       lnRegistro = lnRegistro + 1
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                '
                lnPrecio = 0
                Set rsTmp13 = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodigo, 1)
                If rsTmp13.RecordCount > 0 Then
                   lnPrecio = rsTmp13.Fields!PrecioUnitario
                End If
                rsTmp13.Close
                '
                If InStr(lc_AlmacenesParaICI, "/") > 0 Then
                   lnIdAlmacen = Val(Left(lc_AlmacenesParaICI, InStr(lc_AlmacenesParaICI, "/") - 1)) 'toma la primera farmacia para las FECHAS DE VENCIMIENTO
                Else
                   lnIdAlmacen = Val(lc_AlmacenesParaICI)
                End If
                '
                lbContinua = True
                If mb_ConsiderarSinMovimientos = False Then
                    If txtPaciente.Text <> "" Or txtCajero.Text <> "" Then
                       lnTotSalidas = LnVentas1 + lnCreditoH1 + lnSis1 + lnSoat1 + LnExonerac1 + LnConvenio1 + LnCreditoP1
                    Else
                       lnTotSalidas = LnVentas + lnCreditoH + lnSis + lnSoat + LnExonerac + LnConvenio + LnCreditoP
                    End If
                    If lnTotSalidas = 0 Then
                      lbContinua = False
                    End If
                End If
                '
                If lbContinua Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!codigo = lcCodigo
                    mrs_Tmp.Fields!Nombre = lcNombre
                    mrs_Tmp.Fields!Precio = lnPrecio
                    mrs_Tmp.Fields!saldoI = lnSaldoInicial
                    mrs_Tmp.Fields!ingresos = lnIngresos
                    mrs_Tmp.Fields!DevolucionesP = LnDevolucionesP
                    mrs_Tmp.Fields!TotIngresos = TotIngresos
                    mrs_Tmp.Fields!Ventas = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", LnVentas1, LnVentas)
                    mrs_Tmp.Fields!sis = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", lnSis1, lnSis)
                    mrs_Tmp.Fields!soat = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", lnSoat1, lnSoat)
                    mrs_Tmp.Fields!convenio = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", LnConvenio1, LnConvenio)
                    mrs_Tmp.Fields!creditoH = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", lnCreditoH1, lnCreditoH)
                    mrs_Tmp.Fields!defensaN = lnDefensaN
                    mrs_Tmp.Fields!OsDevol = LnOsDevol
                    mrs_Tmp.Fields!OsVencim = LnOsVencim
                    mrs_Tmp.Fields!OsMerma = LnOsMerma
                    mrs_Tmp.Fields!Exonerac = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", LnExonerac1, LnExonerac)
                    mrs_Tmp.Fields!IntervencionS = LnIntervencionS
                    mrs_Tmp.Fields!otrasS = LnOtrasS
                    mrs_Tmp.Fields!TotSalidas = lnTotSalidas
                    mrs_Tmp.Fields!FechaVencimiento = ldFechaVencimiento
                    mrs_Tmp.Fields!creditop = IIf(txtPaciente.Text <> "" Or txtCajero.Text <> "", LnCreditoP1, LnCreditoP)
                    mrs_Tmp.Fields!ImporteSal = Round(lnTotSalidas * lnPrecio, 2)
                    mrs_Tmp.Fields!saldoF = lnSaldoInicial + TotIngresos - (LnVentas + lnSis + lnSoat + LnConvenio + lnCreditoH + lnDefensaN + LnOsDevol + LnOsVencim + LnOsMerma + LnExonerac + LnIntervencionS + LnOtrasS + LnCreditoP)
                    mrs_Tmp.Update
                    lnVentaTotal = lnVentaTotal + Round(lnTotSalidas * lnPrecio, 2)
                End If
                'Graba Datos en Temporal
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
            mrs_Tmp.Sort = "nombre"
       End If
       oConexion.Close
       rsReporte.Close
        Set rsTmp11 = Nothing
        Set rsTmp12 = Nothing
        Set rsTmp13 = Nothing
        Set rsTmp14 = Nothing
        Set rsTmp15 = Nothing
        Set oFarmMovimientoDetalle = Nothing
        Set oBuscaMovimientos = Nothing
        Set mo_ReglasFacturacion = Nothing
        Set rsReporte = Nothing
        Set rsErrores = Nothing
        Set mo_ReglasCaja = Nothing
        Set oRsTmp984 = Nothing
        Exit Sub
ErrParteDia:
     MsgBox Err.Description
     Exit Sub
     Resume
End Sub

Sub GenerarRecordsetTemporalICI()
    If mrs_Tmp.State = 1 Then Set mrs_Tmp = Nothing
    With mrs_Tmp
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 250, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "saldoI", adInteger, 4, adFldIsNullable
          .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
          .Fields.Append "DevolucionesP", adInteger, 4, adFldIsNullable
          .Fields.Append "TotIngresos", adInteger, 4, adFldIsNullable
          .Fields.Append "TotSalidas", adInteger, 4, adFldIsNullable
          .Fields.Append "ventas", adInteger, 4, adFldIsNullable
          .Fields.Append "creditoH", adInteger, 4, adFldIsNullable
          .Fields.Append "sis", adInteger, 4, adFldIsNullable
          .Fields.Append "soat", adInteger, 4, adFldIsNullable
          .Fields.Append "Exonerac", adInteger, 4, adFldIsNullable
          .Fields.Append "convenio", adInteger, 4, adFldIsNullable
          .Fields.Append "creditoP", adInteger, 4, adFldIsNullable
          .Fields.Append "defensaN", adInteger, 4, adFldIsNullable
          .Fields.Append "OsDevol", adInteger, 4, adFldIsNullable
          .Fields.Append "OsVencim", adInteger, 4, adFldIsNullable
          .Fields.Append "OsMerma", adInteger, 4, adFldIsNullable
          .Fields.Append "IntervencionS", adInteger, 4, adFldIsNullable
          .Fields.Append "otrasS", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
          .Fields.Append "tipo", adVarChar, 15, adFldIsNullable
          .Fields.Append "ImporteSal", adDouble
          .Fields.Append "saldoF", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub


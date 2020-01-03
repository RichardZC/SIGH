VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form rMovimientoES 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos de Entrada y Salida"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rMovimientoES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   15735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6900
      Left            =   15
      TabIndex        =   2
      Top             =   30
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   12171
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "rMovimientoES.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(1)=   "ProgressBar1"
      Tab(0).Control(2)=   "grdMovimientos"
      Tab(0).Control(3)=   "optIngrSalidas"
      Tab(0).Control(4)=   "fraDatosHistoria"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Otros Reportes"
      TabPicture(1)   =   "rMovimientoES.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6270
         Left            =   165
         TabIndex        =   27
         Top             =   375
         Width           =   15165
         Begin VB.OptionButton optExoneracion 
            Caption         =   "Exoneraciones del mes"
            Height          =   390
            Left            =   375
            TabIndex        =   35
            Top             =   4440
            Width           =   4200
         End
         Begin VB.OptionButton optCreditoPendiente 
            Caption         =   "Créditos pendientes de pago del mes"
            Height          =   390
            Left            =   390
            TabIndex        =   34
            Top             =   2790
            Width           =   4200
         End
         Begin VB.OptionButton optCreditoOtorgado 
            Caption         =   "Créditos otorgados y cancelados "
            Height          =   390
            Left            =   390
            TabIndex        =   33
            Top             =   270
            Width           =   6240
         End
         Begin VB.OptionButton optDetalladoIngresos 
            Caption         =   "Ingresos por Proveedor"
            Height          =   390
            Left            =   9195
            TabIndex        =   32
            Top             =   345
            Width           =   2310
         End
         Begin VB.Frame Frame 
            Height          =   975
            Left            =   1950
            TabIndex        =   29
            Top             =   1470
            Width           =   4215
            Begin VB.OptionButton optMismoMes 
               Caption         =   "En el mismo rango de FECHAS"
               Height          =   390
               Left            =   120
               TabIndex        =   31
               Top             =   120
               Value           =   -1  'True
               Width           =   3000
            End
            Begin VB.OptionButton optOtroMes 
               Caption         =   "Fuera del rango de FECHAS"
               Height          =   390
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   3000
            End
         End
         Begin VB.OptionButton optVentasPorProductos 
            Caption         =   "Ventas por producto según forma pago"
            Height          =   390
            Left            =   9195
            TabIndex        =   28
            Top             =   1530
            Width           =   3660
         End
         Begin MSMask.MaskEdBox txtFexon1 
            Height          =   315
            Left            =   1935
            TabIndex        =   36
            Top             =   4920
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtExon2 
            Height          =   315
            Left            =   4815
            TabIndex        =   37
            Top             =   4920
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtFpend1 
            Height          =   315
            Left            =   1950
            TabIndex        =   38
            Top             =   3630
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtFpend2 
            Height          =   315
            Left            =   4830
            TabIndex        =   39
            Top             =   3630
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtFmovOtor1 
            Height          =   315
            Left            =   1950
            TabIndex        =   40
            Top             =   990
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
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
         Begin MSMask.MaskEdBox txtFmovOtor2 
            Height          =   315
            Left            =   4830
            TabIndex        =   41
            Top             =   990
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   750
            TabIndex        =   49
            Top             =   1020
            Width           =   1080
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   210
            Left            =   4350
            TabIndex        =   48
            Top             =   1020
            Width           =   435
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   750
            TabIndex        =   47
            Top             =   3660
            Width           =   1080
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   210
            Left            =   4350
            TabIndex        =   46
            Top             =   3660
            Width           =   435
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "F.Exoneración"
            Height          =   330
            Left            =   735
            TabIndex        =   45
            Top             =   4950
            Width           =   1140
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   210
            Left            =   4335
            TabIndex        =   44
            Top             =   4950
            Width           =   435
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Pacientes: pagantes tomará F.BOLETA,  SEGUROS tomará F.REEMBOLSO"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   750
            TabIndex        =   43
            Top             =   630
            Width           =   6000
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Pacientes: pagantes tomará F.BOLETA,  SEGUROS tomará F.REEMBOLSO"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   750
            TabIndex        =   42
            Top             =   3270
            Width           =   6000
         End
      End
      Begin VB.Frame fraDatosHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74955
         TabIndex        =   5
         Top             =   390
         Width           =   15555
         Begin VB.CheckBox chkDetallado 
            Alignment       =   1  'Right Justify
            Caption         =   "Rep con ITEMS"
            Height          =   225
            Left            =   12240
            TabIndex        =   54
            Top             =   1005
            Width           =   1635
         End
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   14040
            Picture         =   "rMovimientoES.frx":0D02
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   960
            Width           =   1305
         End
         Begin VB.ComboBox cmbAlmacen 
            Height          =   330
            Left            =   960
            TabIndex        =   13
            Top             =   210
            Width           =   3450
         End
         Begin VB.ComboBox cmbAlmacenOrigen 
            Height          =   330
            Left            =   5880
            TabIndex        =   12
            Top             =   600
            Width           =   3765
         End
         Begin VB.ComboBox cmbAlmacenDestino 
            Height          =   330
            Left            =   5880
            TabIndex        =   11
            Top             =   990
            Width           =   3765
         End
         Begin VB.ComboBox cmbEstado 
            Height          =   330
            ItemData        =   "rMovimientoES.frx":394B
            Left            =   5865
            List            =   "rMovimientoES.frx":3958
            TabIndex        =   10
            Top             =   225
            Width           =   3765
         End
         Begin VB.ComboBox cmbmovTipo 
            Height          =   330
            ItemData        =   "rMovimientoES.frx":3988
            Left            =   960
            List            =   "rMovimientoES.frx":3995
            TabIndex        =   9
            Top             =   975
            Width           =   3450
         End
         Begin VB.ComboBox cmbConcepto 
            Height          =   330
            Left            =   960
            TabIndex        =   8
            Top             =   600
            Width           =   3450
         End
         Begin VB.CheckBox chkExcel 
            Alignment       =   1  'Right Justify
            Caption         =   "En Excel"
            Height          =   315
            Left            =   10125
            Picture         =   "rMovimientoES.frx":39CF
            TabIndex        =   7
            Top             =   960
            Width           =   1020
         End
         Begin VB.ComboBox cmbUsuario 
            Height          =   330
            ItemData        =   "rMovimientoES.frx":3CE1
            Left            =   10965
            List            =   "rMovimientoES.frx":3CEB
            TabIndex        =   6
            Text            =   "cmbUsuario"
            Top             =   585
            Width           =   4410
         End
         Begin MSMask.MaskEdBox txtFdesde 
            Height          =   315
            Left            =   10965
            TabIndex        =   14
            Top             =   195
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtFhasta 
            Height          =   315
            Left            =   13275
            TabIndex        =   15
            Top             =   195
            Width           =   1350
            _ExtentX        =   2381
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
         Begin MSMask.MaskEdBox txtHrInicio 
            Height          =   315
            Left            =   12345
            TabIndex        =   16
            Top             =   195
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
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
         Begin MSMask.MaskEdBox txtHrFin 
            Height          =   315
            Left            =   14610
            TabIndex        =   17
            Top             =   195
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   210
            Left            =   13185
            TabIndex        =   26
            Top             =   210
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   9885
            TabIndex        =   25
            Top             =   255
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   210
            Left            =   5265
            TabIndex        =   22
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Almacén Origen"
            Height          =   210
            Left            =   4530
            TabIndex        =   21
            Top             =   675
            Width           =   1290
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Almacén Destino"
            Height          =   210
            Left            =   4455
            TabIndex        =   20
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Mov"
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   1035
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor Farm"
            Height          =   210
            Left            =   9705
            TabIndex        =   18
            Top             =   645
            Width           =   1260
         End
      End
      Begin VB.OptionButton optIngrSalidas 
         Caption         =   "Consolidado de Ingresos y Salidas"
         Height          =   420
         Left            =   -74880
         TabIndex        =   4
         Top             =   60
         Value           =   -1  'True
         Width           =   3780
      End
      Begin UltraGrid.SSUltraGrid grdMovimientos 
         Height          =   4665
         Left            =   -74955
         TabIndex        =   50
         Top             =   1815
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   8229
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "rMovimientoES.frx":3D07
         Caption         =   "Movimientos"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   -69915
         TabIndex        =   51
         Top             =   6585
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Pulse DOBLE CLIC PARA ver detalle del Documento"
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
         Left            =   -74895
         TabIndex        =   52
         Top             =   6570
         Width           =   4680
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
      Height          =   1110
      Left            =   30
      TabIndex        =   0
      Top             =   6885
      Width           =   15660
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rMovimientoES.frx":3D5A
         DownPicture     =   "rMovimientoES.frx":41BA
         Height          =   700
         Left            =   6420
         Picture         =   "rMovimientoES.frx":462F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rMovimientoES.frx":4AA4
         DownPicture     =   "rMovimientoES.frx":4F68
         Height          =   700
         Left            =   7950
         Picture         =   "rMovimientoES.frx":5454
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rMovimientoES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte Movimientos de Entrada y Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbAlmacen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenDestino As New SIGHEntidades.ListaDespleglable
Dim mo_cmbConceptos As New SIGHEntidades.ListaDespleglable
Dim mo_cmbUsuario As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim mrs_Tmp As New Recordset
Dim ms_MensajeError As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    'If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRep As New RepMovimientoES
        Dim orsTmp123 As New Recordset
        Dim mrs_tmp1 As New Recordset
        Dim oConexion As New Connection
        If SSTab1.Tab = 0 Then
            If chkDetallado.Value = 1 Then
               If mrs_Tmp.RecordCount > 0 Then
                    oConexion.CommandTimeout = 900
                    oConexion.CursorLocation = adUseClient
                    oConexion.Open SIGHEntidades.CadenaConexion
                    Set mrs_tmp1 = CopyRecordset(mrs_Tmp, "")
                    mrs_Tmp.MoveFirst
                    Do While Not mrs_Tmp.EOF
                       Set orsTmp123 = mo_ReglasFarmacia.farmMovimientoDetalleSeleccionarPorMovNumeroTipo(mrs_Tmp!movNumero, _
                                                                                        mrs_Tmp!MovTipo, oConexion)
                       If orsTmp123.RecordCount > 0 Then
                          orsTmp123.MoveFirst
                          Do While Not orsTmp123.EOF
                                mrs_tmp1.AddNew
                                mrs_tmp1!fechaCreacion = mrs_Tmp!fechaCreacion
                                mrs_tmp1!HoraCreacion = mrs_Tmp!HoraCreacion
                                mrs_tmp1!MovTipo = mrs_Tmp!MovTipo
                                mrs_tmp1!movNumero = mrs_Tmp!movNumero
                                mrs_tmp1!ingresos = mrs_Tmp!ingresos
                                mrs_tmp1!salidas = mrs_Tmp!salidas
                                mrs_tmp1!saldo = mrs_Tmp!saldo
                                mrs_tmp1!Abreviatura = mrs_Tmp!Abreviatura
                                mrs_tmp1!DocumentoNumero = mrs_Tmp!DocumentoNumero
                                mrs_tmp1!Concepto = mrs_Tmp!Concepto
                                mrs_tmp1!fOrigen = mrs_Tmp!fOrigen
                                mrs_tmp1!Lote = mrs_Tmp!Lote
                                mrs_tmp1!FechaVencimiento = mrs_Tmp!FechaVencimiento
                                mrs_tmp1!fDestino = mrs_Tmp!fDestino
                                mrs_tmp1!Estado = mrs_Tmp!Estado
                                mrs_tmp1!total = mrs_Tmp!total
                                mrs_tmp1!Item = Trim(orsTmp123!Nombre) & " (" & Trim(orsTmp123!codigo) & ")"
                                mrs_tmp1!Precio = orsTmp123!Precio
                                mrs_tmp1!Cantidad = orsTmp123!Cantidad
                                mrs_tmp1.Update
                                orsTmp123.MoveNext
                          Loop
                       End If
                       orsTmp123.Close
                       mrs_Tmp.MoveNext
                    Loop
                    oConexion.Close
               End If
            End If
            If chkExcel.Value = 1 Then
                mo_ReglasReportes.ExportarRecordSetAexcel IIf(chkDetallado.Value = 1, mrs_tmp1, mrs_Tmp), optIngrSalidas.Caption, ml_TextoDelFiltro, "", Me.hwnd
            Else
                Dim oRptClaseCry As New rCrystal
                oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
                oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
                oRptClaseCry.IdAlmacenDestino = Val(mo_cmbAlmacenDestino.BoundText)
                oRptClaseCry.IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
                oRptClaseCry.Concepto = Val(mo_cmbConceptos.BoundText)
                oRptClaseCry.Estado = cmbEstado.ListIndex
                oRptClaseCry.MovTipo = IIf(cmbmovTipo.ListIndex = 0, "E", "S")
                oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
                oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
                oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
                oRptClaseCry.TipoReporte = Me.Name
                oRptClaseCry.idUsuario = Val(mo_cmbUsuario.BoundText)
                oRptClaseCry.ConsiderarDetalle = IIf(chkDetallado.Value = 1, True, False)
                Set oRptClaseCry.oRsRecord = IIf(chkDetallado.Value = 1, mrs_tmp1, mrs_Tmp)
                oRptClaseCry.Show vbModal
                Set oRptClaseCry = Nothing
            End If
        Else
            If optExoneracion.Value = True Then
                oRep.ReporteExoneraciones txtFexon1.Text, txtExon2.Text, Me.hwnd
            ElseIf optCreditoOtorgado.Value = True Then
                oRep.ReporteCreditosCancelados txtFmovOtor1.Text, txtFmovOtor2.Text & " 23:59:59", Me.hwnd, optMismoMes.Value
            ElseIf optCreditoPendiente.Value = True Then
                oRep.ReporteCreditosPendientes txtFpend1.Text, txtFpend2.Text & " 23:59:59", Me.hwnd
            End If
        End If
        Set oRep = Nothing
        Set orsTmp123 = Nothing
        Set mrs_tmp1 = Nothing
        Set oConexion = Nothing
        
        Me.MousePointer = 1
    'End If
End Sub
Private Sub btnBuscar_Click()
  If ValidaDatosObligatorios Then
    FiltrarDatosMovimientos Val(mo_cmbAlmacen.BoundText), IIf(cmbmovTipo.ListIndex = 0, "E", IIf(cmbmovTipo.ListIndex = 1, "S", "")), _
                 CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                 CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                 Val(mo_cmbConceptos.BoundText), cmbEstado.ListIndex, Val(mo_cmbAlmacenOrigen.BoundText), _
                 Val(mo_cmbAlmacenDestino.BoundText), Val(mo_cmbUsuario.BoundText)
  End If
End Sub


Private Sub FiltrarDatosMovimientos(lnIdAlmacen As Long, ml_MovTipo1 As String, mda_FechaInicio As Date, _
                            mda_FechaFin As Date, ml_IdConcepto As Long, ml_IdEstado As Long, _
                            lnIdAlmacenOrigen As Long, lnIdAlmacenDestino As Long, ml_idUsuario As Long)
        Dim oConexion As New Connection
        Dim rsReporte As New Recordset
        Dim mrs_Tmp3 As New Recordset
        Dim mrs_Tmp99 As New Recordset
        Dim lcTexto1 As String, lbContinuar As Boolean, lcTexto3 As String, lnFor1 As Integer, lnFor As Integer
        Dim ml_MovTipo As String
        Me.MousePointer = 11
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        lnFor1 = IIf(ml_MovTipo1 = "", 2, 1)
        For lnFor = 1 To lnFor1
            ml_MovTipo = ml_MovTipo1
            If lnFor1 = 2 Then
               If lnFor = 1 Then
                  ml_MovTipo = "E"
               Else
                  ml_MovTipo = "S"
               End If
            End If
            Set rsReporte = mo_ReglasFarmacia.FarmDevuelveMovimientos(lnIdAlmacen, ml_MovTipo, mda_FechaInicio, mda_FechaFin)
            lcTexto1 = ""
            
            If ml_IdConcepto > 0 Then
               lcTexto1 = lcTexto1 & " IdTipoConcepto=" & ml_IdConcepto & " and "
            End If
            If ml_IdEstado <> 2 Then
               lcTexto1 = lcTexto1 & " IdEstadoMovimiento=" & ml_IdEstado & " and "
            End If
            If lnIdAlmacenOrigen > 0 Then
               lcTexto1 = lcTexto1 & " IdAlmacenOrigen=" & lnIdAlmacenOrigen & " and "
            End If
            If lnIdAlmacenDestino > 0 Then
               lcTexto1 = lcTexto1 & " IdAlmacenDestino=" & lnIdAlmacenDestino & " and "
            End If
            If ml_idUsuario > 0 Then
               lcTexto1 = lcTexto1 & " IdUsuario=" & ml_idUsuario & " and "
            End If
            If lcTexto1 <> "" Then
               lcTexto1 = Left(lcTexto1, Len(lcTexto1) - 5)
               rsReporte.Filter = lcTexto1
            End If
            If lnFor = 1 Then
                If mrs_Tmp.State = 1 Then
                   Set mrs_Tmp = Nothing
                End If
                With mrs_Tmp
                      .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                      .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                      .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
                      .Fields.Append "MovNumero", adVarChar, 10, adFldIsNullable
                      .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
                      .Fields.Append "salidas", adInteger, 4, adFldIsNullable
                      .Fields.Append "saldo", adInteger, 4, adFldIsNullable
                      .Fields.Append "Abreviatura", adVarChar, 10, adFldIsNullable
                      .Fields.Append "DocumentoNumero", adVarChar, 20, adFldIsNullable
                      .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
                      .Fields.Append "fOrigen", adVarChar, 100, adFldIsNullable
                      .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
                      .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
                      .Fields.Append "fDestino", adVarChar, 100, adFldIsNullable
                      .Fields.Append "Estado", adVarChar, 30, adFldIsNullable
                      .Fields.Append "Total", adDouble
                      .Fields.Append "Item", adVarChar, 200, adFldIsNullable
                      .Fields.Append "Precio", adDouble
                      .Fields.Append "Cantidad", adInteger
                      .LockType = adLockOptimistic
                      .Open
                End With
            
                With mrs_Tmp99
                      .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
                      .Fields.Append "salidas", adInteger, 4, adFldIsNullable
                      .Fields.Append "saldo", adInteger, 4, adFldIsNullable
                      .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
                      .LockType = adLockOptimistic
                      .Open
                End With
            End If
            If rsReporte.RecordCount > 0 Then
                Me.ProgressBar1.Min = 0
                Me.ProgressBar1.Max = rsReporte.RecordCount
                Me.ProgressBar1.Value = 0
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                   DoEvents: Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: Me.Refresh
                
                   lbContinuar = True
                   lcTexto1 = ""
                   If rsReporte.Fields!MovTipo = "S" Then
                         Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoVentasFiltrarMovnumero(rsReporte.Fields!movNumero)
                         If mrs_Tmp3.RecordCount > 0 Then
                            lcTexto1 = " (Pac: " & Trim(mrs_Tmp3.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp3.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp3.Fields!PrimerNombre) & ")"
                         Else
                            lcTexto1 = ""
                         End If
                         mrs_Tmp3.Close
                    End If
                    If lbContinuar = True Then
                        lcTexto3 = ""
                        If rsReporte.Fields!MovTipo = "E" Then
                            Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoNotaIngresoSeleccionarXmovimiento(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                            If mrs_Tmp3.RecordCount > 0 Then
                               If Not IsNull(mrs_Tmp3!Abreviatura) Then
                                  lcTexto3 = Trim(mrs_Tmp3!Abreviatura)
                               End If
                               If Not IsNull(mrs_Tmp3!oRigenNumero) Then
                                  lcTexto3 = lcTexto3 & " " & Trim(mrs_Tmp3!oRigenNumero)
                               End If
                               If lcTexto3 <> "" Then
                                  lcTexto3 = " (" & lcTexto3 & ")"
                               End If
                            End If
                        End If
                        '
                        
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
                        mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                        mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                        mrs_Tmp.Fields!Abreviatura = rsReporte.Fields!Abreviatura
                        mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                        mrs_Tmp.Fields!Concepto = rsReporte.Fields!Concepto
                        mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fOrigen & lcTexto3, 100)
                        mrs_Tmp.Fields!fDestino = Trim(rsReporte.Fields!fDestino) & lcTexto1
                        mrs_Tmp.Fields!Estado = rsReporte.Fields!Estado
                        mrs_Tmp.Fields!total = rsReporte.Fields!total
                        mrs_Tmp.Update
                    End If
                    rsReporte.MoveNext
                Loop
                
            End If
        Next
        mrs_Tmp.Sort = "fechaCreacion desc,horaCreacion desc"
        Set grdMovimientos.DataSource = mrs_Tmp
        mo_Apariencia.ConfigurarFilasBiColores Me.grdMovimientos, SIGHEntidades.GrillaConFilasBicolor
        On Error Resume Next
        If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No hay información con esos datos", vbInformation, Me.Caption
        End If
        
        Me.MousePointer = 1
        Set oConexion = Nothing
        Set rsReporte = Nothing
        Set mrs_Tmp3 = Nothing
        Set mrs_Tmp99 = Nothing
End Sub



Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    If optIngrSalidas.Value = True Then
        ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")     F.Movimiento: (" & txtFdesde.Text & " al " & txtFhasta.Text & ")     Tipo Movimiento: (" & Trim(cmbmovTipo.Text) & ")     Estado: (" & Trim(cmbEstado.Text) & ")"
        If cmbConcepto.Text <> "" Then
           ml_TextoDelFiltro = ml_TextoDelFiltro & "     Concepto: (" & Trim(cmbConcepto.Text) & ")"
        End If
        If cmbAlmacenOrigen.Text <> "" Then
           ml_TextoDelFiltro = ml_TextoDelFiltro & "     Alm.Origen: (" & Trim(cmbAlmacenOrigen.Text) & ")"
        End If
        If cmbAlmacenDestino.Text <> "" Then
           ml_TextoDelFiltro = ml_TextoDelFiltro & "     Alm.Destino: (" & Trim(cmbAlmacenDestino.Text) & ")"
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(Val(mo_cmbUsuario.BoundText) > 0, "     (Vendedor: " & Trim(cmbUsuario.Text) & ")", "")
        If mo_cmbAlmacen.BoundText = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén" + Chr(13)
            cmbAlmacen.SetFocus
        End If
        If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
           Exit Function
        End If
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function







Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub



















Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub






Private Sub cmbAlmacenDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacenDestino

End Sub

Private Sub cmbAlmacenOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacenOrigen

End Sub

Private Sub cmbConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConcepto

End Sub



Private Sub cmbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbEstado

End Sub

Private Sub cmbmovTipo_Click()
    If cmbmovTipo.ListIndex = 0 Then
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenOrigen, True   'por ser Ingresos
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenDestino, False   'por ser Ingresos
    Else
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenOrigen, False   'por ser salidas
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenDestino, True   'por ser salidas
    End If
End Sub

Private Sub cmbmovTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbmovTipo

End Sub





Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmacenOrigen
    Set mo_cmbAlmacenDestino.MiComboBox = cmbAlmacenDestino
    Set mo_cmbConceptos.MiComboBox = cmbConcepto
    Set mo_cmbUsuario.MiComboBox = cmbUsuario
End Sub


Sub InicializaFechaHora()
    txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267)
    txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268)
    txtFexon1.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtExon2.Text = Date
    txtFpend1.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFpend2.Text = Date
    txtFmovOtor1.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFmovOtor2.Text = Date
End Sub
Private Sub Form_Load()
    InicializaFechaHora
    
    cmbEstado.ListIndex = 2
    cmbmovTipo.ListIndex = 0    'Ingresos
    cmbmovTipo_Click
    
    '
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    '
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("")
    '
    mo_cmbAlmacenDestino.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenDestino.ListField = "Descripcion"
    Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("")
    '
    mo_cmbConceptos.BoundColumn = "IdTipoConcepto"
    mo_cmbConceptos.ListField = "Concepto"
    Set mo_cmbConceptos.RowSource = mo_ReglasFarmacia.FarmTipoConceptosDevuelveTodos
    '
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
    End If
    '
    mo_cmbUsuario.BoundColumn = "IdEmpleado"
    mo_cmbUsuario.ListField = "DEmpleado"
    Set mo_cmbUsuario.RowSource = mo_ReglasComunes.EmpleadosSeleccionarTodos
    
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub




Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdMovimientos_DblClick()
    On Error GoTo ErrDbl
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = grdMovimientos.DataSource
    If oRsTmp1!MovTipo = "S" Or oRsTmp1!MovTipo = "E" Then
       If oRsTmp1!MovTipo = "S" Then
            Dim oFarmNotaSalida As New FarmNotaSalida
            oFarmNotaSalida.Opcion = sghConsultar
            oFarmNotaSalida.movNumero = oRsTmp1!movNumero
            oFarmNotaSalida.Show 1
            If InStr(ml_TextoDelFiltro, "Especializado") > 0 Then     'debb1212
               oFarmNotaSalida.lnIdTablaLISTBARITEMS = 1305
            Else
               oFarmNotaSalida.lnIdTablaLISTBARITEMS = 1358
            End If
            Set oFarmNotaSalida = Nothing
       Else
            Dim oFarmNotaIngreso As New FarmNotaIngreso
            oFarmNotaIngreso.Opcion = sghConsultar
            oFarmNotaIngreso.movNumero = oRsTmp1!movNumero
            If InStr(ml_TextoDelFiltro, "Especializado") > 0 Then
               oFarmNotaIngreso.lnIdTablaLISTBARITEMS = 1304              'debb1212
            Else
               oFarmNotaIngreso.lnIdTablaLISTBARITEMS = 1357
            End If
            oFarmNotaIngreso.Show 1
            Set oFarmNotaIngreso = Nothing
       End If
    End If
ErrDbl:
End Sub

Private Sub grdMovimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    On Error Resume Next
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdMovimientos.Bands(0).Columns("ingresos").Hidden = True
    grdMovimientos.Bands(0).Columns("salidas").Hidden = True
    grdMovimientos.Bands(0).Columns("saldo").Hidden = True
    grdMovimientos.Bands(0).Columns("lote").Hidden = True
    grdMovimientos.Bands(0).Columns("fechaVencimiento").Hidden = True
    grdMovimientos.Bands(0).Columns("FechaCreacion").Header.Caption = "Fecha"
    grdMovimientos.Bands(0).Columns("FechaCreacion").Width = 800
    grdMovimientos.Bands(0).Columns("FechaCreacion").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("HoraCreacion").Header.Caption = "Hora"
    grdMovimientos.Bands(0).Columns("HoraCreacion").Width = 500
    grdMovimientos.Bands(0).Columns("HoraCreacion").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("movTipo").Header.Caption = "Tipo"
    grdMovimientos.Bands(0).Columns("movTipo").Width = 200
    grdMovimientos.Bands(0).Columns("movTipo").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("MovNumero").Header.Caption = "N° Registro"
    grdMovimientos.Bands(0).Columns("MovNumero").Width = 1000
    grdMovimientos.Bands(0).Columns("MovNumero").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("Abreviatura").Header.Caption = "Doc.Tipo"
    grdMovimientos.Bands(0).Columns("Abreviatura").Width = 500
    grdMovimientos.Bands(0).Columns("Abreviatura").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("DocumentoNumero").Header.Caption = "Doc.N°"
    grdMovimientos.Bands(0).Columns("DocumentoNumero").Width = 1200
    grdMovimientos.Bands(0).Columns("DocumentoNumero").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("Concepto").Header.Caption = "Concepto"
    grdMovimientos.Bands(0).Columns("Concepto").Width = 3000
    grdMovimientos.Bands(0).Columns("Concepto").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("fOrigen").Header.Caption = "Origen"
    grdMovimientos.Bands(0).Columns("fOrigen").Width = 3000
    grdMovimientos.Bands(0).Columns("fOrigen").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("fDestino").Header.Caption = "Destino"
    grdMovimientos.Bands(0).Columns("fDestino").Width = 3000
    grdMovimientos.Bands(0).Columns("fDestino").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("estado").Header.Caption = "Estado"
    grdMovimientos.Bands(0).Columns("estado").Width = 700
    grdMovimientos.Bands(0).Columns("estado").Activation = ssActivationActivateNoEdit
    grdMovimientos.Bands(0).Columns("total").Header.Caption = "Importe"
    grdMovimientos.Bands(0).Columns("total").Width = 1000
    grdMovimientos.Bands(0).Columns("total").Activation = ssActivationActivateNoEdit

End Sub

Private Sub optDetalladoIngresos_Click()
    Dim orProductosIngresados As New rProductosIngresados
    orProductosIngresados.NroReporte = 1
    orProductosIngresados.Show 1
    Set orProductosIngresados = Nothing

End Sub



Private Sub optVentasPorProductos_Click()
    Dim lcMensajeLicencia As String
    'If mo_ReglasComunes.EESSconDerechosAmejoras(2, "61008", lcMensajeLicencia) = True Then
        Dim orProductosIngresados As New rProductosIngresados
        orProductosIngresados.NroReporte = 2
        orProductosIngresados.Show 1
        Set orProductosIngresados = Nothing
   ' End If
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbAlmacenDestino = Nothing
    Set mo_cmbConceptos = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
    
End Sub



Private Sub txtHrFin_LostFocus()
         If Not SIGHEntidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

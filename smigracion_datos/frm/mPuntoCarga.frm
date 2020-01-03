VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mVariosProcesos 
   Caption         =   "Varios Procesos"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   62
      Top             =   9120
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "mPuntoCarga.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdErrores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdActualiza"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame36"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdActualizaNroAutomatico"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdActualizaDNI"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "mPuntoCarga.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Farmacia"
      TabPicture(2)   =   "mPuntoCarga.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame17"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame33"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame12"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame31"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame1(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame1(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "HIS"
      TabPicture(3)   =   "mPuntoCarga.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame21"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame20"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame15"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame22"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame10"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdActEdades"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Frame32"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Frame37"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Frame38"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Frame11"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Impresion de Boleta (CAJA)"
      TabPicture(4)   =   "mPuntoCarga.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frDocumentos"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame27"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame26"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame25"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame24"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "btnGuardarConfiguracionComprobante"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Frame28"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Frame34"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      Begin VB.Frame Frame13 
         BackColor       =   &H8000000D&
         Caption         =   "Cambia el Nro de Historia Clnica por el DNI"
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   6435
         TabIndex        =   242
         Top             =   5355
         Width           =   6285
         Begin VB.CommandButton cmdProcesaHistoriasConDNI 
            Caption         =   "Procesar   (luego verifique REPORTE->ARCH.CLINICOS->  las historias que hoy pasaron a ser DNI)"
            Height          =   495
            Left            =   120
            TabIndex        =   244
            Top             =   660
            Width           =   5985
         End
         Begin VB.TextBox txtNroHistoriasXdia 
            Height          =   375
            Left            =   2475
            TabIndex        =   243
            Text            =   "40"
            Top             =   225
            Width           =   585
         End
         Begin VB.Label Label62 
            Caption         =   "N° Historias a procesar por día"
            Height          =   285
            Left            =   90
            TabIndex        =   245
            Top             =   270
            Width           =   2340
         End
      End
      Begin VB.CommandButton Command 
         Height          =   210
         Left            =   12450
         TabIndex        =   241
         Top             =   4035
         Width           =   165
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualiza Id de RESULTADOS LABORATORIO"
         Height          =   615
         Left            =   6480
         TabIndex        =   240
         Top             =   4635
         Width           =   6135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000D&
         Height          =   1185
         Index           =   3
         Left            =   -74805
         TabIndex        =   236
         Top             =   7620
         Width           =   6270
         Begin VB.CommandButton cmdActPreciosFarm 
            Caption         =   "Proceso para actualizar PRECIO FARMACIA"
            Height          =   450
            Left            =   165
            TabIndex        =   239
            Top             =   660
            Width           =   5850
         End
         Begin VB.TextBox txtNewFF 
            Height          =   375
            Left            =   3930
            TabIndex        =   238
            Top             =   225
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Financiamiento(PRODUCTO/PLAN)   (nueva):"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   237
            Top             =   285
            Width           =   3645
         End
      End
      Begin VB.CommandButton cmdActualizaDNI 
         Caption         =   "Actualiza HISTORIA=DNI a 9 digitos, si parametro=351=S"
         Height          =   615
         Left            =   6555
         TabIndex        =   234
         Top             =   3945
         Width           =   5745
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         Caption         =   "Repara items con TIPOSALIDA errados"
         Height          =   3120
         Index           =   0
         Left            =   -68460
         TabIndex        =   212
         Top             =   4140
         Width           =   6030
         Begin VB.Frame Frame5 
            Height          =   1455
            Left            =   150
            TabIndex        =   215
            Top             =   1335
            Width           =   5625
            Begin VB.TextBox txtNtipoSalida 
               Height          =   285
               Left            =   1665
               TabIndex        =   222
               Top             =   615
               Width           =   300
            End
            Begin VB.CommandButton cmdCodigo 
               Caption         =   "Cambia TIPO DE SALIDA"
               Height          =   375
               Left            =   135
               TabIndex        =   220
               Top             =   1005
               Width           =   5235
            End
            Begin VB.TextBox txtClaveCodigo 
               Height          =   345
               IMEMode         =   3  'DISABLE
               Left            =   4335
               PasswordChar    =   "*"
               TabIndex        =   218
               Top             =   180
               Width           =   1185
            End
            Begin VB.TextBox txtCodigoSismed 
               Height          =   345
               Left            =   765
               TabIndex        =   216
               Top             =   150
               Width           =   1185
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "1) Ventas    2)Interveción Sanitaria"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   2025
               TabIndex        =   223
               Top             =   675
               Width           =   2445
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Nuevo Tipo Salida"
               Height          =   195
               Left            =   165
               TabIndex        =   221
               Top             =   675
               Width           =   1320
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Clave"
               Height          =   195
               Left            =   3825
               TabIndex        =   219
               Top             =   255
               Width           =   405
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código"
               Height          =   195
               Left            =   150
               TabIndex        =   217
               Top             =   225
               Width           =   495
            End
         End
         Begin VB.CommandButton cmdListaTipoSalida 
            Caption         =   "Agrega al EXCEL lista de Items con TIPO SALIDA a verificar con los ALMACENEROS"
            Height          =   510
            Left            =   150
            TabIndex        =   213
            Top             =   780
            Width           =   5640
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe existir    'c:\excel.xls'  y    'hoja1'    vacio"
            Height          =   315
            Left            =   120
            TabIndex        =   214
            Top             =   360
            Width           =   5670
         End
      End
      Begin VB.Frame Frame34 
         Height          =   1695
         Left            =   -64920
         TabIndex        =   209
         Top             =   6600
         Width           =   2145
         Begin VB.CommandButton cmdImprimeBoleta 
            Caption         =   "Imprime BOLETA de prueba"
            Height          =   705
            Left            =   120
            TabIndex        =   210
            Top             =   840
            Width           =   1845
         End
         Begin VB.Label Label61 
            Alignment       =   2  'Center
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   90
            TabIndex        =   211
            Top             =   330
            Width           =   1845
         End
      End
      Begin VB.Frame Frame28 
         Height          =   2265
         Left            =   -69480
         TabIndex        =   192
         Top             =   6600
         Width           =   4395
         Begin VB.TextBox txtPieAlto 
            Height          =   345
            Left            =   1590
            TabIndex        =   200
            Top             =   630
            Width           =   855
         End
         Begin VB.TextBox txtCabeceraAlto 
            Height          =   345
            Left            =   1590
            TabIndex        =   199
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txtMargenInferior 
            Height          =   345
            Left            =   3480
            TabIndex        =   198
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtMargenSuperior 
            Height          =   345
            Left            =   1590
            TabIndex        =   197
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtMargenDerecha 
            Height          =   345
            Left            =   3480
            TabIndex        =   196
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtMargenIzquierda 
            Height          =   345
            Left            =   3480
            TabIndex        =   195
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboReporteador 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   1485
            Width           =   2295
         End
         Begin VB.ComboBox cboPapel 
            Height          =   315
            ItemData        =   "mPuntoCarga.frx":008C
            Left            =   1590
            List            =   "mPuntoCarga.frx":008E
            Sorted          =   -1  'True
            TabIndex        =   193
            Text            =   "cboPapel"
            Top             =   1875
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label54 
            Caption         =   "Pie de Pagina (Alto)"
            Height          =   255
            Left            =   150
            TabIndex        =   208
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label53 
            Caption         =   "Cabecera (Alto)"
            Height          =   255
            Left            =   150
            TabIndex        =   207
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "M. Inferior"
            Height          =   375
            Left            =   2520
            TabIndex        =   206
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "M. Superior"
            Height          =   375
            Left            =   120
            TabIndex        =   205
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "M. Derecha"
            Height          =   375
            Left            =   2520
            TabIndex        =   204
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "M. Izquierdo"
            Height          =   375
            Left            =   2520
            TabIndex        =   203
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Nombre Hoja"
            Height          =   255
            Left            =   120
            TabIndex        =   202
            Top             =   1920
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Reporteador"
            Height          =   255
            Left            =   120
            TabIndex        =   201
            Top             =   1560
            Width           =   1335
         End
      End
      Begin VB.CommandButton btnGuardarConfiguracionComprobante 
         Caption         =   "Grabar Configuración"
         Enabled         =   0   'False
         Height          =   555
         Left            =   -64920
         TabIndex        =   191
         Top             =   8370
         Width           =   2085
      End
      Begin VB.Frame Frame11 
         Caption         =   "Carga movimientos de ATENCIONES HISTORICAS"
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   -68520
         TabIndex        =   187
         Top             =   6300
         Width           =   6015
         Begin VB.CommandButton cmdCargaAtencionesHist 
            Caption         =   "Procesar"
            Height          =   915
            Left            =   4200
            TabIndex        =   188
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   190
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HistCab.dbf, histDet.dbf  deben estar en la carpeta de ODBC HIS"
            Height          =   525
            Left            =   120
            TabIndex        =   189
            Top             =   600
            Width           =   3975
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Carga valores del archivo SETUP_CAJA.INI"
         Height          =   1125
         Left            =   -74940
         TabIndex        =   168
         Top             =   1260
         Width           =   12165
         Begin VB.CommandButton cmdCargaINICajaServicios 
            Caption         =   "Boleta SERVICIOS                    Carga Valores desde: 'c:\.....\archivos\setup_caja_boleta.ini'"
            Height          =   465
            Left            =   2640
            TabIndex        =   171
            Top             =   550
            Width           =   4635
         End
         Begin VB.TextBox txtRutaINI 
            Height          =   315
            Left            =   3300
            TabIndex        =   170
            Text            =   "Text4"
            Top             =   180
            Width           =   8715
         End
         Begin VB.CommandButton cmdCargaINICajaFarmacia 
            Caption         =   "Boleta FARMACIA                Carga Valores desde: 'c:\.....\archivos\setup_caja_boleta.ini'"
            Height          =   465
            Left            =   7380
            TabIndex        =   169
            Top             =   550
            Width           =   4635
         End
         Begin VB.Label lblRutaArchivo 
            Caption         =   "Ruta del archivo:   setup_caja_boleta.ini"
            Height          =   225
            Left            =   120
            TabIndex        =   172
            Top             =   255
            Width           =   3045
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Datos de Cabecera"
         Height          =   4185
         Left            =   -74940
         TabIndex        =   138
         Top             =   2400
         Width           =   5265
         Begin VB.TextBox txtDireccionY 
            Height          =   345
            Left            =   2820
            TabIndex        =   252
            Top             =   3675
            Width           =   885
         End
         Begin VB.TextBox txtDireccionX 
            Height          =   345
            Left            =   1920
            TabIndex        =   251
            Top             =   3675
            Width           =   855
         End
         Begin VB.TextBox txtDireccionProv 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   250
            Text            =   "jr. apurimac N 88"
            Top             =   3675
            Width           =   1365
         End
         Begin VB.TextBox txtRucY 
            Height          =   345
            Left            =   2820
            TabIndex        =   248
            Top             =   3255
            Width           =   885
         End
         Begin VB.TextBox txtRucX 
            Height          =   345
            Left            =   1920
            TabIndex        =   247
            Top             =   3255
            Width           =   855
         End
         Begin VB.TextBox txtRucProv 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   246
            Text            =   "12345678901"
            Top             =   3255
            Width           =   1365
         End
         Begin VB.TextBox txtHistoriaY 
            Height          =   345
            Left            =   2820
            TabIndex        =   178
            Top             =   2830
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox txtHistoriaX 
            Height          =   345
            Left            =   1920
            TabIndex        =   177
            Top             =   2830
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtHistoriaValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   176
            Text            =   "100050"
            Top             =   2830
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtNumeroSerieY 
            Height          =   345
            Left            =   2820
            TabIndex        =   159
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox txtNumeroSerieX 
            Height          =   345
            Left            =   1920
            TabIndex        =   158
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtNumeroSerieValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   157
            Text            =   "123-123456"
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txtEstadoValor 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   156
            Text            =   "Anulado"
            Top             =   690
            Width           =   1365
         End
         Begin VB.TextBox txtEstadoX 
            Height          =   345
            Left            =   1920
            TabIndex        =   155
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox txtEstadoY 
            Height          =   345
            Left            =   2820
            TabIndex        =   154
            Top             =   690
            Width           =   885
         End
         Begin VB.TextBox txtTipoValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   153
            Text            =   "Devolución"
            Top             =   1050
            Width           =   1365
         End
         Begin VB.TextBox txtTipoX 
            Height          =   345
            Left            =   1920
            TabIndex        =   152
            Top             =   1050
            Width           =   855
         End
         Begin VB.TextBox txtTipoY 
            Height          =   345
            Left            =   2820
            TabIndex        =   151
            Top             =   1050
            Width           =   885
         End
         Begin VB.TextBox txtRzSocialValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   150
            Text            =   "Perez Pereira Jorge"
            Top             =   1410
            Width           =   1365
         End
         Begin VB.TextBox txtRzSocialX 
            Height          =   345
            Left            =   1920
            TabIndex        =   149
            Top             =   1410
            Width           =   855
         End
         Begin VB.TextBox txtRzSocialY 
            Height          =   345
            Left            =   2820
            TabIndex        =   148
            Top             =   1410
            Width           =   885
         End
         Begin VB.TextBox txtFechaVAlor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   147
            Text            =   "21/11/2000"
            Top             =   1740
            Width           =   1365
         End
         Begin VB.TextBox txtFechaX 
            Height          =   345
            Left            =   1920
            TabIndex        =   146
            Top             =   1740
            Width           =   855
         End
         Begin VB.TextBox txtFechaY 
            Height          =   345
            Left            =   2820
            TabIndex        =   145
            Top             =   1740
            Width           =   885
         End
         Begin VB.TextBox txtServicioValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   144
            Text            =   "Medicina General"
            Top             =   2100
            Width           =   1365
         End
         Begin VB.TextBox txtServicioX 
            Height          =   345
            Left            =   1920
            TabIndex        =   143
            Top             =   2100
            Width           =   855
         End
         Begin VB.TextBox txtServicioY 
            Height          =   345
            Left            =   2820
            TabIndex        =   142
            Top             =   2100
            Width           =   885
         End
         Begin VB.TextBox txtObservacionesValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3750
            TabIndex        =   141
            Text            =   "Observaciones"
            Top             =   2460
            Width           =   1365
         End
         Begin VB.TextBox txtObservacionesX 
            Height          =   345
            Left            =   1920
            TabIndex        =   140
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtObservacionesY 
            Height          =   345
            Left            =   2820
            TabIndex        =   139
            Top             =   2460
            Width           =   885
         End
         Begin VB.Label Label63 
            Caption         =   "Dirección (factura)"
            Height          =   285
            Left            =   120
            TabIndex        =   253
            Top             =   3735
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "Ruc (factura)"
            Height          =   285
            Left            =   120
            TabIndex        =   249
            Top             =   3315
            Width           =   1815
         End
         Begin VB.Label lblHistoria 
            Caption         =   "Nº Historia:"
            Height          =   285
            Left            =   120
            TabIndex        =   179
            Top             =   2880
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label29 
            Caption         =   "Nro Serie y Documento:"
            Height          =   285
            Left            =   120
            TabIndex        =   167
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label30 
            Caption         =   "Fila(X)      Columna(Y)     VALOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1920
            TabIndex        =   166
            Top             =   150
            Width           =   3255
         End
         Begin VB.Label Label36 
            Caption         =   "Estado:"
            Height          =   285
            Left            =   120
            TabIndex        =   165
            Top             =   750
            Width           =   1815
         End
         Begin VB.Label Label37 
            Caption         =   "Tipo:"
            Height          =   285
            Left            =   120
            TabIndex        =   164
            Top             =   1110
            Width           =   1815
         End
         Begin VB.Label Label38 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   120
            TabIndex        =   163
            Top             =   1470
            Width           =   1815
         End
         Begin VB.Label lblFechaDoc 
            Caption         =   "Fecha Boleta:"
            Height          =   285
            Left            =   120
            TabIndex        =   162
            Top             =   1800
            Width           =   1845
         End
         Begin VB.Label Label40 
            Caption         =   "Servicio Hospital:"
            Height          =   285
            Left            =   120
            TabIndex        =   161
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label41 
            Caption         =   "Observaciones:"
            Height          =   285
            Left            =   120
            TabIndex        =   160
            Top             =   2520
            Width           =   1815
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Datos del Detalle"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   122
         Top             =   6600
         Width           =   5265
         Begin VB.TextBox txtProductoAncho 
            Height          =   345
            Left            =   1200
            TabIndex        =   173
            Top             =   720
            Width           =   675
         End
         Begin VB.TextBox txtCodigoY 
            Height          =   345
            Left            =   2910
            TabIndex        =   132
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox txtCodigoValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   131
            Text            =   "12345"
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txtProductoValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   130
            Text            =   "Cpt/Medicamento de prueba"
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtProductoY 
            Height          =   345
            Left            =   2910
            TabIndex        =   129
            Top             =   720
            Width           =   885
         End
         Begin VB.TextBox txtCantidadValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   128
            Text            =   "2"
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txtCantidadY 
            Height          =   345
            Left            =   2910
            TabIndex        =   127
            Top             =   1080
            Width           =   885
         End
         Begin VB.TextBox txtPrecioValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   126
            Text            =   "5"
            Top             =   1440
            Width           =   1365
         End
         Begin VB.TextBox txtPrecioY 
            Height          =   345
            Left            =   2910
            TabIndex        =   125
            Top             =   1440
            Width           =   885
         End
         Begin VB.TextBox txtImporteValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   124
            Text            =   "10"
            Top             =   1800
            Width           =   1365
         End
         Begin VB.TextBox txtImporteY 
            Height          =   345
            Left            =   2910
            TabIndex        =   123
            Top             =   1800
            Width           =   885
         End
         Begin VB.Label Label34 
            Caption         =   "Ancho     Fila(X)     Columna(Y)    VALOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1200
            TabIndex        =   174
            Top             =   150
            Width           =   3975
         End
         Begin VB.Label Label31 
            Caption         =   "Codigo:"
            Height          =   285
            Left            =   150
            TabIndex        =   137
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label42 
            Caption         =   "Producto:"
            Height          =   285
            Left            =   150
            TabIndex        =   136
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label43 
            Caption         =   "Cantidad:"
            Height          =   285
            Left            =   150
            TabIndex        =   135
            Top             =   1140
            Width           =   1815
         End
         Begin VB.Label Label44 
            Caption         =   "Precio:"
            Height          =   285
            Left            =   150
            TabIndex        =   134
            Top             =   1500
            Width           =   1815
         End
         Begin VB.Label Label45 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   150
            TabIndex        =   133
            Top             =   1860
            Width           =   1815
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Datos del Pie de Página"
         Height          =   4095
         Left            =   -69510
         TabIndex        =   88
         Top             =   2490
         Width           =   6735
         Begin VB.TextBox txtIGV 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   186
            Text            =   "1.62"
            Top             =   3630
            Width           =   1365
         End
         Begin VB.TextBox txtIGVX 
            Height          =   345
            Left            =   2010
            TabIndex        =   185
            Top             =   3630
            Width           =   855
         End
         Begin VB.TextBox txtIGVY 
            Height          =   345
            Left            =   2910
            TabIndex        =   184
            Top             =   3630
            Width           =   885
         End
         Begin VB.TextBox txtSubTotalY 
            Height          =   345
            Left            =   2910
            TabIndex        =   182
            Top             =   3270
            Width           =   885
         End
         Begin VB.TextBox txtSubTotalX 
            Height          =   345
            Left            =   2010
            TabIndex        =   181
            Top             =   3270
            Width           =   855
         End
         Begin VB.TextBox txtSubTotal 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   180
            Text            =   "9"
            Top             =   3270
            Width           =   1365
         End
         Begin VB.TextBox txtTotalLetrasAncho 
            Height          =   345
            Left            =   5250
            TabIndex        =   175
            ToolTipText     =   "Ancho del texto"
            Top             =   2520
            Width           =   885
         End
         Begin VB.TextBox txtTotalY 
            Height          =   345
            Left            =   2910
            TabIndex        =   112
            Top             =   2910
            Width           =   885
         End
         Begin VB.TextBox txtTotalX 
            Height          =   345
            Left            =   2010
            TabIndex        =   111
            Top             =   2910
            Width           =   855
         End
         Begin VB.TextBox txtValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   110
            Text            =   "9"
            Top             =   2880
            Width           =   1365
         End
         Begin VB.TextBox txtCajeroValor 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   109
            Text            =   "Debb"
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txtCajeroX 
            Height          =   345
            Left            =   2010
            TabIndex        =   108
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCajeroY 
            Height          =   345
            Left            =   2910
            TabIndex        =   107
            Top             =   360
            Width           =   885
         End
         Begin VB.TextBox txtCajaVAlor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   106
            Text            =   "Caja Principal"
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtCajaX 
            Height          =   345
            Left            =   2010
            TabIndex        =   105
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtCajaY 
            Height          =   345
            Left            =   2910
            TabIndex        =   104
            Top             =   720
            Width           =   885
         End
         Begin VB.TextBox txtAdelantosValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   103
            Text            =   "0"
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txtAdelantosX 
            Height          =   345
            Left            =   2010
            TabIndex        =   102
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtAdelantosY 
            Height          =   345
            Left            =   2910
            TabIndex        =   101
            Top             =   1080
            Width           =   885
         End
         Begin VB.TextBox txtTotalPagarVAlor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   100
            Text            =   "10"
            Top             =   1440
            Width           =   1365
         End
         Begin VB.TextBox txtTotalPagarX 
            Height          =   345
            Left            =   2010
            TabIndex        =   99
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtTotalPagarY 
            Height          =   345
            Left            =   2910
            TabIndex        =   98
            Top             =   1440
            Width           =   885
         End
         Begin VB.TextBox txtCuentaValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   97
            Text            =   "1234567"
            Top             =   1800
            Width           =   1365
         End
         Begin VB.TextBox txtCuentaX 
            Height          =   345
            Left            =   2010
            TabIndex        =   96
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txtCuentaY 
            Height          =   345
            Left            =   2910
            TabIndex        =   95
            Top             =   1800
            Width           =   885
         End
         Begin VB.TextBox txtExoneracionesVAlor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   94
            Text            =   "1"
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox txtExoneracionesX 
            Height          =   345
            Left            =   2010
            TabIndex        =   93
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtExoneracionesY 
            Height          =   345
            Left            =   2910
            TabIndex        =   92
            Top             =   2160
            Width           =   885
         End
         Begin VB.TextBox txtTotalEnLetrasValor 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   91
            Text            =   "Son 9 con 00/100 nuevos soles"
            Top             =   2520
            Width           =   1365
         End
         Begin VB.TextBox txtTotalEnLetrasX 
            Height          =   345
            Left            =   2010
            TabIndex        =   90
            Top             =   2550
            Width           =   855
         End
         Begin VB.TextBox txtTotalEnLetrasY 
            Height          =   345
            Left            =   2910
            TabIndex        =   89
            Top             =   2550
            Width           =   885
         End
         Begin VB.Label lblSubTotal 
            Caption         =   " SubTotal por Pagar:                                                  IGV:"
            Height          =   645
            Left            =   120
            TabIndex        =   183
            Top             =   3330
            Width           =   1815
         End
         Begin VB.Label lblTotalDoc 
            Caption         =   "Total Boleta:"
            Height          =   285
            Left            =   150
            TabIndex        =   121
            Top             =   2970
            Width           =   1815
         End
         Begin VB.Label Label35 
            Caption         =   "Fila(X)     Columna(Y)      VALOR       Ancho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2040
            TabIndex        =   120
            Top             =   180
            Width           =   4605
         End
         Begin VB.Label Label46 
            Caption         =   "Cajero:"
            Height          =   285
            Left            =   150
            TabIndex        =   119
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label lblCaja 
            Caption         =   "Caja:"
            Height          =   285
            Left            =   150
            TabIndex        =   118
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label48 
            Caption         =   "Adelantos:"
            Height          =   285
            Left            =   150
            TabIndex        =   117
            Top             =   1140
            Width           =   1815
         End
         Begin VB.Label Label49 
            Caption         =   "Total por Pagar:"
            Height          =   285
            Left            =   150
            TabIndex        =   116
            Top             =   1500
            Width           =   1815
         End
         Begin VB.Label Label50 
            Caption         =   "Cuenta de Atencion:"
            Height          =   285
            Left            =   150
            TabIndex        =   115
            Top             =   1860
            Width           =   1815
         End
         Begin VB.Label Label51 
            Caption         =   "Exoneraciones:"
            Height          =   285
            Left            =   150
            TabIndex        =   114
            Top             =   2220
            Width           =   1815
         End
         Begin VB.Label lblTotalDocLetras 
            Caption         =   "Total Boleta (letras):"
            Height          =   285
            Left            =   150
            TabIndex        =   113
            Top             =   2610
            Width           =   1815
         End
      End
      Begin VB.Frame frDocumentos 
         Caption         =   "Tipo de comprobante"
         ForeColor       =   &H000000FF&
         Height          =   880
         Left            =   -74940
         TabIndex        =   82
         Top             =   360
         Width           =   12165
         Begin VB.TextBox txtPasos 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   3330
            MultiLine       =   -1  'True
            TabIndex        =   83
            Text            =   "mPuntoCarga.frx":0090
            Top             =   210
            Width           =   8685
         End
         Begin Threed.SSOption optBoleta 
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   280
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   344
            _Version        =   262144
            Caption         =   "BOLETA"
            Value           =   -1
         End
         Begin Threed.SSOption optFactura 
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   580
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   344
            _Version        =   262144
            Caption         =   "FACTURA"
         End
         Begin Threed.SSOption optRecibo 
            Height          =   195
            Left            =   1530
            TabIndex        =   86
            Top             =   280
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   344
            _Version        =   262144
            Caption         =   "RECIBO"
         End
         Begin Threed.SSOption optTicket 
            Height          =   195
            Left            =   1530
            TabIndex        =   87
            Top             =   580
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   344
            _Version        =   262144
            Caption         =   "TICKET"
         End
      End
      Begin VB.Frame Frame38 
         BackColor       =   &H8000000D&
         Caption         =   "Llena tabla UPServicios"
         ForeColor       =   &H000000FF&
         Height          =   1185
         Left            =   -74850
         TabIndex        =   79
         Top             =   6030
         Width           =   6165
         Begin VB.CommandButton Command8 
            Caption         =   "Actualiza UPS en GalenHos desde el HIS"
            Height          =   885
            Left            =   3330
            TabIndex        =   80
            Top             =   210
            Width           =   2655
         End
         Begin VB.Label Label68 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS                                                                   tabla: TABLAUPS.DBF"
            Height          =   615
            Left            =   120
            TabIndex        =   81
            Top             =   330
            Width           =   2595
         End
      End
      Begin VB.Frame Frame37 
         Caption         =   "Agrega LAB desde el HIS"
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   -68520
         TabIndex        =   75
         Top             =   3690
         Width           =   6315
         Begin VB.CommandButton cmdAgregaLAB 
            Caption         =   "Agrega nuevos LAB"
            Height          =   915
            Left            =   4200
            TabIndex        =   76
            Top             =   210
            Width           =   2055
         End
         Begin VB.Label Label67 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Situacio.dbf  debe estar en la carpeta de ODBC HIS"
            Height          =   525
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label60 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame32 
         BackColor       =   &H8000000D&
         Caption         =   "Llena datos simulados en tabla: hisa.dbf y laboratorio.dbf"
         ForeColor       =   &H000000FF&
         Height          =   1185
         Left            =   -68520
         TabIndex        =   72
         Top             =   2490
         Width           =   6165
         Begin VB.CommandButton Command6 
            Caption         =   "Llena datos simulados (laboratorio, Materno, niño Sano)"
            Enabled         =   0   'False
            Height          =   735
            Left            =   4290
            TabIndex        =   73
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label66 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   74
            Top             =   330
            Width           =   2565
         End
      End
      Begin VB.Frame Frame31 
         Caption         =   "Actualiza Cant y Prec en: farmVentasDetalle, FacturacionBienesFinanciamiento"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -68490
         TabIndex        =   64
         Top             =   2490
         Width           =   6105
         Begin VB.CommandButton cmdActTablasFacturacion 
            Caption         =   "Procesar en base a tabla farmMovimientoDetalle (solo SALIDAS con SEGURO)"
            Height          =   405
            Left            =   90
            TabIndex        =   66
            Top             =   1080
            Width           =   5895
         End
         Begin VB.TextBox txtCodigoItem 
            Height          =   345
            Left            =   1230
            TabIndex        =   65
            Top             =   330
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtFinicial 
            Height          =   345
            Left            =   1230
            TabIndex        =   67
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFfinal 
            Height          =   345
            Left            =   3540
            TabIndex        =   68
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   150
            TabIndex        =   71
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimientos"
            Height          =   195
            Left            =   150
            TabIndex        =   70
            Top             =   780
            Width           =   1020
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   195
            Left            =   3060
            TabIndex        =   69
            Top             =   780
            Width           =   390
         End
      End
      Begin VB.CommandButton cmdActEdades 
         Caption         =   "Actualiza EDAD calculada en cada atención"
         Height          =   375
         Left            =   -68550
         TabIndex        =   63
         Top             =   4980
         Width           =   5955
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H8000000D&
         Caption         =   "Actualiza check RECIEN NACIDO en tabla ATENCIONES"
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   90
         TabIndex        =   60
         Top             =   4350
         Width           =   6105
         Begin VB.CommandButton cmdRecienNacido 
            Caption         =   "Procesa informacion"
            Height          =   345
            Left            =   60
            TabIndex        =   61
            Top             =   300
            Width           =   5925
         End
      End
      Begin VB.CommandButton cmdActualizaNroAutomatico 
         Caption         =   "..."
         Height          =   255
         Left            =   150
         TabIndex        =   59
         ToolTipText     =   "Actualiza último N° Automático de HC"
         Top             =   8325
         Width           =   255
      End
      Begin VB.Frame Frame36 
         BackColor       =   &H8000000D&
         Caption         =   "Depura Puntos de Carga (32 y 38)"
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   6480
         TabIndex        =   57
         Top             =   3090
         Width           =   6105
         Begin VB.CommandButton cmdDepuraPtoCarga 
            Caption         =   "Procesa informacion"
            Height          =   345
            Left            =   60
            TabIndex        =   58
            Top             =   300
            Width           =   5925
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Lista Items que se despacharon por INTERVENCION SANITARIA"
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   -68460
         TabIndex        =   55
         Top             =   1740
         Width           =   6105
         Begin VB.CommandButton cmdListaItemsIS 
            Caption         =   "Procesa informacion (debe existir c:\excel.xls)"
            Height          =   285
            Left            =   90
            TabIndex        =   56
            Top             =   240
            Width           =   5925
         End
      End
      Begin VB.Frame Frame33 
         BackColor       =   &H8000000D&
         Caption         =   "Actualiza Id Vendedor en tabla Movimiento (solo  Preventas)"
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   -68490
         TabIndex        =   53
         Top             =   810
         Width           =   6165
         Begin VB.CommandButton btnActualizaIdVendedor 
            Caption         =   "Procesar"
            Height          =   435
            Left            =   90
            TabIndex        =   54
            Top             =   300
            Width           =   5895
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Lee Archivo Excel y graba datos en Galenhos (CIE10 'descripcion corta')"
         ForeColor       =   &H000000FF&
         Height          =   1965
         Left            =   -68520
         TabIndex        =   48
         Top             =   390
         Width           =   6165
         Begin VB.CommandButton cmdCie10DescripcionCorta 
            Caption         =   "Procesar"
            Height          =   405
            Left            =   210
            TabIndex        =   51
            Top             =   1470
            Width           =   5895
         End
         Begin VB.TextBox txtExcel2 
            Height          =   315
            Left            =   1320
            TabIndex        =   50
            Text            =   "c:\cie10.xls"
            Top             =   210
            Width           =   4725
         End
         Begin VB.TextBox Text5 
            Height          =   885
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   49
            Text            =   "mPuntoCarga.frx":0175
            Top             =   540
            Width           =   5835
         End
         Begin VB.Label Label17 
            Caption         =   "Archivo Excel:"
            Height          =   285
            Left            =   210
            TabIndex        =   52
            Top             =   270
            Width           =   1245
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Agrega DX desde el HIS (N digitos)"
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   -74880
         TabIndex        =   44
         Top             =   4800
         Width           =   6315
         Begin VB.CommandButton cmdActualizaCie10 
            Caption         =   "Actualiza 'descripcion corta',  'nuevos CIE10' , codigoCIE10sinPUNTOS"
            Height          =   915
            Left            =   4200
            TabIndex        =   45
            Top             =   210
            Width           =   2055
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cie.dbf y Codienf.dbf debe estar en la carpeta de ODBC HIS"
            Height          =   525
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   3975
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H8000000D&
         Caption         =   "Perinatal: llena CODIGOS desde el mismo GalenHos"
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   -68550
         TabIndex        =   42
         Top             =   5460
         Width           =   6285
         Begin VB.CommandButton cmdLlenaCodigoHIS 
            Caption         =   "Cpt y Cie10 en tabla de CATALOGOS PERINATALES (columna 'CodigoHIS')"
            Height          =   435
            Left            =   120
            TabIndex        =   43
            Top             =   270
            Width           =   5985
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Migra ultimos PAISES desde el HIS"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -74880
         TabIndex        =   38
         Top             =   1500
         Width           =   6315
         Begin VB.CommandButton cmdActualizaAbrevPaises 
            Caption         =   "Actualiza ABREVIATURAS DE PAISES"
            Enabled         =   0   'False
            Height          =   405
            Left            =   90
            TabIndex        =   39
            Top             =   1080
            Width           =   6105
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Luego de migrar todos las ABREVIATURAS, eliminar tabla 'HIS_paises'"
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   690
            Width           =   6075
         End
         Begin VB.Label Label23 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe importar al SQL la tabla 'paises.dbf' al SQL como 'HIS_paises'"
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   330
            Width           =   6075
         End
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H8000000D&
         Caption         =   "Migra ultimos ESTABLECIMIENTOS desde el HIS"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -74910
         TabIndex        =   34
         Top             =   3090
         Width           =   6315
         Begin VB.CommandButton cmdNuevosEstablecimientos 
            Caption         =   "Agrega NUEVOS ESTABLECIMIENTOS"
            Height          =   405
            Left            =   90
            TabIndex        =   35
            Top             =   1080
            Width           =   6105
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " la tabla 'establec.dbf' debe estar en ....\galenhos\archivos"
            Height          =   315
            Left            =   120
            TabIndex        =   37
            Top             =   330
            Width           =   6075
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe existir ODBC llamado HIS que apunte a ..\galenhos\archivos"
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   690
            Width           =   6075
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Actualiza columna de la tabla -->Diagnosticos.ClaseDxHis (AGREGA CPT HIS)"
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   -74880
         TabIndex        =   30
         Top             =   420
         Width           =   6315
         Begin VB.CommandButton cmdActualizaDxHis 
            Caption         =   "Actualiza columna ClaseDxHis de tabla Diagnosticos (ADD CPTS HIS)"
            Height          =   735
            Left            =   4200
            TabIndex        =   31
            Top             =   210
            Width           =   2055
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TabDiag.dbf, Cpt.dbf debe estar en la carpeta de ODBC HIS"
            Height          =   435
            Left            =   105
            TabIndex        =   33
            Top             =   525
            Width           =   3975
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS"
            Height          =   270
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Agrega MEDICAMENTOS desde SISMEDV2"
         ForeColor       =   &H000000FF&
         Height          =   4065
         Left            =   -74820
         TabIndex        =   29
         Top             =   3480
         Width           =   6315
         Begin VB.Frame Frame1 
            Caption         =   "Desde tabla ZIP (debe contener archivo: medicame.dbf)"
            Height          =   1950
            Index           =   2
            Left            =   150
            TabIndex        =   228
            Top             =   2010
            Width           =   5955
            Begin VB.CheckBox chkClaveDig 
               Caption         =   "ZIP con clave DIGEMID"
               Height          =   270
               Left            =   180
               TabIndex        =   235
               Top             =   1005
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.TextBox txtRutaGalenhos 
               Height          =   345
               Left            =   1950
               TabIndex        =   233
               Text            =   "C:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
               Top             =   255
               Width           =   3825
            End
            Begin VB.TextBox txtZipArchivo 
               Height          =   345
               Left            =   480
               TabIndex        =   231
               Text            =   "c:\CAT_18062014_125119.zip"
               Top             =   630
               Width           =   5280
            End
            Begin VB.CommandButton MedicamentosDesdeZIP 
               Caption         =   "Agrega Medicamentos que faltan en GalenHos (actualiza datos si son NULL p' los items que ya existen)"
               Height          =   405
               Left            =   105
               TabIndex        =   229
               Top             =   1485
               Width           =   5655
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ZIP"
               Height          =   195
               Left            =   150
               TabIndex        =   232
               Top             =   705
               Width           =   255
            End
            Begin VB.Label Label56 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Debe tener ODBC: HIS"
               Height          =   315
               Left            =   150
               TabIndex        =   230
               Top             =   270
               Width           =   1755
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Desde tabla SISMED"
            Height          =   1635
            Index           =   1
            Left            =   165
            TabIndex        =   224
            Top             =   300
            Width           =   5955
            Begin VB.CommandButton cmdAgregaMedicamentos 
               Caption         =   "Agrega Medicamentos que faltan en GalenHos (actualiza datos si son NULL p' los items que ya existen)"
               Height          =   405
               Left            =   120
               TabIndex        =   225
               Top             =   990
               Width           =   5655
            End
            Begin VB.Label Label20 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "mMedicam.DBF"
               Height          =   315
               Left            =   150
               TabIndex        =   227
               Top             =   600
               Width           =   5655
            End
            Begin VB.Label Label21 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Debe tener ODBC: HIS"
               Height          =   315
               Left            =   135
               TabIndex        =   226
               Top             =   225
               Width           =   5655
            End
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Agrega ESTABLECIMIENTOS desde SISMEDV2"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -74850
         TabIndex        =   25
         Top             =   840
         Width           =   6315
         Begin VB.CommandButton cmdEstablecimientosNew 
            Caption         =   "Agrega Establecimientos que faltan en GalenHos"
            Height          =   405
            Left            =   90
            TabIndex        =   26
            Top             =   1080
            Width           =   6105
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "mAlmacen.DBF, mDistrito.DBF, mProvinC.dbf"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   690
            Width           =   6075
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC: HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   330
            Width           =   6075
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H8000000D&
         Caption         =   "Saldos x Estratégicos/Ventas/Donaciones"
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   -74850
         TabIndex        =   23
         Top             =   2580
         Width           =   6285
         Begin VB.CommandButton cmdEstratVtas 
            Caption         =   "Pone idTipoSalidaBienInsumo=4 para Almacen-Donaciones, Farmacia-Donaciones"
            Height          =   435
            Left            =   120
            TabIndex        =   24
            Top             =   300
            Width           =   5985
         End
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   $"mPuntoCarga.frx":01F7
         Height          =   555
         Left            =   120
         TabIndex        =   22
         Top             =   780
         Width           =   6285
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabla OPCs  vs Procedimientos_CPT"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   90
         TabIndex        =   18
         Top             =   2730
         Width           =   6315
         Begin VB.CommandButton cmdActualizaOPCs 
            Caption         =   "Actualiza Tabla OPCs y columna FactCatalogoServicios.idOpcs"
            Height          =   405
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   6105
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OPCs.dbf debe estar en la carpeta de ODBC HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   690
            Width           =   6075
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Top             =   330
            Width           =   6075
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000D&
         Caption         =   "Cambia el Nro de Historia Clnica"
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   6285
         Begin VB.CommandButton Command1 
            Caption         =   "Cambia el N° Historia ACTUAL por la NUEVA (que no existe aun)"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   780
            Width           =   5985
         End
         Begin VB.TextBox txtNroHistoriaActual 
            Height          =   375
            Left            =   1410
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox TxtNroHistoriaNew 
            Height          =   375
            Left            =   4950
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label6 
            Caption         =   "N° Historia Actual"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   17
            Top             =   270
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "N° Historia Nueva"
            Height          =   285
            Left            =   3480
            TabIndex        =   16
            Top             =   270
            Width           =   1485
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Pasa los movimientos de una Historia Clínica a otra"
         ForeColor       =   &H000000FF&
         Height          =   2115
         Left            =   6450
         TabIndex        =   2
         Top             =   720
         Width           =   6105
         Begin VB.TextBox txtHcNueva 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Text            =   "0"
            Top             =   1380
            Width           =   795
         End
         Begin VB.CommandButton cmdPasaMovimientos 
            Caption         =   "Procesar "
            Height          =   375
            Left            =   90
            TabIndex        =   7
            Top             =   1710
            Width           =   5925
         End
         Begin VB.Frame Frame7 
            Height          =   1035
            Left            =   1590
            TabIndex        =   3
            Top             =   210
            Width           =   4485
            Begin VB.TextBox txtHcActual 
               Height          =   375
               Left            =   90
               TabIndex        =   5
               Text            =   "0"
               Top             =   570
               Width           =   795
            End
            Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
               Left            =   90
               TabIndex        =   4
               Top             =   180
               Width           =   4215
            End
            Begin VB.Label lblActual 
               Caption         =   "El N° HISTORIA quedará libre para usarlo en otro Paciente"
               Height          =   375
               Left            =   930
               TabIndex        =   6
               Top             =   600
               Width           =   3405
            End
         End
         Begin VB.Label Label10 
            Caption         =   "N° Historia Destino"
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "N° Historia Origen"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblNueva 
            Caption         =   "Es el que queda en el SISTEMA y ARCHIVO "
            Height          =   285
            Left            =   2550
            TabIndex        =   9
            Top             =   1410
            Width           =   3405
         End
      End
      Begin UltraGrid.SSUltraGrid grdErrores 
         Height          =   225
         Left            =   6030
         TabIndex        =   1
         Top             =   465
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   397
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "SSUltraGrid1"
      End
   End
End
Attribute VB_Name = "mVariosProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Varios Procesos
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Dim lnIdPacienteActual As Long
Dim lnIdPacienteNuevo As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New ListaDespleglable
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsGrid As New Recordset
Dim lcSql As String
Dim lbProcesaVAriosDBF As Boolean
Dim lcTipoComprobanteCaja As String
Dim lcTipoServicioFarmacia As String
Const lcTicket As String = "Ticket"
Const lcFactura As String = "Factura"

Private Sub btnActualizaIdVendedor_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
  
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
  
        Me.MousePointer = 11
                
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FarmPreVentaFarmMovimientoVentasSeleccionarPorIdEstadoPreventa2"
          Set oRsTmp1 = .Execute
          Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        
        If oRsTmp1.RecordCount > 0 Then
          ProgressBar1.Min = 0: ProgressBar1.Max = oRsTmp1.RecordCount + 2
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              DoEvents:            ProgressBar1.Value = ProgressBar1.Value + 1:          Me.Refresh
                
                
                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "FarmMovimientoActualizarIdUsuarioPorMovTipoS"
                  Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oRsTmp1.Fields!idVendedor): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oRsTmp1.Fields!movNumero): .Parameters.Append oParameter
                  Set oRsTmp2 = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
                oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        Me.MousePointer = 1
        oConexion.Close
        Set oConexion = Nothing
        Unload Me
    End If
End Sub


Private Sub btnGuardarConfiguracionComprobante_Click()
    If Not verifyPath(txtRutaINI.Text) Then
        MsgBox "Ruta Especificada para Archivo de Configuración No Existe", vbExclamation
        Exit Sub
    End If
    
    'mgaray
    If optTicket.Value = False Then
        If Not validarAltoDeSeccionesBoleta("¿Desea Guardar Los Valores?") Then
            Exit Sub
        End If
        If Not validarAnchoDeSeccionesBoleta("¿Desea Guardar Los Valores?") Then
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Desea Grabar los Cambios de la Configuración de Impresión de Comprobantes de Pago", vbQuestion + vbYesNo, "") = vbYes Then
        asignarValorDeControlesAVariables
        grabarSetup_Caja txtRutaINI.Text, lcTipoServicioFarmacia, wxIdTipoComprobanteDefault
    End If

End Sub

Private Sub cmdActEdades_Click()
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Dim oRsServ As New Recordset
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesActualizarEdadPorTipoEdad"
        Set oRsServ = .Execute
    End With
    
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Unload Me
End Sub

Private Sub cmdActPreciosFarm_Click()
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    

    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       Dim wxConexionRed As New Connection
       wxConexionRed.CommandTimeout = 900
       wxConexionRed.CursorLocation = adUseClient
       wxConexionRed.Open sighentidades.CadenaConexion

       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wrs_GalenHos3 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
       Dim wrsGalenHosTemp As New ADODB.Recordset
       Dim wrs_LolCli As New ADODB.Recordset
       Dim lcFechaNac As String: Dim lnTipoSexo As Long
       Dim lcPrimerNombre As String
       Dim lcSegundoNombre As String
       Dim lnNroHistoriaClinica As Long
       Dim lcSql As String
       Dim lntotReg As Long
       Dim lnRegAct As Long
       Dim lnIdPaciente As Long
       Dim lcAutogenerado As String
       Dim lcFechaAnt As Date
       Dim lntipoOcupacion As Long
       Dim lnIdDepartamentoDomicilio As Long
       Dim LnIdProvinciaDomicilio As Long
       Dim lnIdDistritoDomicilio As Long
       Dim lnIdDepartamentoNacimiento As Long
       Dim LnIdProvinciaNacimiento As Long
       Dim lnIdDistritoNacimiento As Long
       Dim lnIdEstadoCivil  As Long
       Dim lbNuevoHC As Boolean
       Dim lcApellidoPaterno As String, lcApellidoMaterno As String
       Dim lbContinuarProceso As Boolean, lnFor As Integer, lnCantGuiones As Integer, lnHistoriaAutogenerada As Long
       Dim wFec1 As String, wFec2 As String, wFFecha As Date, lnTipoNumeracion As Long
       
       lcSql = "select * from tipoSFinanciamIento where idTipoFinanciamiento= " & Me.txtNewFF.Text
       wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       If wrs_GalenHos.RecordCount = 0 Then
          MsgBox "No existe ese ID de TIPO FINANCIAMIENTO"
       Else
          
          lcSql = "select * from factCatalogoBienesInsumosHosp where idTipoFinanciamiento=1"
          wrs_GalenHos2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
          If wrs_GalenHos2.RecordCount > 0 Then
             prog2.Text = wrs_GalenHos2.RecordCount
             lnFor = 1
             wrs_GalenHos2.MoveFirst
             Do While Not wrs_GalenHos2.EOF
                lcSql = "select * from factCatalogoBienesInsumosHosp where idTipoFinanciamiento=" & Me.txtNewFF.Text & _
                        " and idProducto= " & wrs_GalenHos2.Fields!idProducto
                wrs_GalenHos3.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                If wrs_GalenHos3.RecordCount = 0 Then
                   wrs_GalenHos3.AddNew
                   wrs_GalenHos3.Fields!IdTipoFinanciamiento = Val(Me.txtNewFF.Text)
                   wrs_GalenHos3.Fields!idProducto = wrs_GalenHos2.Fields!idProducto
                   wrs_GalenHos3.Fields!Activo = wrs_GalenHos2.Fields!Activo
                End If
                 wrs_GalenHos3.Fields!PrecioUnitario = wrs_GalenHos2.Fields!PrecioUnitario
                wrs_GalenHos3.Update
                wrs_GalenHos3.Close
                wrs_GalenHos2.MoveNext
             Loop
          End If
       End If
    End If
    Unload Me
    Exit Sub
err_proceso:
    MsgBox Err.Description

    Unload Me
    Resume

End Sub

Private Sub cmdActTablasFacturacion_Click()

    If Val(txtCodigoItem) <= 0 Then
       MsgBox "Ingrese el CODIGO SISMED a procesar"
       Exit Sub
    End If
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       'On Error GoTo errActTab
       Dim oRsServ As New Recordset
       Dim oRsServ1 As New Recordset
       Dim lcSql As String
       Dim lnIdProducto As Long
       Dim lcMovNumero As String
       Dim lnPrecio As Double
       Dim lnCantidad As Long
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
  
       
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = wxConexionRed
            .CommandTimeout = 150
            .CommandText = "FarmMovimientoDetalleFactCatalogoBienesInsumosSeleccionarPorCodigo"
            Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 7, Trim(Me.txtCodigoItem.Text)): .Parameters.Append oParameter
            Set oRsServ = .Execute
            Set oRsServ.ActiveConnection = Nothing
       End With
       Set oCommand = Nothing
       Set oParameter = Nothing
       
       If oRsServ.RecordCount > 0 Then
          lnIdProducto = oRsServ.Fields!idProducto
          oRsServ.MoveFirst
          Do While Not oRsServ.EOF
             If oRsServ.Fields!FechaCreacion >= CDate(Me.txtFinicial.Text) And oRsServ.Fields!FechaCreacion <= CDate(Me.txtFfinal.Text) Then
                lcMovNumero = oRsServ.Fields!movNumero
                lnPrecio = oRsServ.Fields!precio
                lnCantidad = 0
                Do While Not oRsServ.EOF And lcMovNumero = oRsServ.Fields!movNumero
                   lnCantidad = lnCantidad + oRsServ.Fields!cantidad
                   oRsServ.MoveNext
                   If oRsServ.EOF Then
                      Exit Do
                   End If
                Loop
               
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "FarmMovimientoVentasDetalleActualizaCantidadPrecioTotal"
                    Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, FormatCurrency(lnPrecio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, Round(lnCantidad * lnPrecio, 2)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                    Set oRsServ1 = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
                
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "FacturacionBienesFinanciamientosActualizaCantidadPrecioTotal"
                    Set oParameter = .CreateParameter("@CantidadFinanciada", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@PrecioFinanciado", adCurrency, adParamInput, 0, FormatCurrency(lnPrecio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@TotalFinanciado", adCurrency, adParamInput, 0, Round(lnCantidad * lnPrecio, 2)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                    Set oRsServ1 = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
  
             Else
                 oRsServ.MoveNext
             End If
          Loop
       End If
       Unload Me
    End If

End Sub

Private Sub cmdActualiza_Click()

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcHE As String
    On Error GoTo ErrAct
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
         Dim oRsServ As New Recordset
         Dim oRsPtoCarga As New Recordset
         Dim oRsPtoCarga1 As New Recordset

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = wxConexionRed
            .CommandTimeout = 150
            .CommandText = "ServiciosSeleccionarIdTipoServicio2y3"
            Set oRsServ = .Execute
            Set oRsServ.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
         
        If oRsServ.RecordCount > 0 Then
            oRsServ.MoveFirst
            Do While Not oRsServ.EOF
               lcHE = IIf(oRsServ.Fields!IdTipoServicio = 3, " (H)", IIf(oRsServ.Fields!IdTipoServicio = 1, " (CE)", " (E)"))
               
               With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "FactPuntosCargaAgregarDesdeServiciosIdTipoServicio2y3" 'Agregar o Consulta segun encuentre el IdServicio
                    Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, oRsServ.Fields!IdServicio + 500): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Descripcion", adChar, adParamInput, 50, Left(Trim(oRsServ.Fields!nombre) & lcHE, 50)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdUPS", adInteger, adParamInput, 0, 5): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oRsServ.Fields!IdServicio): .Parameters.Append oParameter
                    .Execute
               End With
               Set oCommand = Nothing
               Set oParameter = Nothing
               
               oRsServ.MoveNext
            Loop
            
         End If
         
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = wxConexionRed
          .CommandTimeout = 150
          .CommandText = "FactPuntosCargaSeleccionarTodos"
          Set oRsPtoCarga = .Execute
          Set oRsPtoCarga.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
         
         'Elimina Puntos de Carga, cuyos IdServicio no se encuentra en tabla SERVICIOS
         oRsServ.Close
         oRsPtoCarga.MoveFirst
         Do While Not oRsPtoCarga.EOF
            If oRsPtoCarga.Fields!IdServicio > 0 Then
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "ServiciosSeleccionarPorId"
                    Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oRsPtoCarga.Fields!IdServicio): .Parameters.Append oParameter
                    Set oRsServ = .Execute
                    Set oRsServ.ActiveConnection = Nothing
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
               
               If oRsServ.RecordCount = 0 Then


                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosPtosEliminaXidPuntoCarga"
                        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsPtoCarga.Fields!idPuntoCarga): .Parameters.Append oParameter
                        Set oRsPtoCarga1 = .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    

                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "FactPuntosCargaEliminar"
                        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, oRsPtoCarga.Fields!idPuntoCarga): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                  
               End If
               oRsServ.Close
            End If
            oRsPtoCarga.MoveNext
         Loop
         Unload Me
    End If
    Exit Sub
ErrAct:
   MsgBox Err.Description
End Sub

Private Sub cmdActualizaAbrevPaises_Click()

       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim oConexion As New ADODB.Connection
  
       On Error Resume Next
       
        With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "PaisesSeleccionarTodosCampos"
              Set oRsTmp1 = .Execute
              Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing

       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             lcSql = "select * from his_paises where rtrim(upper(pais))='" & Trim(UCase(oRsTmp1.Fields!nombre)) & "'"
             oRsTmp2.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
             If oRsTmp2.RecordCount > 0 Then

                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = oConexion
                    .CommandTimeout = 150
                    .CommandText = "PaisesModificarCodigo"
                    Set oParameter = .CreateParameter("@IdPais", adInteger, adParamInput, 0, oRsTmp1.Fields!IdPais): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 3, oRsTmp2.Fields!cod): .Parameters.Append oParameter
                    .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
  
             End If
             oRsTmp2.Close
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       Unload Me
End Sub

'mgaray09
Private Function getAditionalDataCIEFromDBF(sCodCat As String, sCodEnf As String, oConexionFox As Connection) As ADODB.Recordset
    Dim oRsTmp As New Recordset
    lcSql = "select * from Codienf WHERE cod_Cat = '" & sCodCat & "' and cod_enf = '" & sCodEnf & "'"
    oRsTmp.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
    Set getAditionalDataCIEFromDBF = oRsTmp
End Function

Private Function getEdadEnDias(iEdad As Integer, cTipo As String) As Long
    Dim iEdadEnDias As Long
    cTipo = UCase(cTipo)
    Select Case cTipo
        Case "D":
            iEdadEnDias = CLng(iEdad)
        Case "M":
            iEdadEnDias = CLng(iEdad) * 30
        Case "A":
            iEdadEnDias = CLng(iEdad) * 365
        Case Else:
            iEdadEnDias = 0
    End Select
    getEdadEnDias = iEdadEnDias
End Function

Private Sub cmdActualizaCie10_Click()

        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        'mgaray09
        Dim oRsDetallesCie As New ADODB.Recordset
        Dim cCodCat As String, cCodEnf As String
        Dim bEsActivo As Boolean
        Dim lIdTipoSexo As Integer
        Dim dFechaInicioVigencia As Date
        Dim lEdadMaxDias As Long, lEdadMinDias As Long
        Dim oDiagnostico As New DODiagnostico
        
        dFechaInicioVigencia = getFechaInicioVigenciaParaNuevosRegistros()
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        lcSql = "select * from cie"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        ProgressBar1.Max = oRsFox.RecordCount
        cmdActualizaCie10.Enabled = False
        If ProgressBar1.Max > 0 Then
           ProgressBar1.Min = 0
           ProgressBar1.Value = 0
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
               ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
               DoEvents
             ' If Len(Trim(oRsFox.Fields!cod_enf)) = 1 Then
                lcCodDx = Trim(oRsFox.Fields!cod_Cat) & "." & Trim(oRsFox.Fields!cod_enf)
                'mgaray09
                cCodCat = oRsFox.Fields!cod_Cat
                cCodEnf = oRsFox.Fields!cod_enf
                Set oRsDetallesCie = getAditionalDataCIEFromDBF(cCodCat, cCodEnf, oConexionFox)
                
                lEdadMinDias = 0
                lEdadMaxDias = 0
                lIdTipoSexo = 0
                
                If oRsDetallesCie.RecordCount > o Then
                
                    lIdTipoSexo = getIdTipoSexo(oRsDetallesCie.Fields!sexo)
                    lEdadMinDias = getEdadEnDias(oRsDetallesCie.Fields!Min_edad, oRsDetallesCie.Fields!Min_tipo)
                    lEdadMaxDias = getEdadEnDias(oRsDetallesCie.Fields!Max_edad, oRsDetallesCie.Fields!Max_tipo)
                End If
                
                oDiagnostico.EdadMaxDias = lEdadMaxDias
                oDiagnostico.EdadMinDias = lEdadMinDias
                oDiagnostico.EsActivo = True
                If oRsDetallesCie.RecordCount > 0 Then
                    oDiagnostico.EsActivo = getIsActive(oRsDetallesCie.Fields!est)
                End If
                'oDiagnostico.EsActivo = IIf(oRsDetallesCie.RecordCount > 0, getIsActive(oRsDetallesCie.Fields!est), True)
                oDiagnostico.idTipoSexo = lIdTipoSexo
                oDiagnostico.Descripcion = Left(oRsFox.Fields!desc_enf, 250)
                oDiagnostico.CodigoCIE2004 = lcCodDx
                oDiagnostico.FechaInicioVigencia = dFechaInicioVigencia
                
                With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = wxConexionRed
                      .CommandTimeout = 150
                      .CommandText = "DiagnosticosSeleccionarTodoCamposPorCodigoCie2004"
                      Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 7, lcCodDx): .Parameters.Append oParameter
                      Set oRsTmp = .Execute
                      Set oRsTmp.ActiveConnection = Nothing
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
                If oRsTmp.RecordCount = 0 Then
                    With oCommand
                          .CommandType = adCmdStoredProc
                          Set .ActiveConnection = wxConexionRed
                          .CommandTimeout = 150
                          .CommandText = "DiagnosticosAgregarPorCodigoDescripcionDatosCompletos"
                          Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 7, lcCodDx): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, oDiagnostico.Descripcion): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(lEdadMaxDias = 0, Null, lEdadMaxDias)): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(lEdadMinDias = 0, Null, lEdadMinDias)): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(lIdTipoSexo = 0, Null, lIdTipoSexo)): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@FechaInicioVigencia", adDBTimeStamp, adParamInput, 0, dFechaInicioVigencia): .Parameters.Append oParameter 'Actualizado 23092014
                          Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, oDiagnostico.EsActivo): .Parameters.Append oParameter
                          .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                
                Else
                   If IsNull(oRsTmp.Fields!DescripcionMINSA) And (Not IsNull(oRsFox.Fields!desc_enf)) Then
                        With oCommand
                                .CommandType = adCmdStoredProc
                                Set .ActiveConnection = wxConexionRed
                                .CommandTimeout = 150
                                .CommandText = "DiagnosticosModificarDescripcionYOtroDatos"
                                Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oRsTmp.Fields!IdDiagnostico): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, oDiagnostico.Descripcion): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(lEdadMaxDias = 0, Null, lEdadMaxDias)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(lEdadMinDias = 0, Null, lEdadMinDias)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(lIdTipoSexo = 0, Null, lIdTipoSexo)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@FechaInicioVigencia", adDBTimeStamp, adParamInput, 0, dFechaInicioVigencia): .Parameters.Append oParameter 'Actualizado 23092014
                                Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, oDiagnostico.EsActivo): .Parameters.Append oParameter
                                .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                    
                   Else
                   End If
                   
                End If
                oRsTmp.Close
                'mgaray09
                'DiagnosticosDasactivarDescripcionesPasadas oRsFox, oDiagnostico
             ' End If
              oRsFox.MoveNext
           Loop
        End If
        'Actualiza Dx SIN PUNTO
        mo_AdminArchivoClinico.ActualizaCampoDxSINpunto wxConexionRed
'
        Me.MousePointer = 11
        oConexionFox.Close
        MsgBox "Actualizó correctamente....ahora PROCESAR:" & Chr(13) & Chr(13) & "Actualiza columna de la tabla -->Diagnosticos.ClaseDxHis"
        cmdActualizaCie10.Enabled = True
        Unload Me
End Sub




'mgaray09
Private Function getIsActive(cEst As String) As Boolean
    Dim bReturnValue As Boolean
    bReturnValue = False
    If UCase(cEst) = "A" Then
        bReturnValue = True
    End If
    getIsActive = bReturnValue
End Function

Private Function getIdTipoSexo(cSexo As String) As Integer
    Dim idTipoSexo As Integer
    
    If cSexo = "M" Then
        idTipoSexo = 1
    End If
    If cSexo = "F" Then
        idTipoSexo = 2
    End If
        
    getIdTipoSexo = idTipoSexo
End Function

Public Function getFechaInicioVigenciaParaNuevosRegistros() As Date
    getFechaInicioVigenciaParaNuevosRegistros = lcBuscaParametro.RetornaFechaServidorSQL
End Function

Private Sub cmdActualizaDxHis_Click()
        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        'mgaray09
        Dim lcSqlEstado As String
        Dim oRsEstado As New ADODB.Recordset
        
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        lcSql = "select * from tabdiag"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              If Len(Trim(oRsFox.Fields!diag)) = 4 And Val(oRsFox.Fields!clase) > 0 Then
                    lcCodDx = Left(oRsFox.Fields!diag, 3) & "." & Mid(oRsFox.Fields!diag, 4, 1)
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "DiagnosticosActualizaClaseDxHis"
                        Set oParameter = .CreateParameter("@claseDxHis", adVarChar, adParamInput, 1, oRsFox.Fields!clase): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@codigoCie2004", adVarChar, adParamInput, 7, lcCodDx): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, getIsActive(oRsFox.Fields!est)): .Parameters.Append oParameter
                        Set oRsTmp = .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
  
              ElseIf Len(Trim(oRsFox.Fields!diag)) = 5 And Val(oRsFox.Fields!clase) > 0 Then
                    lcCodDx = Left(oRsFox.Fields!diag, 3) & "." & Mid(oRsFox.Fields!diag, 4, 1)
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "DiagnosticosActualizaClaseDxHis"
                        Set oParameter = .CreateParameter("@claseDxHis", adVarChar, adParamInput, 1, oRsFox.Fields!clase): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@codigoCie2004", adVarChar, adParamInput, 7, lcCodDx): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, getIsActive(oRsFox.Fields!est)): .Parameters.Append oParameter
                        Set oRsTmp = .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
              End If
              'Añade CPT HIS
              If Val(oRsFox.Fields!clase) = 4 Then
              

                
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "FactCatalogoServiciosAgregarCodigoNombre"
                    Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, Left(Trim(oRsFox.Fields!diag), 20)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(oRsFox.Fields!descrip, 255)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oRsFox.Fields!est = "A", 1, 0)): .Parameters.Append oParameter
                    .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                'mgaray09
                'FactCatalogoServiciosDasactivarDescripcionesPasadas oRsFox, False
                 
                 
'                 oRsTmp.Close
              End If
              oRsFox.MoveNext
           Loop
        End If
        ActualizaCPTDesdeDBFRegiones oConexionFox 'Mario 16092014
        ActualizaCPTDesdeDBFCPT oConexionFox
        Me.MousePointer = 11
        oConexionFox.Close
        Unload Me

End Sub

'mgaray09
Private Sub ActualizaCPTDesdeDBFCPT(oConexionFox As Connection)
        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        
        Dim lcSql As String ', lcCodDx As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        Me.MousePointer = 1
        lcSql = "select * from Cpt"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              
              'Añade CPT HIS
              If Val(oRsFox.Fields!clase) = 4 Then
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "FactCatalogoServiciosAgregarCodigoNombre"
                    Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, Left(Trim(oRsFox.Fields!Cod_cpt), 20)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(oRsFox.Fields!desc_cpt, 255)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oRsFox.Fields!est = "A", 1, 0)): .Parameters.Append oParameter
                    
                    .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                 
                'FactCatalogoServiciosDasactivarDescripcionesPasadas oRsFox, True
'                 oRsTmp.Close
              End If
              oRsFox.MoveNext
           Loop
        End If
        Me.MousePointer = 11
End Sub

Private Function FactCatalogoServiciosDasactivarDescripcionesPasadas(oRsFox As Recordset, FromFileCpt As Boolean)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = wxConexionRed
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosDasactivarDescripcionesPasadas"
        If FromFileCpt = True Then
            Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, Left(Trim(oRsFox.Fields!Cod_cpt), 20)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(oRsFox.Fields!desc_cpt, 255)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oRsFox.Fields!est = "A", 1, 0)): .Parameters.Append oParameter
        Else
            Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, Left(Trim(oRsFox.Fields!diag), 20)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(oRsFox.Fields!descrip, 255)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oRsFox.Fields!est = "A", 1, 0)): .Parameters.Append oParameter
        End If
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Private Function DiagnosticosDasactivarDescripcionesPasadas(oRsFox As Recordset, oDiagnostico As DODiagnostico)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = wxConexionRed
        .CommandTimeout = 150
        .CommandText = "DiagnosticosDasactivarDescripcionesPasadas"
        
        Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 7, oDiagnostico.CodigoCIE2004): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, oDiagnostico.Descripcion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(oDiagnostico.EdadMaxDias = 0, Null, oDiagnostico.EdadMaxDias)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(oDiagnostico.EdadMinDias = 0, Null, oDiagnostico.EdadMinDias)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oDiagnostico.idTipoSexo = 0, Null, oDiagnostico.idTipoSexo)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicioVigencia", adDBTimeStamp, adParamInput, 0, oDiagnostico.FechaInicioVigencia): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, oDiagnostico.EsActivo): .Parameters.Append oParameter
    
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

Private Sub cmdActualizaNroAutomatico_Click()
    Me.MousePointer = 11
    mo_AdminArchivoClinico.ActualizaDatosConProblemas True
    Me.MousePointer = 1
    Unload Me
End Sub


Private Sub cmdActualizaOPCs_Click()
'MODIFICADO POR FRANKLIN CACHAY 07/11/2013 - se cambio a store procedure


        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 11
'        lcSql = "update FactCatalogoServicios set idOpcs=null"
'        oRsTmpCat.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
        
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = wxConexionRed
          .CommandTimeout = 150
          .CommandText = "FactCatalogoServiciosActualizarIdOpcsNull"
          Set oRsTmpCat = .Execute
        End With
        Set oCommand = Nothing
        
'        lcSql = "select * from Opcs"
'        oRsTmpOpc.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic

'        With oCommand
'            .CommandType = adCmdStoredProc
'            Set .ActiveConnection = wxConexionRed
'            .CommandTimeout = 150
'            .CommandText = "OpcsSeleccionarTodo"
'            Set oRsTmpOpc = .Execute
'            Set oRsTmpOpc.ActiveConnection = Nothing
'        End With
'        Set oCommand = Nothing
'        lnIdOpc = oRsTmpOpc.RecordCount + 1
        
        lcSql = "select * from Opcs order by opcs"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              lcOPC = Trim(oRsFox.Fields!opcs)
              
'MODIFICADO POR FRANKLIN CACHAY 11/11/2013 - se cambio a store procedure

'              'Actualiza tabla: OPCs
'              lbNuevo = True
'              If oRsTmpOpc.RecordCount > 0 Then
'                 oRsTmpOpc.MoveFirst
'                 lcSql = "codigo='" & Trim(oRsFox.Fields!opcs) & "'"
'                 oRsTmpOpc.Find lcSql
'                 If Not oRsTmpOpc.EOF Then
'                    lbNuevo = False
'                 End If
'              End If
'              If lbNuevo = True Then
'                 oRsTmpOpc.AddNew
'                 oRsTmpOpc.Fields!idOpcs = lnIdOpc
'                 oRsTmpOpc.Fields!Codigo = Trim(oRsFox.Fields!opcs)
'                 lnIdOpc = lnIdOpc + 1
'              End If
'              oRsTmpOpc.Fields!Descripcion = Trim(oRsFox.Fields!descripcio)
'              oRsTmpOpc.Update
              
               With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "OPCsActualizarDesdeOpcsBDFoxPorCodigo"
                    Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 7, Trim(oRsFox.Fields!opcs)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 255, Trim(oRsFox.Fields!descripcio)): .Parameters.Append oParameter
                    .Execute
               End With
               Set oCommand = Nothing
               Set oParameter = Nothing
               
               With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "OpcsSeleccionarTodo"
                    Set oRsTmpOpc = .Execute
                    Set oRsTmpOpc.ActiveConnection = Nothing
               End With
               Set oCommand = Nothing
               lcSql = "codigo='" & Trim(oRsFox.Fields!opcs) & "'"
               oRsTmpOpc.Find lcSql
'
              Do While Not oRsFox.EOF And lcOPC = Trim(oRsFox.Fields!opcs)
'                 lcSql = "update FactCatalogoServicios set idOpcs=" & oRsTmpOpc.Fields!idOpcs & " where codigo='" & Trim(oRsFox.Fields!cpt99) & "'"
'                 oRsTmpCat.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosActualiIaridOpcsPorCodigo"
                        Set oParameter = .CreateParameter("@IdOpcs", adInteger, adParamInput, 0, oRsTmpOpc.Fields!idOpcs): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, Trim(oRsFox.Fields!cpt99)): .Parameters.Append oParameter
                        Set oRsTmpCat = .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                 
                 oRsFox.MoveNext
                 If oRsFox.EOF Then
                    Exit Do
                 End If
              Loop
           Loop
        End If
        Me.MousePointer = 1
        oConexionFox.Close
        Unload Me
End Sub

Private Sub cmdAgregaLAB_Click()

        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        lcSql = "select * from situacio"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
                lcCodDx = Trim(oRsFox.Fields!valores)
                
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "His_situacioSeleccionarPorValores"
                    Set oParameter = .CreateParameter("@Valores", adVarChar, adParamInput, 3, lcCodDx): .Parameters.Append oParameter
                    Set oRsTmp = .Execute
                    Set oRsTmp.ActiveConnection = Nothing
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
                
                If oRsTmp.RecordCount = 0 Then
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "HIS_situacioAgregar"
                        Set oParameter = .CreateParameter("@IdHisSituacio", adInteger, adParamInput, 0, 0): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Valores", adChar, adParamInput, 3, lcCodDx): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@descripcio", adChar, adParamInput, 40, oRsFox!descripcio): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Codigo", adInteger, adParamInput, 0, oRsFox!Codigo): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@est", adVarChar, adParamInput, 3, oRsFox!est): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                Else
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "HIS_situacioModificar"
                        Set oParameter = .CreateParameter("@IdHisSituacio", adInteger, adParamInput, 0, oRsTmp!IdHisSituacio): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Valores", adChar, adParamInput, 3, Trim(oRsTmp.Fields!valores)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@descripcio", adChar, adParamInput, 40, Trim(oRsFox!descripcio)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Codigo", adInteger, adParamInput, 0, oRsFox!Codigo): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@est", adChar, adParamInput, 1, Trim(oRsFox!est)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                End If
                oRsTmp.Close
                oRsFox.MoveNext
           Loop
        End If
        oConexionFox.Close
        Set oRsTmp = Nothing
        Set oRsFox = Nothing
        Set oConexionFox = Nothing
        Unload Me

End Sub

Private Sub cmdAgregaMedicamentos_Click()
    On Error GoTo ErrActItems
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 11
        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim oCommand As New ADODB.Command
        Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos, oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim oParameter As ADODB.Parameter
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long
        Dim lcCodigo As String, lcTipoProductoSismed As String
        
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=HIS"
        '
        Me.MousePointer = 1
        'Medicamentos
        lcSql = "select * from mMedicam"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           Set oCatalogoBienesInsumos.Conexion = wxConexionRed
           oDOCatalogoBienesInsumos.IdUsuarioAuditoria = 0
           Me.ProgressBar1.Max = oRsFox.RecordCount
           Me.ProgressBar1.Min = 0
           Me.ProgressBar1.Value = 0
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
                Me.Refresh: ProgressBar1.Value = ProgressBar1.Value + 1: DoEvents
                lcCodigo = Trim(oRsFox.Fields!medCod)
                lcTipoProductoSismed = Trim(oRsFox!medEst)
                Set oRsTmpOpc = oCatalogoBienesInsumos.SeleccionarPorCodigo(lcCodigo, wxConexionRed)
                If oRsTmpOpc.RecordCount = 0 And Len(lcCodigo) > 4 Then
                   oDOCatalogoBienesInsumos.IdSubGrupoFarmacologico = 999
                   oDOCatalogoBienesInsumos.IdGrupoFarmacologico = 999
                   oDOCatalogoBienesInsumos.nombre = Left(Trim(oRsFox.Fields!medNom) & " " & Trim(oRsFox.Fields!medPres) & " " & _
                                                     Trim(oRsFox.Fields!medcnc) & " " & oRsFox.Fields!medFF, 300)
                   oDOCatalogoBienesInsumos.Codigo = lcCodigo
                   
                   If UCase(Trim(oRsFox.Fields!medEst)) = "E" Then
                      oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = IIf(oRsFox.Fields!medEstVta = 1, 3, 2)
                   Else
                      oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = 1
                   End If
                   oDOCatalogoBienesInsumos.TipoProducto = IIf(oRsFox.Fields!medTip = "M", 0, 1)
                   oDOCatalogoBienesInsumos.denominacion = oRsFox.Fields!medNom
                   oDOCatalogoBienesInsumos.Concentracion = oRsFox.Fields!medcnc
                   oDOCatalogoBienesInsumos.Presentacion = oRsFox.Fields!medPres
                   oDOCatalogoBienesInsumos.FormaFarmaceutica = oRsFox.Fields!medFF
                   oDOCatalogoBienesInsumos.TipoProductoSismed = lcTipoProductoSismed
                   oDOCatalogoBienesInsumos.Petitorio = IIf(oRsFox.Fields!MedPet = "P", 1, 0)
                   
                   If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) = False Then
                      MsgBox oCatalogoBienesInsumos.MensajeError
                      GoTo ErrActItems
                   End If
                Else
                   oDOCatalogoBienesInsumos.idProducto = oRsTmpOpc!idProducto
                   If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) = True Then
                   End If
                   oDOCatalogoBienesInsumos.denominacion = oRsFox.Fields!medNom
                   oDOCatalogoBienesInsumos.Concentracion = oRsFox.Fields!medcnc
                   oDOCatalogoBienesInsumos.Presentacion = oRsFox.Fields!medPres
                   oDOCatalogoBienesInsumos.FormaFarmaceutica = oRsFox.Fields!medFF
                   oDOCatalogoBienesInsumos.TipoProductoSismed = lcTipoProductoSismed
                   oDOCatalogoBienesInsumos.Petitorio = IIf(oRsFox.Fields!MedPet = "P", 1, 0)
                   If oCatalogoBienesInsumos.Modificar(oDOCatalogoBienesInsumos) = False Then
                      MsgBox oCatalogoBienesInsumos.MensajeError
                      GoTo ErrActItems
                   End If
                End If
                oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        mo_ReglasArchivoClinico.ActualizaNULLenIdPartidaIdCCostoCon999
        Me.MousePointer = 11
        Unload Me
    End If
    Exit Sub
ErrActItems:
    Resume Next
End Sub






'Private Sub cmdCargaAtencionesHist_Click()
'        Dim oRsTmp As New Recordset
'        Dim oRsFox As New Recordset
'        Dim oRsFox1 As New Recordset
'        Dim oConexionFox As New Connection
'        Dim oConexion As New Connection
'        Dim oDOhis_historicoAten As New DOhis_historicoAten, oHIS_historicoAten As New HIS_historicoAten
'        Dim oDOPaciente As New DOPaciente, oPacientes As New Pacientes
'        Dim oDOHistoriaClinica As New DOHistoriaClinica, oHistoriasClinicas As New HistoriasClinicas
'        Dim lnIdUsuario As Long, lbContinuar As Boolean, ldFechaNac As Date
'        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
'        Dim lcSegundoNombre As String, lnTipoSexo As Long
'        Dim lcAutogenerado As String, lcDNI As String, lnNroHistoriaClinica As Long
'        Dim lcSql As String, lcCodDx As String, lcCod2000 As String, lbEsNuevaHC As Boolean
'        Const lnIdTipoNumeracion As Long = 2
'        On Error GoTo ErrProAtHis
'        '
'        oConexion.CommandTimeout = 300
'        oConexion.CursorLocation = adUseClient
'        oConexion.Open sighentidades.CadenaConexion
'        oConexion.BeginTrans
'        '
'        oConexionFox.CommandTimeout = 300
'        oConexionFox.Open "DSN=his"
'        '
'        Me.MousePointer = 1
'        mo_ReglasAdmision.his_historicoAtencionesEliminarTodas oConexion
'        lcCod2000 = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
'        lcSql = "select * from histdet where cod_2000='" & lcCod2000 & "' order by dni"
'        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'        ProgressBar1.Max = oRsFox.RecordCount
'        If ProgressBar1.Max > 0 Then
'           ProgressBar1.Min = 0
'           ProgressBar1.Value = 0
'
'           Set oHIS_historicoAten.Conexion = oConexion
'           Set oPacientes.Conexion = oConexion
'           Set oHistoriasClinicas.Conexion = oConexion
'           oRsFox.MoveFirst
'           Do While Not oRsFox.EOF
'               ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
'               lcDNI = Right("        " & Trim(oRsFox.Fields!DNI), 8)
'               lnNroHistoriaClinica = Val(oRsFox.Fields!fichaFam)
'               lcSql = "select * from histcab where dni='" & lcDNI & "'"
'               If oRsFox1.State = 1 Then oRsFox1.Close
'               oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'               If oRsFox1.RecordCount > 0 Then
'                    lbContinuar = True
'                    lbEsNuevaHC = True
'                    Set oRsTmp = mo_ReglasAdmision.PacientesXdni(lcDNI, oConexion)
'                    If oRsTmp.RecordCount = 0 Then
'                       Set oRsTmp = mo_ReglasAdmision.PacientesXnroHistoriaTipoNumeracion(lnNroHistoriaClinica, lnIdTipoNumeracion, oConexion)
'                       If oRsTmp.RecordCount > 0 Then
'                          lbEsNuevaHC = False
'                          lnIdPaciente = oRsTmp.Fields!idPaciente
'                       End If
'                    Else
'                       lbEsNuevaHC = False
'                       lnIdPaciente = oRsTmp.Fields!idPaciente
'                    End If
'                    If lbEsNuevaHC = True Then
'                        lcApellidoPaterno = UCase(Left(Trim(oRsFox1.Fields!pApellido), 20))
'                        lcApellidoMaterno = UCase(Left(Trim(oRsFox1.Fields!sApellido), 20))
'                        lcPrimerNombre = UCase(RetornaPrimerNombre(oRsFox1.Fields!nombres))
'                        lcSegundoNombre = UCase(RetornaSegundoNombre(oRsFox1.Fields!nombres))
'                        lnTipoSexo = IIf(Val(oRsFox1.Fields!Sexo) = 0, 1, Val(oRsFox1.Fields!Sexo))
'                        ldFechaNac = IIf(IsNull(oRsFox1.Fields!fnac), CDate("01/01/1970"), oRsFox1.Fields!fnac)
'                        lcAutogenerado = PacienteCrearNroAutogenerado(Format(ldFechaNac, sighentidades.DevuelveFechaSoloFormato_DMY), lcApellidoPaterno, _
'                                         lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
'                        oPacientes.SetDefaults oDOPaciente
'                        oDOPaciente.NroHistoriaClinica = lnNroHistoriaClinica
'                        oDOPaciente.ApellidoMaterno = lcApellidoMaterno
'                        oDOPaciente.ApellidoPaterno = lcApellidoPaterno
'                        oDOPaciente.Autogenerado = lcAutogenerado
'                        If IsNull(oRsFox1.Fields!direccion) Or Trim(oRsFox1.Fields!direccion) = "" Then
'                           oDOPaciente.DireccionDomicilio = ""
'                        Else
'                           oDOPaciente.DireccionDomicilio = Left(oRsFox1.Fields!direccion, 50)
'                        End If
'                        oDOPaciente.FechaNacimiento = ldFechaNac
'                        If Val(oRsFox1.Fields!ubigeo) > 0 Then
'                           oDOPaciente.IdDistritoDomicilio = Val(oRsFox1.Fields!ubigeo)
'                        Else
'                           oDOPaciente.IdDistritoDomicilio = 0
'                        End If
'                        oDOPaciente.IdDocIdentidad = 1
'                        oDOPaciente.IdTipoNumeracion = lnIdTipoNumeracion
'                        oDOPaciente.idTipoSexo = lnTipoSexo
'                        oDOPaciente.NroDocumento = lcDNI
'                        oDOPaciente.PrimerNombre = lcPrimerNombre
'                        oDOPaciente.SegundoNombre = lcSegundoNombre
'                        oDOPaciente.TercerNombre = RetornaTercerNombre(oRsFox1.Fields!nombres)
'                       If oPacientes.Insertar(oDOPaciente) = True Then
'                            oDOHistoriaClinica.FechaCreacion = Date
'                            oDOHistoriaClinica.IdEstadoHistoria = 1
'                            oDOHistoriaClinica.idPaciente = oDOPaciente.idPaciente
'                            oDOHistoriaClinica.IdTipoHistoria = 1
'                            oDOHistoriaClinica.IdTipoNumeracion = lnIdTipoNumeracion
'                            oDOHistoriaClinica.IdUsuarioAuditoria = lnIdUsuario
'                            oDOHistoriaClinica.NroHistoriaClinica = lnNroHistoriaClinica
'                            If oHistoriasClinicas.Insertar(oDOHistoriaClinica) = False Then
'                               lbContinuar = False
'                            Else
'                               lnIdPaciente = oDOPaciente.idPaciente
'                            End If
'                       Else
'                            lbContinuar = False
'                       End If
'
'                    End If
'                    If lbContinuar = True Then
'                        Do While Not oRsFox.EOF And Val(lcDNI) = Val(oRsFox.Fields!DNI)
'                           If IsNull(oRsFox.Fields!cpt) Or Trim(oRsFox.Fields!cpt) = "" Then
'                              oDOhis_historicoAten.cpt = ""
'                           Else
'                              oDOhis_historicoAten.cpt = oRsFox.Fields!cpt
'                           End If
'                           If IsNull(oRsFox.Fields!diagnost) Or Trim(oRsFox.Fields!diagnost) = "" Then
'                              oDOhis_historicoAten.diagnost = oRsFox.Fields!diagnost
'                           Else
'                              oDOhis_historicoAten.diagnost = oRsFox.Fields!diagnost
'                           End If
'                           oDOhis_historicoAten.fecha = CDate(oRsFox.Fields!dia & "/" & oRsFox.Fields!mes & "/" & oRsFox.Fields!ano)
'                           oDOhis_historicoAten.IdUsuarioAuditoria = lnIdUsuario
'                           oDOhis_historicoAten.idPaciente = lnIdPaciente
'                           If IsNull(oRsFox.Fields!ups) Or Trim(oRsFox.Fields!ups) = "" Then
'                              oDOhis_historicoAten.ups = ""
'                           Else
'                              oDOhis_historicoAten.ups = oRsFox.Fields!ups
'                           End If
'                           If oHIS_historicoAten.Insertar(oDOhis_historicoAten) = True Then
'                              lcSql = ""
'                           End If
'                           oRsFox.MoveNext
'                           If oRsFox.EOF Then
'                              Exit Do
'                           End If
'                        Loop
'                    End If
'               Else
'                    lbContinuar = False
'               End If
'               If lbContinuar = False Then
'                    Do While Not oRsFox.EOF And Val(lcDNI) = Val(oRsFox.Fields!DNI)
'                        oRsFox.MoveNext
'                        If oRsFox.EOF Then
'                           Exit Do
'                        End If
'                    Loop
'               End If
'           Loop
'           oConexion.CommitTrans
'        End If
'        Me.MousePointer = 11
'        oConexionFox.Close
'        oConexion.Close
'
'        Set oRsTmp = Nothing
'        Set oRsFox = Nothing
'        Set oRsFox1 = Nothing
'        Set oConexionFox = Nothing
'        Set oConexion = Nothing
'        Set oDOhis_historicoAten = Nothing
'        Set oHIS_historicoAten = Nothing
'        Set oDOPaciente = Nothing
'        Set oPacientes = Nothing
'        Set oDOHistoriaClinica = Nothing
'        Set oHistoriasClinicas = Nothing
'        Unload Me
'        Exit Sub
'ErrProAtHis:
'        MsgBox Err.Description
'
'        oConexion.RollbackTrans
'        Set oRsTmp = Nothing
'        Set oRsFox = Nothing
'        Set oRsFox1 = Nothing
'        Set oConexionFox = Nothing
'        Set oConexion = Nothing
'        Set oDOhis_historicoAten = Nothing
'        Set oHIS_historicoAten = Nothing
'        Set oDOPaciente = Nothing
'        Set oPacientes = Nothing
'        Set oDOHistoriaClinica = Nothing
'        Set oHistoriasClinicas = Nothing
'        Unload Me
'End Sub
'
'Function RetornaPrimerNombre(lcPrimerSegundoNombreJuntos As String) As String
'    Dim ln As Integer
'    RetornaPrimerNombre = ""
'    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
'    If ln > 0 Then
'       RetornaPrimerNombre = Trim(Left(lcPrimerSegundoNombreJuntos, ln))
'    Else
'       RetornaPrimerNombre = lcPrimerSegundoNombreJuntos
'    End If
'End Function
'
'Function RetornaSegundoNombre(lcPrimerSegundoNombreJuntos As String) As String
'    Dim ln As Integer
'    RetornaSegundoNombre = ""
'    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
'    If ln > 0 Then
'       RetornaSegundoNombre = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
'       ln = InStr(RetornaSegundoNombre, " ")
'       If ln > 0 Then
'          RetornaSegundoNombre = Trim(Left(RetornaSegundoNombre, ln))
'       End If
'    End If
'End Function
'
'Function RetornaTercerNombre(lcPrimerSegundoNombreJuntos As String) As String
'    Dim ln As Integer, lcNombre1 As String, lcNombre2 As String, lcNombre3 As String
'    RetornaTercerNombre = ""
'    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
'    If ln > 0 Then
'       lcNombre1 = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
'       ln = InStr(lcNombre1, " ")
'       If ln > 0 Then
'          lcNombre2 = Trim(Left(lcNombre1, ln))
'          RetornaTercerNombre = Trim(Mid(lcNombre1, ln + 1, 100))
'       End If
'    End If
'End Function
'Function PacienteCrearNroAutogenerado(lcFechaNacimiento As String, lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String, lcSegundoNombre As String, lnIdTipoSexo As Long)
'Dim P1 As String    'Primer digito del apellido paterno
'Dim P4 As String    'Cuarto Digito del apellido paterno
'Dim M1 As String    'Primer digito del apellido materno
'Dim M4 As String    'Cuarto digito del apellido materno
'Dim N11 As String   'Primer digito del primer nombre
'Dim N41 As String   'Cuarto digito del primer materno
'Dim N12 As String   'Primer digito del Ultimo materno
'Dim N42 As String   'Cuarto digito del Ultimo materno
'Dim D As String     'Digito de verificacion
'Dim DD As String
'Dim MM As String
'Dim AAA As String
'Dim sTemp  As String
'
'        DD = Left(lcFechaNacimiento, 2)
'        MM = Mid(lcFechaNacimiento, 4, 2)
'        AAA = Mid(lcFechaNacimiento, 8, 3)
'        DevuelvePrimeryCuartoCaracter lcApellidoPaterno, P1, P4
'        DevuelvePrimeryCuartoCaracter lcApellidoMaterno, M1, M4
'        DevuelvePrimeryCuartoCaracter lcPrimerNombre, N11, N41
'        DevuelvePrimeryCuartoCaracter lcSegundoNombre, N12, N42
'        sTemp = AAA + MM + DD & lnIdTipoSexo & P1 + P4 + M1 + M4 + N11 + N41 + N12 + N42
'        PacienteCrearNroAutogenerado = sTemp & Modulo10(sTemp)
'
'End Function
'
'Function Modulo10(sValor As String) As Integer
'Dim sTemp As String
'Dim i As Integer
'Dim k As Integer
'Dim iTotal As Integer
'
'    sTemp = ""
'
'    For i = 1 To Len(sValor)
'        If IsNumeric(Mid(sValor, i, 1)) Then
'            sTemp = sTemp + Mid(sValor, i, 1)
'        Else
'            sTemp = sTemp + DevuelveValorEnNumeros(Mid(sValor, i, 1))
'        End If
'    Next i
'
'    'Acumula total de digitos
'    iTotal = 0
'    For i = 1 To Len(sTemp)
'        If i Mod 2 <> 0 Then
'            k = CInt(Mid(sTemp, i, 1)) * 2
'            iTotal = iTotal + (k - (k Mod 10)) / 10 + (k Mod 10)
'        Else
'            iTotal = iTotal + CInt(Mid(sTemp, i, 1))
'        End If
'    Next i
'
'    If (iTotal Mod 10) = 0 Then
'        Modulo10 = 0
'    Else
'        Modulo10 = 10 - (iTotal Mod 10)
'    End If
'
'
'
'End Function
'Function DevuelveValorEnNumeros(sCaracter As String) As String
'
'    Select Case sCaracter
'    Case "A" To "N"
'        DevuelveValorEnNumeros = Asc(sCaracter) - 55
'    Case "Ñ"
'        DevuelveValorEnNumeros = 24
'    Case "O" To "Z"
'        DevuelveValorEnNumeros = Asc(sCaracter) - 54
'    End Select
'
'End Function
'
'Sub DevuelvePrimeryCuartoCaracter(sPalabra As String, C1 As String, C2 As String)
'Dim sTemp As String
'        If sPalabra <> "" Then
'            sTemp = ObtenerUltimaPalabra(EliminarConjunciones(sPalabra))
'            C1 = Left(sTemp, 1)
'            C2 = DevuelveCuartoCaracter(sTemp)
'        Else
'            C1 = "X"
'            C2 = "X"
'        End If
'End Sub
'
'Function DevuelveCuartoCaracter(sPalabra) As String
'    If Len(sPalabra) <= 4 Then
'        DevuelveCuartoCaracter = Right(sPalabra, 1)
'    Else
'        DevuelveCuartoCaracter = Mid(sPalabra, 4, 1)
'    End If
'End Function
'Function ObtenerUltimaPalabra(sTexto As String) As String
'Dim p As String
'Dim iUltBlanco As Integer
'Dim sTemp As String
'
'
'    sTemp = Trim(sTexto)
'
'    p = InStr(sTemp, " ")
'    iUltBlanco = 0
'    Do While p > 0
'        iUltBlanco = p
'        p = InStr(p + 1, sTemp, " ")
'    Loop
'    If iUltBlanco > 0 Then
'        ObtenerUltimaPalabra = Mid(sTemp, iUltBlanco + 1)
'    Else
'        ObtenerUltimaPalabra = sTemp
'    End If
'End Function
'
'Function EliminarConjunciones(sPalabra As String)
'Dim sTemp As String
'
'        sTemp = ReemplazarCadena(sPalabra, " DE ", " ")
'        sTemp = ReemplazarCadena(sTemp, " DEL ", " ")
'        sTemp = ReemplazarCadena(sTemp, " EL ", " ")
'        sTemp = ReemplazarCadena(sTemp, " LA ", " ")
'        sTemp = ReemplazarCadena(sTemp, " LOS ", " ")
'        sTemp = ReemplazarCadena(sTemp, " LAS ", " ")
'
'        EliminarConjunciones = sTemp
'
'End Function






Private Sub cmdCargaINICajaFarmacia_Click()
    lcTipoServicioFarmacia = "FARMACIA"
    lblFechaDoc.Caption = "Fecha " & lcTipoComprobanteCaja & ":"
    lblTotalDocLetras.Caption = "Total " & lcTipoComprobanteCaja & " (letras):"
    lblTotalDoc.Caption = "Total " & lcTipoComprobanteCaja & ":"
    cmdImprimeBoleta.Caption = "Imprime " & UCase(lcTipoComprobanteCaja) & " de prueba"
    Label61.Caption = UCase(lcTipoComprobanteCaja) & " FARMACIA"
    
    CargaSetup_Caja txtRutaINI.Text, wxIdTipoComprobanteDefault, False
    
    Me.txtNumeroSerieX.Text = WxLnNumeroSerieX_F
    Me.txtNumeroSerieY.Text = WxLnNumeroSerieY_F
    Me.txtEstadoX.Text = WxLnEstadoX_F
    Me.txtEstadoY.Text = WxLnEstadoY_F
    Me.txtTipoX.Text = WxLnTipoX_F
    Me.txtTipoY.Text = WxLnTipoY_F
    Me.txtRzSocialX.Text = WxLnRzSocialX_F
    Me.txtRzSocialY.Text = WxLnRzSocialY_F
    Me.txtFechaX.Text = WxLnFechaX_F
    Me.txtFechaY.Text = WxLnFechaY_F
    Me.txtServicioX.Text = WxLnServicioX_F
    Me.txtServicioY.Text = WxLnServicioY_F
    Me.txtObservacionesX.Text = WxLnObservacionesX_F
    Me.txtObservacionesY.Text = WxLnObservacionesY_F
    
    txtRucX.Text = ""
    txtRucY.Text = ""
    txtDireccionX.Text = ""
    txtDireccionY.Text = ""
    If lcTipoComprobanteCaja = lcFactura Then
        txtRucX.Text = WxLnCabRucX_F
        txtRucY.Text = WxLnCabRucY_F
        txtDireccionX.Text = WxLnCabDireccionX_F
        txtDireccionY.Text = WxLnCabDireccionY_F
    End If
    
    '
    Me.txtCodigoY.Text = WxLnCodigoY_F
    Me.txtProductoY.Text = WxLnProductoY_F
    Me.txtProductoAncho.Text = WxLnProductoWidhtY_F
    Me.txtCantidadY.Text = WxLnCantidadY_F
    Me.txtPrecioY.Text = WxLnPrecioY_F
    Me.txtImporteY.Text = WxLnImporteY_F
    '
    Me.txtCajeroX.Text = WxLnCajeroX_F
    Me.txtCajeroY.Text = WxLnCajeroY_F
    Me.txtCajaX.Text = WxLnCajaX_F
    Me.txtCajaY.Text = WxLnCajaY_F
    Me.txtAdelantosX.Text = WxLnAdelantosX_F
    Me.txtAdelantosY.Text = WxLnAdelantosY_F
    Me.txtTotalPagarX.Text = WxLnTotalPagarX_F
    Me.txtTotalPagarY.Text = WxLnTotalPagarY_F
    Me.txtCuentaX.Text = WxLnCuentaX_F
    Me.txtCuentaY.Text = WxLnCuentaY_F
    Me.txtExoneracionesX.Text = WxLnExoneracionesX_F
    Me.txtExoneracionesY.Text = WxLnExoneracionesY_F
    Me.txtTotalEnLetrasX.Text = WxLnTotalEnLetrasX_F
    Me.txtTotalEnLetrasY.Text = WxLnTotalEnLetrasY_F
    txtTotalLetrasAncho.Text = WxLnTotalLetrasWidhtY_F
    Me.txtTotalX.Text = WxLnTotalX_F
    Me.txtTotalY.Text = WxLnTotalY_F
    Me.txtSubTotalX.Text = WxLnSubTotalX_F
    Me.txtSubTotalY.Text = WxLnSubTotalY_F
    Me.txtIGVX.Text = WxLnIGVX_F
    Me.txtIGVY.Text = WxLnIGVY_F
    '
    Me.txtCabeceraAlto.Text = WxLnCabeceraAlto_F
    Me.txtPieAlto.Text = WxLnPieAlto_F
    
    'mgaray
    cboReporteador.ListIndex = WxLnTipoReporteador_F
    cboPapel.Text = WxLnNombreHoja_F
    txtMargenIzquierda.Text = WxLnMargenIzquierdoX_F
    txtMargenDerecha.Text = WxLnMargenDerechoX_F
    txtMargenSuperior.Text = WxLnMargenSuperiorY_F
    txtMargenInferior.Text = WxLnMargenInferiorY_F
    'mgaray
    activarGrabarConfiguracionComprobante
End Sub

Private Sub cmdCargaINICajaServicios_Click()
    lcTipoServicioFarmacia = "SERVICIOS"
    lblFechaDoc.Caption = "Fecha " & lcTipoComprobanteCaja & ":"
    lblTotalDocLetras.Caption = "Total " & lcTipoComprobanteCaja & " (letras):"
    lblTotalDoc.Caption = "Total " & lcTipoComprobanteCaja & ":"
    cmdImprimeBoleta.Caption = "Imprime " & UCase(lcTipoComprobanteCaja) & " de prueba"
    Label61.Caption = UCase(lcTipoComprobanteCaja) & " SERVICIOS"
    
    CargaSetup_Caja txtRutaINI.Text, wxIdTipoComprobanteDefault, False
    
    Me.txtNumeroSerieX.Text = WxLnNumeroSerieX
    Me.txtNumeroSerieY.Text = WxLnNumeroSerieY
    Me.txtEstadoX.Text = WxLnEstadoX
    Me.txtEstadoY.Text = WxLnEstadoY
    Me.txtTipoX.Text = WxLnTipoX
    Me.txtTipoY.Text = WxLnTipoY
    Me.txtRzSocialX.Text = WxLnRzSocialX
    Me.txtRzSocialY.Text = WxLnRzSocialY
    Me.txtFechaX.Text = WxLnFechaX
    Me.txtFechaY.Text = WxLnFechaY
    Me.txtServicioX.Text = WxLnServicioX
    Me.txtServicioY.Text = WxLnServicioY
    Me.txtObservacionesX.Text = WxLnObservacionesX
    Me.txtObservacionesY.Text = WxLnObservacionesY
    Me.txtHistoriaX.Text = WxLnHistoriaX
    Me.txtHistoriaY.Text = WxLnHistoriaY
    
    txtRucX.Text = ""
    txtRucY.Text = ""
    txtDireccionX.Text = ""
    txtDireccionY.Text = ""
    If lcTipoComprobanteCaja = lcFactura Then
        txtRucX.Text = WxLnCabRucX
        txtRucY.Text = WxLnCabRucY
        txtDireccionX.Text = WxLnCabDireccionX
        txtDireccionY.Text = WxLnCabDireccionY
    End If
    
'    IIf(lcTipoComprobanteCaja = lcTicket, True, False)
    '
    Me.txtCodigoY.Text = WxLnCodigoY
    Me.txtProductoY.Text = WxLnProductoY
    Me.txtProductoAncho.Text = WxLnProductoWidhtY
    Me.txtCantidadY.Text = WxLnCantidadY
    Me.txtPrecioY.Text = WxLnPrecioY
    Me.txtImporteY.Text = WxLnImporteY    '
    Me.txtCajeroX.Text = WxLnCajeroX
    Me.txtCajeroY.Text = WxLnCajeroY
    Me.txtCajaX.Text = WxLnCajaX
    Me.txtCajaY.Text = WxLnCajaY
    Me.txtAdelantosX.Text = WxLnAdelantosX
    Me.txtAdelantosY.Text = WxLnAdelantosY
    Me.txtTotalPagarX.Text = WxLnTotalPagarX
    Me.txtTotalPagarY.Text = WxLnTotalPagarY
    Me.txtCuentaX.Text = WxLnCuentaX
    Me.txtCuentaY.Text = WxLnCuentaY
    Me.txtExoneracionesX.Text = WxLnExoneracionesX
    Me.txtExoneracionesY.Text = WxLnExoneracionesY
    Me.txtTotalEnLetrasX.Text = WxLnTotalEnLetrasX
    Me.txtTotalEnLetrasY.Text = WxLnTotalEnLetrasY
    txtTotalLetrasAncho.Text = WxLnTotalLetrasWidhtY
    Me.txtTotalX.Text = WxLnTotalX
    Me.txtTotalY.Text = WxLnTotalY
    Me.txtSubTotalX.Text = WxLnSubTotalX
    Me.txtSubTotalY.Text = WxLnSubTotalY
    Me.txtIGVX.Text = WxLnIGVX
    Me.txtIGVY.Text = WxLnIGVY
    '
    Me.txtCabeceraAlto.Text = WxLnCabeceraAlto
    Me.txtPieAlto.Text = WxLnPieAlto

    'mgaray
    cboReporteador.ListIndex = WxLnTipoReporteador
    cboPapel.Text = WxLnNombreHoja
    
    txtMargenIzquierda.Text = WxLnMargenIzquierdoX
    txtMargenDerecha.Text = WxLnMargenDerechoX
    txtMargenSuperior.Text = WxLnMargenSuperiorY
    txtMargenInferior.Text = WxLnMargenInferiorY

    'mgaray
    activarGrabarConfiguracionComprobante
End Sub

Private Sub cmdCie10DescripcionCorta_Click()

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
  
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtExcel2.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    Dim lcCPTcorta As String, lnPrecioPagante As Double, lnIdProducto As Long, lbContinuar As Boolean
    lnFila = 1
    lnFilaFinal = 10000
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    For lnFor = lnFila To lnFilaFinal
        lcRango = "E" + Trim(Str(lnFor))
        lcCodigo = Trim(s.Range(lcRango).Value)
        lcCodigo = Left(lcCodigo, 3) & "." & Mid(lcCodigo, 4, 10)   'Le incluye el Punto
        lcRango = "C" + Trim(Str(lnFor))
        lnPrecioPagante = Val(s.Range(lcRango).Value)
        lcRango = "B" + Trim(Str(lnFor))
        lcCPTcorta = Trim(s.Range(lcRango).Value)
        If Len(Trim(lcCodigo)) > 0 And Trim(lcCPTcorta) <> "" Then
            With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "DiagnosticosSeleccionarTodoCamposPorCodigoCie2004"
                  Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
                  Set oRsTmp = .Execute
                  Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
  
            If oRsTmp.RecordCount > 0 Then
            
                
                With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = oConexion
                      .CommandTimeout = 150
                      .CommandText = "DiagnosticosModificarDescripcion"
                      Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oRsTmp.Fields!IdDiagnostico): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, Left(lcCPTcorta, 250)): .Parameters.Append oParameter
                      .Execute
                End With
                lnIdProducto = oRsTmp.Fields!IdDiagnostico
                Set oCommand = Nothing
                Set oParameter = Nothing
            
            End If
            oRsTmp.Close
     
        Else
           lcSql = ""
        End If
    Next
    Set s = Nothing
    W.Save
    W.Close
    
    oConexion.Close
    Set oConexion = Nothing
    Set W = Nothing
    Set EXL = Nothing
    Unload Me
End Sub


Private Sub cmdCodigo_Click()
    If txtCodigoSismed.Text = "" Then
       MsgBox "Ingrese el código Sismed", vbCritical
       Exit Sub
    End If
    If Not (Val(txtNtipoSalida.Text) = 1 Or Val(txtNtipoSalida.Text) = 2) Then
       MsgBox "El nuevo TIPO SALIDA solo puede ser 1 o 2", vbCritical
       Exit Sub
    End If
    If txtClaveCodigo.Text = Format(Date, "ddmmyyyy") Then
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim lnIdProducto As Long, lbContinuar As Boolean
       'Set oRsTmp1 = mo_ReglasFarmacia.FarmMovimientoDetalleSeleccionarXcodigo(txtCodigoSismed.Text)
       Set oRsTmp1 = FarmMovimientoDetalleSeleccionarXcodigo(txtCodigoSismed.Text)
       If oRsTmp1.RecordCount > 0 Then
          lnIdProducto = oRsTmp1!idProducto
          oRsTmp1.Filter = "idTipoSalidaBienInsumo=" & txtNtipoSalida.Text
          lbContinuar = True
          If oRsTmp1.RecordCount > 0 Then
             If oRsTmp1!idTipoSalidaBienInsumo = Val(txtNtipoSalida.Text) Then
                MsgBox "El NUEVO TIPO SALIDA ya exite", vbCritical
                lbContinuar = False
             End If
          End If
          If lbContinuar = True Then
             If Val(txtNtipoSalida.Text) = 2 Then
                'Set oRsTmp2 = mo_ReglasFarmacia.farmMovimientoVentasDetalleXidProducto(lnIdProducto)
                Set oRsTmp2 = farmMovimientoVentasDetalleXidProducto(lnIdProducto)
                If oRsTmp2.RecordCount > 0 Then
                   MsgBox "El NUEVO TIPO SALIDA ya tiene DOCUMENTOS registrados en la opción VENTAS", vbCritical
                Else
                   'mo_ReglasFarmacia.FarmaciaActualizaTipoSalida lnIdProducto, Val(txtNtipoSalida.Text)
                   FarmaciaActualizaTipoSalida lnIdProducto, Val(txtNtipoSalida.Text)
                End If
                oRsTmp2.Close
             Else
                oRsTmp1.Filter = "MovTipo='S' and idTipoConcepto=16 and idEstadoMovimiento=1"
                If oRsTmp1.RecordCount > 0 Then
                   MsgBox "El NUEVO TIPO SALIDA ya tiene DOCUMENTOS registrados en la opción INTERVENCIONES SANITARIAS", vbCritical
                Else
                   'mo_ReglasFarmacia.FarmaciaActualizaTipoSalida lnIdProducto, Val(txtNtipoSalida.Text)
                   FarmaciaActualizaTipoSalida lnIdProducto, Val(txtNtipoSalida.Text)
                End If
                
             End If
             Unload Me
          End If
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
    End If
End Sub


Private Sub cmdDepuraPtoCarga_Click()

  Dim oCommand As New ADODB.Command
  Dim oConexion As New ADODB.Connection

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
    'Depura Puntos de Carga (32 y 38)
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "DepuraPtoCarga"
            Set oRsTmp1 = .Execute
        End With
        
        Unload Me
    End If
    
End Sub

'elimina pacientes sin movimientos
'txtNpacientesElim.text <---nro de pacientes a eliminar
Private Sub cmdEliminaPacientes_Click()
    If Val(txtNpacientesElim.Text) > 0 Then
        On Error GoTo ErrEP
        Dim oRsPacientes As New Recordset
        Dim oRsHistorias As New Recordset
        Dim oRsAtenciones As New Recordset
        Dim lnContador As Long
        Dim lnError As Long
        oRsPacientes.Open "SELECT * from Pacientes order by apellidoMaterno desc", wxConexionRed, adOpenKeyset, adLockOptimistic
        lnContador = 0
        oRsPacientes.MoveFirst
        Do While Not oRsPacientes.EOF
           oRsAtenciones.Open "select idPaciente from atenciones  where idPaciente=" & oRsPacientes.Fields!idPaciente, wxConexionRed, adOpenKeyset, adLockOptimistic
           If oRsAtenciones.RecordCount = 0 Then
                lnError = 0
                oRsHistorias.Open "delete from HistoriasClinicas where idPaciente= " & oRsPacientes.Fields!idPaciente, wxConexionRed, adOpenKeyset, adLockOptimistic
                If lnError = 0 Then
                      oRsPacientes.Delete
                      oRsPacientes.Update
                End If
           End If
           oRsAtenciones.Close
           oRsPacientes.MoveNext
           lnContador = lnContador + 1
           If lnContador >= Val(txtNpacientesElim.Text) Then
              Exit Do
           End If
        Loop
        oRsPacientes.Close
    End If
    Unload Me
    Exit Sub
ErrEP:
    lnError = 1
    Resume Next
    
End Sub


Private Sub cmdEstablecimientosNew_Click()

    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 1
        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=HIS"
        '
        Me.MousePointer = 1
        'Provincias
        lcSql = "select * from mProvinc"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
                lcTipo = Trim(Str(Val(oRsFox.Fields!dptoCod & oRsFox.Fields!provCod)))
                If Val(lcTipo) > 0 Then
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "ProvinciasActualizarInformacionFox"
                        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, CLng(lcTipo)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, Left(oRsFox.Fields!provDes, 50)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, Val(oRsFox.Fields!dptoCod)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                End If
                
                oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        'Distritos
        lcSql = "select * from mDistrito"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
                lcTipo = Trim(Str(Val(oRsFox.Fields!dptoCod & oRsFox.Fields!provCod & oRsFox.Fields!distCod)))
                                
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "DistritosSeleccionarPorId"
                    Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, CLng(lcTipo)): .Parameters.Append oParameter
                    Set oRsTmpOpc = .Execute
                    Set oRsTmpOpc.ActiveConnection = Nothing
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                 
                If oRsTmpOpc.RecordCount = 0 And Val(lcTipo) > 0 Then
                   lcTipo = Trim(Str(Val(oRsFox.Fields!dptoCod & oRsFox.Fields!provCod)))
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "ProvinciasSeleccionarPorId"
                        Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, CLng(lcTipo)): .Parameters.Append oParameter
                        Set oRsTmpCat = .Execute
                        Set oRsTmpCat.ActiveConnection = Nothing
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                
                   If oRsTmpCat.RecordCount > 0 Then
                        lcTipo = Trim(Str(Val(oRsFox.Fields!dptoCod & oRsFox.Fields!provCod & oRsFox.Fields!distCod)))

                        
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = wxConexionRed
                            .CommandTimeout = 150
                            .CommandText = "DistritoAgregar"
                            Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, CLng(lcTipo)): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, Left(oRsFox.Fields!DistDes, 50)): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, Val(oRsFox.Fields!dptoCod & oRsFox.Fields!provCod)): .Parameters.Append oParameter
                            .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                    
                   End If
                   oRsTmpCat.Close
                End If
                oRsTmpOpc.Close
                oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        'Establecimeintos
        
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = wxConexionRed
          .CommandTimeout = 150
          .CommandText = "EstablecimientosSeleccionarTodos"
          Set oRsTmpOpc = .Execute
          Set oRsTmpOpc.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
              
        lnIdEstablecimiento = oRsTmpOpc.Fields!IdEstablecimiento + 1
        oRsTmpOpc.Close
        '
        lcSql = "select * from mAlmacen"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              If (Val(oRsFox.Fields!almcod) > 0 And Val(oRsFox.Fields!almcod) < 15000) And Len(Trim(oRsFox.Fields!almcod)) <= 6 Then
              
                
                With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "EstablecimientosSeleccionarTodoCamposXCodigo"
                        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, Trim(oRsFox.Fields!almcod)): .Parameters.Append oParameter
                        Set oRsTmpOpc = .Execute
                        Set oRsTmpOpc.ActiveConnection = Nothing
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                    
                If oRsTmpOpc.RecordCount = 0 Then
                   oRsTmpOpc.Close
                   
                   
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = wxConexionRed
                        .CommandTimeout = 150
                        .CommandText = "DistritosSeleccionarPorId"
                        Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, CLng(Trim(Str(Val(oRsFox.Fields!almUbigeo))))): .Parameters.Append oParameter
                        Set oRsTmpCat = .Execute
                        Set oRsTmpCat.ActiveConnection = Nothing
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                   
                   If oRsTmpCat.RecordCount > 0 Then
                        lcTipo = "0"
                        If oRsFox.Fields!almTipo = "C" Then
                           lcTipo = "2"
                        ElseIf oRsFox.Fields!almTipo = "P" Then
                           lcTipo = "3"
                        ElseIf oRsFox.Fields!almTipo = "H" Then
                           lcTipo = "1"
                        End If
                        If lcTipo <> "0" And Len(Trim(oRsFox.Fields!almcod)) = 5 And oRsFox.Fields!almSit = "1" Then
                             
                            With oCommand
                                .CommandType = adCmdStoredProc
                                Set .ActiveConnection = wxConexionRed
                                .CommandTimeout = 150
                                .CommandText = "EstablecimientosAgregar"
                                Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, lnIdEstablecimiento): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, Trim(oRsFox.Fields!almcod)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, Left(oRsFox.Fields!almDes, 150)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, Val(oRsFox.Fields!almUbigeo)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@IdTipo", adInteger, adParamInput, 0, Val(lcTipo)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                                Set oRsTmpOpc = .Execute
                            End With
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                             
                             lnIdEstablecimiento = lnIdEstablecimiento + 1
                        End If
                   End If
                   oRsTmpCat.Close
                Else
                   oRsTmpOpc.Close
                End If
              End If
              oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
         
        Me.MousePointer = 11
        Unload Me
    End If
End Sub


Private Sub cmdEstratVtas_Click()

If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnIdProducto As Long, lnIdTipoSalidaBienInsumo As Long
    Dim lbEsEstrategico As Boolean, lbEsVenta As Boolean
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim mo_ReporteUtil As New ReporteUtil
    Dim iFila As Integer
    '
    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    Set oWorkBook = oExcel.Workbooks.Add
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\HojaLibre.xls")
    oWorkBookPlantilla.Worksheets("Hoja_libre").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    Set oWorkSheet = oWorkBook.Sheets(1)
    iFila = 5
        
    Me.MousePointer = 11
    Set mo_conexion = New Connection
    mo_conexion.CommandTimeout = 300
    mo_conexion.CursorLocation = adUseClient
    mo_conexion.Open sighentidades.CadenaConexion
    
    
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDetalleFarmMovimientoSeleccionarMovTipoS"
        Set oRsTmp1 = .Execute
    End With
    Set oCommand = Nothing
    
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lnIdProducto = oRsTmp1.Fields!idProducto
          lcCodigo = oRsTmp1.Fields!Codigo
          lcNombre = oRsTmp1.Fields!nombre
          lnIdTipoSalidaBienInsumo = 1
          lbEsEstrategico = False: lbEsVenta = False
          
          Do While Not oRsTmp1.EOF And lnIdProducto = oRsTmp1.Fields!idProducto
             If oRsTmp1.Fields!idTipoConcepto = 16 Then
                lbEsEstrategico = True
             Else
                lbEsVenta = True
             End If
             oRsTmp1.MoveNext
             If oRsTmp1.EOF Then
                Exit Do
             End If
          Loop
          
          If lbEsEstrategico = True And lbEsVenta = True Then
             oWorkSheet.Cells(iFila, 1).Value = "consumo: Venta/Estrategico"
             oWorkSheet.Cells(iFila, 2).Value = lcCodigo
             oWorkSheet.Cells(iFila, 3).Value = lcNombre
             iFila = iFila + 1
          ElseIf lbEsEstrategico = False And lbEsVenta = True Then
          

             With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = mo_conexion
                    .CommandTimeout = 150
                    .CommandText = "FarmSaldoFarmAlmacenSeleccionarPorIdProducto"
                    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                    Set oRsTmp2 = .Execute
                    Set oRsTmp2.ActiveConnection = Nothing
             End With
             Set oCommand = Nothing
             Set oParameter = Nothing
    
             lcSql = ""
             If oRsTmp2.RecordCount > 0 Then
                oRsTmp2.MoveFirst
                Do While Not oRsTmp2.EOF
                    If oRsTmp2.Fields!idTipoSalidaBienInsumo <> 1 Then
                        oWorkSheet.Cells(iFila, 1).Value = "consumo Venta, pero los saldos como " & Trim(Str(oRsTmp2.Fields!idTipoSalidaBienInsumo)) & " (almacen:" & Trim(oRsTmp2.Fields!Descripcion) & ")"
                        oWorkSheet.Cells(iFila, 2).Value = lcCodigo
                        oWorkSheet.Cells(iFila, 3).Value = lcNombre
                        iFila = iFila + 1
                        lcSql = "problemas"
                        Exit Do
                    End If
                    oRsTmp2.MoveNext
                Loop
             End If
             oRsTmp2.Close
             If lcSql = "" Then
                With oCommand
                       .CommandType = adCmdStoredProc
                       Set .ActiveConnection = mo_conexion
                       .CommandTimeout = 150
                       .CommandText = "FactCatalogoBienesInsumosActualizarIdTipoSalidaBienInsumo"
                       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                       Set oParameter = .CreateParameter("@IdTipoSalidaBienInsumo", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                       Set oRsTmp2 = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
             End If
          ElseIf lbEsEstrategico = True And lbEsVenta = False Then

             With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = mo_conexion
                    .CommandTimeout = 150
                    .CommandText = "FarmSaldoFarmAlmacenSeleccionarPorIdProducto"
                    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                    Set oRsTmp2 = .Execute
                    Set oRsTmp2.ActiveConnection = Nothing
             End With
             Set oCommand = Nothing
             Set oParameter = Nothing
             
             lcSql = ""
             If oRsTmp2.RecordCount > 0 Then
                oRsTmp2.MoveFirst
                Do While Not oRsTmp2.EOF
                    If oRsTmp2.Fields!idTipoSalidaBienInsumo <> 2 Then
                        oWorkSheet.Cells(iFila, 1).Value = "consumo Estrategico, pero los saldos como " & Trim(Str(oRsTmp2.Fields!idTipoSalidaBienInsumo)) & " (almacen:" & Trim(oRsTmp2.Fields!Descripcion) & ")"
                        oWorkSheet.Cells(iFila, 2).Value = lcCodigo
                        oWorkSheet.Cells(iFila, 3).Value = lcNombre
                        iFila = iFila + 1
                        lcSql = "problemas"
                        Exit Do
                    End If
                    oRsTmp2.MoveNext
                Loop
             End If
             oRsTmp2.Close
             If lcSql = "" Then
                
                With oCommand
                       .CommandType = adCmdStoredProc
                       Set .ActiveConnection = mo_conexion
                       .CommandTimeout = 150
                       .CommandText = "FactCatalogoBienesInsumosActualizarIdTipoSalidaBienInsumo"
                       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                       Set oParameter = .CreateParameter("@IdTipoSalidaBienInsumo", adInteger, adParamInput, 0, 2): .Parameters.Append oParameter
                       Set oRsTmp2 = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
             End If
          End If
          
       Loop
    End If
    oRsTmp1.Close
    mo_conexion.Close
    Me.MousePointer = 1
    
    
    oExcel.Visible = True
    oWorkSheet.PrintPreview
    
    
    Unload Me
End If

End Sub



'*****(En Lima) Compara archivo "tablasypa.mdb" con ultima estructura bd
Private Sub cmdEstrucCompara_Click()
    Dim oCatalogo As ADOX.Catalog
    Dim oTabla As ADOX.Table
    Dim oConexionMDB As New ADODB.Connection
    Dim oRsTablasMDB As New Recordset
    Dim oRsEstructuraMDB As New Recordset
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    Dim lnFor As Integer
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim mo_ReporteUtil As New ReporteUtil
    Dim iFila As Integer
    '
    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    Set oWorkBook = oExcel.Workbooks.Add
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\HojaLibre.xls")
    oWorkBookPlantilla.Worksheets("Hoja_libre").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    Set oWorkSheet = oWorkBook.Sheets(1)
    iFila = 5
    '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    '
    oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
    '
    oWorkSheet.Cells(iFila, 1).Value = "Tabla"
    oWorkSheet.Cells(iFila, 2).Value = "Campo"
    oWorkSheet.Cells(iFila, 3).Value = "Tipo"
    oWorkSheet.Cells(iFila, 4).Value = "Long"
    oWorkSheet.Cells(iFila, 5).Value = "Observación"
    iFila = iFila + 2
    'Carga Tablas
    Set oCatalogo = New ADOX.Catalog
    oCatalogo.ActiveConnection = sighentidades.CadenaConexion
    For Each oTabla In oCatalogo.Tables
        If oTabla.Type = "TABLE" Then
           lcSql = "select * from estructura  where tabla='" & oTabla.Name & "'"
           oRsTablasMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
           If oRsTablasMDB.RecordCount = 0 Then
              oRsTablasMDB.Close
              oWorkSheet.Cells(iFila, 1).Value = oTabla.Name
              oWorkSheet.Cells(iFila, 5).Value = "NO existe la Tabla "
              iFila = iFila + 1
           Else
              oRsTablasMDB.Close
              For lnFor = 1 To oTabla.Columns.Count - 1
                    lcSql = "select * from estructura where tabla='" & oTabla.Name & "' and campo='" & oTabla.Columns(lnFor).Name & "'"
                    oRsTablasMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
                    If oRsTablasMDB.RecordCount = 0 Then
                       oWorkSheet.Cells(iFila, 1).Value = oTabla.Name
                       oWorkSheet.Cells(iFila, 2).Value = oTabla.Columns(lnFor).Name
                       oWorkSheet.Cells(iFila, 3).Value = ObtenerTipoDeStoreProcedure(oTabla.Columns(lnFor).Type)
                       oWorkSheet.Cells(iFila, 4).Value = oTabla.Columns(lnFor).DefinedSize
                       oWorkSheet.Cells(iFila, 5).Value = "NO existe el campo"
                       iFila = iFila + 1
                    ElseIf ObtenerTipoDeStoreProcedure(oTabla.Columns(lnFor).Type) <> oRsTablasMDB.Fields!Tipo Then
                       oWorkSheet.Cells(iFila, 1).Value = oTabla.Name
                       oWorkSheet.Cells(iFila, 2).Value = oTabla.Columns(lnFor).Name
                       oWorkSheet.Cells(iFila, 3).Value = oRsTablasMDB.Fields!Tipo
                       oWorkSheet.Cells(iFila, 4).Value = oRsTablasMDB.Fields!longitud
                       lcSql = "El TIPO debería ser " & ObtenerTipoDeStoreProcedure(oTabla.Columns(lnFor).Type)
                       If oTabla.Columns(lnFor).DefinedSize <> oRsTablasMDB.Fields!longitud Then
                          lcSql = lcSql & " y La LONGITUD debería ser " & oTabla.Columns(lnFor).DefinedSize
                       End If
                       oWorkSheet.Cells(iFila, 5).Value = lcSql
                       iFila = iFila + 1
                    ElseIf oTabla.Columns(lnFor).DefinedSize <> oRsTablasMDB.Fields!longitud Then
                       oWorkSheet.Cells(iFila, 1).Value = oTabla.Name
                       oWorkSheet.Cells(iFila, 2).Value = oTabla.Columns(lnFor).Name
                       oWorkSheet.Cells(iFila, 3).Value = oRsTablasMDB.Fields!Tipo
                       oWorkSheet.Cells(iFila, 4).Value = oRsTablasMDB.Fields!longitud
                       oWorkSheet.Cells(iFila, 5).Value = "La LONGITUD debería ser " & oTabla.Columns(lnFor).DefinedSize
                       iFila = iFila + 1
                    End If
                    oRsTablasMDB.Close
              Next
           End If
        End If
    Next
    '
    iFila = iFila + 2
    oRsEstructuraMDB.Open "select * from ultimoRegistro", oConexionMDB, adOpenKeyset, adLockOptimistic
    oRsTmp.Open "select * from parametros order by idparametro desc", oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsEstructuraMDB.MoveFirst
       oRsEstructuraMDB.Find "tabla='parametros'"
       If oRsTmp.Fields!IdParametro <> Val(oRsEstructuraMDB.Fields!ultimoId) Then
            oWorkSheet.Cells(iFila, 1).Value = "Tabla Parametros, el Hospital tiene el id hasta " & oRsTmp.Fields!IdParametro
            iFila = iFila + 1
       End If
    End If
    oRsTmp.Close
    oRsTmp.Open "select * from ListBarItems order by idListItem desc", oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsEstructuraMDB.MoveFirst
       oRsEstructuraMDB.Find "tabla='listbaritems'"
       If oRsTmp.Fields!IdListItem <> Val(oRsEstructuraMDB.Fields!ultimoId) Then
            oWorkSheet.Cells(iFila, 1).Value = "Tabla ListBarItems, el Hospital tiene el id hasta " & oRsTmp.Fields!IdListItem
            iFila = iFila + 1
       End If
    End If
    oRsTmp.Close
    oRsTmp.Open "select * from ListBarReporte order by IdReporte desc", oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsEstructuraMDB.MoveFirst
       oRsEstructuraMDB.Find "tabla='listbarreporte'"
       If oRsTmp.Fields!idReporte <> Val(oRsEstructuraMDB.Fields!ultimoId) Then
            oWorkSheet.Cells(iFila, 1).Value = "Tabla ListBarReporte, el Hospital tiene el id hasta " & oRsTmp.Fields!idReporte
            iFila = iFila + 1
       End If
    End If
    oRsTmp.Close
    '
    oConexionMDB.Close
    '
    oExcel.Visible = True
    oWorkSheet.PrintPreview

End Sub

'********(En Hospital) Llena archivo "tablasYpa.mdb" con estructura actual BD
Private Sub cmdEstrucGenera_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    
    Dim oCatalogo As ADOX.Catalog
    Dim oTabla As ADOX.Table
    Dim oConexionMDB As New ADODB.Connection
    Dim oRsTablasMDB As New Recordset
    Dim oRsEstructuraMDB As New Recordset
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    Dim lnFor As Integer
    '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    '
    oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
    oRsTablasMDB.Open "delete from tablas", oConexionMDB, adOpenKeyset, adLockOptimistic
    oRsTablasMDB.Open "select * from tablas", oConexionMDB, adOpenKeyset, adLockOptimistic
    oRsEstructuraMDB.Open "delete from estructura", oConexionMDB, adOpenKeyset, adLockOptimistic
    oRsEstructuraMDB.Open "select * from estructura", oConexionMDB, adOpenKeyset, adLockOptimistic
    'Carga Tablas
    Set oCatalogo = New ADOX.Catalog
    oCatalogo.ActiveConnection = sighentidades.CadenaConexion
    For Each oTabla In oCatalogo.Tables
        If oTabla.Type = "TABLE" Then
           oRsTablasMDB.AddNew
           oRsTablasMDB.Fields!nombre = oTabla.Name
           oRsTablasMDB.Update
           For lnFor = 1 To oTabla.Columns.Count - 1
                oRsEstructuraMDB.AddNew
                oRsEstructuraMDB.Fields!tabla = oTabla.Name
                oRsEstructuraMDB.Fields!campo = oTabla.Columns(lnFor).Name
                oRsEstructuraMDB.Fields!Tipo = ObtenerTipoDeStoreProcedure(oTabla.Columns(lnFor).Type)
                oRsEstructuraMDB.Fields!longitud = oTabla.Columns(lnFor).DefinedSize
                'oRsEstructuraMDB.Fields!decimales = X
                oRsEstructuraMDB.Update
               
           Next lnFor
        End If
    Next
    oRsTablasMDB.Close
    oRsEstructuraMDB.Close
    'Por tabla
    oRsEstructuraMDB.Open "delete from ultimoRegistro", oConexionMDB, adOpenKeyset, adLockOptimistic
    oRsEstructuraMDB.Open "select * from ultimoRegistro", oConexionMDB, adOpenKeyset, adLockOptimistic
    oRsTmp.Open "select * from parametros order by idparametro desc", oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsEstructuraMDB.AddNew
       oRsEstructuraMDB.Fields!tabla = "parametros"
       oRsEstructuraMDB.Fields!ultimoId = oRsTmp.Fields!IdParametro
       oRsEstructuraMDB.Update
    End If
    oRsTmp.Close
    oRsTmp.Open "select * from listBarItems order by idListItem desc", oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsEstructuraMDB.AddNew
       oRsEstructuraMDB.Fields!tabla = "listbaritems"
       oRsEstructuraMDB.Fields!ultimoId = oRsTmp.Fields!IdListItem
       oRsEstructuraMDB.Update
    End If
    oRsTmp.Close
    oRsTmp.Open "select * from listBarReporte order by idReporte desc", oConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsEstructuraMDB.AddNew
       oRsEstructuraMDB.Fields!tabla = "listbarreporte"
       oRsEstructuraMDB.Fields!ultimoId = oRsTmp.Fields!idReporte
       oRsEstructuraMDB.Update
    End If
    oRsTmp.Close
    oRsEstructuraMDB.Close
    oConexionMDB.Close
    Unload Me
    End If
End Sub

Function ObtenerTipoDeStoreProcedure(myType As ADOX.DataTypeEnum)
Dim sType As String

    Select Case myType
    Case adTinyInt
        sType = "adTinyInt"
    Case adSmallInt
        sType = "adSmallInt"
    Case adInteger
        sType = "Int"
    Case adBigInt
        sType = "adBigInt"
    Case adUnsignedTinyInt
        sType = "adUnsignedTinyInt"
    Case adUnsignedSmallInt
        sType = "adUnsignedSmallInt"
    Case adUnsignedInt
        sType = "adUnsignedInt"
    Case adUnsignedBigInt
        sType = "adUnsignedBigInt"
    Case adSingle
        sType = "adSingle"
    Case adDouble
        sType = "Float"
    Case adCurrency
        sType = "Money"
    Case adDecimal
        sType = "adDecimal"
    Case adNumeric
        sType = "Decimal"
    Case adBoolean
        sType = "Bit"
    Case adUserDefined
        sType = "adUserDefined"
    Case adVariant
        sType = "adVariant"
    Case adGUID
        sType = "adGUID"
    Case adDate
        sType = "adDate"
    Case adDBDate
        sType = "adDBDate"
    Case adDBTime
        sType = "adDBTime"
    Case adDBTimeStamp
        sType = "DateTime"
    Case adBSTR
        sType = "adBSTR"
    Case adChar
        sType = "Char"
    Case adVarChar
        sType = "VarChar"
    Case adLongVarChar
        sType = "Text"
    Case adWChar
        sType = "adWChar"
    Case adVarWChar
        sType = "nVarChar"
    Case adLongVarWChar
        sType = "adLongVarWChar"
    Case adBinary
        sType = "adBinary"
    Case adVarBinary
        sType = "adVarBinary"
    Case adLongVarBinary
        sType = "adLongVarBinary"
    End Select
    
    ObtenerTipoDeStoreProcedure = sType

End Function


Private Sub cmdGrabaDescripcionCortaCPT_Click()

    Dim lnPrecioSIS As Double, lnPrecioSOAT As Double, lnPrecioConvenio As Double, lnPrecioESSSALUD As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtExcel1.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    Dim lcCPTcorta As String, lnPrecioPagante As Double, lnIdProducto As Long, lbContinuar As Boolean
    Dim lbCont2 As Boolean
    lnFila = 1
    lnFilaFinal = 10000
    Me.ProgressBar1.Min = lnFila
    Me.ProgressBar1.Max = lnFilaFinal
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    For lnFor = lnFila To lnFilaFinal
        DoEvents: Me.ProgressBar1.Value = lnFor: Me.Refresh
        lbCont2 = True
        lcRango = "A" + Trim(Str(lnFor))
        lcCodigo = Trim(s.Range(lcRango).Value)
        lcRango = "B" + Trim(Str(lnFor))
        lcCPTcorta = Trim(s.Range(lcRango).Value)
        lcRango = "C" + Trim(Str(lnFor))
        lnPrecioPagante = Val(s.Range(lcRango).Value)
        lcRango = "D" + Trim(Str(lnFor))
        lnPrecioSIS = Val(s.Range(lcRango).Value)
        lcRango = "E" + Trim(Str(lnFor))
        lnPrecioSOAT = Val(s.Range(lcRango).Value)
        lcRango = "F" + Trim(Str(lnFor))
        lnPrecioConvenio = Val(s.Range(lcRango).Value)
        lcRango = "G" + Trim(Str(lnFor))
        lnPrecioESSSALUD = Val(s.Range(lcRango).Value)
        If Len(Trim(lcCodigo)) > 0 And Len(lcCodigo) < 8 And Trim(lcCPTcorta) <> "" Then
            If oRsTmp.State = 1 Then
               oRsTmp.Close
            End If
            
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoServiciosXcodigo"
                Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
                Set oRsTmp = .Execute
                Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            
            lbContinuar = True
            If chkSoloActualiza.Value = 0 And oRsTmp.RecordCount = 0 Then
               lbContinuar = False
            End If

            
            
            If lbContinuar = True Then
            
                lbCont2 = True
                If chkSoloSisSoat.Value = 1 Then
                   lbCont2 = False
                End If
            
                If oRsTmp.RecordCount = 0 Then
                    If chkNohallado.Value = 1 Then
                       lcRango = "G" + Trim(Str(lnFor))
                       s.Range(lcRango).Value = "NO HALLADO"
                    End If
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosAgregarInformacion"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamOutput, 0, 1): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(lcCPTcorta, 255)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdServicioGrupo", adInteger, adParamInput, 0, 5): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdServicioSubGrupo", adInteger, adParamInput, 0, 24): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdServicioSeccion", adInteger, adParamInput, 0, 78): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@EsCPT", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@NombreMINSA", adVarChar, adParamInput, 255, Left(lcCPTcorta, 255)): .Parameters.Append oParameter
                        .Execute
                        lnIdProducto = .Parameters("@IdProducto")
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                Else
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosActualizarInformacion"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmp.Fields!idProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(lcCPTcorta, 255)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    lnIdProducto = oRsTmp.Fields!idProducto
                End If
                
                oRsTmp.Close
                
                If lbCont2 = True Then                  'particular
               
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioPagante, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
               End If
               ' If lnPrecioSIS > 0 Then
                
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProductoSIS"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 2): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioSIS, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
                'End If
                'If lnPrecioSOAT > 0 Then
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 3): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioSOAT, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
               ' End If
               If lbCont2 = True Then       'convenio
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 4): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioConvenio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
               End If
               If lnPrecioESSSALUD > 0 And Val(txtCodigoESSALUD.Text) > 10 And lbCont2 = True Then
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, CLng(txtCodigoESSALUD.Text)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioESSSALUD, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
                End If
            End If
        End If
    Next
    Set s = Nothing
    'W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    Unload Me
End Sub

Private Sub cmdImprimeBoleta_Click()
    'mgaray
    If optTicket.Value = False Then
        If Not validarAltoDeSeccionesBoleta("¿Desea Probar Impresión de todas formas?") Then
            Exit Sub
        End If
    End If

   Dim oRptBoleta As New RptBoleta
   Dim rsReporte As New Recordset
   Dim lbElDocumentoEsTicket As Boolean, lcRucDirecccionProveedor As String
   Dim lbElDocumentoEsFactura As Boolean
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   
   lcRucDirecccionProveedor = "RUC: " & Me.txtRucProv.Text & " DIRECCION: " & Me.txtDireccionProv.Text
   'lcFactura
   lbElDocumentoEsTicket = IIf(lcTipoComprobanteCaja = lcTicket, True, False)
   lbElDocumentoEsFactura = IIf(lcTipoComprobanteCaja = lcFactura, True, False)
   With rsReporte
        .Fields.Append "TotalBoleta", adCurrency
        .Fields.Append "dctos", adCurrency
        .Fields.Append "idCuentaAtencion", adInteger
        .Fields.Append "IdOrden", adInteger
        .Fields.Append "IdOrdenPago", adInteger
        .Fields.Append "IdPreVenta", adInteger
        .Fields.Append "IdCajero", adInteger
        .Fields.Append "idTipoPago", adInteger
        .Fields.Append "idEstadoComprobante", adInteger
        .Fields.Append "razonSocial", adVarChar, 50, adFldIsNullable
        .Fields.Append "NroHistoriaClinica", adInteger
        .Fields.Append "FechaCobranza", adDate
        .Fields.Append "totalPorPagar", adCurrency
        .Fields.Append "cantidad", adInteger
        .Fields.Append "precioUnitario", adCurrency
        .Fields.Append "NombreProducto", adVarChar, 150, adFldIsNullable
        .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
        .Fields.Append "exoneraciones", adCurrency
        .Fields.Append "adelantos", adCurrency
        .Fields.Append "Observaciones", adVarChar, 100, adFldIsNullable
        .Fields.Append "nombreCaja", adVarChar, 100, adFldIsNullable
        .Fields.Append "idCaja", adInteger
        .LockType = adLockOptimistic
        .Open
        .AddNew
        .Fields!TotalBoleta = Val(txtValor.Text)
        .Fields!Dctos = 0
        .Fields!idCuentaAtencion = Val(txtCuentaValor.Text)
        .Fields!IdOrden = 2
        .Fields!idOrdenPago = 3
        .Fields!idPreventa = 4
        .Fields!IdCajero = 738
        .Fields!IdTipoPago = 1
        .Fields!IdEstadoComprobante = IIf(Left(txtEstadoValor.Text, 1) = "A", 0, 1)
        .Fields!RazonSocial = txtRzSocialValor.Text
        .Fields!NroHistoriaClinica = 100100
        .Fields!FechaCobranza = CDate(txtFechaVAlor.Text)
        .Fields!TotalPorPagar = Val(txtTotalPagarVAlor.Text)
        .Fields!cantidad = Val(txtCantidadValor.Text)
        .Fields!PrecioUnitario = Val(txtPrecioValor.Text)
        .Fields!NombreProducto = txtProductoValor.Text
        .Fields!Codigo = txtCodigoValor.Text
        .Fields!Exoneraciones = Val(txtExoneracionesVAlor.Text)
        .Fields!Adelantos = Val(txtAdelantosValor.Text)
        .Fields!Observaciones = txtObservacionesValor.Text
        .Fields!nombreCaja = txtCajaVAlor.Text
        .Fields!idCaja = 0
        .Update
   End With
   '
   'mgaray
   asignarValorDeControlesAVariables
   If lcTipoServicioFarmacia = "SERVICIOS" Then
        oRptBoleta.ImprimeBoletaEnFormatoUsuario False, lcRucDirecccionProveedor, "", rsReporte, txtCajeroValor.Text, txtServicioValor.Text, Val(txtValor.Text), _
                   Val(txtExoneracionesVAlor.Text), txtCuentaValor.Text, txtAdelantosValor.Text, "123-234", "123-456", True, "901-123123", False, _
                   lbElDocumentoEsTicket, Val(txtValor.Text), 0, lbElDocumentoEsFactura, oConexion, False
   Else
        oRptBoleta.ImprimeBoletaEnFormatoUsuario False, lcRucDirecccionProveedor, "", rsReporte, txtCajeroValor.Text, txtServicioValor.Text, Val(txtValor.Text), _
                    Val(txtExoneracionesVAlor.Text), txtCuentaValor.Text, txtAdelantosValor.Text, "123-234", "123-456", True, "901-123123", True, _
                    lbElDocumentoEsTicket, Val(txtValor.Text), 0, lbElDocumentoEsFactura, oConexion, False
   End If
   '
   'oRptBoleta.ImprimeBoletaEnFormatoUsuario rsReporte, txtCajeroValor.Text, txtServicioValor.Text, Val(txtValor.Text), Val(txtExoneracionesVAlor.Text), txtCuentaValor.Text, txtAdelantosValor.Text, True, "901-123123", False
   oConexion.Close
   Set oConexion = Nothing
   Set oRptBoleta = Nothing

End Sub

Private Sub cmdListaItemsIS_Click()

    Dim oRsTmp As New Recordset
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim mo_ReporteUtil As New ReporteUtil
    Dim iFila As Integer
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    '
    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    Set oWorkBook = oExcel.Workbooks.Add
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "c:\excel.xls")
    oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    Set oWorkSheet = oWorkBook.Sheets(1)
    iFila = 5
    '
     
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "FarmMovimientoProgramasFarmAlmacenSeleccionarMovNumeroNULL"
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
     
     If oRsTmp.RecordCount > 0 Then
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           oWorkSheet.Cells(iFila, 1).Value = oRsTmp.Fields!Descripcion
           oWorkSheet.Cells(iFila, 4).Value = oRsTmp.Fields!Codigo
           oWorkSheet.Cells(iFila, 5).Value = oRsTmp.Fields!nombre
           iFila = iFila + 1
           lcNombre = oRsTmp.Fields!nombre
           lcAlmacen = oRsTmp.Fields!Descripcion
           Do While Not oRsTmp.EOF And lcNombre = oRsTmp.Fields!nombre And lcAlmacen = oRsTmp.Fields!Descripcion
              oRsTmp.MoveNext
              If oRsTmp.EOF Then
                 Exit Do
              End If
           Loop
        Loop
    End If
    oRsTmp.Close
    '
    oExcel.Visible = True
    oWorkSheet.PrintPreview
End Sub

Private Sub cmdLlenaCodigoHIS_Click()
If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
     Dim oRsTmp1 As New Recordset
     Dim oRsTmp2 As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     'Cpt
  
     With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = wxConexionRed
            .CommandTimeout = 150
            .CommandText = "PerinatalCatalogoDeProcedimientosSeleccionarTodos"
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
     End With
     Set oCommand = Nothing
     
     If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF

            With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = wxConexionRed
                  .CommandTimeout = 150
                  .CommandText = "FactCatalogoServiciosSeleccionarTodoPorIdProducto"
                  Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmp1.Fields!idProducto): .Parameters.Append oParameter
                  Set oRsTmp2 = .Execute
                  Set oRsTmp2.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
           
           If oRsTmp2.RecordCount > 0 Then

              
                With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = wxConexionRed
                      .CommandTimeout = 150
                      .CommandText = "PerinatalCatalogoCptModificarCodigoHIS"
                      Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, oRsTmp1.Fields!idModulo): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@IdLista", adInteger, adParamInput, 0, oRsTmp1.Fields!idLista): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmp1.Fields!idProducto): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 7, oRsTmp2.Fields!Codigo): .Parameters.Append oParameter
                      .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
           End If
           oRsTmp2.Close
           oRsTmp1.MoveNext
        Loop
     End If
     oRsTmp1.Close
     'Cie10
     With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = wxConexionRed
            .CommandTimeout = 150
            .CommandText = "PerinatalCatalogoDiagnosticosSeleccionarTodos"
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
     End With
     Set oCommand = Nothing
     
     If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = wxConexionRed
                  .CommandTimeout = 150
                  .CommandText = "DiagnosticosSeleccionarPorId"
                  Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oRsTmp1.Fields!IdDiagnostico): .Parameters.Append oParameter
                  Set oRsTmp2 = .Execute
                  Set oRsTmp2.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
           
           If oRsTmp2.RecordCount > 0 Then
                With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = wxConexionRed
                      .CommandTimeout = 150
                      .CommandText = "PerinatalCatalogoCie10ModificarCodigoHis"
                      Set oParameter = .CreateParameter("@IdModulo", adInteger, adParamInput, 0, oRsTmp1.Fields!idModulo): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@IdLista", adInteger, adParamInput, 0, oRsTmp1.Fields!idLista): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oRsTmp1.Fields!IdDiagnostico): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 7, Trim(Left(oRsTmp2.Fields!CodigoCIE2004, 3) + Mid(oRsTmp2.Fields!CodigoCIE2004, 5, 10))): .Parameters.Append oParameter
                      .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                
           End If
           oRsTmp2.Close
           oRsTmp1.MoveNext
        Loop
     End If
     oRsTmp1.Close
     Unload Me
End If

End Sub

Private Sub cmdNuevosEstablecimientos_Click()

        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long, lnIdEstablecimiento As Long
        Dim lcCodigoE As String
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=HIS"
        '
        Me.MousePointer = 1
        On Error Resume Next
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = wxConexionRed
            .CommandTimeout = 150
            .CommandText = "EstablecimientosSeleccionarTodos"
            Set oRsTmpOpc = .Execute
            Set oRsTmpOpc.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        
        lnIdEstablecimiento = oRsTmpOpc.Fields!IdEstablecimiento + 1
        oRsTmpOpc.Close
        lcSql = "select * from establec"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              lcCodigoE = Right("00000" & Trim(Str(oRsFox.Fields!cod_2000)), 5)
              
               With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "EstablecimientosSeleccionarTodoCamposXCodigo"
                    Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, lcCodigoE): .Parameters.Append oParameter
                    Set oRsTmpOpc = .Execute
                    Set oRsTmpOpc.ActiveConnection = Nothing
               End With
               Set oCommand = Nothing
               Set oParameter = Nothing
              
              If oRsTmpOpc.RecordCount = 0 Then

                 
                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = wxConexionRed
                     .CommandTimeout = 150
                     .CommandText = "EstablecimientosAgregar"
                     Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, lnIdEstablecimiento): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, lcCodigoE): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, Left(oRsFox.Fields!desc_estab, 150)): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, Val(oRsFox.Fields!COD_DPTO & oRsFox.Fields!COD_PROV & oRsFox.Fields!COD_DIST)): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@IdTipo", adInteger, adParamInput, 0, Val(oRsFox.Fields!TIPOESTAB)): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                     .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                 
                 lnIdEstablecimiento = lnIdEstablecimiento + 1
              End If
              oRsTmpOpc.Close
              oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        Me.MousePointer = 11
        Unload Me

End Sub

Private Sub cmdPasaMovimientos_Click()
        Me.MousePointer = 1
        On Error GoTo ErrorProceso
        If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                Dim oRsTmp As New Recordset
                '
                If lnIdPacienteActual > 0 And lnIdPacienteNuevo > 0 Then
                    mo_AdminArchivoClinico.ActualizaIdPacienteEnTodasLasTablas lnIdPacienteNuevo, lnIdPacienteActual, Val(txtHcActual.Text), 0, "", Val(txtHcNueva.Text), 0
                Else
                    MsgBox "Problemas con las HC, deben ingresar el NUMERO y pulsar ENTER"
                End If
                Me.MousePointer = 11
                Unload Me
        End If
        Exit Sub
ErrorProceso:
    oConexion.RollbackTrans
    MsgBox Err.Description
    Resume
End Sub




Private Sub cmdProcesaSip2000_Click()
    On Error GoTo ErrSip2000
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open("d:\barrantes\sisMt.xls")
    Dim s As Excel.Worksheet
    Dim lcImadre As String, lcHC As String, lcEstablecim As String, lcEdad As String, lcDistrito As String
    Dim lcEstudios As String, lcTalla As String, lcFecha As String, lcTalla_rn As String, lcPeso_rn As String
    Dim Peso1 As String, Peso2 As String, Peso3 As String, Peso4 As String, Peso5 As String
    Dim Peso6 As String, Peso7 As String, Peso8 As String, Peso9 As String, Peso10 As String
    Dim Peso11 As String, Peso12 As String, Peso13 As String, Peso14 As String, Peso15 As String
    Dim Peso16 As String, Peso17 As String, Peso18 As String, Peso19 As String, Peso20 As String
    Dim Peso21 As String, Peso22 As String, Peso23 As String, Peso24 As String, Peso25 As String
    Dim Peso26 As String, Peso27 As String, Peso28 As String, Peso29 As String, Peso30 As String
    Dim Peso31 As String, Peso32 As String, Peso33 As String, Peso34 As String, Peso35 As String
    Dim Peso36 As String, Peso37 As String, Peso38 As String, Peso39 As String, Peso40 As String
    Dim Peso41 As String, Peso42 As String, Peso43 As String, Peso44 As String, Peso45 As String
    Dim lbNuevo As Boolean, lnFor1 As Integer
    Dim lnFor As Long, lnFila As Long, lcRango As String, lnFilaFinal As Long
    Dim oConexionFox As New ADODB.Connection, oRsFox1 As New Recordset
    
    
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    '
    For lnFor1 = 1 To 2
         If lnFor1 = 1 Then
            Set s = W.Sheets("Parte 01")
         Else
            Set s = W.Sheets("Parte 02")
         End If
         lnFila = 2
         lnFilaFinal = 65500
         ProgressBar1.Min = 0: ProgressBar1.Max = lnFilaFinal
         For lnFor = lnFila To lnFilaFinal
             ProgressBar1.Value = lnFor
             lcRango = "C" + Trim(Str(lnFor))
             lcImadre = Left(Trim(s.Range(lcRango).Value) & Space(100), 30)
             If Trim(lcImadre) = "" Then
                Exit For
             End If
             lcRango = "D" + Trim(Str(lnFor))
             lcHC = Trim(s.Range(lcRango).Value)
             lcRango = "E" + Trim(Str(lnFor))
             lcEstablecim = Trim(s.Range(lcRango).Value)
             lcRango = "F" + Trim(Str(lnFor))
             lcEdad = Trim(s.Range(lcRango).Value)
             lcRango = "G" + Trim(Str(lnFor))
             lcDistrito = Trim(s.Range(lcRango).Value)
             lcRango = "H" + Trim(Str(lnFor))
             lcEstudios = Trim(s.Range(lcRango).Value)
             lcRango = "J" + Trim(Str(lnFor))
             lcTalla = Trim(s.Range(lcRango).Value)
             'lcRango = "L" + Trim(Str(lnFor))
             'lcFecha = Trim(s.Range(lcRango).Value)
             lcRango = "R" + Trim(Str(lnFor))
             lcTalla_rn = Trim(s.Range(lcRango).Value)
             lcRango = "S" + Trim(Str(lnFor))
             lcPeso_rn = Trim(s.Range(lcRango).Value)
             lcRango = "M" + Trim(Str(lnFor))
             lcEdadG = Trim(s.Range(lcRango).Value)
             Peso1 = 0: Peso2 = 0: Peso3 = 0: Peso4 = 0: Peso5 = 0
             Peso6 = 0: Peso7 = 0: Peso8 = 0: Peso9 = 0: Peso10 = 0
             Peso11 = 0: Peso12 = 0: Peso13 = 0: Peso14 = 0: Peso15 = 0
             Peso16 = 0: Peso17 = 0: Peso18 = 0: Peso19 = 0: Peso20 = 0
             Peso21 = 0: Peso22 = 0: Peso23 = 0: Peso24 = 0: Peso25 = 0
             Peso26 = 0: Peso27 = 0: Peso28 = 0: Peso29 = 0: Peso30 = 0
             Peso31 = 0: Peso32 = 0: Peso33 = 0: Peso34 = 0: Peso35 = 0
             Peso36 = 0: Peso37 = 0: Peso38 = 0: Peso39 = 0: Peso40 = 0
             Peso41 = 0: Peso42 = 0: Peso43 = 0: Peso44 = 0: Peso45 = 0
             Select Case lcEdadG
             Case "1"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso1 = Trim(s.Range(lcRango).Value)
             Case "2"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso2 = Trim(s.Range(lcRango).Value)
             Case "3"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso3 = Trim(s.Range(lcRango).Value)
             Case "4"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso4 = Trim(s.Range(lcRango).Value)
             Case "5"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso5 = Trim(s.Range(lcRango).Value)
             Case "6"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso6 = Trim(s.Range(lcRango).Value)
             Case "7"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso7 = Trim(s.Range(lcRango).Value)
             Case "8"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso8 = Trim(s.Range(lcRango).Value)
             Case "9"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso9 = Trim(s.Range(lcRango).Value)
             Case "10"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso10 = Trim(s.Range(lcRango).Value)
             Case "11"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso11 = Trim(s.Range(lcRango).Value)
             Case "12"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso12 = Trim(s.Range(lcRango).Value)
             Case "13"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso13 = Trim(s.Range(lcRango).Value)
             Case "14"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso14 = Trim(s.Range(lcRango).Value)
             Case "15"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso15 = Trim(s.Range(lcRango).Value)
             Case "16"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso16 = Trim(s.Range(lcRango).Value)
             Case "17"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso17 = Trim(s.Range(lcRango).Value)
             Case "18"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso18 = Trim(s.Range(lcRango).Value)
             Case "19"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso19 = Trim(s.Range(lcRango).Value)
             Case "20"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso20 = Trim(s.Range(lcRango).Value)
             Case "21"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso21 = Trim(s.Range(lcRango).Value)
             Case "22"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso22 = Trim(s.Range(lcRango).Value)
             Case "23"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso23 = Trim(s.Range(lcRango).Value)
             Case "24"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso24 = Trim(s.Range(lcRango).Value)
             Case "25"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso25 = Trim(s.Range(lcRango).Value)
             Case "26"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso26 = Trim(s.Range(lcRango).Value)
             Case "27"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso27 = Trim(s.Range(lcRango).Value)
             Case "28"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso28 = Trim(s.Range(lcRango).Value)
             Case "29"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso29 = Trim(s.Range(lcRango).Value)
             Case "30"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso30 = Trim(s.Range(lcRango).Value)
             Case "31"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso31 = Trim(s.Range(lcRango).Value)
             Case "32"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso32 = Trim(s.Range(lcRango).Value)
             Case "33"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso33 = Trim(s.Range(lcRango).Value)
             Case "34"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso34 = Trim(s.Range(lcRango).Value)
             Case "35"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso35 = Trim(s.Range(lcRango).Value)
             Case "36"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso36 = Trim(s.Range(lcRango).Value)
             Case "37"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso37 = Trim(s.Range(lcRango).Value)
             Case "38"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso38 = Trim(s.Range(lcRango).Value)
             Case "39"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso39 = Trim(s.Range(lcRango).Value)
             Case "40"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso40 = Trim(s.Range(lcRango).Value)
             Case "41"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso41 = Trim(s.Range(lcRango).Value)
             Case "42"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso42 = Trim(s.Range(lcRango).Value)
             Case "43"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso43 = Trim(s.Range(lcRango).Value)
             Case "44"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso44 = Trim(s.Range(lcRango).Value)
             Case "45"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso45 = Trim(s.Range(lcRango).Value)
             End Select
If Trim(lcImadre) = "M00KJO1E9935" Then
lcSql = ""
End If
             lbNuevo = False
             lcSql = "select * from sip2000 where iMadre='" & lcImadre & "'"
             If oRsFox1.State = 1 Then oRsFox1.Close
             oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
             If oRsFox1.RecordCount = 0 Then
                lbNuevo = True
             End If
            
             If lbNuevo = True Then
                oRsFox1.AddNew
                oRsFox1.Fields!iMadre = lcImadre
                oRsFox1.Fields!HC = lcHC
                oRsFox1.Fields!ESTABLECIM = lcEstablecim
                oRsFox1.Fields!Edad = Val(lcEdad)
                oRsFox1.Fields!Distrito = lcDistrito
                oRsFox1.Fields!ESTUDIOS = lcEstudios
                oRsFox1.Fields!Talla = Val(lcTalla)
              '  oRsFox1.Fields!Fecha = CDate(lcFecha)
                oRsFox1.Fields!talla_rn = Val(lcTalla_rn)
                oRsFox1.Fields!peso_rn = CDbl(lcPeso_rn)
             End If
             oRsFox1.Fields!Peso1 = CDbl(Peso1)
             oRsFox1.Fields!Peso2 = CDbl(Peso2)
             oRsFox1.Fields!Peso3 = CDbl(Peso3)
             oRsFox1.Fields!Peso4 = CDbl(Peso4)
             oRsFox1.Fields!Peso5 = CDbl(Peso5)
             oRsFox1.Fields!Peso6 = CDbl(Peso6)
             oRsFox1.Fields!Peso7 = CDbl(Peso7)
             oRsFox1.Fields!Peso8 = CDbl(Peso8)
             oRsFox1.Fields!Peso9 = CDbl(Peso9)
             oRsFox1.Fields!Peso10 = CDbl(Peso10)
             oRsFox1.Fields!Peso11 = CDbl(Peso11)
             oRsFox1.Fields!Peso12 = CDbl(Peso12)
             oRsFox1.Fields!Peso13 = CDbl(Peso13)
             oRsFox1.Fields!Peso14 = CDbl(Peso14)
             oRsFox1.Fields!Peso15 = CDbl(Peso15)
             oRsFox1.Fields!Peso16 = CDbl(Peso16)
             oRsFox1.Fields!Peso17 = CDbl(Peso17)
             oRsFox1.Fields!Peso18 = CDbl(Peso18)
             oRsFox1.Fields!Peso19 = CDbl(Peso19)
             oRsFox1.Fields!Peso20 = CDbl(Peso20)
             oRsFox1.Fields!Peso21 = CDbl(Peso21)
             oRsFox1.Fields!Peso22 = CDbl(Peso22)
             oRsFox1.Fields!Peso23 = CDbl(Peso23)
             oRsFox1.Fields!Peso24 = CDbl(Peso24)
             oRsFox1.Fields!Peso25 = CDbl(Peso25)
             oRsFox1.Fields!Peso26 = CDbl(Peso26)
             oRsFox1.Fields!Peso27 = CDbl(Peso27)
             oRsFox1.Fields!Peso28 = CDbl(Peso28)
             oRsFox1.Fields!Peso29 = CDbl(Peso29)
             oRsFox1.Fields!Peso30 = CDbl(Peso30)
             oRsFox1.Fields!Peso31 = CDbl(Peso31)
             oRsFox1.Fields!Peso32 = CDbl(Peso32)
             oRsFox1.Fields!Peso33 = CDbl(Peso33)
             oRsFox1.Fields!Peso34 = CDbl(Peso34)
             oRsFox1.Fields!Peso35 = CDbl(Peso35)
             oRsFox1.Fields!Peso36 = CDbl(Peso36)
             oRsFox1.Fields!Peso37 = CDbl(Peso37)
             oRsFox1.Fields!Peso38 = CDbl(Peso38)
             oRsFox1.Fields!Peso39 = CDbl(Peso39)
             oRsFox1.Fields!Peso40 = CDbl(Peso40)
             oRsFox1.Fields!Peso41 = CDbl(Peso41)
             oRsFox1.Fields!Peso42 = CDbl(Peso42)
             oRsFox1.Fields!Peso43 = CDbl(Peso43)
             oRsFox1.Fields!Peso44 = CDbl(Peso44)
             oRsFox1.Fields!Peso45 = CDbl(Peso45)
             oRsFox1.Update
        Next
   Next
   Unload Me
   Exit Sub
ErrSip2000:
   MsgBox Err.Description
   Resume
End Sub

Private Sub cmdProcesaHistoriasConDNI_Click()
    If lcBuscaParametro.SeleccionaFilaParametro(351) = "S" Then
       If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            Dim oConexion As New Connection
            Dim oConexionExterna As New Connection
            Dim oRsTmp0 As New Recordset
            Dim oRsTmp1 As New Recordset
            Dim wxNueve As String, lcNuevaHC As String
            wxNueve = "9"
            Me.MousePointer = 11
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            oConexionExterna.CommandTimeout = 900
            oConexionExterna.CursorLocation = adUseClient
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            lcSql = "select top " & txtNroHistoriasXdia.Text & " *  from pacientes where IdTipoNumeracion<3 and IdDocIdentidad=1 and " & _
                    " len(NroDocumento)>2  and convert(int,NroDocumento)<>(NroHistoriaClinica-900000000) "
            oRsTmp0.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            If oRsTmp0.RecordCount > 0 Then
               oRsTmp0.MoveFirst
               Do While Not oRsTmp0.EOF
                  lcNuevaHC = wxNueve & Right("00000000" & Trim(oRsTmp0!NroDocumento), 8)
                  Set oRsTmp1 = mo_ReglasAdmision.PacientesXnroHistoriaTipoNumeracion(Val(lcNuevaHC), oRsTmp0!IdTipoNumeracion, oConexion)
                  If oRsTmp1.RecordCount = 0 Then
                     mo_ReglasAdmision.PasaNuevaHC Trim(Str(oRsTmp0!NroHistoriaClinica)), lcNuevaHC, oConexion, oConexionExterna

                  End If
                  oRsTmp1.Close
                  '
                  oRsTmp0.MoveNext
               Loop
            End If
            oRsTmp0.Close
            oConexion.Close
            oConexionExterna.Close
            Set oConexion = Nothing
            Set oConexionExterna = Nothing
            Set oRsTmp0 = Nothing
            Set oRsTmp1 = Nothing
            Unload Me
       End If
    Else
       MsgBox "El parametro 351 debe ser S", vbInformation, ""
    End If

End Sub

Private Sub cmdRecienNacido_Click()
    Me.MousePointer = 11
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    mo_AdminArchivoClinico.ActualizaIdRecienNacido 0, oConexion
    oConexion.Close
    Set oConexion = Nothing
    Me.MousePointer = 1
    Unload Me
End Sub

Private Sub Command_Click()
    mo_ReglasAdmision.ActualizaHistoriaIgualDNI ""
    Unload Me

End Sub

Private Sub Command1_Click()
    If Val(txtNroHistoriaActual.Text) = 0 Then
       MsgBox "ingrese el N° Historia ACTUAL"
       Exit Sub
    End If
    If Val(TxtNroHistoriaNew.Text) = 0 Then
       MsgBox "ingrese el N° Historia NUEVA"
       Exit Sub
    End If
    
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_ReglasAdmision.PacientesSeleccionarPorNroHistoria(Val(TxtNroHistoriaNew.Text))
    If oRsTmp.RecordCount > 0 Then
       
       MsgBox "La nueva HISTORIA ya existe para: " & oRsTmp!ApellidoPaterno & " " & oRsTmp!ApellidoMaterno & " " & oRsTmp!PrimerNombre
       oRsTmp.Close
       Exit Sub
    End If
    oRsTmp.Close
    
    On Error GoTo ErrorC7
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        'PasaNuevaHC
        Dim oConexion As New Connection
        Dim oConexionExterna As New Connection
        oConexionExterna.CommandTimeout = 900
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)

        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
        mo_ReglasAdmision.PasaNuevaHC txtNroHistoriaActual.Text, TxtNroHistoriaNew.Text, oConexion, oConexionExterna
        oConexion.Close
        oConexionExterna.Close
        Set oConexionExterna = Nothing
        Set oConexion = Nothing
        
        Unload Me
    End If
    Exit Sub
ErrorC7:
    MsgBox Err.Description
End Sub

'Sub PasaNuevaHC()
'        Dim oCommand As New ADODB.Command
'        Dim oParameter As ADODB.Parameter
'        Dim oRsTmp As New Recordset
'        Dim oConexion As New Connection
'        oConexion.Open sighentidades.CadenaConexion
'        oConexion.BeginTrans
'
'        With oCommand
'            .CommandType = adCmdStoredProc
'            Set .ActiveConnection = oConexion
'            .CommandTimeout = 150
'            .CommandText = "HistoriasClinicasActualizarNroHistoriaClinica"
'            Set oParameter = .CreateParameter("@NroHistoriaNew", adInteger, adParamInput, 0, CLng(TxtNroHistoriaNew.Text)): .Parameters.Append oParameter
'            Set oParameter = .CreateParameter("@NroHistoriaActual", adInteger, adParamInput, 0, CLng(txtNroHistoriaActual.Text)): .Parameters.Append oParameter
'            Set oRsTmp = .Execute
'        End With
'        Set oCommand = Nothing
'        Set oParameter = Nothing
'
'        With oCommand
'            .CommandType = adCmdStoredProc
'            Set .ActiveConnection = oConexion
'            .CommandTimeout = 150
'            .CommandText = "PacientesActualizarNroHistoriaClinica"
'            Set oParameter = .CreateParameter("@NroHistoriaNew", adInteger, adParamInput, 0, CLng(TxtNroHistoriaNew.Text)): .Parameters.Append oParameter
'            Set oParameter = .CreateParameter("@NroHistoriaActual", adInteger, adParamInput, 0, CLng(txtNroHistoriaActual.Text)): .Parameters.Append oParameter
'            Set oRsTmp = .Execute
'        End With
'        Set oCommand = Nothing
'        Set oParameter = Nothing
'
'        oConexion.CommitTrans
'        Set oConexion = Nothing
'
'End Sub


'amarra Partidas vs CPT



Private Sub CommandCuatro_Click()
    On Error GoTo ErSip2000
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Set W = EXL.Workbooks.Open(App.Path & "\archivos\percentiles.xls")       'usa
    Set s = W.Sheets("IMC")
    '
    Dim oConexionMDB As New Connection, oRsMDB As New Recordset
    oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
    '
    Dim lcImadre As String, lcHC As String, lcEstablecim As String, lcEdad As String, lcDistrito As String
    Dim lcEstudios As String, lcTalla As String, lcFecha As String, lcTalla_rn As String, lcPeso_rn As String
    Dim Peso1 As String, Peso2 As String, Peso3 As String, Peso4 As String, Peso5 As String
    Dim Peso6 As String, Peso7 As String, Peso8 As String, Peso9 As String, Peso10 As String
    Dim Peso11 As String, Peso12 As String, Peso13 As String, Peso14 As String, Peso15 As String
    Dim Peso16 As String, Peso17 As String, Peso18 As String, Peso19 As String, Peso20 As String
    Dim Peso21 As String, Peso22 As String, Peso23 As String, Peso24 As String, Peso25 As String
    Dim Peso26 As String, Peso27 As String, Peso28 As String, Peso29 As String, Peso30 As String
    Dim Peso31 As String, Peso32 As String, Peso33 As String, Peso34 As String, Peso35 As String
    Dim Peso36 As String, Peso37 As String, Peso38 As String, Peso39 As String, Peso40 As String
    Dim Peso41 As String, Peso42 As String, Peso43 As String, Peso44 As String, Peso45 As String
    Dim lbNuevo As Boolean, lnFor1 As Integer
    Dim lnFor As Long, lnFila As Long, lcRango As String, lnFilaFinal As Long
    Dim oRsFox1 As New Recordset, lnPercentilIMC As Double, lcPercentilIMC As String
    '
    Dim oConexionFox As New Connection
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    
    
    lcSql = "delete from sip"
    If oRsFox1.State = 1 Then oRsFox1.Close
    oRsFox1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    '
         lcSql = "select * from princip"
         oRsMDB.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
         '
         lnFila = 1
         lnFilaFinal = oRsMDB.RecordCount
         oRsMDB.MoveFirst
         ProgressBar1.Min = 0: ProgressBar1.Max = lnFilaFinal
         For lnFor = lnFila To lnFilaFinal
             ProgressBar1.Value = lnFor
             lcImadre = Left(oRsMDB.Fields!IDMADRE & Space(30), 30)
If Trim(lcImadre) = "000X2AWB4763" Then
lcSql = ""
End If
             If Trim(lcImadre) = "" Then
                Exit For
             End If
             lcHC = oRsMDB.Fields!HC
             lcEstablecim = oRsMDB.Fields!ESTABLECIM
             lcEdad = oRsMDB.Fields!Edad
             lcDistrito = IIf(IsNull(oRsMDB.Fields!Distrito), "", oRsMDB.Fields!Distrito)
             lcEstudios = oRsMDB.Fields!ESTUDIOS
             lcTalla = oRsMDB.Fields!Talla
             lcTalla_rn = oRsMDB.Fields!talla_ra
             lcPeso_rn = oRsMDB.Fields!peso_rn
             lcEdadG = oRsMDB.Fields!EDAD_GESTA
             If sighentidades.EsFecha(oRsMDB.Fields!Fecha, "DD/MM/AAAA") = True Then
                ldFecha = oRsMDB.Fields!Fecha
             End If
             Peso1 = 0: Peso2 = 0: Peso3 = 0: Peso4 = 0: Peso5 = 0
             Peso6 = 0: Peso7 = 0: Peso8 = 0: Peso9 = 0: Peso10 = 0
             Peso11 = 0: Peso12 = 0: Peso13 = 0: Peso14 = 0: Peso15 = 0
             Peso16 = 0: Peso17 = 0: Peso18 = 0: Peso19 = 0: Peso20 = 0
             Peso21 = 0: Peso22 = 0: Peso23 = 0: Peso24 = 0: Peso25 = 0
             Peso26 = 0: Peso27 = 0: Peso28 = 0: Peso29 = 0: Peso30 = 0
             Peso31 = 0: Peso32 = 0: Peso33 = 0: Peso34 = 0: Peso35 = 0
             Peso36 = 0: Peso37 = 0: Peso38 = 0: Peso39 = 0: Peso40 = 0
             Peso41 = 0: Peso42 = 0: Peso43 = 0: Peso44 = 0: Peso45 = 0
             '
             lcSql = ".."
             lnPercentilIMC = 0
             lcPercentilIMC = "ERR"
             If oRsMDB.Fields!Peso > 0 And oRsMDB.Fields!Talla > 0 Then
                s.Cells(203, 6).Value = oRsMDB.Fields!Peso
                s.Cells(205, 6).Value = Round(oRsMDB.Fields!Talla / 100, 2)
                s.Cells(209, 6).Value = lcEdadG
                lcSql = "percentil"
                lcPercentilIMC = s.Cells(211, 6).Value
                lcSql = ".."
                lnPercentilIMC = IIf(UCase(Left(lcPercentilIMC, 3)) = "ERR", 0, Val(lcPercentilIMC))
             End If
             '
             ldFecEmbarazo = ldFecha - (Val(lcEdadG) * 7)
             Select Case lcEdadG
             Case "1"
                     Peso1 = lnPercentilIMC         'oRsMDB.Fields!Peso
                     
             Case "2"
                     Peso2 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "3"
                     Peso3 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "4"
                     Peso4 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "5"
                     Peso5 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "6"
                     Peso6 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "7"
                     Peso7 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "8"
                     Peso8 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "9"
                     Peso9 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "10"
                     Peso10 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "11"
                     Peso11 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "12"
                     Peso12 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "13"
                     Peso13 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "14"
                     Peso14 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "15"
                     Peso15 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "16"
                     Peso16 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "17"
                     Peso17 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "18"
                     Peso18 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "19"
                     Peso19 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "20"
                     Peso20 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "21"
                     Peso21 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "22"
                     Peso22 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "23"
                     Peso23 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "24"
                     Peso24 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "25"
                     Peso25 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "26"
                     Peso26 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "27"
                     Peso27 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "28"
                     Peso28 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "29"
                     Peso29 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "30"
                     Peso30 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "31"
                     Peso31 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "32"
                     Peso32 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "33"
                     Peso33 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "34"
                     Peso34 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "35"
                     Peso35 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "36"
                     Peso36 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "37"
                     Peso37 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "38"
                     Peso38 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "39"
                     Peso39 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "40"
                     Peso40 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "41"
                     Peso41 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "42"
                     Peso42 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "43"
                     Peso43 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "44"
                     Peso44 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "45"
                     Peso45 = lnPercentilIMC         'oRsMDB.Fields!Peso
             End Select
If Trim(lcImadre) = "1513KULW7812" Then
lcSql = ""
End If
             lbNuevo = False
             lcSql = "select * from sip where iMadre='" & lcImadre & "'"
             If oRsFox1.State = 1 Then oRsFox1.Close
             oRsFox1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
             If oRsFox1.RecordCount = 0 Then
                lbNuevo = True
             End If
            
             If lbNuevo = True Then
                oRsFox1.AddNew
                oRsFox1.Fields!iMadre = Left(lcImadre, 30)
                oRsFox1.Fields!HC = Left(lcHC, 20)
                oRsFox1.Fields!ESTABLECIM = Left(lcEstablecim, 100)
                oRsFox1.Fields!Edad = Val(lcEdad)
                oRsFox1.Fields!Distrito = Left(lcDistrito, 10)
                oRsFox1.Fields!ESTUDIOS = Left(lcEstudios, 60)
                oRsFox1.Fields!Talla = Val(lcTalla)
              '  oRsFox1.Fields!Fecha = CDate(lcFecha)
                oRsFox1.Fields!talla_rn = Val(lcTalla_rn)
                oRsFox1.Fields!peso_rn = CDbl(lcPeso_rn)
                If Year(ldFecEmbarazo) > 1980 Then
                   oRsFox1.Fields!Fembarazo = ldFecEmbarazo
                End If
             Else
                If ldFecEmbarazo < oRsFox1.Fields!Fembarazo Then
                   If Year(ldFecEmbarazo) > 1980 Then
                      oRsFox1.Fields!Fembarazo = ldFecEmbarazo
                   End If
                End If
             End If
             If Peso1 > 0 Then
                oRsFox1.Fields!Peso1 = CDbl(Peso1)
             End If
             If Peso1 > 0 Then
                oRsFox1.Fields!Peso2 = CDbl(Peso2)
             End If
             If Peso3 > 0 Then
                oRsFox1.Fields!Peso3 = CDbl(Peso3)
             End If
             If Peso4 > 0 Then
                oRsFox1.Fields!Peso4 = CDbl(Peso4)
             End If
             If Peso5 > 0 Then
                oRsFox1.Fields!Peso5 = CDbl(Peso5)
             End If
             If Peso6 > 0 Then
                oRsFox1.Fields!Peso6 = CDbl(Peso6)
             End If
             If Peso7 > 0 Then
                oRsFox1.Fields!Peso7 = CDbl(Peso7)
             End If
             If Peso8 > 0 Then
                oRsFox1.Fields!Peso8 = CDbl(Peso8)
             End If
             If Peso9 > 0 Then
                oRsFox1.Fields!Peso9 = CDbl(Peso9)
             End If
             If Peso10 > 0 Then
                oRsFox1.Fields!Peso10 = CDbl(Peso10)
             End If
             If Peso11 > 0 Then
                oRsFox1.Fields!Peso11 = CDbl(Peso11)
             End If
             If Peso12 > 0 Then
                oRsFox1.Fields!Peso12 = CDbl(Peso12)
             End If
             If Peso13 > 0 Then
                oRsFox1.Fields!Peso13 = CDbl(Peso13)
             End If
             If Peso14 > 0 Then
                oRsFox1.Fields!Peso14 = CDbl(Peso14)
             End If
             If Peso15 > 0 Then
                oRsFox1.Fields!Peso15 = CDbl(Peso15)
             End If
             If Peso16 > 0 Then
                oRsFox1.Fields!Peso16 = CDbl(Peso16)
             End If
             If Peso17 > 0 Then
                oRsFox1.Fields!Peso17 = CDbl(Peso17)
             End If
             If Peso18 > 0 Then
                oRsFox1.Fields!Peso18 = CDbl(Peso18)
             End If
             If Peso19 > 0 Then
                oRsFox1.Fields!Peso19 = CDbl(Peso19)
             End If
             If Peso20 > 0 Then
                oRsFox1.Fields!Peso20 = CDbl(Peso20)
             End If
             If Peso21 > 0 Then
                oRsFox1.Fields!Peso21 = CDbl(Peso21)
             End If
             If Peso22 > 0 Then
                oRsFox1.Fields!Peso22 = CDbl(Peso22)
             End If
             If Peso23 > 0 Then
                oRsFox1.Fields!Peso23 = CDbl(Peso23)
             End If
             If Peso24 > 0 Then
                oRsFox1.Fields!Peso24 = CDbl(Peso24)
             End If
             If Peso25 > 0 Then
                oRsFox1.Fields!Peso25 = CDbl(Peso25)
             End If
             If Peso26 > 0 Then
                oRsFox1.Fields!Peso26 = CDbl(Peso26)
             End If
             If Peso27 > 0 Then
                oRsFox1.Fields!Peso27 = CDbl(Peso27)
             End If
             If Peso28 > 0 Then
                oRsFox1.Fields!Peso28 = CDbl(Peso28)
             End If
             If Peso29 > 0 Then
                oRsFox1.Fields!Peso29 = CDbl(Peso29)
             End If
             If Peso30 > 0 Then
                oRsFox1.Fields!Peso30 = CDbl(Peso30)
             End If
             If Peso31 > 0 Then
                oRsFox1.Fields!Peso31 = CDbl(Peso31)
             End If
             If Peso32 > 0 Then
                oRsFox1.Fields!Peso32 = CDbl(Peso32)
             End If
             If Peso33 > 0 Then
                oRsFox1.Fields!Peso33 = CDbl(Peso33)
             End If
             If Peso34 > 0 Then
                oRsFox1.Fields!Peso34 = CDbl(Peso34)
             End If
             If Peso35 > 0 Then
                oRsFox1.Fields!Peso35 = CDbl(Peso35)
             End If
             If Peso36 > 0 Then
                oRsFox1.Fields!Peso36 = CDbl(Peso36)
             End If
             If Peso37 > 0 Then
                oRsFox1.Fields!Peso37 = CDbl(Peso37)
             End If
             If Peso38 > 0 Then
                oRsFox1.Fields!Peso38 = CDbl(Peso38)
             End If
             If Peso39 > 0 Then
                oRsFox1.Fields!Peso39 = CDbl(Peso39)
             End If
             If Peso40 > 0 Then
                oRsFox1.Fields!Peso40 = CDbl(Peso40)
             End If
             If Peso41 > 0 Then
                oRsFox1.Fields!Peso41 = CDbl(Peso41)
             End If
             If Peso42 > 0 Then
                oRsFox1.Fields!Peso42 = CDbl(Peso42)
             End If
             If Peso43 > 0 Then
                oRsFox1.Fields!Peso43 = CDbl(Peso43)
             End If
             If Peso44 > 0 Then
                oRsFox1.Fields!Peso44 = CDbl(Peso44)
             End If
             If Peso45 > 0 Then
                 oRsFox1.Fields!Peso45 = CDbl(Peso45)
             End If
             oRsFox1.Update
             oRsMDB.MoveNext
        Next
   Unload Me
   Exit Sub
ErSip2000:
   If Err.Number = 13 And lcSql = "percentil" Then
      Resume Next
   Else
      MsgBox Err.Description
      Resume
   End If

End Sub




Private Sub Command2_Click()
        If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            Me.MousePointer = 11
            Dim oConexion As New ADODB.Connection
            '
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            mo_AdminArchivoClinico.ActualizaIDenTablaLabResultadoPorItems oConexion, 0
            oConexion.Close
            Set oConexion = Nothing
            Unload Me
        End If
End Sub

Private Sub Command6_Click()

    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 11
        On Error GoTo ErrXXX
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oRsFoxHis1 As New Recordset
        Dim oRsFoxHisA As New Recordset
        Dim oRsFoxLaboratorio As New Recordset
        Dim oConexionFox As New Connection
        Const lcMaterno As String = "301605"
        Const lcNino As String = "303712"
        Const lnDNImaximo As Integer = 15
        Dim lcMamaDNI As String, lnNrosDni As Integer, lcVariasMamaDNI As String
        Dim lnDCI As Double, lnDGI As Double, lnDAI As Double
        Dim lnEdadG As Integer, lcDNI As String
        Dim lnAguaSegura As Integer, lnLME As Integer, lnLM As Integer, lnAC As Integer
        Dim lnRciu As Integer, lnAcidoFolico As Integer, lnSulfatoFerroso As Integer, lnClampaje As Integer, lnPinstitucional As Integer
        Dim lnControl As Integer, ldFechaParto As Date, ldFechaFur As Date, lcUltimo As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        '
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
  
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
''
'     'crea tmp para Errores
''
        lcVariasMamaDNI = "": lnNrosDni = 0
       
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "PacientesSeleccionarPorLongitudNroDocumento8IdTipoSexo2"
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           lnNrosDni = 0
           Do While Not oRsTmp1.EOF
              If (Year(Date) - Year(oRsTmp1.Fields!FechaNacimiento)) > 15 And (Year(Date) - Year(oRsTmp1.Fields!FechaNacimiento) <= 40) Then
                 lcVariasMamaDNI = lcVariasMamaDNI & oRsTmp1.Fields!NroDocumento & "/"
                 lnNrosDni = lnNrosDni + 1
              End If
              If lnNrosDni >= lnDNImaximo Then
                 Exit Do
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        lcSql = "select * from his1"
        oRsFoxHis1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        
        lcSql = "select * from laboratorio"
        oRsFoxLaboratorio.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFoxHis1.RecordCount > 0 Then
           Me.ProgressBar1.Max = oRsFoxHis1.RecordCount
           Me.ProgressBar1.Min = 0
           Me.ProgressBar1.Value = 0
           oRsFoxHis1.MoveFirst
           Do While Not oRsFoxHis1.EOF
              Me.Refresh: Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
              If oRsFoxHis1.Fields!cod_servsa = lcMaterno Or oRsFoxHis1.Fields!cod_servsa = lcNino Then
                 lcSql = "select * from hisa where ano=" & oRsFoxHis1.Fields!ano & _
                          "  and mes=" & oRsFoxHis1.Fields!Mes & " and nom_lote='" & oRsFoxHis1.Fields!nom_lote & "'" & _
                          " and (ncontrol is null) and (edadg is null)"
                 If oRsFoxHisA.State = 1 Then oRsFoxHisA.Close
                 oRsFoxHisA.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
                 lnRegHisA = oRsFoxHisA.RecordCount
              
              
                 If lnRegHisA > 0 Then
                    lnRegHisA = 0
                    oRsFoxHisA.MoveFirst
                    Do While Not oRsFoxHisA.EOF
                       lnRegHisA = lnRegHisA + 1
                       lcSql = "Error"
                       If oRsFoxHisA.Fields!ano = oRsFoxHis1.Fields!ano And _
                                           oRsFoxHisA.Fields!Mes = oRsFoxHis1.Fields!Mes And _
                                           oRsFoxHisA.Fields!nom_lote = oRsFoxHis1.Fields!nom_lote Then
                          If lcSql <> "SIGUIENTE" Then
                               lcSql = ""
                               If oRsFoxHis1.Fields!cod_servsa = lcNino And oRsFoxHisA.Fields!Edad <= 9 Then
                                  'niño
                                  lcDNI = Trim(Mid(oRsFoxHisA.Fields!DNI, 5, 8))
                                  lcMamaDNI = ""
                                  If Len(lcDNI) <> 8 Then
                                     lcMamaDNI = EligeDNImamaYnroHijo(lnNrosDni, lcVariasMamaDNI, lnDNImaximo)
                                  End If
                                  '
                                  lnDCI = 0: lnDGI = 0: lnDAI = 0
                                  LlenaPercentiles lnDCI, lnDGI, lnDAI, oRsFoxHisA.Fields!Edad, oRsFoxHisA.Fields!tip_edad, _
                                                   lnNrosDni, lnDNImaximo, oRsFoxHisA.Fields!sexo
                                  '
                                  lnEdadG = RetornaEdadGestacionalEnMeses(oRsFoxHisA.Fields!tip_edad, oRsFoxHisA.Fields!Edad)
                                  LlenaFlags lnAguaSegura, lnLME, lnLM, lnAC, lnNrosDni, lnDNImaximo
                                  '
                                  oRsFoxHisA.Fields!dniMadre = lcMamaDNI
                                  oRsFoxHisA.Fields!dci = lnDCI
                                  oRsFoxHisA.Fields!dgi = lnDGI
                                  oRsFoxHisA.Fields!dai = lnDAI
                                  oRsFoxHisA.Fields!edadg = lnEdadG
                                  oRsFoxHisA.Fields!aguaSegura = lnAguaSegura
                                  oRsFoxHisA.Fields!lme = lnLME
                                  oRsFoxHisA.Fields!lm = lnLM
                                  oRsFoxHisA.Fields!ac = lnAC
                                  oRsFoxHisA.Fields!lm = lnLM
                                  oRsFoxHisA.Update
                                  llenaLaboratorio oRsFoxLaboratorio, lnNrosDni, lcDNI, lcMamaDNI, oRsFoxHisA.Fields!cod_2000, _
                                                   oRsFoxHisA.Fields!ano, oRsFoxHisA.Fields!Mes, oRsFoxHisA.Fields!nom_lote, _
                                                   oRsFoxHisA.Fields!dia, False
                               ElseIf oRsFoxHis1.Fields!cod_servsa = lcMaterno And oRsFoxHisA.Fields!Edad >= 15 And oRsFoxHisA.Fields!Edad <= 30 Then
                                  'maternidad
                                  lcDNI = Trim(Mid(oRsFoxHisA.Fields!DNI, 5, 8))
                                  lcMamaDNI = ""
                                  If Len(lcDNI) <> 8 Then
                                     lcMamaDNI = EligeDNImamaYnroHijo(lnNrosDni, lcVariasMamaDNI, lnDNImaximo)
                                  End If
                                  RetornaDatosM lnControl, ldFechaParto, ldFechaFur, lnNrosDni, oRsFoxHisA.Fields!dia, _
                                                oRsFoxHisA.Fields!Mes, oRsFoxHisA.Fields!ano
                                  RetornaFlasM lnRciu, lnAcidoFolico, lnSulfatoFerroso, lnClampaje, lnPinstitucional, _
                                               lnNrosDni, lnDNImaximo
                                  '
                                  oRsFoxHisA.Fields!fechaParto = ldFechaParto
                                  oRsFoxHisA.Fields!fecha_Fur = ldFechaFur
                                  oRsFoxHisA.Fields!nControl = lnControl
                                  oRsFoxHisA.Fields!rCiu = lnRciu
                                  oRsFoxHisA.Fields!acidoFolic = lnAcidoFolico
                                  oRsFoxHisA.Fields!sulfatoFer = lnSulfatoFerroso
                                  oRsFoxHisA.Fields!clampAje = lnClampaje
                                  oRsFoxHisA.Fields!pInstituci = lnPinstitucional
                                  oRsFoxHisA.Update
                                  llenaLaboratorio oRsFoxLaboratorio, lnNrosDni, lcDNI, lcMamaDNI, oRsFoxHisA.Fields!cod_2000, _
                                                   oRsFoxHisA.Fields!ano, oRsFoxHisA.Fields!Mes, oRsFoxHisA.Fields!nom_lote, _
                                                   oRsFoxHisA.Fields!dia, True
                               End If
                            End If
                       End If
                       oRsFoxHisA.MoveNext
                    Loop
                 End If
              End If
              oRsFoxHis1.MoveNext
           Loop
        End If
        oRsFoxHis1.Close
        oRsFoxHisA.Close
        oRsFoxLaboratorio.Close
        oConexionFox.Close
        oConexion.Close
        Set oConexion = Nothing
        Unload Me
    End If
    Exit Sub
ErrXXX:
    If lcSql = "Error" Then
        lcSql = "SIGUIENTE"
        Resume Next
    Else
        MsgBox Err.Description
        Resume
    End If
End Sub
Function EligeDNImamaYnroHijo(ByRef lnNrosDni As Integer, lcMamaDNI As String, lnDNImaximo As Integer) As String
     Dim lnFor As Integer
     lnFor = 1
     lnpos = 1
     Do While True
        EligeDNImamaYnroHijo = Mid(lcMamaDNI, lnpos, 8) & Left(Trim(Str(lnNrosDni)), 1)
        If lnNrosDni = lnFor Then
           Exit Do
        End If
        lnpos = lnpos + 9
        lnFor = lnFor + 1
     Loop
     If lnNrosDni < lnDNImaximo Then
        lnNrosDni = lnNrosDni + 1
     Else
        lnNrosDni = 1
     End If
End Function
Sub LlenaPercentiles(ByRef lnDCI As Double, ByRef lnDGI As Double, ByRef lnDAI As Double, lnEdad As Integer, lcTipoEdad As String, _
                     ByRef lnNrosDni As Integer, lnDNImaximo As Integer, lnSexo As String)
     Const lcPercentilDCIf As String = "57/29/99/77/36"
     Const lcPercentilDCIm As String = "48/99/45/77/26"
     Const lcPercentilDGIf As String = "58/57/34/40/45"
     Const lcPercentilDGIm As String = "53/48/33/40/45"
     Const lcPercentilDAIf As String = "07/05/01/09/03"
     Const lcPercentilDAIm As String = "06/09/02/01/03"
     If lcTipoEdad = "A" And lnEdad > 4 Then
     Else
        Dim lnFor As Integer
        lnFor = 1
        lnpos = 1
        Do While True
           If lnSexo = "M" Then
              lnDCI = Val(Mid(lcPercentilDCIm, lnpos, 2))
              lnDGI = Val(Mid(lcPercentilDGIm, lnpos, 2))
              lnDAI = Val(Mid(lcPercentilDAIm, lnpos, 2))
           Else
              lnDCI = Val(Mid(lcPercentilDCIf, lnpos, 2))
              lnDGI = Val(Mid(lcPercentilDGIf, lnpos, 2))
              lnDAI = Val(Mid(lcPercentilDAIf, lnpos, 2))
           End If
           If lnNrosDni = lnFor Then
              Exit Do
           End If
           lnpos = lnpos + 3
           lnFor = lnFor + 1
        Loop
     End If
     If lnNrosDni < lnDNImaximo Then
        lnNrosDni = lnNrosDni + 1
     Else
        lnNrosDni = 1
     End If
End Sub
Function RetornaEdadGestacionalEnMeses(lcTip_edad As String, lnEdad As Integer) As Integer
    RetornaEdadGestacionalEnMeses = 0
    Select Case lcTip_edad
    Case "A"
         RetornaEdadGestacionalEnMeses = lnEdad * 12
    Case "M"
         RetornaEdadGestacionalEnMeses = lnEdad
    Case "D"
         RetornaEdadGestacionalEnMeses = 1
    End Select

End Function
Sub LlenaFlags(ByRef lnAguaSegura As Integer, ByRef lnLME As Integer, ByRef lnLM As Integer, ByRef lnAC As Integer, _
               ByRef lnNrosDni As Integer, lnDNImaximo As Integer)
     lnAguaSegura = 1: lnLME = 2: lnLM = 1: lnAC = 2
     Dim lnFor As Integer
     If lnNrosDni <> lnDNImaximo Then
        For lnFor = 1 To (lnDNImaximo - lnNrosDni)
            lnAguaSegura = IIf(lnAguaSegura = 1, 2, 1)
            lnLME = IIf(lnLME = 1, 2, 1)
            lnLM = IIf(lnLM = 1, 2, 1)
            lnAC = IIf(lnAC = 1, 2, 1)
        Next
     End If
     
     If lnNrosDni < lnDNImaximo Then
        lnNrosDni = lnNrosDni + 1
     Else
        lnNrosDni = 1
     End If
End Sub
Sub RetornaDatosM(ByRef lnControl As Integer, ByRef ldFechaParto As Date, ByRef ldFechaFur As Date, lnNrosDni As Integer, _
                  lnDia As Integer, lnMes As Integer, lnAnio As Integer)
    lnControl = lnNrosDni
    If lnNrosDni > 9 Then
       lnControl = 1
    End If
    Dim ldAtencion As Date
    
    ldAtencion = CDate(Right("0" & Trim(Str(lnDia)), 2) & "/" & Right("0" & Trim(Str(lnMes)), 2) & "/" & Trim(Str(lnAnio)))
    ldFechaParto = ldAtencion + ((9 - lnControl) * 30)
    ldFechaFur = ldAtencion - (lnControl * 30)
End Sub
Sub RetornaFlasM(ByRef lnRciu As Integer, ByRef lnAcidoFolico As Integer, ByRef lnSulfatoFerroso As Integer, ByRef lnClampaje As Integer, _
                 ByRef lnPinstitucional As Integer, ByRef lnNrosDni As Integer, lnDNImaximo As Integer)

     lnRciu = 2: lnAcidoFolico = 0: lnSulfatoFerroso = 1: lnClampaje = 0: lnPinstitucional = 2
     Dim lnFor As Integer
     If lnNrosDni <> lnDNImaximo Then
        For lnFor = 1 To (lnDNImaximo - lnNrosDni)
            lnRciu = IIf(lnRciu = 1, 2, 1)
            lnAcidoFolico = IIf(lnAcidoFolico = 1, 0, 1)
            lnSulfatoFerroso = IIf(lnSulfatoFerroso = 1, 0, 1)
            lnClampaje = IIf(lnClampaje = 1, 0, 1)
            lnPinstitucional = IIf(lnPinstitucional = 1, 2, 1)
        Next
     End If
     
     If lnNrosDni < lnDNImaximo Then
        lnNrosDni = lnNrosDni + 1
     Else
        lnNrosDni = 1
     End If
End Sub
Sub llenaLaboratorio(ByRef oRsFoxLaboratorio As Recordset, lnNrosDni As Integer, lcDNI As String, lcDNImadre As String, _
                     lcCod_2000 As String, lnAnio As Integer, lnMes As Integer, lcNom_lote As String, lnDia As Integer, _
                     lbEsMaterno As Boolean)
    If lnNrosDni <= 3 Then
    
        Dim lcCpts As String
        Dim lcCPT As String
        Dim lnFor As Integer
        If lbEsMaterno = True Then
           lcCpts = "82310/86689/84402"
        Else
           lcCpts = "85018/89127/84590/"
        End If
        lnFor = 1
        lnpos = 1
        Do While True
           lcCPT = Mid(lcCpts, lnpos, 5)
           If lnNrosDni = lnFor Then
              Exit Do
           End If
           lnpos = lnpos + 6
           lnFor = lnFor + 1
        Loop
        oRsFoxLaboratorio.AddNew
        oRsFoxLaboratorio.Fields!cod_2000 = lcCod_2000
        oRsFoxLaboratorio.Fields!ano = lnAnio
        oRsFoxLaboratorio.Fields!Mes = lnMes
        oRsFoxLaboratorio.Fields!nom_lote = lcNom_lote
        oRsFoxLaboratorio.Fields!dia = lnDia
        oRsFoxLaboratorio.Fields!DNI = IIf(lcDNImadre <> "", lcDNImadre, lcDNI)
        oRsFoxLaboratorio.Fields!prueba_lab = lcCPT
        oRsFoxLaboratorio.Update
    End If
End Sub



Private Sub Command8_Click()
        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long, lnIdEstablecimiento As Long
        Dim lcCodigo As String, lcDescripcion As String
        Dim mo_ReglasComunes As New ReglasComunes
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=HIS"
        '
        Me.MousePointer = 1
        lcSql = "select * from tablaUps"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              lcCodigo = oRsFox!cod_servsa
              lcDescripcion = Left(oRsFox!desc_servs, 50)
              mo_ReglasComunes.UPServiciosActualizar lcCodigo, lcDescripcion, wxConexionRed
              oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        Me.MousePointer = 11
        Unload Me

End Sub

Private Sub Form_Initialize()
       On Error Resume Next
       Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
End Sub

Private Sub Form_Load()
        On Error Resume Next
        '
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        mo_cmbIdTipoGenHistoriaClinica.BoundText = "2"
        txtRutaINI.Text = "C:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
        lbProcesaVAriosDBF = False
        '
        
        '
        txtFinicial.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
        txtFfinal.Text = Date
        '
        Me.Caption = Me.Caption & "      EESS= " & lcBuscaParametro.SeleccionaFilaParametro(205)
        '
        optBoleta_Click 1

        wxParametro102 = lcBuscaParametro.SeleccionaFilaParametro(102)
        wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
        wxParametro206 = lcBuscaParametro.SeleccionaFilaParametro(206)
        wxParametro207 = lcBuscaParametro.SeleccionaFilaParametro(207)
        wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
        wxParametro221 = lcBuscaParametro.SeleccionaFilaParametro(221)
        wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
        wxParametro285 = lcBuscaParametro.SeleccionaFilaParametro(285)
        wxParametro286 = lcBuscaParametro.SeleccionaFilaParametro(286)
        wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
        wxParametro339 = lcBuscaParametro.SeleccionaFilaParametro(339)
        
        'mgaray
        Call cargarDatosIniciales
End Sub

Function DevuelveUltimaCuentaEnParametros() As Long
'        Dim oRsTmp As New Recordset
'        oRsTmp.Open "select VAlorInt from Parametros where idParametro=1", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        DevuelveUltimaCuentaEnParametros = oRsTmp.Fields!ValorInt
'        oRsTmp.Close
'        Set oRsTmp = Nothing
End Function







Private Sub MedicamentosDesdeZIP_Click()
 On Error GoTo ErrActItems1
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 11
       
        Dim oCrypKey As New CrypKey.Util
        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim oCommand As New ADODB.Command
        Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos, oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim oParameter As ADODB.Parameter
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long, lcClaveDIGEMID As String
        Dim lcCodigo As String, lcTipoProductoSismed As String
        '
        If chkClaveDig.Value <> 1 Then
            lcClaveDIGEMID = ""
        Else
            
            lcClaveDIGEMID = lcBuscaParametro.SeleccionaFilaParametro(350)
            lcClaveDIGEMID = oCrypKey.DecryptString(lcClaveDIGEMID)
        End If
        sighentidades.DescomprimeArchivoZIP lcClaveDIGEMID, Me.txtZipArchivo.Text, Me.txtRutaGalenhos.Text, False
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=HIS"
        '
        Me.MousePointer = 1
        'Medicamentos
        lcSql = "select * from medicame"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           Set oCatalogoBienesInsumos.Conexion = wxConexionRed
           oDOCatalogoBienesInsumos.IdUsuarioAuditoria = 0
           Me.ProgressBar1.Max = oRsFox.RecordCount
           Me.ProgressBar1.Min = 0
           Me.ProgressBar1.Value = 0
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
                Me.Refresh: ProgressBar1.Value = ProgressBar1.Value + 1: DoEvents
                lcCodigo = Trim(oRsFox.Fields!CODIGO_MED)
                lcTipoProductoSismed = Trim(oRsFox!estrategic)
                Set oRsTmpOpc = oCatalogoBienesInsumos.SeleccionarPorCodigo(lcCodigo, wxConexionRed)
                If oRsTmpOpc.RecordCount = 0 And Len(lcCodigo) > 4 Then
                   oDOCatalogoBienesInsumos.IdSubGrupoFarmacologico = 999
                   oDOCatalogoBienesInsumos.IdGrupoFarmacologico = 999
                   oDOCatalogoBienesInsumos.nombre = Left(Trim(oRsFox.Fields!medicament) & " " & Trim(oRsFox.Fields!presentaci) & " " & _
                                                     Trim(oRsFox.Fields!concentrac) & " " & oRsFox.Fields!FF, 300)
                   oDOCatalogoBienesInsumos.Codigo = lcCodigo
                   
                   If UCase(Trim(oRsFox.Fields!estrategic)) = "E" Then
                      oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = IIf(oRsFox.Fields!EstVta = 1, 3, 2)
                   Else
                      oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = 1
                   End If
                   oDOCatalogoBienesInsumos.TipoProducto = IIf(oRsFox.Fields!Tipo = "M", 0, 1)
                   oDOCatalogoBienesInsumos.denominacion = Left(oRsFox.Fields!medicament, 100)
                   oDOCatalogoBienesInsumos.Concentracion = oRsFox.Fields!concentrac
                   oDOCatalogoBienesInsumos.Presentacion = oRsFox.Fields!presentaci
                   oDOCatalogoBienesInsumos.FormaFarmaceutica = Left(oRsFox.Fields!FF, 10)
                   oDOCatalogoBienesInsumos.TipoProductoSismed = lcTipoProductoSismed
                   oDOCatalogoBienesInsumos.Petitorio = IIf(oRsFox.Fields!Petitorio = "P", 1, 0)
                   
                   If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) = False Then
                      MsgBox oCatalogoBienesInsumos.MensajeError
                      GoTo ErrActItems1
                   End If
                Else
                   oDOCatalogoBienesInsumos.idProducto = oRsTmpOpc!idProducto
                   If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) = True Then
                   End If
                   oDOCatalogoBienesInsumos.denominacion = Left(oRsFox.Fields!medicament, 100)
                   oDOCatalogoBienesInsumos.Concentracion = oRsFox.Fields!concentrac
                   oDOCatalogoBienesInsumos.Presentacion = oRsFox.Fields!presentaci
                   oDOCatalogoBienesInsumos.FormaFarmaceutica = Left(oRsFox.Fields!FF, 10)
                   oDOCatalogoBienesInsumos.TipoProductoSismed = lcTipoProductoSismed
                   oDOCatalogoBienesInsumos.Petitorio = IIf(oRsFox.Fields!Tipo = "P", 1, 0)
                   If oCatalogoBienesInsumos.Modificar(oDOCatalogoBienesInsumos) = False Then
                      MsgBox oCatalogoBienesInsumos.MensajeError
                      GoTo ErrActItems1
                   End If
                End If
                oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        mo_ReglasArchivoClinico.ActualizaNULLenIdPartidaIdCCostoCon999
        Me.MousePointer = 11
        Unload Me
    End If
    Exit Sub
ErrActItems1:
    Resume
End Sub

Private Sub optBoleta_Click(Value As Integer)
    SelccionaOpcionComprobante ("Boleta")
    wxIdTipoComprobanteDefault = 3
End Sub

Private Sub optFactura_Click(Value As Integer)
    SelccionaOpcionComprobante ("Factura")
    wxIdTipoComprobanteDefault = 2
End Sub

Private Sub optRecibo_Click(Value As Integer)
    SelccionaOpcionComprobante ("Recibo")
    wxIdTipoComprobanteDefault = 1
End Sub

Private Sub optTicket_Click(Value As Integer)
    SelccionaOpcionComprobante (lcTicket)
    wxIdTipoComprobanteDefault = 4
End Sub

Sub SelccionaOpcionComprobante(ByVal lcOpcion As String)
    txtPasos.Text = "Pasos: * Cargar Valores del archivo SETUP_CAJA_" & UCase(lcOpcion) & ".INI   * Cuadrar " & lcOpcion & " en Vista Previa e Imprimirla  " & vbCrLf & _
                    "           * Los valores X,Y finales deberá guardarlo en SETUP_CAJA_" & UCase(lcOpcion) & ".INI " & vbCrLf & _
                    "- Los MEDICOS ya se ingresaron antes"
    lcTipoComprobanteCaja = lcOpcion
    lblRutaArchivo.Caption = "Ruta del archivo:   setup_caja_" & LCase(lcOpcion) & ".ini"
    cmdCargaINICajaServicios.Caption = lcOpcion & " SERVICIOS                    Carga Valores desde: 'c:\.....\archivos\setup_caja_" & LCase(lcOpcion) & ".ini'"
    cmdCargaINICajaFarmacia.Caption = lcOpcion & " FARMACIA                Carga Valores desde: 'c:\.....\archivos\setup_caja_" & LCase(lcOpcion) & ".ini'"
    
    txtCodigoY.BackColor = vbWhite
    txtPrecioY.BackColor = vbWhite
    'OCultando Historia y Caja en Cabecera
    lblHistoria.Visible = False
    txtHistoriaX.Visible = False
    txtHistoriaY.Visible = False
    txtHistoriaValor.Visible = False
    
    lblSubTotal.Visible = False
    txtSubTotal.Visible = False
    txtSubTotalX.Visible = False
    txtSubTotalY.Visible = False
    txtIGV.Visible = False
    txtIGVX.Visible = False
    txtIGVY.Visible = False
    
    If lcOpcion = "Ticket" Then
        txtCodigoY.BackColor = vbBlue
        txtPrecioY.BackColor = vbBlue
        'Viendo Historia y Caja en Cabecera
        lblHistoria.Visible = True
        txtHistoriaX.Visible = True
        txtHistoriaY.Visible = True
        txtHistoriaValor.Visible = True
    End If
    
    If lcOpcion = "Factura" Then
        lblSubTotal.Visible = True
        txtSubTotal.Visible = True
        txtSubTotalX.Visible = True
        txtSubTotalY.Visible = True
        txtIGV.Visible = True
        txtIGVX.Visible = True
        txtIGVY.Visible = True
    End If
    'mgaray
    activarGrabarConfiguracionComprobante False
End Sub





Private Sub txtHcActual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRsTmp As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        oRsTmp.Open "select * from Pacientes where idTipoNumeracion=" & mo_cmbIdTipoGenHistoriaClinica.BoundText & " and NroHistoriaClinica=" & Me.txtHcActual.Text, oConexion, adOpenKeyset, adLockOptimistic
        lblActual.Caption = ""
        lnIdPacienteActual = 0
        If oRsTmp.RecordCount > 0 Then
           lblActual.Caption = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre)
           lnIdPacienteActual = oRsTmp.Fields!idPaciente
        End If
        oRsTmp.Close
        oConexion.Close
    End If
End Sub


Private Sub txtHcNueva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRsTmp As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        oRsTmp.Open "select * from Pacientes where idTipoNumeracion=" & mo_cmbIdTipoGenHistoriaClinica.BoundText & " and NroHistoriaClinica=" & Me.txtHcNueva.Text, oConexion, adOpenKeyset, adLockOptimistic
        lblNueva.Caption = ""
        lnIdPacienteNuevo = 0
        If oRsTmp.RecordCount > 0 Then
           lblNueva.Caption = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre)
           lnIdPacienteNuevo = oRsTmp.Fields!idPaciente
        End If
        oRsTmp.Close
        oConexion.Close
    End If
End Sub




Private Sub activarGrabarConfiguracionComprobante(Optional activar As Boolean = True)
    btnGuardarConfiguracionComprobante.Enabled = activar
End Sub

Private Sub formatoDefectoControlesImprimirBoleta()
    txtPasos.Locked = True
    txtRutaINI.Locked = True
End Sub

Private Sub asignarValorDeControlesAVariables()
   If lcTipoServicioFarmacia = "SERVICIOS" Then
        WxLnNumeroSerieX = Val(Me.txtNumeroSerieX.Text)
        WxLnNumeroSerieY = Val(Me.txtNumeroSerieY.Text)
        WxLnEstadoX = Val(Me.txtEstadoX.Text)
        WxLnEstadoY = Val(Me.txtEstadoY.Text)
        WxLnTipoX = Val(Me.txtTipoX.Text)
        WxLnTipoY = Val(Me.txtTipoY.Text)
        WxLnRzSocialX = Val(Me.txtRzSocialX.Text)
        WxLnRzSocialY = Val(Me.txtRzSocialY.Text)
        WxLnFechaX = Val(Me.txtFechaX.Text)
        WxLnFechaY = Val(Me.txtFechaY.Text)
        WxLnServicioX = Val(Me.txtServicioX.Text)
        WxLnServicioY = Val(Me.txtServicioY.Text)
        WxLnObservacionesX = Val(Me.txtObservacionesX.Text)
        WxLnObservacionesY = Val(Me.txtObservacionesY.Text)
        WxLnHistoriaX = Val(Me.txtHistoriaX.Text)
        WxLnHistoriaY = Val(Me.txtHistoriaY.Text)
        WxLnCodigoY = Val(Me.txtCodigoY.Text)
        WxLnProductoY = Val(Me.txtProductoY.Text)
        WxLnProductoWidhtY = Val(Me.txtProductoAncho.Text)
        WxLnCantidadY = Val(Me.txtCantidadY.Text)
        WxLnPrecioY = Val(Me.txtPrecioY.Text)
        WxLnImporteY = Val(Me.txtImporteY.Text)
        WxLnCajeroX = Val(Me.txtCajeroX.Text)
        WxLnCajeroY = Val(Me.txtCajeroY.Text)
        WxLnCajaX = Val(Me.txtCajaX.Text)
        WxLnCajaY = Val(Me.txtCajaY.Text)
        WxLnAdelantosX = Val(Me.txtAdelantosX.Text)
        WxLnAdelantosY = Val(Me.txtAdelantosY.Text)
        WxLnTotalPagarX = Val(Me.txtTotalPagarX.Text)
        WxLnTotalPagarY = Val(Me.txtTotalPagarY.Text)
        WxLnCuentaX = Val(Me.txtCuentaX.Text)
        WxLnCuentaY = Val(Me.txtCuentaY.Text)
        WxLnExoneracionesX = Val(Me.txtExoneracionesX.Text)
        WxLnExoneracionesY = Val(Me.txtExoneracionesY.Text)
        WxLnTotalEnLetrasX = Val(Me.txtTotalEnLetrasX.Text)
        WxLnTotalEnLetrasY = Val(Me.txtTotalEnLetrasY.Text)
        WxLnTotalLetrasWidhtY = Val(txtTotalLetrasAncho.Text)
        WxLnTotalX = Val(Me.txtTotalX.Text)
        WxLnTotalY = Val(Me.txtTotalY.Text)
        WxLnSubTotalX = Val(Me.txtSubTotalX.Text)
        WxLnSubTotalY = Val(Me.txtSubTotalY.Text)
        WxLnIGVX = Val(Me.txtIGVX.Text)
        WxLnIGVY = Val(Me.txtIGVY.Text)
        WxLnCabeceraAlto = Val(Me.txtCabeceraAlto.Text)
        WxLnPieAlto = Val(Me.txtPieAlto.Text)
        
        WxLnNombreHoja = cboPapel.Text
        WxLnTipoReporteador = Val(Me.cboReporteador.ListIndex)
        WxLnMargenIzquierdoX = Val(txtMargenIzquierda.Text)
        WxLnMargenDerechoX = Val(txtMargenDerecha.Text)
        WxLnMargenSuperiorY = Val(txtMargenSuperior.Text)
        WxLnMargenInferiorY = Val(txtMargenInferior.Text)
        
        WxLnCabRucX = Val(txtRucX.Text)
        WxLnCabRucY = Val(txtRucY.Text)
        WxLnCabDireccionX = Val(txtDireccionX.Text)
        WxLnCabDireccionY = Val(txtDireccionY.Text)
        
   Else
        WxLnNumeroSerieX_F = Val(Me.txtNumeroSerieX.Text)
        WxLnNumeroSerieY_F = Val(Me.txtNumeroSerieY.Text)
        WxLnEstadoX_F = Val(Me.txtEstadoX.Text)
        WxLnEstadoY_F = Val(Me.txtEstadoY.Text)
        WxLnTipoX_F = Val(Me.txtTipoX.Text)
        WxLnTipoY_F = Val(Me.txtTipoY.Text)
        WxLnRzSocialX_F = Val(Me.txtRzSocialX.Text)
        WxLnRzSocialY_F = Val(Me.txtRzSocialY.Text)
        WxLnFechaX_F = Val(Me.txtFechaX.Text)
        WxLnFechaY_F = Val(Me.txtFechaY.Text)
        WxLnServicioX_F = Val(Me.txtServicioX.Text)
        WxLnServicioY_F = Val(Me.txtServicioY.Text)
        WxLnObservacionesX_F = Val(Me.txtObservacionesX.Text)
        WxLnObservacionesY_F = Val(Me.txtObservacionesY.Text)
        WxLnHistoriaX_F = Val(Me.txtHistoriaX.Text)
        WxLnHistoriaY_F = Val(Me.txtHistoriaY.Text)
        WxLnCodigoY_F = Val(Me.txtCodigoY.Text)
        WxLnProductoY_F = Val(Me.txtProductoY.Text)
        WxLnProductoWidhtY_F = Val(Me.txtProductoAncho.Text)
        WxLnCantidadY_F = Val(Me.txtCantidadY.Text)
        WxLnPrecioY_F = Val(Me.txtPrecioY.Text)
        WxLnImporteY_F = Val(Me.txtImporteY.Text)
        WxLnCajeroX_F = Val(Me.txtCajeroX.Text)
        WxLnCajeroY_F = Val(Me.txtCajeroY.Text)
        WxLnCajaX_F = Val(Me.txtCajaX.Text)
        WxLnCajaY_F = Val(Me.txtCajaY.Text)
        WxLnAdelantosX_F = Val(Me.txtAdelantosX.Text)
        WxLnAdelantosY_F = Val(Me.txtAdelantosY.Text)
        WxLnTotalPagarX_F = Val(Me.txtTotalPagarX.Text)
        WxLnTotalPagarY_F = Val(Me.txtTotalPagarY.Text)
        WxLnCuentaX_F = Val(Me.txtCuentaX.Text)
        WxLnCuentaY_F = Val(Me.txtCuentaY.Text)
        WxLnExoneracionesX_F = Val(Me.txtExoneracionesX.Text)
        WxLnExoneracionesY_F = Val(Me.txtExoneracionesY.Text)
        WxLnTotalEnLetrasX_F = Val(Me.txtTotalEnLetrasX.Text)
        WxLnTotalEnLetrasY_F = Val(Me.txtTotalEnLetrasY.Text)
        WxLnTotalLetrasWidhtY_F = Val(txtTotalLetrasAncho.Text)
        WxLnTotalX_F = Val(Me.txtTotalX.Text)
        WxLnTotalY_F = Val(Me.txtTotalY.Text)
        WxLnSubTotalX_F = Val(Me.txtSubTotalX.Text)
        WxLnSubTotalY_F = Val(Me.txtSubTotalY.Text)
        WxLnIGVX_F = Val(Me.txtIGVX.Text)
        WxLnIGVY_F = Val(Me.txtIGVY.Text)
        WxLnCabeceraAlto_F = Val(Me.txtCabeceraAlto.Text)
        WxLnPieAlto_F = Val(Me.txtPieAlto.Text)
        
        WxLnNombreHoja_F = cboPapel.Text
        WxLnTipoReporteador_F = Val(Me.cboReporteador.ListIndex)
        WxLnMargenIzquierdoX_F = Val(txtMargenIzquierda.Text)
        WxLnMargenDerechoX_F = Val(txtMargenDerecha.Text)
        WxLnMargenSuperiorY_F = Val(txtMargenSuperior.Text)
        WxLnMargenInferiorY_F = Val(txtMargenInferior.Text)
   
        WxLnCabRucX_F = Val(txtRucX.Text)
        WxLnCabRucY_F = Val(txtRucY.Text)
        WxLnCabDireccionX_F = Val(txtDireccionX.Text)
        WxLnCabDireccionY_F = Val(txtDireccionY.Text)
   End If
End Sub

'===========================================

'mgaray
Public Function cargarDatosIniciales()
    Call cargarTamanioPapel
    Call cargarTiposReporteador
End Function

Public Function cargarTiposReporteador()
    cboReporteador.AddItem "Data Report", reporteadorAUsarEnum.DataReport
    cboReporteador.AddItem "Crystal Report", reporteadorAUsarEnum.CrystalReport
End Function

Public Sub cargarTamanioPapel()
    Dim rsTamanioPapel As New Recordset
    Set rsTamanioPapel = listFormPrinter()
    cboPapel.Clear
'    cboPapel.Sorted = True
    If Not (rsTamanioPapel.EOF And rsTamanioPapel.BOF) Then
        rsTamanioPapel.MoveFirst
        While rsTamanioPapel.EOF = False
            cboPapel.AddItem rsTamanioPapel!FormName
            rsTamanioPapel.MoveNext
        Wend
    End If
End Sub

Public Function setPrinterDefault(PrinterName As String, Optional ByRef printerNameDefault As String)
    Dim impresora As Printer
    
    printerNameDefault = Printer.DeviceName
    
    For Each impresora In Printers
        'Si es igual al contenido señalado en el combobox
        If impresora.DeviceName = PrinterName Then
            'Lo seteamos así.
            Set Printer = impresora
        End If
    Next
End Function

Public Function validarAltoAreaImpresionBoleta() As Boolean
    Dim oRptBoleta As New RptBoleta
    
    validarAltoAreaImpresionBoleta = False
    
    Dim altoCabecera As Long, altoPie As Long
    Dim TopMargin As Integer, BottomMargin  As Integer
    Dim totalAltoImpresion As Long, altoPagina As Long
    
    TopMargin = Val(txtMargenSuperior.Text)
    BottomMargin = Val(txtMargenInferior.Text)
    
    altoCabecera = Val(txtCabeceraAlto.Text)
    altoPie = Val(txtPieAlto.Text)
    
    totalAltoImpresion = TopMargin + BottomMargin + altoCabecera + altoPie
    altoPagina = getPrinterHeight
    
    If oRptBoleta.getDetailMinimunHeight(cboReporteador.ListIndex) <= altoPagina - totalAltoImpresion Then
        validarAltoAreaImpresionBoleta = True
    End If
End Function

'Validar que las coordenas para la configuracion de los datos en el pie de pagina
'No superen el alto establecido para el pie de pagina, se puede permitir esto pero
'al imprimir el reporte, el alto de pie de pagina sera diferente al configurado
'porque este tomara el valor de control que esta mas hacia abajo + el alto de ese control
Private Function validarCoordenadaYCabeceraBoleta() As Boolean
    Dim maximoValorY As Long
    
    validarCoordenadaYCabeceraBoleta = False
    
    maximoValorY = validarCoordenadaMayor("Frame25", "X")
    
    If maximoValorY < Val(txtCabeceraAlto.Text) Then
        validarCoordenadaYCabeceraBoleta = True
    End If
End Function

Private Function validarCoordenadaYPieBoleta() As Boolean
    Dim maximoValorY As Long
    
    validarCoordenadaYPieBoleta = False
    
    maximoValorY = validarCoordenadaMayor("Frame27", "X")
    
    If maximoValorY < Val(txtPieAlto.Text) Then
        validarCoordenadaYPieBoleta = True
    End If
End Function

Private Function validarCoordenadaXCabeceraBoleta(anchoHoja As Long, margenes As Integer) As Boolean
    Dim maximoValorX As Long
    
    validarCoordenadaXCabeceraBoleta = False
    
    maximoValorX = validarCoordenadaMayor("Frame25", "Y")
    
    If maximoValorX < (anchoHoja - margenes) Then
        validarCoordenadaXCabeceraBoleta = True
    End If
End Function

Private Function validarCoordenadaXDetalleBoleta(anchoHoja As Long, margen As Integer) As Boolean
    Dim maximoValorX As Long
    
    validarCoordenadaXDetalleBoleta = False
    
    maximoValorX = validarCoordenadaMayor("Frame25", "Y")
    
    If maximoValorX < Val(txtProductoY.Text) + Val(txtProductoAncho.Text) Then
        maximoValorX = Val(txtProductoY.Text) + Val(txtProductoAncho.Text)
    End If
    
    If maximoValorX < anchoHoja - margen Then
        validarCoordenadaXDetalleBoleta = True
    End If
End Function

Private Function validarCoordenadaXPieBoleta(anchoHoja As Long, margen As Integer) As Boolean
    Dim maximoValorX As Long
    validarCoordenadaXPieBoleta = False
    
    maximoValorX = validarCoordenadaMayor("Frame27", "Y")
    
    If maximoValorX < Val(txtTotalEnLetrasY.Text) + Val(txtTotalLetrasAncho.Text) Then
        maximoValorX = Val(txtTotalEnLetrasY.Text) + Val(txtTotalLetrasAncho.Text)
    End If
    
    If maximoValorX < anchoHoja - margen Then
        validarCoordenadaXPieBoleta = True
    End If
End Function

'ContainerName es el nombre del control que contiene los controles que se quiere
'ultimoCaracterNombreCoordenada Es el ultimo caracter del nombre del grupo de controles
'que se quiere verificar, se observo que el nombre de los controles que definien las coordenada de impresion
'tiene al final la letra x o y
Private Function validarCoordenadaMayor(containerName As String, _
            ultimoCaracterNombreCoordenada As String) As Long
    Dim maximaCoordenada As Long, coordenaActual As Long
    
    maximaCoordenada = 0
    
    For Each oControl In Me.Controls
        If TypeOf oControl Is TextBox Then
            If UCase(oControl.Container.Name) = UCase(containerName) _
                    And UCase(Right(oControl.Name, 1)) = UCase(ultimoCaracterNombreCoordenada) _
                    And oControl.Visible = True Then
                If IsNumeric(oControl.Text) Then
                    coordenaActual = Val(oControl.Text) + Val(oControl.Tag)
                    If coordenaActual > maximaCoordenada Then
                        maximaCoordenada = coordenaActual
                    End If
                End If
            End If
        End If
    Next
    validarCoordenadaMayor = maximaCoordenada
End Function

'Creados para corregir valor de controles que se ocultan dependiendo del
'tipo de comprobante, a veces esos controles tiene valores establecidos que el usuario
'no ve (posiblemenete vienen del ini de la etapa inicial de prueba)
Private Sub ajustarValoresDeControlesOcultos()
    If txtHistoriaX.Visible = False Then txtHistoriaX.Text = ""
    If txtHistoriaY.Visible = False Then txtHistoriaY.Text = ""
    
    If txtSubTotalX.Visible = False Then txtSubTotalX.Text = ""
    If txtSubTotalY.Visible = False Then txtSubTotalY.Text = ""
    
    If txtIGVX.Visible = False Then txtIGVX.Text = ""
    If txtIGVY.Visible = False Then txtIGVY.Text = ""
    
End Sub

Private Function validarAltoDeSeccionesBoleta(Optional pregunta As String = "") As Boolean
On errro GoTo miError
    validarAltoDeSeccionesBoleta = True
    Exit Function
    validarAltoDeSeccionesBoleta = False
    
    If Not validarCoordenadaYCabeceraBoleta() Then
        If MsgBox("Los valores para las coordenadas X de la sección Datos de Cabecera, superan el alto establecido para la Cabecera." & vbCrLf & pregunta, vbInformation + vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    
    If Not validarCoordenadaYPieBoleta() Then
        If MsgBox("Los valores para las coordenadas X de la sección Datos de Pie de Página, superan el alto establecido para el pie de Pagina." & vbCrLf & pregunta, vbInformation + vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    
    If Not validarAltoAreaImpresionBoleta() Then
        Dim message As String
        Dim oRptBoleta As New RptBoleta
        Dim totalAltoImpresion As Long, altoPagina As Long
        
        altoPagina = getPrinterHeight
        
        totalAltoImpresion = Val(txtMargenSuperior.Text) + Val(txtCabeceraAlto.Text) + Val(txtPieAlto.Text) + Val(txtMargenInferior.Text)
        
        message = "Los Valores Definidos para el Area de Impresión Superan el Alto de la hoja Configurada." & vbCrLf
        message = message & vbCrLf & "Datos Referenciales:"
        message = message & vbCrLf & "Alto Pagina:" & altoPagina
        message = message & vbCrLf & "Margen Superior + Cabecera + Pie + Margen Inferior:" & totalAltoImpresion
        
        message = message & vbCrLf & "Alto Detalle, se recomienda un minimo de " _
                        & oRptBoleta.getDetailMinimunHeight(cboReporteador.ListIndex) & ":" & altoPagina - totalAltoImpresion
        message = message & vbCrLf & vbCrLf & pregunta
    
        If MsgBox(message, vbInformation + vbYesNo, "Advertencia") = vbNo Then
            Exit Function
        End If
    End If
    validarAltoDeSeccionesBoleta = True
miError:
    If Err Then
        MsgBox Err.Description, vbCritical, "Error"
    End If
End Function

Private Function validarAnchoDeSeccionesBoleta(Optional pregunta As String = "") As Boolean
On Error GoTo miError
    validarAnchoDeSeccionesBoleta = True
    Exit Function
    validarAnchoDeSeccionesBoleta = False
    Dim anchoHoja As Long, margen As Integer
    Dim oRptBoleta As New RptBoleta
    Dim arraySecciones() As String
    
    anchoHoja = getPrinterWidth()
    margen = Val(txtMargenIzquierda.Text) + Val(txtMargenDerecha.Text) + 4 '$ es un valor para los bordes de los controles
    
    ReDim arraySecciones(0)
    If Not validarCoordenadaXCabeceraBoleta(anchoHoja, margen) Then
        Call agregarElementoArray(arraySecciones, "Datos de Cabecera")
    End If
    
    If Not validarCoordenadaXDetalleBoleta(anchoHoja, margen) Then
        Call agregarElementoArray(arraySecciones, "Datos del detalle")
    End If
    
    If Not validarCoordenadaXPieBoleta(anchoHoja, margen) Then
        Call agregarElementoArray(arraySecciones, "Datos de Pie de Página")
    End If
    
    Dim nombresSeccionesFueraRango As String
    nombresSeccionesFueraRango = ""
    If arraySecciones(0) <> "" Then
        nombresSeccionesFueraRango = Join(arraySecciones, ", ")
    End If
    
    If nombresSeccionesFueraRango <> "" Then
        Dim message As String
        
        message = "Los Valores Definidos para Y de la(s) seccion(es) :"
        message = message & vbCrLf & nombresSeccionesFueraRango & " Superan el ancho de la hoja Configurada." & vbCrLf
        
        message = message & vbCrLf & "Datos Referenciales:"
        message = message & vbCrLf & "Ancho Pagina:" & anchoHoja
        message = message & vbCrLf & "Margen Izquierdo + Margen Derecho:" & Val(txtMargenIzquierda.Text) + Val(txtMargenDerecha.Text)
        
        
        message = message & vbCrLf & "Ancho Disponible para Impresion : " _
                        & anchoHoja - margen
        message = message & vbCrLf & vbCrLf & pregunta
    
        If MsgBox(message, vbInformation + vbYesNo, "Advertencia") = vbNo Then
            Exit Function
        End If
    End If
    validarAnchoDeSeccionesBoleta = True
miError:
    If Err Then
        MsgBox Err.Description, vbCritical, "Error"
    End If
End Function

Public Sub leerAltoControlesEnBoleta()
    If elDocumentoEsTicket = False Then
        Dim oReporte As New FormBoleta
        
        txtNumeroSerieX.Tag = leerAltoControlEnReporte("BoletaNumeroSerie", "cabecera", oReporte)
        txtEstadoX.Tag = leerAltoControlEnReporte("BoletaEstado", "cabecera", oReporte)
        txtTipoX.Tag = leerAltoControlEnReporte("BoletaTipo", "cabecera", oReporte)
        txtRzSocialX.Tag = leerAltoControlEnReporte("RazonSocial", "cabecera", oReporte)
        txtFechaX.Tag = leerAltoControlEnReporte("FechaCobranza", "cabecera", oReporte)
        txtServicioX.Tag = leerAltoControlEnReporte("Servicio", "cabecera", oReporte)
        txtObservacionesX.Tag = leerAltoControlEnReporte("Observaciones", "cabecera", oReporte)
        
        
        txtCajeroX.Tag = leerAltoControlEnReporte("Cajero", "pie", oReporte)
        txtCajaX.Tag = leerAltoControlEnReporte("nombreCaja", "pie", oReporte)
        txtAdelantosX.Tag = leerAltoControlEnReporte("Adelantos", "pie", oReporte)
        txtTotalPagarX.Tag = leerAltoControlEnReporte("TotalPorPagar", "pie", oReporte)
        txtCuentaX.Tag = leerAltoControlEnReporte("idCuentaAtencion", "pie", oReporte)
        txtExoneracionesX.Tag = leerAltoControlEnReporte("Exoneraciones", "pie", oReporte)
        
        txtTotalEnLetrasX.Tag = leerAltoControlEnReporte("TotalEnLetras", "pie", oReporte)
        txtTotalX.Tag = leerAltoControlEnReporte("TotalBoleta", "pie", oReporte)
        txtSubTotalX.Tag = leerAltoControlEnReporte("SubTotal", "pie", oReporte)
        txtIGVX.Tag = leerAltoControlEnReporte("IGV", "pie", oReporte)
    End If
End Sub

Public Sub leerAnchoControlesEnBoleta()
    If elDocumentoEsTicket = False Then
        Dim oReporte As New FormBoleta
        
        txtNumeroSerieY.Tag = leerAnchoControlEnReporte("BoletaNumeroSerie", "cabecera", oReporte)
        txtEstadoY.Tag = leerAnchoControlEnReporte("BoletaEstado", "cabecera", oReporte)
        txtTipoY.Tag = leerAnchoControlEnReporte("BoletaTipo", "cabecera", oReporte)
        txtRzSocialY.Tag = leerAnchoControlEnReporte("RazonSocial", "cabecera", oReporte)
        txtFechaY.Tag = leerAnchoControlEnReporte("FechaCobranza", "cabecera", oReporte)
        txtServicioY.Tag = leerAnchoControlEnReporte("Servicio", "cabecera", oReporte)
        txtObservacionesY.Tag = leerAnchoControlEnReporte("Observaciones", "cabecera", oReporte)
        
        txtCodigoY.Tag = leerAnchoControlEnReporte("txtCodigo", "Detalle", oReporte)
        txtCodigoY.ToolTipText = "Ancho :" & txtCodigoY.Tag
        
        txtProductoY.Tag = leerAnchoControlEnReporte("txtNombreProducto", "Detalle", oReporte)
        txtCantidadY.Tag = leerAnchoControlEnReporte("txtCantidad", "Detalle", oReporte)
        txtCantidadY.ToolTipText = "Ancho :" & txtCantidadY.Tag
        
        txtPrecioY.Tag = leerAnchoControlEnReporte("txtPrecioUnitario", "Detalle", oReporte)
        txtPrecioY.ToolTipText = "Ancho :" & txtPrecioY.Tag
        txtImporteY.Tag = leerAnchoControlEnReporte("txtTotalPorPagar", "Detalle", oReporte)
        txtImporteY.ToolTipText = "Ancho :" & txtImporteY.Tag
        
        
        txtCajeroY.Tag = leerAnchoControlEnReporte("Cajero", "pie", oReporte)
        txtCajaY.Tag = leerAnchoControlEnReporte("nombreCaja", "pie", oReporte)
        txtAdelantosY.Tag = leerAnchoControlEnReporte("Adelantos", "pie", oReporte)
        txtTotalPagarY.Tag = leerAnchoControlEnReporte("TotalPorPagar", "pie", oReporte)
        txtCuentaY.Tag = leerAnchoControlEnReporte("idCuentaAtencion", "pie", oReporte)
        txtExoneracionesY.Tag = leerAnchoControlEnReporte("Exoneraciones", "pie", oReporte)
        txtTotalEnLetrasY.Tag = leerAnchoControlEnReporte("TotalEnLetras", "pie", oReporte)
        txtTotalY.Tag = leerAnchoControlEnReporte("TotalBoleta", "pie", oReporte)
        txtSubTotalY.Tag = leerAnchoControlEnReporte("SubTotal", "pie", oReporte)
        txtIGVY.Tag = leerAnchoControlEnReporte("IGV", "pie", oReporte)
    End If
End Sub

'devuelve el alto del control en un dataReport
Private Function leerAltoControlEnReporte(controlName As String, _
        seccionName As String, oReporte As DataReport) As Integer
On Error GoTo miError
    leerAltoControlEnReporte = oReporte.Sections(seccionName).Controls(controlName).Height
miError:
    If Err Then
        MsgBox "Error al leer el alto del control en el reporte :" & Err.Description, vbExclamation, "Advertencia"
    End If
End Function

'devuelve el alto del control en un dataReport
Private Function leerAnchoControlEnReporte(controlName As String, _
        seccionName As String, oReporte As DataReport) As Integer
On Error GoTo miError
    leerAnchoControlEnReporte = oReporte.Sections(seccionName).Controls(controlName).Width
miError:
    If Err Then
        MsgBox "Error al leer el ancho del control en el reporte :" & Err.Description, vbExclamation, "Advertencia"
    End If
End Function


Private Function elDocumentoEsTicket() As Boolean
    elDocumentoEsTicket = IIf(lcTipoComprobanteCaja = lcTicket, True, False)
End Function

Public Function getPrinterHeight() As Long
On Error GoTo miError
    Dim rsTamanioPapel As New Recordset
    Dim pageHeight As Long
    
    pageHeight = 0
    
    Set rsTamanioPapel = listFormPrinter()
    If buscarHoja(cboPapel.Text, rsTamanioPapel) = True Then
        pageHeight = rsTamanioPapel!formHeight * 56.7 'Se convierte a twip
    End If
    'getPrinterHeight = 5783
    getPrinterHeight = pageHeight
miError:
    If Err Then
        MsgBox Err.Description, vbCritical, "Error"
    End If
End Function


Public Function getPrinterWidth() As Long
    Dim rsTamanioPapel As New Recordset
    Dim pageWidth As Long
    
    pageWidth = 0
    
    Set rsTamanioPapel = listFormPrinter()
    If buscarHoja(cboPapel.Text, rsTamanioPapel) = True Then
        pageWidth = rsTamanioPapel!formWidth * 56.7 'Se convierte a twip
    End If
'    getPrinterWidth = 10149
    getPrinterWidth = pageWidth
End Function

Public Function buscarHoja(nombreHoja As String, ByRef rsTamanioPapel As Recordset) As Boolean
    buscarHoja = False
    If nombreHoja <> "" Then
        If Not (rsTamanioPapel.EOF And rsTamanioPapel.BOF) Then
            rsTamanioPapel.MoveFirst
            rsTamanioPapel.Find "formName='" & cboPapel.Text & "'"
            If Not rsTamanioPapel.EOF Then
                buscarHoja = True
            End If
        End If
    Else
        MsgBox "Seleccione Hoja de Papel para poder determinar el alto o ancho del papel", vbInformation, "Advertencia"
    End If
End Function


Public Sub agregarElementoArray(ByRef nombreArray() As String, valor As String)
    If nombreArray(UBound(nombreArray, 1)) <> "" Then
        ReDim Preserve nombreArray(UBound(nombreArray, 1) + 1)
    End If
    nombreArray(UBound(nombreArray, 1)) = valor
End Sub




Private Sub cmdListaTipoSalida_Click()
    On Error GoTo ErrRptHuelga
    'On Error Resume Next
    Dim ml_EdadEnMeses As Long
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Dim W1 As Excel.Workbook
    Dim s1 As Excel.Worksheet
    Dim oRsTmp1 As New Recordset
    Dim oFila As Long, ldFecha As Date, lbNuevo As Boolean
    Dim ldFechaInicialHist As Date, ldFechaFinalHist As Date
    Dim lnNroConsultas As Long, lcFecha As String, lcHoraAtencion As String, lcTexto As String
    Dim oConexion As New Connection
    Dim ml_idTipoSexo As Integer, ldFechaNacimiento As Date, ldFechaAtencion As Date
    Dim lnPeso As Double, lnTalla As Double, lnEdadGest As Integer
    Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
    Dim lnTallaEnCmMasPuntoCinco As Double
    Dim lnPercentilPE As Double, lnPercentilTE As Double, lnPercentilPT As Double
    Dim lnPercentilIMC As Double, lcPercentilIMC As String, lcCodigo As String

    Const lnPercentilNull As Long = 0
    '
    Set W = EXL.Workbooks.Open("c:\excel.xls")
    Set s = W.Sheets("hoja1")
    lcSql = "SELECT      dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, dbo.farmSaldo.cantidad,  " & _
            "                      dbo.farmSaldo.idTipoSalidaBienInsumo , dbo.FactCatalogoBienesInsumos.TipoProductoSismed" & _
            " FROM         dbo.farmSaldo INNER JOIN" & _
            "                      dbo.FactCatalogoBienesInsumos ON dbo.farmSaldo.idProducto = dbo.FactCatalogoBienesInsumos.IdProducto" & _
            " ORDER BY dbo.FactCatalogoBienesInsumos.Codigo"
    oRsTmp1.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oFila = 1
       s.Cells(oFila, 1).Value = "codigo"
       s.Cells(oFila, 2).Value = "TipoSismed"
       s.Cells(oFila, 3).Value = "descripciòn"
       s.Cells(oFila, 4).Value = "SoloVentas"
       s.Cells(oFila, 5).Value = "SoloInterv.Sanitarias"
       s.Cells(oFila, 6).Value = "Ventas e Interv.Sanitarias"
       oFila = oFila + 2
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          s.Cells(oFila, 1).Value = oRsTmp1!Codigo
          s.Cells(oFila, 2).Value = oRsTmp1!TipoProductoSismed
          s.Cells(oFila, 3).Value = oRsTmp1!nombre
          lcCodigo = oRsTmp1!Codigo
          Do While Not oRsTmp1.EOF And lcCodigo = oRsTmp1!Codigo
             Select Case oRsTmp1!idTipoSalidaBienInsumo
             Case 1
               s.Cells(oFila, 4).Value = "X"
               If oRsTmp1!TipoProductoSismed <> "_" Then
                  s.Cells(oFila, 7).Value = "Chequear a detalle"
               End If
             Case 2
               s.Cells(oFila, 5).Value = "X"
               If oRsTmp1!TipoProductoSismed = "_" Then
                  s.Cells(oFila, 7).Value = "Chequear a detalle"
               End If
             Case 3
               s.Cells(oFila, 6).Value = "X"
             End Select
             
             oRsTmp1.MoveNext
             If oRsTmp1.EOF Then
                Exit Do
             End If
          Loop
          oFila = oFila + 1
       Loop
    End If
    oRsTmp1.Close
    EXL.Visible = True
    W.PrintPreview
    Set s = Nothing
    Set s1 = Nothing
    Set W = Nothing
    Set W1 = Nothing
    Set EXL = Nothing
    MsgBox "procesó sin problemas"
    Exit Sub
ErrRptHuelga:
    MsgBox Err.Description
    Resume

End Sub




Function FarmMovimientoDetalleSeleccionarXcodigo(lcCodigo As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDetalleSeleccionarXcodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoDetalleSeleccionarXcodigo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub FarmaciaActualizaTipoSalida(lnIdProducto As Long, lnNuevoTipo As Long)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaActualizaTipoSalida"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nuevoTipo", adInteger, adParamInput, 0, lnNuevoTipo): .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub


Function farmMovimientoVentasDetalleXidProducto(lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasDetalleXidProducto"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasDetalleXidProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Private Sub ActualizaCPTDesdeDBFRegiones(oConexionFox As Connection)
        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        
        Dim lcSql As String ', lcCodDx As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        Me.MousePointer = 1
        lcSql = "select * from cpt_regiones"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              
              'Añade CPT HIS
'              If Val(oRsFox.Fields!clase) = 4 Then
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = wxConexionRed
                    .CommandTimeout = 150
                    .CommandText = "FactCatalogoServiciosAgregarCodigoNombre"
                    Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, Left(Trim(oRsFox.Fields!Cod_cpt), 20)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(oRsFox.Fields!desc_cpt, 255)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                    
                    .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
'                 oRsTmp.Close
 '             End If
              oRsFox.MoveNext
           Loop
        End If
        Me.MousePointer = 11
End Sub
Private Sub cmdHuecosCtas_Click()
'MODIFICADO POR FRANKLIN CACHAY 07/11/2013 - se cambio a store procedure

  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Dim mo_ReglasFacturacion  As New SIGHNegocios.ReglasFacturacion
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open("c:\excel.xls")
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim oRsTmpCtas As New Recordset, lnMaximo As Long
    Dim lnFor As Long, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    
'    lcSql = "select * from FacturacionCuentasAtencionPtos order by idCuentaAtencion"
'    oRsTmp.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionPtosSeleccionarTodo"
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    
    lnMaximo = oRsTmp.RecordCount
    If lnMaximo > 0 Then
       ProgressBar1.Min = 0
       ProgressBar1.Max = lnMaximo
       ProgressBar1.Value = 0
       lnFila = 3
       lcRango = "B" + Trim(Str(lnFila))
       s.Range(lcRango).Value = "Cta"
       lcRango = "C" + Trim(Str(lnFila))
       s.Range(lcRango).Value = "Cta Sgte"
       oRsTmp.MoveFirst
       lnIdCuentaAnterior = oRsTmp.Fields!idCuentaAtencion
       Do While Not oRsTmp.EOF
          lnIdCuentaAtencionSiguiente = oRsTmp.Fields!idCuentaAtencion
          If (lnIdCuentaAtencionSiguiente - lnIdCuentaAnterior) > 1 Then
'            lcSql = "select idCuentaAtencion from FacturacionCuentasAtencion where (idCuentaAtencion>=" & lnIdCuentaAnterior & " and idCuentaAtencion<=" & lnIdCuentaAtencionSiguiente & ") and idEstado<>5 and idEstado<>9"
            If oRsTmpCtas.State = 1 Then oRsTmpCtas.Close
'            oRsTmpCtas.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
            
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FacturacionCuentasAtencionSeleccionarPorRangoCuentas"
                Set oParameter = .CreateParameter("@IdCuentaAnterior", adInteger, adParamInput, 0, lnIdCuentaAnterior): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdCuentaAtencionSiguiente", adInteger, adParamInput, 0, lnIdCuentaAtencionSiguiente): .Parameters.Append oParameter
                Set oRsTmpCtas = .Execute
                Set oRsTmpCtas.ActiveConnection = Nothing
            End With
            
            Set oCommand = Nothing
            Set oParameter = Nothing
            If oRsTmpCtas.RecordCount > 2 Then
                lnFila = lnFila + 1
                lcRango = "B" + Trim(Str(lnFila))
                s.Range(lcRango).Value = lnIdCuentaAnterior
                lcRango = "C" + Trim(Str(lnFila))
                s.Range(lcRango).Value = lnIdCuentaAtencionSiguiente
                'Procesa dichos Rangos
                For lnFor = lnIdCuentaAnterior To lnIdCuentaAtencionSiguiente
                     mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar lnFor, False, 0
                Next
                
             End If
          End If
          lnIdCuentaAnterior = oRsTmp.Fields!idCuentaAtencion
          Do While Not oRsTmp.EOF And lnIdCuentaAtencionSiguiente = oRsTmp.Fields!idCuentaAtencion
             DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
             oRsTmp.MoveNext
             If oRsTmp.EOF Then
                Exit Do
             End If
          Loop
       Loop
    End If
    Set s = Nothing
    W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    MsgBox "Cuentas Huecas en archivo: c:\excel.xls"
    Unload Me
End If
End Sub


Private Sub cmdActualizaDNI_Click()
    If lcBuscaParametro.SeleccionaFilaParametro(351) = "S" Then
        Me.MousePointer = 11
        Dim oConexion As New Connection
        Dim oConexionExterna As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        oConexionExterna.CommandTimeout = 900
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        PasaHCigualDNI oConexion, oConexionExterna
        oConexion.Close
        oConexionExterna.Close
        Set oConexion = Nothing
        Set oConexionExterna = Nothing
        Unload Me
    End If
End Sub

Sub PasaHCigualDNI(oConexion As Connection, oConexionExterna As Connection)
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oRsTmp As New Recordset
        Dim oRsTmp2 As New Recordset
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "PacientesIgualDNI"
            .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
        
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexionExterna
            .CommandTimeout = 150
            .CommandText = "AtencionesCEactualizaHistoriaDesdeSIGH"
            .Execute
        End With
        Set oCommand = Nothing
        
        
        Set oRsTmp = Nothing
        Set oRsTmp2 = Nothing
   

End Sub



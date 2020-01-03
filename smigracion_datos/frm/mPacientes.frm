VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mPacientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migración de Pacientes hacia SisGalenPlus"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   Icon            =   "mPacientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   12630
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   285
      Left            =   60
      TabIndex        =   78
      Top             =   8610
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Private Sub cmdActualizaMayusculasAuto_Click()"
      TabPicture(0)   =   "mPacientes.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame30"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame29"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Varios Proc2"
      TabPicture(1)   =   "mPacientes.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame22"
      Tab(1).Control(1)=   "Frame11"
      Tab(1).Control(2)=   "Frame1(4)"
      Tab(1).Control(3)=   "Frame1(3)"
      Tab(1).Control(4)=   "Frame18"
      Tab(1).Control(5)=   "Frame1(2)"
      Tab(1).Control(6)=   "Frame17"
      Tab(1).Control(7)=   "Frame16"
      Tab(1).Control(8)=   "Frame7"
      Tab(1).Control(9)=   "Frame5"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Varios Proc3"
      TabPicture(2)   =   "mPacientes.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdNuevosCptVsGrupo"
      Tab(2).Control(1)=   "Frame21"
      Tab(2).Control(2)=   "Frame2(3)"
      Tab(2).Control(3)=   "Frame2(2)"
      Tab(2).Control(4)=   "Frame35"
      Tab(2).Control(5)=   "cmdArreglaENE"
      Tab(2).Control(6)=   "Frame2(1)"
      Tab(2).Control(7)=   "Frame8"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Kimbiri"
      TabPicture(3)   =   "mPacientes.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtClave1"
      Tab(3).Control(1)=   "Frame2(4)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "CS Nazareno(Ayac)"
      TabPicture(4)   =   "mPacientes.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkFichaFamiliar1"
      Tab(4).Control(1)=   "txtFFinCSN"
      Tab(4).Control(2)=   "txtFIniCSN"
      Tab(4).Control(3)=   "List6"
      Tab(4).Control(4)=   "cmdCSnazarena"
      Tab(4).Control(5)=   "Label10"
      Tab(4).Control(6)=   "Label9"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "CS San Juan (Ayacucho)"
      TabPicture(5)   =   "mPacientes.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chkFF2"
      Tab(5).Control(1)=   "txtFinCSsb"
      Tab(5).Control(2)=   "cmdSanJuanAyacucho"
      Tab(5).Control(3)=   "List7"
      Tab(5).Control(4)=   "txtIniCSsb"
      Tab(5).Control(5)=   "Label12"
      Tab(5).Control(6)=   "Label11"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "CS Sta Elena"
      TabPicture(6)   =   "mPacientes.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "chkActualizaFechaREg"
      Tab(6).Control(1)=   "txtINIse"
      Tab(6).Control(2)=   "List8"
      Tab(6).Control(3)=   "cmdProcesaStaElena"
      Tab(6).Control(4)=   "txtFINse"
      Tab(6).Control(5)=   "Label14"
      Tab(6).Control(6)=   "Label13"
      Tab(6).ControlCount=   7
      TabCaption(7)   =   "HRC -Cajamarca"
      TabPicture(7)   =   "mPacientes.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "chkkPacientes1"
      Tab(7).Control(1)=   "Text2"
      Tab(7).Control(2)=   "cmdMigraHRC"
      Tab(7).Control(3)=   "List9"
      Tab(7).Control(4)=   "Text1"
      Tab(7).Control(5)=   "Label16"
      Tab(7).Control(6)=   "Label15"
      Tab(7).ControlCount=   7
      Begin VB.CommandButton cmdNuevosCptVsGrupo 
         Caption         =   "Agrega nuevos GRUPO a CPT de Laboratorio (c:\excel.xls    libro1   a=cpt, b=grupo)"
         Height          =   735
         Left            =   -68580
         TabIndex        =   217
         Top             =   5205
         Width           =   5745
      End
      Begin VB.Frame Frame21 
         Caption         =   "Agrega todos Puntos de Carga (Hosp/Emerg/CE) para Cuentas Corrientes"
         ForeColor       =   &H000000FF&
         Height          =   1530
         Left            =   -68625
         TabIndex        =   211
         Top             =   3615
         Width           =   6105
         Begin VB.TextBox txtTipoServicio 
            Height          =   375
            Left            =   1200
            TabIndex        =   213
            Text            =   "0"
            Top             =   300
            Width           =   795
         End
         Begin VB.CommandButton cmdTodosPuntosCarga 
            Caption         =   "Procesa "
            Height          =   645
            Left            =   105
            TabIndex        =   212
            Top             =   705
            Width           =   5925
         End
         Begin VB.Label lblPto 
            Caption         =   "...."
            Height          =   360
            Left            =   3585
            TabIndex        =   216
            Top             =   255
            Width           =   1725
         End
         Begin VB.Label lblTotPtos 
            Alignment       =   1  'Right Justify
            Caption         =   "....."
            Height          =   270
            Left            =   2205
            TabIndex        =   215
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label Label51 
            Caption         =   "Tipo Servicio"
            Height          =   285
            Left            =   180
            TabIndex        =   214
            Top             =   345
            Width           =   975
         End
      End
      Begin VB.TextBox txtClave1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74835
         PasswordChar    =   "*"
         TabIndex        =   210
         Text            =   "Text"
         Top             =   510
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Height          =   6135
         Index           =   4
         Left            =   -74835
         TabIndex        =   203
         Top             =   750
         Visible         =   0   'False
         Width           =   11685
         Begin VB.TextBox txtCartillas 
            Height          =   285
            Left            =   4095
            TabIndex        =   206
            Text            =   "1"
            Top             =   5565
            Width           =   435
         End
         Begin VB.CommandButton cmdCartillas 
            Caption         =   "Procesar"
            Height          =   315
            Left            =   4770
            TabIndex        =   205
            Top             =   5520
            Width           =   1365
         End
         Begin VB.PictureBox Picture 
            AutoSize        =   -1  'True
            Height          =   4830
            Left            =   150
            Picture         =   "mPacientes.frx":0522
            ScaleHeight     =   4770
            ScaleWidth      =   5940
            TabIndex        =   204
            Top             =   270
            Width           =   6000
         End
         Begin MSDataGridLib.DataGrid grdCartillas 
            Height          =   5610
            Left            =   6255
            TabIndex        =   207
            Top             =   270
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   9895
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   23
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "nro"
               Caption         =   "N°"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "ganador"
               Caption         =   "Eempate),  L(gana local),  V(gana visita)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "fijo"
               Caption         =   "x (certeza)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   360
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3435.024
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   840.189
               EndProperty
            EndProperty
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "N° cartillas"
            Height          =   195
            Index           =   3
            Left            =   3180
            TabIndex        =   209
            Top             =   5580
            Width           =   750
         End
         Begin VB.Label lblNro 
            AutoSize        =   -1  'True
            Caption         =   "..."
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   4
            Left            =   4125
            TabIndex        =   208
            Top             =   5895
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cambia de Cuenta - LABORATORIO"
         Height          =   1305
         Index           =   3
         Left            =   -68340
         TabIndex        =   197
         Top             =   2220
         Width           =   5715
         Begin VB.CommandButton cmbCambiaCtaLab 
            Caption         =   "procesar"
            Height          =   360
            Left            =   255
            TabIndex        =   202
            Top             =   780
            Width           =   5310
         End
         Begin VB.TextBox txtNcuentaNew 
            Height          =   330
            Left            =   4455
            TabIndex        =   201
            Top             =   300
            Width           =   1080
         End
         Begin VB.TextBox txtNmovimiento 
            Height          =   330
            Left            =   2055
            TabIndex        =   199
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Cuenta"
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   200
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "N° movimiento Laboratorio"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   198
            Top             =   360
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Actualiza resultados de LABORATORIO desde HUARAL"
         Height          =   1545
         Index           =   2
         Left            =   -68385
         TabIndex        =   193
         Top             =   375
         Width           =   5760
         Begin VB.CommandButton cmdProcesaLabHuaral 
            Caption         =   "Procesar"
            Height          =   765
            Left            =   120
            TabIndex        =   195
            Top             =   615
            Width           =   5400
         End
         Begin VB.TextBox txtClave 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   4320
            PasswordChar    =   "*"
            TabIndex        =   194
            Text            =   "Text1"
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "...después del proceso eliminar cpt q empiezen con ER"
            Height          =   195
            Left            =   90
            TabIndex        =   196
            Top             =   270
            Width           =   3885
         End
      End
      Begin VB.Frame Frame35 
         Caption         =   "Actualiza IMPORTES de cada CUENTA para Reporte: CONSUMO PTO CARGA"
         ForeColor       =   &H000000FF&
         Height          =   2895
         Left            =   -74880
         TabIndex        =   180
         Top             =   2760
         Width           =   6105
         Begin VB.TextBox txtTiempoProcesaCtas 
            Height          =   315
            Left            =   2850
            TabIndex        =   188
            Text            =   "5"
            Top             =   1740
            Width           =   795
         End
         Begin VB.CommandButton cmdProcesaCuentas 
            Caption         =   "Procesa "
            Height          =   645
            Left            =   120
            TabIndex        =   187
            Top             =   2160
            Width           =   5925
         End
         Begin VB.TextBox txtCtaInicial 
            Height          =   375
            Left            =   2460
            TabIndex        =   186
            Text            =   "0"
            Top             =   510
            Width           =   795
         End
         Begin VB.TextBox txtCta1 
            Height          =   375
            Left            =   2460
            TabIndex        =   183
            Text            =   "0"
            Top             =   1320
            Width           =   1005
         End
         Begin VB.TextBox txtCta2 
            Height          =   375
            Left            =   4890
            TabIndex        =   182
            Text            =   "0"
            Top             =   1290
            Width           =   1035
         End
         Begin VB.CommandButton cmdHuecosCtas 
            Caption         =   "..."
            Height          =   315
            Left            =   5550
            TabIndex        =   181
            ToolTipText     =   "Detecta HUECOS ENTRE CUENTAS de la tabla FacturacionCuentasAtencionPTOS (debe existir C:\EXCEL.XLS)"
            Top             =   240
            Width           =   465
         End
         Begin Threed.SSOption optActImpRangoCtas 
            Height          =   285
            Left            =   180
            TabIndex        =   184
            Top             =   1050
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   503
            _Version        =   262144
            Caption         =   "Rango de Cuentas"
         End
         Begin Threed.SSOption optActImpUnaCuenta 
            Height          =   255
            Left            =   180
            TabIndex        =   185
            Top             =   300
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   262144
            Caption         =   "Desde una Cuenta"
            Value           =   -1
         End
         Begin VB.Label Label63 
            Caption         =   "N° Minutos que durará este proceso"
            Height          =   285
            Left            =   210
            TabIndex        =   192
            Top             =   1800
            Width           =   2625
         End
         Begin VB.Label Label64 
            Caption         =   "N° Cta inicial"
            Height          =   285
            Left            =   1470
            TabIndex        =   191
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label62 
            Caption         =   "N° Cta inicial"
            Height          =   285
            Left            =   1500
            TabIndex        =   190
            Top             =   1410
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "N° Cta Final"
            Height          =   285
            Left            =   3930
            TabIndex        =   189
            Top             =   1350
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdArreglaENE 
         Caption         =   "Arregla EÑE, |  en Pacientes (Apellidos, nombres, direccion,documento)"
         Height          =   495
         Left            =   -74880
         TabIndex        =   179
         Top             =   2040
         Width           =   6015
      End
      Begin VB.Frame Frame2 
         Caption         =   "Farmacia: Actualiza PRECIO VENTA para la NUEVA TARIFA"
         Height          =   1545
         Index           =   1
         Left            =   -74880
         TabIndex        =   175
         Top             =   450
         Width           =   5760
         Begin VB.TextBox txtNewFF 
            Height          =   345
            Left            =   4320
            TabIndex        =   177
            Text            =   "Text1"
            Top             =   240
            Width           =   1185
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Proceso para actualizar PRECIO FARMACIA"
            Height          =   765
            Left            =   120
            TabIndex        =   176
            Top             =   615
            Width           =   5400
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Financiamiento(PRODUCTO/PLAN)   (nueva):"
            Height          =   195
            Left            =   90
            TabIndex        =   178
            Top             =   270
            Width           =   3645
         End
      End
      Begin VB.Frame Frame22 
         BackColor       =   &H8000000D&
         Caption         =   "Actualiza MEDICAMENTOS, origen=GALENHOS, destino=Cubo formato ICI (odbc)"
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   -74910
         TabIndex        =   173
         Top             =   7350
         Width           =   6285
         Begin VB.CommandButton cmdActualizaCuboGalenhos2008 
            Caption         =   "Actualiza MEDICAMENTOS del Cubo (Odbc=GalenHosSql2008)"
            Enabled         =   0   'False
            Height          =   435
            Left            =   120
            TabIndex        =   174
            Top             =   300
            Width           =   5985
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "HIS Lluyllucucha"
         Height          =   2865
         Left            =   7140
         TabIndex        =   167
         Top             =   375
         Width           =   3555
         Begin VB.CommandButton cmdProcesaHISlluyDET 
            Caption         =   "Procesar (DET - desde HIS hacia Galenhos)"
            Height          =   465
            Left            =   120
            TabIndex        =   171
            Top             =   1530
            Width           =   3315
         End
         Begin VB.CommandButton cmdProcesaHISlluy1 
            Caption         =   "Procesar (desde GALENHOS hacia HIS/DET)"
            Height          =   465
            Left            =   120
            TabIndex        =   170
            Top             =   2310
            Width           =   3315
         End
         Begin VB.CommandButton cmdProcesaHISlluy 
            Caption         =   "Procesar (CAB/DET - desde HIS hacia Galenhos)"
            Height          =   465
            Left            =   120
            TabIndex        =   169
            Top             =   1020
            Width           =   3315
         End
         Begin VB.TextBox Text3 
            Height          =   615
            Left            =   1290
            MultiLine       =   -1  'True
            TabIndex        =   168
            Text            =   "mPacientes.frx":F66E
            Top             =   330
            Width           =   2145
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Consideraciones"
            Height          =   195
            Left            =   120
            TabIndex        =   172
            Top             =   420
            Width           =   1170
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "solo Hosp SICUANI"
         Height          =   420
         Left            =   7395
         TabIndex        =   166
         Top             =   7935
         Width           =   1935
      End
      Begin VB.Frame Frame29 
         Caption         =   "Actualiza a MAYUSCULAS el campo AUTOGENERADO de la tabla PACIENTES"
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   7065
         TabIndex        =   164
         Top             =   7005
         Width           =   5415
         Begin VB.CommandButton cmdActualizaMayusculasAuto 
            Caption         =   "Apellido Paterno, Materno, primer nombre a Mayusculas asi como el autogenerado"
            Height          =   405
            Left            =   90
            TabIndex        =   165
            Top             =   300
            Width           =   5250
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H8000000D&
         Caption         =   "Cuentas con Problemas de Tarifa (no corresponde al PLAN)"
         ForeColor       =   &H000000FF&
         Height          =   2085
         Left            =   -68580
         TabIndex        =   161
         Top             =   5760
         Width           =   5985
         Begin VB.CommandButton cmdCuentasYtarifas 
            Caption         =   "Ejecuta proceso (debe existir c:\excel.xls)"
            Height          =   375
            Left            =   135
            TabIndex        =   162
            Top             =   1605
            Width           =   5745
         End
         Begin UltraGrid.SSUltraGrid grdCuentasYtarifas 
            Height          =   1275
            Left            =   120
            TabIndex        =   163
            Top             =   270
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   2249
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108864
            Caption         =   "Lista de CUENTAS a revisar"
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Actualiza tabla: FactPartidasPresupuestalesXMes  para REPORTES X PARTIDA"
         ForeColor       =   &H000000FF&
         Height          =   1200
         Index           =   4
         Left            =   -74895
         TabIndex        =   155
         Top             =   6075
         Width           =   6105
         Begin VB.TextBox txtAnioProc 
            Height          =   375
            Left            =   2715
            TabIndex        =   158
            Top             =   315
            Width           =   795
         End
         Begin VB.CommandButton cmdProcesaPartidas 
            Caption         =   "Procesa "
            Height          =   600
            Left            =   3570
            TabIndex        =   157
            Top             =   315
            Width           =   2445
         End
         Begin VB.TextBox txtMinutosProc 
            Height          =   315
            Left            =   2730
            TabIndex        =   156
            Text            =   "5"
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "N° Minutos que durará este proceso"
            Height          =   195
            Left            =   90
            TabIndex        =   160
            Top             =   780
            Width           =   2550
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Año a procesar"
            Height          =   195
            Left            =   90
            TabIndex        =   159
            Top             =   330
            Width           =   1080
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Cantidad Afiliados"
         Height          =   795
         Left            =   7035
         TabIndex        =   153
         Top             =   6135
         Width           =   5445
         Begin VB.CommandButton cmdAfiliados 
            Caption         =   "Imprime AFILIADOS (total, por género, por grupo edad)"
            Height          =   360
            Left            =   60
            TabIndex        =   154
            Top             =   255
            Width           =   5295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Lee Archivo Excel y graba datos en Galenhos (Distrito, Provincia y Departamento)"
         ForeColor       =   &H000000FF&
         Height          =   2970
         Index           =   3
         Left            =   -68535
         TabIndex        =   147
         Top             =   2760
         Width           =   5950
         Begin VB.TextBox Text7 
            Height          =   1305
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   150
            Text            =   "mPacientes.frx":F69F
            Top             =   540
            Width           =   5625
         End
         Begin VB.TextBox txtExcelrptUbegeo 
            Height          =   315
            Left            =   1320
            TabIndex        =   149
            Text            =   "c:\rptUbigeo_08082014.xls"
            Top             =   240
            Width           =   4520
         End
         Begin VB.CommandButton cmdGrabaDistritoProvinciaDepartamento 
            Caption         =   "Procesar"
            Height          =   405
            Left            =   195
            TabIndex        =   148
            Top             =   2445
            Width           =   5655
         End
         Begin VB.Label Label46 
            Caption         =   "Empezar desde la 2 fila"
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   255
            TabIndex        =   152
            Top             =   1920
            Width           =   5520
         End
         Begin VB.Label Label44 
            Caption         =   "Archivo Excel:"
            Height          =   285
            Left            =   210
            TabIndex        =   151
            Top             =   270
            Width           =   1245
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FF8080&
         Caption         =   "Pasar a un solo Excel varios Excel (con varias hojas)"
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   -68520
         TabIndex        =   143
         Top             =   600
         Width           =   5955
         Begin VB.TextBox txtNroHojas 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   146
            Text            =   "Text5"
            Top             =   1680
            Width           =   5655
         End
         Begin VB.CommandButton cmdProceVariasExcel 
            Caption         =   "procesar"
            Height          =   405
            Left            =   120
            TabIndex        =   144
            Top             =   1080
            Width           =   5745
         End
         Begin VB.Label Label45 
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"mPacientes.frx":F6F8
            Height          =   675
            Left            =   120
            TabIndex        =   145
            Top             =   330
            Width           =   5715
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Agrega columna descripción CPT GALENHOS a una tabla CPT trabajado por el SIS"
         ForeColor       =   &H000000FF&
         Height          =   1755
         Index           =   2
         Left            =   -74880
         TabIndex        =   137
         Top             =   4230
         Width           =   6285
         Begin VB.CommandButton cmdProcesaCPT 
            Caption         =   "Procesar"
            Height          =   945
            Left            =   4860
            TabIndex        =   138
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label Label43 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cpt_sis.c  (descripción CPT)"
            Height          =   315
            Left            =   90
            TabIndex        =   142
            Top             =   1290
            Width           =   4665
         End
         Begin VB.Label Label42 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cpt_sis.b  (codigo SIS)"
            Height          =   315
            Left            =   90
            TabIndex        =   141
            Top             =   960
            Width           =   4665
         End
         Begin VB.Label Label41 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cpt_sis.a  (codigo CPT)"
            Height          =   315
            Left            =   90
            TabIndex        =   140
            Top             =   630
            Width           =   4665
         End
         Begin VB.Label Label40 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS y el archivo ..............\galenhos\archivos\Cpt_sis.dbf"
            Height          =   315
            Left            =   90
            TabIndex        =   139
            Top             =   270
            Width           =   6105
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Elimina Historias Clinicas"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -74850
         TabIndex        =   133
         Top             =   510
         Width           =   6315
         Begin VB.CommandButton cmdEliminaHC 
            Caption         =   "Elimina Historias Clinicas (solamente aquellas  que no tengan ninguna atencion)"
            Height          =   405
            Left            =   90
            TabIndex        =   134
            Top             =   1080
            Width           =   6105
         End
         Begin VB.Label Label38 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "c:\barrantes\hc.dbf........campo: nroHistori"
            Height          =   315
            Left            =   120
            TabIndex        =   136
            Top             =   690
            Width           =   6075
         End
         Begin VB.Label Label39 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC: SISMEDV2"
            Height          =   315
            Left            =   120
            TabIndex        =   135
            Top             =   330
            Width           =   6075
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FF8080&
         Caption         =   "Lee Archivo Excel y graba datos en EXCEL"
         ForeColor       =   &H000000FF&
         Height          =   1995
         Left            =   -74850
         TabIndex        =   128
         Top             =   2160
         Width           =   6165
         Begin VB.TextBox Text4 
            Height          =   915
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   131
            Text            =   "mPacientes.frx":F79E
            Top             =   540
            Width           =   5835
         End
         Begin VB.TextBox txtExcel 
            Height          =   315
            Left            =   1320
            TabIndex        =   130
            Text            =   "c:\sis.xls"
            Top             =   210
            Width           =   4725
         End
         Begin VB.CommandButton ProcesaCPTtarapoto 
            Caption         =   "Procesar"
            Height          =   405
            Left            =   180
            TabIndex        =   129
            Top             =   1500
            Width           =   5895
         End
         Begin VB.Label Label37 
            Caption         =   "Archivo Excel:"
            Height          =   285
            Left            =   210
            TabIndex        =   132
            Top             =   270
            Width           =   1245
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Actualiza columna de la tabla -->AtencionesDatosAdicionales.SisCodigo"
         ForeColor       =   &H000000FF&
         Height          =   1305
         Left            =   7050
         TabIndex        =   126
         Top             =   3330
         Width           =   5445
         Begin VB.CommandButton Command5 
            Caption         =   "Actualiza columna de la tabla -->AtencionesDatosAdicionales.SisCodigo (con NULL)"
            Enabled         =   0   'False
            Height          =   855
            Left            =   240
            TabIndex        =   127
            Top             =   300
            Width           =   5025
         End
      End
      Begin VB.Frame Frame30 
         BackColor       =   &H00FF0000&
         Caption         =   "Migra ultimos ESTABLECIMIENTOS NO MINSA desde el SIS"
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   7050
         TabIndex        =   123
         Top             =   4680
         Width           =   5445
         Begin VB.CommandButton cmdEstabNewDesdeSIS 
            Caption         =   "Agrega NUEVOS ESTABLECIMIENTOS, (tambien Dpto, prov, dist nuevos)"
            Height          =   405
            Left            =   120
            TabIndex        =   124
            Top             =   660
            Width           =   5085
         End
         Begin VB.Label Label56 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe existir la BD: sigh_sis"
            Height          =   315
            Left            =   120
            TabIndex        =   125
            Top             =   300
            Width           =   5055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF0000&
         Caption         =   "Carga movimientos de ATENCIONES HISTORICAS"
         Height          =   7995
         Index           =   1
         Left            =   120
         TabIndex        =   79
         Top             =   420
         Width           =   6915
         Begin VB.Frame Frame10 
            Caption         =   "Para SIS"
            Height          =   1815
            Left            =   120
            TabIndex        =   113
            Top             =   5670
            Width           =   6705
            Begin VB.TextBox txtSISCodigoUDR 
               Height          =   285
               Left            =   5070
               TabIndex        =   120
               Top             =   540
               Width           =   1335
            End
            Begin VB.TextBox txtSISptoDigitacion 
               Height          =   285
               Left            =   5070
               TabIndex        =   118
               Top             =   210
               Width           =   1335
            End
            Begin VB.TextBox txtSisCatEESS 
               Height          =   285
               Left            =   1290
               TabIndex        =   116
               Top             =   540
               Width           =   1335
            End
            Begin VB.TextBox txtSisDisa 
               Height          =   285
               Left            =   1290
               TabIndex        =   114
               Top             =   180
               Width           =   1335
            End
            Begin MSDataGridLib.DataGrid grdSIS 
               Height          =   795
               Left            =   240
               TabIndex        =   122
               Top             =   930
               Visible         =   0   'False
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   1402
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   3082
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   3082
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000009&
               DrawMode        =   6  'Mask Pen Not
               X1              =   3270
               X2              =   3270
               Y1              =   120
               Y2              =   780
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Codigo UDR"
               Height          =   195
               Left            =   4110
               TabIndex        =   121
               Top             =   570
               Width           =   900
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pto Digitación"
               Height          =   195
               Left            =   4020
               TabIndex        =   119
               Top             =   300
               Width           =   990
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Categoria ESSS"
               Height          =   195
               Left            =   150
               TabIndex        =   117
               Top             =   540
               Width           =   1140
            End
            Begin VB.Label Label30 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Disa SIS"
               Height          =   255
               Left            =   150
               TabIndex        =   115
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdCargaAtencionesHist 
            Caption         =   "Procesar (carga Atenciones Históricas y actualiza datos personales)"
            Height          =   375
            Left            =   210
            TabIndex        =   108
            Top             =   7530
            Width           =   6615
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "       Datos"
            Height          =   1815
            Left            =   120
            TabIndex        =   97
            Top             =   2250
            Width           =   6705
            Begin MSDataGridLib.DataGrid dgrMuestra 
               Height          =   1215
               Left            =   1380
               TabIndex        =   109
               Top             =   540
               Visible         =   0   'False
               Width           =   5205
               _ExtentX        =   9181
               _ExtentY        =   2143
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   3082
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   3082
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtubigeo 
               Height          =   285
               Left            =   1320
               TabIndex        =   102
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox txttelefono 
               Height          =   285
               Left            =   1320
               TabIndex        =   101
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtdireccion 
               Height          =   285
               Left            =   1320
               TabIndex        =   100
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox txtnombre 
               Height          =   285
               Left            =   1320
               TabIndex        =   99
               Top             =   240
               Width           =   4935
            End
            Begin VB.CheckBox chkDejarBuscar 
               Caption         =   "Check1"
               Height          =   255
               Left            =   6360
               TabIndex        =   98
               ToolTipText     =   "Para Dejar de Buscar en la grilla."
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label28 
               BackColor       =   &H00E0E0E0&
               Caption         =   "para HIS"
               Height          =   255
               Left            =   2640
               TabIndex        =   107
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label Label27 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ubigeo"
               Height          =   255
               Left            =   240
               TabIndex        =   106
               Top             =   1320
               Width           =   615
            End
            Begin VB.Label Label26 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Telefono"
               Height          =   255
               Left            =   240
               TabIndex        =   105
               Top             =   960
               Width           =   735
            End
            Begin VB.Label Label25 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Direccion"
               Height          =   255
               Left            =   240
               TabIndex        =   104
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label24 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Nombre"
               Height          =   255
               Left            =   240
               TabIndex        =   103
               Top             =   240
               Width           =   615
            End
            Begin VB.Image Image5 
               Height          =   240
               Left            =   120
               Picture         =   "mPacientes.frx":F812
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "       Codigos"
            Height          =   1485
            Left            =   120
            TabIndex        =   84
            Top             =   4140
            Width           =   6705
            Begin VB.TextBox txtcodrenaes 
               Height          =   285
               Left            =   5040
               TabIndex        =   90
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtcodHIS 
               Height          =   285
               Left            =   5040
               TabIndex        =   89
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtcodmicrored 
               Height          =   285
               Left            =   5040
               TabIndex        =   88
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtcodred 
               Height          =   285
               Left            =   1320
               TabIndex        =   87
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtcoddisa 
               Height          =   285
               Left            =   1320
               TabIndex        =   86
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtcodminsa 
               Height          =   285
               Left            =   1320
               TabIndex        =   85
               Top             =   360
               Width           =   1335
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000009&
               DrawMode        =   6  'Mask Pen Not
               X1              =   3240
               X2              =   3240
               Y1              =   360
               Y2              =   1320
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "RENAES"
               Height          =   195
               Left            =   4320
               TabIndex        =   96
               Top             =   1110
               Width           =   660
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "DEL HOSPITAL (His)"
               Height          =   195
               Left            =   3450
               TabIndex        =   95
               Top             =   750
               Width           =   1515
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "MICRORED"
               Height          =   195
               Left            =   4110
               TabIndex        =   94
               Top             =   390
               Width           =   870
            End
            Begin VB.Label Label20 
               BackColor       =   &H00E0E0E0&
               Caption         =   "RED"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   1080
               Width           =   375
            End
            Begin VB.Label Label19 
               BackColor       =   &H00E0E0E0&
               Caption         =   "DISA"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "MINSA (Sismed)"
               Height          =   195
               Left            =   120
               TabIndex        =   91
               Top             =   360
               Width           =   1155
            End
            Begin VB.Image Image6 
               Height          =   240
               Left            =   120
               Picture         =   "mPacientes.frx":10214
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "        Sistema HIS"
            Height          =   615
            Left            =   120
            TabIndex        =   81
            Top             =   1590
            Width           =   2175
            Begin VB.OptionButton optHISAnterior 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Anterior"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optHISActual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Actual"
               Height          =   195
               Left            =   1200
               TabIndex        =   82
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.Image Image4 
               Height          =   240
               Left            =   120
               Picture         =   "mPacientes.frx":10C16
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.CheckBox chkCentroSalud 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Es Centro de Salud?"
            Height          =   255
            Left            =   4950
            TabIndex        =   80
            Top             =   1620
            Width           =   1815
         End
         Begin VB.Label Label29 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Existir el archivo PARAMETROS.MDB en la misma ruta de MIGRACION.EXE"
            Height          =   315
            Left            =   120
            TabIndex        =   112
            Top             =   1170
            Width           =   6675
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC:  HIS   (Driver Visual Foxpro Driver, tabla libre, que apunte a c:\archiv...\galenhos\archivos)"
            Height          =   495
            Left            =   120
            TabIndex        =   111
            Top             =   270
            Width           =   6675
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HistCab.dbf, histDet.dbf  deben estar en la carpeta de ODBC HIS"
            Height          =   315
            Left            =   120
            TabIndex        =   110
            Top             =   810
            Width           =   6675
         End
      End
      Begin VB.CheckBox chkFF2 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo usa FICHA FAMILIAR "
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   -65040
         TabIndex        =   74
         Top             =   7050
         Width           =   2445
      End
      Begin VB.CheckBox chkFichaFamiliar1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sólo usa FICHA FAMILIAR "
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -65220
         TabIndex        =   73
         Top             =   6960
         Width           =   2565
      End
      Begin VB.CheckBox chkkPacientes1 
         Caption         =   "Usa Tabla ""Pacientes2"" de la BD  ""DbHospi"" ? (solo se agrega Pacientes, ya no limpia datos de Pacientes)"
         Height          =   825
         Left            =   -69660
         TabIndex        =   72
         Top             =   6420
         Width           =   3675
      End
      Begin VB.CheckBox chkActualizaFechaREg 
         Caption         =   "Actualiza Fecha de Registro"
         Height          =   315
         Left            =   -69690
         TabIndex        =   71
         Top             =   7470
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   -72300
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   6510
         Width           =   1185
      End
      Begin VB.CommandButton cmdMigraHRC 
         Caption         =   "Proceso de Migración de Pacientes"
         Enabled         =   0   'False
         Height          =   1275
         Left            =   -74910
         TabIndex        =   67
         Top             =   7020
         Width           =   3165
      End
      Begin VB.ListBox List9 
         Height          =   5910
         Left            =   -74880
         TabIndex        =   66
         Top             =   510
         Width           =   12135
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   -74040
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   6510
         Width           =   1185
      End
      Begin VB.TextBox txtINIse 
         Height          =   345
         Left            =   -74040
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   7410
         Width           =   1185
      End
      Begin VB.ListBox List8 
         Height          =   6690
         Left            =   -74790
         TabIndex        =   61
         Top             =   480
         Width           =   12135
      End
      Begin VB.CommandButton cmdProcesaStaElena 
         Caption         =   "Proceso de Migración de Pacientes"
         Enabled         =   0   'False
         Height          =   465
         Left            =   -74940
         TabIndex        =   60
         Top             =   7890
         Width           =   3165
      End
      Begin VB.TextBox txtFINse 
         Height          =   345
         Left            =   -72330
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   7410
         Width           =   1185
      End
      Begin VB.TextBox txtFFinCSN 
         Height          =   345
         Left            =   -72060
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   6960
         Width           =   1185
      End
      Begin VB.TextBox txtFinCSsb 
         Height          =   345
         Left            =   -72150
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   7050
         Width           =   1185
      End
      Begin VB.CommandButton cmdSanJuanAyacucho 
         Caption         =   "Proceso de Migración de Pacientes"
         Height          =   645
         Left            =   -74730
         TabIndex        =   54
         Top             =   7650
         Width           =   3165
      End
      Begin VB.ListBox List7 
         Height          =   6495
         Left            =   -74910
         TabIndex        =   53
         Top             =   420
         Width           =   12375
      End
      Begin VB.TextBox txtIniCSsb 
         Height          =   345
         Left            =   -73860
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   7050
         Width           =   1185
      End
      Begin VB.TextBox txtFIniCSN 
         Height          =   345
         Left            =   -73830
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   6990
         Width           =   1185
      End
      Begin VB.ListBox List6 
         Height          =   6300
         Left            =   -74700
         TabIndex        =   48
         Top             =   480
         Width           =   12075
      End
      Begin VB.CommandButton cmdCSnazarena 
         Caption         =   "Proceso de Migración de Pacientes"
         Height          =   765
         Left            =   -74820
         TabIndex        =   47
         Top             =   7530
         Width           =   3165
      End
      Begin VB.Frame Frame8 
         Caption         =   "Migra DATOS PERSONALES de los Pacientes"
         Height          =   555
         Left            =   -64500
         TabIndex        =   34
         Top             =   7860
         Visible         =   0   'False
         Width           =   2025
         Begin ComctlLib.ProgressBar ProgressBar4 
            Height          =   285
            Left            =   6090
            TabIndex        =   75
            Top             =   4560
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   503
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.CommandButton cmdMigraSicuani 
            Caption         =   "Proceso de Migración de Pacientes"
            Enabled         =   0   'False
            Height          =   1275
            Left            =   6030
            TabIndex        =   45
            Top             =   3180
            Width           =   3165
         End
         Begin VB.ListBox List4 
            Height          =   2595
            Left            =   150
            TabIndex        =   44
            Top             =   330
            Width           =   9315
         End
         Begin VB.Frame Frame9 
            Height          =   1395
            Left            =   90
            TabIndex        =   37
            Top             =   3090
            Width           =   5565
            Begin VB.TextBox txtCuzcoF2 
               Height          =   345
               Left            =   4230
               TabIndex        =   41
               Text            =   "Text1"
               Top             =   420
               Width           =   1185
            End
            Begin VB.TextBox txtCuzcoF1 
               Height          =   345
               Left            =   2550
               TabIndex        =   40
               Text            =   "Text1"
               Top             =   420
               Width           =   1185
            End
            Begin Threed.SSOption SSOption1 
               Height          =   285
               Left            =   180
               TabIndex        =   38
               Top             =   930
               Visible         =   0   'False
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   503
               _Version        =   262144
               Caption         =   "Con Fechas de Ingreso VACIAS"
            End
            Begin Threed.SSOption SSOption2 
               Height          =   195
               Left            =   150
               TabIndex        =   39
               Top             =   240
               Width           =   3195
               _ExtentX        =   5636
               _ExtentY        =   344
               _Version        =   262144
               Caption         =   "Por Fechas de Ingreso"
               Value           =   -1
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "al"
               Height          =   195
               Left            =   3900
               TabIndex        =   43
               Top             =   510
               Width           =   120
            End
            Begin VB.Label Label7 
               Caption         =   "F.Ingreso Hospital:"
               Height          =   255
               Left            =   960
               TabIndex        =   42
               Top             =   480
               Width           =   1515
            End
         End
         Begin VB.CheckBox chkLimpiaLolcli1 
            Caption         =   "Limpia tabla 'LolCliProblemasHC"
            Height          =   315
            Left            =   90
            TabIndex        =   36
            Top             =   4590
            Width           =   2595
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Pacientes ya migrados en GAlenHos: Graba Historias  con problemas en la tabla 'lolcliProblemasHC'"
            Enabled         =   0   'False
            Height          =   645
            Left            =   90
            TabIndex        =   35
            Top             =   5010
            Width           =   9285
         End
         Begin VB.Label lblProcesando11 
            Caption         =   "..."
            Height          =   345
            Left            =   2820
            TabIndex        =   46
            Top             =   4560
            Width           =   2865
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H8000000D&
         Caption         =   "Migra Atencioens de Citas y Triaje"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -64740
         TabIndex        =   26
         Top             =   8010
         Visible         =   0   'False
         Width           =   2175
         Begin ComctlLib.ProgressBar ProgressBar3 
            Height          =   345
            Left            =   750
            TabIndex        =   76
            Top             =   2340
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   609
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.ListBox List3 
            Height          =   1425
            Left            =   150
            TabIndex        =   33
            Top             =   300
            Width           =   9315
         End
         Begin VB.TextBox txtFCita2 
            Height          =   345
            Left            =   2460
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   1830
            Width           =   1185
         End
         Begin VB.TextBox txtFCita1 
            Height          =   345
            Left            =   780
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1860
            Width           =   1185
         End
         Begin VB.CommandButton cmdMigraAtencionesJamo 
            Caption         =   "Elimina lo Migrado y vuelve a Migrar las CITAS"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   27
            Top             =   2220
            Width           =   3975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Left            =   2160
            TabIndex        =   31
            Top             =   1950
            Width           =   120
         End
         Begin VB.Label Label5 
            Caption         =   "F.Citas"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   30
            Top             =   1920
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Migra DATOS PERSONALES de los Pacientes"
         Height          =   465
         Left            =   -66480
         TabIndex        =   14
         Top             =   8010
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CommandButton cmdProblemasJamo 
            Caption         =   "Pacientes ya migrados en GAlenHos: Graba Historias del con problemas en la tabla 'lolcliProblemasHC'"
            Enabled         =   0   'False
            Height          =   645
            Left            =   60
            TabIndex        =   32
            Top             =   3840
            Width           =   9285
         End
         Begin VB.CheckBox chkLimpiaLolcli 
            Caption         =   "Limpia tabla 'LolCliProblemasHC"
            Height          =   315
            Left            =   90
            TabIndex        =   25
            Top             =   3510
            Width           =   2595
         End
         Begin VB.Frame Frame6 
            Height          =   1395
            Left            =   90
            TabIndex        =   18
            Top             =   2010
            Width           =   5565
            Begin Threed.SSOption optFechasVacias 
               Height          =   285
               Left            =   180
               TabIndex        =   24
               Top             =   930
               Visible         =   0   'False
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   503
               _Version        =   262144
               Caption         =   "Con Fechas de Ingreso VACIAS"
            End
            Begin Threed.SSOption optFechasJamo 
               Height          =   195
               Left            =   150
               TabIndex        =   23
               Top             =   240
               Width           =   3195
               _ExtentX        =   5636
               _ExtentY        =   344
               _Version        =   262144
               Caption         =   "Por Fechas de Ingreso"
               Value           =   -1
            End
            Begin VB.TextBox txtIni1 
               Height          =   345
               Left            =   2550
               TabIndex        =   20
               Text            =   "Text1"
               Top             =   420
               Width           =   1185
            End
            Begin VB.TextBox txtFin1 
               Height          =   345
               Left            =   4230
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   420
               Width           =   1185
            End
            Begin VB.Label Label3 
               Caption         =   "F.Ingreso Hospital:"
               Height          =   255
               Left            =   960
               TabIndex        =   22
               Top             =   480
               Width           =   1515
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "al"
               Height          =   195
               Left            =   3900
               TabIndex        =   21
               Top             =   510
               Width           =   120
            End
         End
         Begin VB.ListBox List2 
            Height          =   1425
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   9315
         End
         Begin VB.CommandButton cmdMigraJamo 
            Caption         =   "Proceso de Migración de Pacientes"
            Enabled         =   0   'False
            Height          =   1275
            Left            =   6030
            TabIndex        =   15
            Top             =   2100
            Width           =   3165
         End
         Begin VB.Label lblProcesando1 
            Caption         =   "..."
            Height          =   345
            Left            =   2820
            TabIndex        =   16
            Top             =   3480
            Width           =   2865
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   11220
         TabIndex        =   1
         Top             =   8130
         Visible         =   0   'False
         Width           =   1245
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Consideraciones:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   5505
            Left            =   90
            TabIndex        =   12
            Top             =   210
            Width           =   9585
            Begin VB.ListBox List1 
               Height          =   5130
               Left            =   120
               TabIndex        =   13
               Top             =   210
               Width           =   9315
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1155
            Left            =   60
            TabIndex        =   5
            Top             =   5760
            Width           =   9555
            Begin ComctlLib.ProgressBar ProgressBar1 
               Height          =   315
               Left            =   3540
               TabIndex        =   77
               Top             =   630
               Width           =   5865
               _ExtentX        =   10345
               _ExtentY        =   556
               _Version        =   327682
               Appearance      =   1
            End
            Begin VB.TextBox txtFechaIni 
               Height          =   345
               Left            =   1830
               TabIndex        =   8
               Text            =   "Text1"
               Top             =   180
               Width           =   1185
            End
            Begin VB.TextBox txtFechaFin 
               Height          =   345
               Left            =   3510
               TabIndex        =   7
               Text            =   "Text1"
               Top             =   180
               Width           =   1185
            End
            Begin VB.CommandButton cmdProcesa 
               Caption         =   "Proceso de Migración de Pacientes"
               Enabled         =   0   'False
               Height          =   345
               Left            =   6150
               TabIndex        =   6
               Top             =   210
               Width           =   3195
            End
            Begin VB.Label Label1 
               Caption         =   "F.Ingreso Hospital:"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   210
               Width           =   1515
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "al"
               Height          =   195
               Left            =   3270
               TabIndex        =   10
               Top             =   240
               Width           =   120
            End
            Begin VB.Label lblProcesando 
               AutoSize        =   -1  'True
               Caption         =   "Label3"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   780
               Width           =   480
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "GalenHos YA MIGRADO"
            Height          =   975
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   6960
            Width           =   9525
            Begin VB.CommandButton cmdHCproblemas 
               Caption         =   "Graba Historias del LolCli con problemas en la tabla 'lolcliProblemasHC'"
               Enabled         =   0   'False
               Height          =   645
               Left            =   120
               TabIndex        =   4
               Top             =   210
               Width           =   3465
            End
            Begin VB.CommandButton cmdActFecNac 
               Caption         =   "Cambia FECHA NACIMIENTO, Autogenerado, vuelve a REPROCESAR LolCliProblemasHC"
               Enabled         =   0   'False
               Height          =   645
               Left            =   5040
               TabIndex        =   3
               Top             =   180
               Width           =   4365
            End
         End
      End
      Begin VB.Label Label16 
         Caption         =   "F.Ingreso:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   70
         Top             =   6540
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   -72570
         TabIndex        =   69
         Top             =   6570
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   -72600
         TabIndex        =   64
         Top             =   7470
         Width           =   120
      End
      Begin VB.Label Label13 
         Caption         =   "F.Ingreso:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   7440
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "F.Ingreso:"
         Height          =   255
         Left            =   -74700
         TabIndex        =   56
         Top             =   7080
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   -72420
         TabIndex        =   55
         Top             =   7110
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   -72390
         TabIndex        =   51
         Top             =   7050
         Width           =   120
      End
      Begin VB.Label Label9 
         Caption         =   "F.Ingreso:"
         Height          =   255
         Left            =   -74670
         TabIndex        =   50
         Top             =   7020
         Width           =   855
      End
   End
End
Attribute VB_Name = "mPacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Migración de Pacientes hacia SisGalenPlus
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcMensajeError As String
Dim lnTipoNumeracion As Long   ' Const lnTipoNumeracion As Long = 2
Dim wxConexionJAMO As New Connection
Dim mo_ReglasAdmision As New ReglasAdmision
Dim mo_ReglasComunes As New ReglasComunes
Dim mo_SIGHProxies As New SIGHProxies.Procesos
Dim oRsGrdSIS As New Recordset
Dim oRsCartillas As New Recordset

Private Sub chkDejarBuscar_Click()
    If chkDejarBuscar.Value = 1 Then
       dgrMuestra.Visible = False
    End If
End Sub

Private Sub cmbCambiaCtaLab_Click()
    If Val(txtNmovimiento.Text) = 0 Then
       MsgBox "Tiene que ingresar Nro Movimiento"
       Exit Sub
    End If
     If Val(txtNcuentaNew.Text) = 0 Then
       MsgBox "Tiene que ingresar Nro Cuenta"
       Exit Sub
    End If
'    Dim oRsTmp1 As New Recordset
'    Dim oRsTmp2 As New Recordset
'    Dim oRsTmp3 As New Recordset
'    Dim oRsTmp4 As New Recordset
'    Dim oConexion As New ADODB.Connection
'    Dim lcSql As String
'
'    oConexion.CursorLocation = adUseClient
'    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
'...FALTA QUE ME CONFIRME CESAR, SI MUEVO FECHA-HORA DEL EXAMEN Y RESULTADO a la nueva FECHA INGRESO DE HOSP
'...FALTA CAMBIAR EL SERVICIO DONDE ESTUVO (SE TOMARA DEL SERV. INGRESO HOSP)
'    lcSql = "select * from LabMovimientoLaboratorio where IdMovimiento=" & txtNmovimiento.Text
'    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'    If oRsTmp1.RecordCount > 0 Then
'        oRsTmp1.MoveFirst
'        Do While Not oRsTmp1.EOF
'            lcSql = "select * from labMovimiento where idMovimiento=" & txtNmovimiento.Text
'            oRsTmp4.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'            If oRsTmp4.RecordCount > 0 Then
'                lcSql = "select * from atenciones where idCuentaAtencion=" & txtNcuentaNew.Text
'                oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                If oRsTmp3.RecordCount > 0 Then
'                    If oRsTmp3!FechaIngreso >= oRsTmp4!fecha Then
'                        lcSql = "update FactOrdenServicio set idcuentaAtencion=" & txtNcuentaNew.Text & " where IdOrden=" & oRsTmp1!IdOrden
'                        oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                        oRsTmp1.Fields!IdCuentaAtencion = Val(txtNcuentaNew.Text)
'                        oRsTmp1.Update
'                    Else
'                        MsgBox "La fecha "
'                    End If
'                End If
'                oRsTmp3.Close
'            End If
'            oRsTmp4.Close
'            oRsTmp1.MoveNext
'        Loop
'    End If
'    oRsTmp1.Close
'    oConexion.Close
    Unload Me
End Sub

Private Sub cmdActFecNac_Click()
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       Dim wrs_LolCli As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim lnNroHistoriaClinica As Long
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim lcSql As String
       Dim lntotReg As Long
       Dim lcHC As String
        Dim lcpacHis As String
        Dim lcpacPat As String
        Dim lcpacMat As String
        Dim lcpacNam As String
        Dim ldpacFin As Date
        Dim lcpacHis1 As String
        Dim lcpacPat1 As String
        Dim lcpacMat1 As String
        Dim lcpacNam1 As String
        Dim ldpacFin1 As Date
        Dim lbNuevo As Boolean
        Dim lcAutogenerado As String, lcAutogenerado1 As String
        Dim lnCant As Long
        Dim lnIdPaciente As Long, lnIdPaciente1 As Long
        Dim lcFechaNac As String
        With wrs_GalenHos
           'Autogenerado Repetidos
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado,dbo.Pacientes.IdPaciente," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre," & _
                   "  dbo.Pacientes.FechaNacimiento, dbo.Pacientes.segundoNombre, dbo.Pacientes.idTipoSexo" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   " Where  not ((dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or (dbo.HistoriasClinicas.HistoriaSistemaAnterior='')) " & _
                   "  order by dbo.Pacientes.autogenerado"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
              Do While Not .EOF
                    lcpacHis = .Fields!HistoriaSistemaAnterior
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = .Fields!PrimerNombre
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = .Fields!Autogenerado
                    lnIdPaciente = .Fields!idPaciente
                    lnCant = 0
                    Do While Not .EOF And lcAutogenerado = .Fields!Autogenerado
                        lnCant = lnCant + 1
                        If lnCant > 1 Then
                            lcpacHis1 = .Fields!HistoriaSistemaAnterior
                            lcpacPat1 = .Fields!ApellidoPaterno
                            lcpacMat1 = .Fields!ApellidoMaterno
                            lcpacNam1 = .Fields!PrimerNombre
                            ldpacFin1 = .Fields!FechaCreacion
                            lnIdPaciente1 = .Fields!idPaciente
                        End If
                        .MoveNext
                        If .EOF Then
                           Exit Do
                        End If
                    Loop
                    If lnCant > 1 Then
                       lcSql = "select * from Pacientes where idPaciente=" & lnIdPaciente
                       wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                       If wrs_GalenHos1.RecordCount > 0 Then
'                          lcFechaNac = wrs_GalenHos1.Fields!fechaNacimiento + 1
'                          lcAutogenerado = PacienteCrearNroAutogenerado(lcFechaNac, wrs_GalenHos1.Fields!apellidoPaterno, wrs_GalenHos1.Fields!apellidoMaterno, wrs_GalenHos1.Fields!primerNombre, wrs_GalenHos1.Fields!segundoNombre, wrs_GalenHos1.Fields!idTipoSexo)
'                          wrs_GalenHos1.Fields!fechaNacimiento = CDate(lcFechaNac)
'                          wrs_GalenHos1.Fields!autogenerado = lcAutogenerado
'                          wrs_GalenHos1.Update
                          lcAutogenerado = Left(wrs_GalenHos1.Fields!Autogenerado & "debb", 20)
                          wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado
                          wrs_GalenHos1.Update
                       End If
                       wrs_GalenHos1.Close
                       lcSql = "select * from Pacientes where idPaciente=" & lnIdPaciente1
                       wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                       If wrs_GalenHos1.RecordCount > 0 Then
'                          lcFechaNac = wrs_GalenHos1.Fields!fechaNacimiento - 1
'                          lcAutogenerado1 = PacienteCrearNroAutogenerado(lcFechaNac, wrs_GalenHos1.Fields!apellidoPaterno, wrs_GalenHos1.Fields!apellidoMaterno, wrs_GalenHos1.Fields!primerNombre, wrs_GalenHos1.Fields!segundoNombre, wrs_GalenHos1.Fields!idTipoSexo)
'                          wrs_GalenHos1.Fields!fechaNacimiento = CDate(lcFechaNac)
'                          wrs_GalenHos1.Fields!autogenerado = lcAutogenerado1
'                          wrs_GalenHos1.Update
                          lcAutogenerado1 = Left(wrs_GalenHos1.Fields!Autogenerado & "kike", 20)
                          wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado1
                          wrs_GalenHos1.Update
                       End If
                       wrs_GalenHos1.Close
                    End If
              Loop
           End If
           .Close
        End With
        'BuscarProblemasHCLolcli
    End If
End Sub



Private Sub cmdAfiliados_Click()
      MsgBox "Este reporte se ha movido a: sisgalenPlus -> reportes -> archivo clinico -> relación de historias clínicas pacientes VIH"

'     Me.MousePointer = 11
'     Dim lcBuscaParametro As New SIGHDatos.Parametros
'     Dim oRsTmp As New Recordset
'     Dim lcSql As String
'     Dim lnTotalAfiliados As Long, lnTotalMasculino As Long, lnTotalFemenino As Long, lnTotalMenorasAanios As Long
'     Dim lnPorFemenino As Double
'     Dim lnMujeresEdadFertil As Long
'     Dim lnGErn As Long, lnGE1anio As Long, lnGE1a4anios As Long, lnGE5a9anios As Long
'     Dim lnGE10a11anios As Long, lnGENinos As Long
'     Dim lnGEadolescente As Long, lnGEjoven As Long, lnGEadulto As Long, lnGEadultoMayor As Long
'     Dim ldHoy As Date, lcFiliacionInicio As String, lcFiliacionFinal As String
'     Dim lnGErnP As Double, lnGE1anioP As Double, lnGE1a4aniosP As Double, lnGE10a11aniosP As Double
'     Dim lnGE5a9aniosP    As Double, lnGENinosP As Double, lnGEadolescenteP As Double
'     Dim lnGEjovenP    As Double, lnGEadultoP As Double, lnGEadultoMayorP As Double, lnPorFertil As Double
'     '
'     ldHoy = lcBuscaParametro.RetornaFechaHoraServidorSQL
'     '
'     lcSql = "select fechaCreacion from HistoriasClinicas order by fechaCreacion"
'     If oRsTmp.State = 1 Then oRsTmp.Close
'     oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'     If oRsTmp.RecordCount = 0 Then
'        MsgBox "No existe ningun Paciente", vbInformation, "Pacientes"
'     Else
'        oRsTmp.MoveFirst
'        lcFiliacionInicio = Format(oRsTmp.Fields!FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
'        oRsTmp.MoveLast
'        lcFiliacionFinal = Format(oRsTmp.Fields!FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
'
'        '
'        'lcSql = "select idPaciente from Pacientes where idTipoSexo=1"
'        lcSql = "SELECT     dbo.Pacientes.idTipoSexo" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " where dbo.Pacientes.idTipoSexo=1"
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnTotalMasculino = oRsTmp.RecordCount
'        '
'        'lcSql = "select idPaciente from Pacientes where idTipoSexo=2"
'        lcSql = "SELECT     dbo.Pacientes.idTipoSexo" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " where dbo.Pacientes.idTipoSexo=2"
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnTotalFemenino = oRsTmp.RecordCount
'
'        '
'        lnTotalAfiliados = lnTotalFemenino + lnTotalMasculino
'        lnPorFemenino = Round((lnTotalFemenino * 100 / lnTotalAfiliados), 1)
'        '
''        lcSql = "select idPaciente from Pacientes where idTipoSexo=2" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=15" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<50 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " where dbo.Pacientes.idTipoSexo=2" & _
'                " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=15" & _
'                " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<50 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnMujeresEdadFertil = oRsTmp.RecordCount
'        lnPorFertil = Round(lnMujeresEdadFertil * 100 / lnTotalFemenino, 1)
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       " DATEDIFF(day, dbo.Pacientes.FechaNacimiento, getdate())<29"
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " where DATEDIFF(day, dbo.Pacientes.FechaNacimiento, getdate())<29"
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGErn = oRsTmp.RecordCount
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(day, dbo.Pacientes.FechaNacimiento, getdate())>=29" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<1 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " where    DATEDIFF(day, dbo.Pacientes.FechaNacimiento, getdate())>=29" & _
'                "          and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<1 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGE1anio = oRsTmp.RecordCount
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=1" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<5 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " WHERE    DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=1" & _
'                "          and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<5 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGE1a4anios = oRsTmp.RecordCount
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=5" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<10 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " WHERE    DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=5" & _
'                "        and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<10 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGE5a9anios = oRsTmp.RecordCount
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=10" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<12 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " WHERE    DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=10" & _
'                "    and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<12 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGE10a11anios = oRsTmp.RecordCount
'        lnGENinos = lnGErn + lnGE1anio + lnGE1a4anios + lnGE5a9anios + lnGE10a11anios
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=12" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<18 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " WHERE    DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=12" & _
'                "     and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<18 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGEadolescente = oRsTmp.RecordCount
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=18" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<30 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " WHERE    DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=18" & _
'                "     and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<30 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGEjoven = oRsTmp.RecordCount
'        '
''        lcSql = "select idPaciente from Pacientes where " & _
''                       "     DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=30" & _
''                       " and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<60 "
'        lcSql = "SELECT     dbo.Pacientes.FechaNacimiento" & _
'                " FROM         dbo.HistoriasClinicas LEFT OUTER JOIN" & _
'                "                      dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
'                " WHERE    DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())>=30" & _
'                "     and DATEDIFF(year, dbo.Pacientes.FechaNacimiento, getdate())<60 "
'        If oRsTmp.State = 1 Then oRsTmp.Close
'        oRsTmp.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'        lnGEadulto = oRsTmp.RecordCount
'        '
'        lnGEadultoMayor = lnTotalAfiliados - (lnGENinos + lnGEadolescente + lnGEjoven + lnGEadulto)
'        lnGErnP = Round(lnGErn * 100 / lnTotalAfiliados, 1)
'        lnGE1anioP = Round(lnGE1anio * 100 / lnTotalAfiliados, 1)
'        lnGE1a4aniosP = Round(lnGE1a4anios * 100 / lnTotalAfiliados, 1)
'        lnGE5a9aniosP = Round(lnGE5a9anios * 100 / lnTotalAfiliados, 1)
'        lnGE10a11aniosP = Round(lnGE10a11anios * 100 / lnTotalAfiliados, 1)
'        lnGENinosP = lnGErnP + lnGE1anioP + lnGE1a4aniosP + lnGE5a9aniosP + lnGE10a11aniosP
'        lnGEadolescenteP = Round(lnGEadolescente * 100 / lnTotalAfiliados, 1)
'        lnGEjovenP = Round(lnGEjoven * 100 / lnTotalAfiliados, 1)
'        lnGEadultoP = Round(lnGEadulto * 100 / lnTotalAfiliados, 1)
'        lnGEadultoMayorP = 100 - (lnGENinosP + lnGEadolescenteP + lnGEjovenP + lnGEadultoP)
'
'        '
'        Set FormPacientes.DataSource = oRsTmp
'        FormPacientes.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
'        FormPacientes.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
'        FormPacientes.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
'        FormPacientes.Sections("cabecera").Controls("lblTitulo").Caption = "Indicadores de Pacientes en el Establecimiento : " & lcBuscaParametro.SeleccionaFilaParametro(205)
'        FormPacientes.Sections("cabecera").Controls("lblFinicial").Caption = lcFiliacionInicio
'        FormPacientes.Sections("cabecera").Controls("lblFfinal").Caption = lcFiliacionFinal
'        FormPacientes.Sections("cabecera").Controls("lblTotalPacientes").Caption = Format(lnTotalAfiliados, "###,###,###")
'        FormPacientes.Sections("cabecera").Controls("lblTotalMujeres").Caption = Format(lnTotalFemenino, "###,###,###")
'        FormPacientes.Sections("cabecera").Controls("lblTotalMujeresP").Caption = Format(lnPorFemenino, "###0.0")
'        FormPacientes.Sections("cabecera").Controls("lblTotalHombres").Caption = Format(lnTotalMasculino, "###,###,###")
'        FormPacientes.Sections("cabecera").Controls("lblTotalHombresP").Caption = Format(100 - lnPorFemenino, "###0.0")
'        FormPacientes.Sections("cabecera").Controls("lblFertil").Caption = IIf(lnMujeresEdadFertil = 0, "0", Format(lnMujeresEdadFertil, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblFertilP").Caption = IIf(lnPorFertil = 0, "0", Format(lnPorFertil, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lblNinos").Caption = IIf(lnGENinos = 0, "0", Format(lnGENinos, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblNinosP").Caption = IIf(lnGENinosP = 0, "0", Format(lnGENinosP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lblRN").Caption = IIf(lnGErn = 0, "0", Format(lnGErn, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblRNp").Caption = IIf(lnGErnP = 0, "0", Format(lnGErnP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lbl1anio").Caption = IIf(lnGE1anio = 0, "0", Format(lnGE1anio, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lbl1anioP").Caption = IIf(lnGE1anioP = 0, "0", Format(lnGE1anioP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lbl1a4").Caption = IIf(lnGE1a4anios = 0, "0", Format(lnGE1a4anios, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lbl1a4P").Caption = IIf(lnGE1a4aniosP = 0, "0", Format(lnGE1a4aniosP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lbl5a9").Caption = IIf(lnGE5a9anios = 0, "0", Format(lnGE5a9anios, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lbl5a9p").Caption = IIf(lnGE5a9aniosP = 0, "0", Format(lnGE5a9aniosP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lbl10a11").Caption = IIf(lnGE10a11anios = 0, "0", Format(lnGE10a11anios, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lbl10a11P").Caption = IIf(lnGE10a11aniosP = 0, "0", Format(lnGE10a11aniosP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lblAdolecente").Caption = IIf(lnGEadolescente = 0, "0", Format(lnGEadolescente, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblAdolecenteP").Caption = IIf(lnGEadolescenteP = 0, "0", Format(lnGEadolescenteP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lblJoven").Caption = IIf(lnGEjoven = 0, "0", Format(lnGEjoven, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblJovenP").Caption = IIf(lnGEjovenP = 0, "0", Format(lnGEjovenP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lblAdulto").Caption = IIf(lnGEadulto = 0, "0", Format(lnGEadulto, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblAdultoP").Caption = IIf(lnGEadultoP = 0, "0", Format(lnGEadultoP, "###0.0"))
'        FormPacientes.Sections("cabecera").Controls("lblAdultoM").Caption = IIf(lnGEadultoMayor = 0, "0", Format(lnGEadultoMayor, "###,###,###"))
'        FormPacientes.Sections("cabecera").Controls("lblAdultoMp").Caption = IIf(lnGEadultoMayorP = 0, "0", Format(lnGEadultoMayorP, "###0.0"))
'        'FormPacientes.Orientation = rptOrientPortrait
'        FormPacientes.Show 1
'        '
'    End If
'    Me.MousePointer = 1
'    Unload Me
End Sub



Private Sub cmdCartillas_Click()
    ProgressBar2.Min = 0
    ProgressBar2.Max = Val(Me.txtCartillas.Text)
    ProgressBar2.Value = 0
    Dim lnJugada As Integer, lcCartilla As String, lnSegundo As Integer
    Dim oRsJugadas As New Recordset
    Dim oRsPie As New Recordset
    Const LcEspacio = ""
    With oRsJugadas
        .Fields.Append "jugada", adInteger
        .Fields.Append "ganador", adVarChar, 200
        .LockType = adLockOptimistic
        .Open
    End With
    '
    lcCartilla = "/"
    oRsCartillas.MoveFirst
    Do While Not oRsCartillas.EOF
       lcCartilla = lcCartilla & LcEspacio & UCase(oRsCartillas!ganador) & LcEspacio & "/"
       oRsCartillas.MoveNext
    Loop
    oRsJugadas.AddNew
    oRsJugadas!jugada = 1
    oRsJugadas!ganador = lcCartilla
    oRsJugadas.Update
    '
    For lnJugada = 2 To Val(Me.txtCartillas.Text)
        DoEvents:  ProgressBar2.Value = ProgressBar2.Value + 1: Me.Refresh
        Do While True
           '
           lcCartilla = "/"
           oRsCartillas.MoveFirst
           Do While Not oRsCartillas.EOF
               If UCase(oRsCartillas!fijo) = "X" Then
                    lcCartilla = lcCartilla & LcEspacio & UCase(oRsCartillas!ganador) & LcEspacio & "/"
               Else
'                    mo_ReglasComunes.WaitSeconds 1
                    lnSegundo = Second(Time)
                    If lnSegundo < 20 Then
                       lcCartilla = lcCartilla & LcEspacio & "L" & LcEspacio & "/"
                    ElseIf lnSegundo < 40 Then
                       lcCartilla = lcCartilla & LcEspacio & "E" & LcEspacio & "/"
                    Else
                       lcCartilla = lcCartilla & LcEspacio & "V" & LcEspacio & "/"
                    End If
               End If
               oRsCartillas.MoveNext
           Loop
           '
           oRsJugadas.Find "ganador='" & lcCartilla & "'"
           If oRsJugadas.EOF Then
              Exit Do
           End If
        Loop
        oRsJugadas.AddNew
        oRsJugadas!jugada = lnJugada
        oRsJugadas!ganador = lcCartilla
        oRsJugadas.Update
        lblNro(4).Caption = lnJugada
    Next
    '
    'Set oRsPie = SIGHEntidades.CopyRecordset(oRsJugadas)
'    Dim mo_AdminReportes11 As New SIGHNegocios.ReglasReportes
'    mo_AdminReportes11.ExportarRecordSetAexcel oRsJugadas, "GANAGOL", "", ".", Me.hwnd, False, True, oRsPie
    Unload Me
End Sub

Private Sub cmdCSnazarena_Click()
    mo_ReglasAdmision.his_historicoAtencionesEliminarTodas wxConexionRed
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter

    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=HIS"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
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
       Dim wFec1 As String, wFec2 As String, wFFecha As Date
       With wrs_LolCli
           wFFecha = CDate(Me.txtFIniCSN.Text)
           wFec1 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           wFFecha = CDate(Me.txtFFinCSN.Text)
           wFec2 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           
           Me.txtIni1.Text = Me.txtFIniCSN.Text
           Me.txtFin1.Text = Me.txtFFinCSN.Text
           'wRsProblemas.Open "delete from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           'elimina historias GalenHos de esas Fechas ---comentado anteriormente
           lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "HistoriasClinicasSeleccionarPorFechaCreacion"
                Set oParameter = .CreateParameter("@FechaInicio", adVarChar, adParamInput, 20, Me.txtIni1.Text): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaFin", adVarChar, adParamInput, 20, Me.txtFin1.Text): .Parameters.Append oParameter
                Set wrs_GalenHos = .Execute
                Set wrs_GalenHos.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           Set oParameter = Nothing
           
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar2.Min = 0
                ProgressBar2.Max = lntotReg
                lnRegAct = 0
                
                Do While Not wrs_GalenHos.EOF
                   DoEvents
                   lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                   Me.Refresh
                   wxConexionRed.BeginTrans
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   
                   'MODIFICADO POR FRANKLIN CACHAY 13/11/2013 - se cambio por problemas en el delete y update por store
'                   wrs_GalenHos.Delete
'                   wrs_GalenHos.Update
                   With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "HistoriasClinicasEliminar"
                         Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, wrs_GalenHos.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                         Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                         .Execute
                   End With
                   Set oCommand = Nothing
                   Set oParameter = Nothing
                   
                   
                   With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "PacientesEliminarPorIdPaciente"
                         Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
                         .Execute
                   End With
                   Set oCommand = Nothing
                   Set oParameter = Nothing
           
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando11.Caption = "Insertando HC en GalenHos"
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "HistoriasClinicasSeleccionarTodos"
                Set wrs_GalenHos = .Execute
                Set wrs_GalenHos.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "PacientesConsultarTodos"
                Set wrs_GalenHos1 = .Execute
                Set wrs_GalenHos1.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           lcSql = "select * from tarjeta " & _
                    "  where  fec_ins>=" & wFec1 & " and fec_ins<=" & wFec2 & _
                     " order by fec_ins"
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en CS Nazareno"
              wxConexionJAMO.Close
              Exit Sub
           End If
           
           If chkFichaFamiliar1.Value = 1 Then
                    lnTipoNumeracion = 1
                    lnHistoriaAutogenerada = 1000001
                    
                    
                    With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "PacientesSeleccionarPorIdTipoNumeracion1"
                         Set wRsProblemas = .Execute
                         Set wRsProblemas.ActiveConnection = Nothing
                    End With
                    Set oCommand = Nothing
           
                    If wRsProblemas.RecordCount > 0 Then
                       lnHistoriaAutogenerada = wRsProblemas.Fields!NroHistoriaClinica + 1
                    End If
                    wRsProblemas.Close
                    .MoveFirst
                    Do While Not .EOF
                       If IsNull(.Fields!fichaf) Then
                          MsgBox "Existe FICHA FAMILIAR vacia, NO SE PODRA SEGUIR AÑADIENDO FICHAS FAMILIARES", vbCritical, "Mensaje"
                          Exit Sub
                       End If
                       lcSegundoNombre = Trim(.Fields!fichaf)
                       lnCantGuiones = 0
                       For lnFor = 1 To Len(lcSegundoNombre)
                           If Mid(lcSegundoNombre, lnFor, 1) = "-" Then
                              lnCantGuiones = lnCantGuiones + 1
                           End If
                       Next
                       If lnCantGuiones <> 2 Then
                          MsgBox "Ficha: " & .Fields!fichaf & Chr(13) & "Todas las Fichas Familiares deben tener 2 GUIONES" & Chr(13) & "El formato de la FICHA FAMILIAR es: sector-NumeroHistoria-NumeroIntegranteFamilia" & Chr(13) & " NO SE PODRA SEGUIR AÑADIENDO FICHAS FAMILIARES", vbCritical, "Mensaje"
                          Exit Sub
                       End If
                       .MoveNext
                    Loop
           End If
           
           
           With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = wxConexionRed
                  .CommandTimeout = 150
                  .CommandText = "lolcliProblemasHCSeleccionarTodos"
                  Set wRsProblemas = .Execute
                  Set wRsProblemas.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 0
           Do While Not .EOF
              DoEvents
              ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
              Me.Refresh
              lbContinuarProceso = True
              If chkFichaFamiliar1.Value = 1 Then
                    If IsNull(.Fields!ape_pat) Or Trim(.Fields!ape_pat) = "" Or IsNull(.Fields!ape_mat) Or Trim(.Fields!ape_mat) = "" Or IsNull(.Fields!nro_his) Or Trim(.Fields!fichaf) = "" Then
                       lbContinuarProceso = False
                    End If
              Else
                    If IsNull(.Fields!ape_pat) Or Trim(.Fields!ape_pat) = "" Or IsNull(.Fields!ape_mat) Or Trim(.Fields!ape_mat) = "" Or IsNull(.Fields!nro_his) Or Trim(.Fields!nro_his) = "" Then
                       lbContinuarProceso = False
                    End If
              End If
              If lbContinuarProceso = True Then
                  lbNuevoHC = False
                  If Not sighentidades.EsFecha(Trim(.Fields!fec_nac), "DD/MM/AAAA") Then
                        If .Fields!Edad > 120 Then
                           lcFechaNac = "01/01/" + Trim(Str(Year(Date)))
                        Else
                           lcFechaNac = "01/01/" + Trim(Str(Year(Date) - .Fields!Edad))
                        End If
                  Else
                        lcFechaNac = .Fields!fec_nac
                  End If
                  lnNroHistoriaClinica = 0
                  If chkFichaFamiliar1.Value = 1 Then
                     lnNroHistoriaClinica = lnHistoriaAutogenerada
                     lnHistoriaAutogenerada = lnHistoriaAutogenerada + 1
                  Else
                     lnNroHistoriaClinica = Val(.Fields!nro_his)
                  End If
                  If lnNroHistoriaClinica = 0 Then
                     'Historia clinica con problemas
                     lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!nro_his)
    '                 lbNuevoHC = True
                  End If
                  If IsNull(.Fields!nombres1) Then
                      lcPrimerNombre = "NN"
                  Else
                      lcPrimerNombre = Trim(.Fields!nombres1)
                  End If
                  If IsNull(.Fields!Nombres2) Then
                     lcSegundoNombre = ""
                  Else
                     lcSegundoNombre = Trim(.Fields!Nombres2)
                  End If
                  
                  lnTipoSexo = IIf(UCase(.Fields!sexo) = "M", 1, 2)
                  If lnNroHistoriaClinica = 0 Then
                     lbNuevoHC = True
                  Else
                        
                        With oCommand
                               .CommandType = adCmdStoredProc
                               Set .ActiveConnection = wxConexionRed
                               .CommandTimeout = 150
                               .CommandText = "PacientesSeleccionarPorNroHistoriaClinica"
                               Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                               Set wrs_GalenHos2 = .Execute
                               Set wrs_GalenHos2.ActiveConnection = Nothing
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
           
                        If wrs_GalenHos2.RecordCount > 0 Then
                           lbNuevoHC = True
                        End If
                  End If
                  
                  If lbNuevoHC = True Then
                          'HC repetida
                          'solo graba como problemas
                          
                          
                              oCommand.CommandType = adCmdStoredProc
                              Set oCommand.ActiveConnection = wxConexionRed
                              oCommand.CommandTimeout = 150
                              oCommand.CommandText = "lolcliProblemasHCAgregar"
                              Set oParameter = oCommand.CreateParameter("@pacHis", adChar, adParamInput, 50, IIf(chkFichaFamiliar1.Value = 1, .Fields!nro_his, .Fields!fichaf)): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacPat", adChar, adParamInput, 30, .Fields!ape_pat): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacMat", adChar, adParamInput, 30, .Fields!ape_mat): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacNam", adChar, adParamInput, 50, .Fields!nombres1): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacFin", adDBTimeStamp, adParamInput, 0, CDate(Me.txtFIniCSN.Text)): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@nroHistoriaGalenHos", adVarChar, adParamInput, 50, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@autogeneradoGalenHos", adVarChar, adParamInput, 50, "*" & Trim(Str(lnRegAct)) & "HcRep-Nazareno"): oCommand.Parameters.Append oParameter
                              oCommand.Execute
                            
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                          '
                          If lnNroHistoriaClinica <> 0 Then
                          
                            With oCommand
                                .CommandType = adCmdStoredProc
                                Set .ActiveConnection = wxConexionRed
                                .CommandTimeout = 150
                                .CommandText = "lolcliProblemasHCAgregarSinNroHistoriaClinica"
                                Set oParameter = .CreateParameter("@pacHis", adChar, adParamInput, 50, wrs_GalenHos2.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacPat", adChar, adParamInput, 30, wrs_GalenHos2.Fields!ApellidoPaterno): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacMat", adChar, adParamInput, 30, wrs_GalenHos2.Fields!ApellidoMaterno): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacNam", adChar, adParamInput, 50, Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@autogeneradoGalenHos", adVarChar, adParamInput, 50, "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"): .Parameters.Append oParameter
                                .Execute
                            End With
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                            '
                            wrs_GalenHos2.Close
                          End If
                  Else
                          wrs_GalenHos2.Close
                          lcApellidoPaterno = UCase(Left(Trim(.Fields!ape_pat), 20))
                          lcApellidoMaterno = UCase(Left(Trim(.Fields!ape_mat), 20))
                          lcPrimerNombre = UCase(lcPrimerNombre)
                          lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                          lcFechaAnt = .Fields!fec_ins
                          'Busca en Tabla xx Equivalencia LolCli
                          lntipoOcupacion = 0
                          lnIdDepartamentoDomicilio = 0
                          LnIdProvinciaDomicilio = 0
                          lnIdDistritoDomicilio = 0
'                          If Not IsNull(.Fields!codgeo) Then
'                            lnIdDepartamentoDomicilio = Val(Left(.Fields!codgeo, 2))
'                            LnIdProvinciaDomicilio = Val(Left(.Fields!codgeo, 4))
'                            lnIdDistritoDomicilio = Val(.Fields!codgeo)
'                          End If
                          lnIdDepartamentoNacimiento = 0
                          LnIdProvinciaNacimiento = 0
                          lnIdDistritoNacimiento = 0
                          lnIdEstadoCivil = 0
'                          If Not IsNull(.Fields!estCivil) Then
'                             Select Case UCase(Left(.Fields!estCivil, 1))
'                             Case "S"   'soltero
'                                  lnIdEstadoCivil = 2
'                             Case "C"   'casado
'                                  lnIdEstadoCivil = 1
'                             End Select
'                          End If
                          'Graba Pacientes
                          wxConexionRed.BeginTrans

                          
                                oCommand.CommandType = adCmdStoredProc
                                Set oCommand.ActiveConnection = wxConexionRed
                                oCommand.CommandTimeout = 150
                                oCommand.CommandText = "PacientesAgregarPorHistoriaClinica"
                                Set oParameter = oCommand.CreateParameter("@IdPaciente", adInteger, adParamOutput, 0, 0): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, lcApellidoPaterno): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, lcApellidoMaterno): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, lcPrimerNombre): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, lcSegundoNombre): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, lnTipoSexo): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, CDate(lcFechaNac)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnTipoNumeracion): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, lcAutogenerado): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, lnIdDistritoDomicilio): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(Not IsNull(.Fields!dom_act), Left(Trim(.Fields!dom_act), 50), "")): oCommand.Parameters.Append oParameter
                                If Not IsNull(.Fields!DNI) Then
                                   If Len(Trim(.Fields!DNI)) = 8 Then
                                        Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, Left(.Fields!DNI, 8)): oCommand.Parameters.Append oParameter
                                        Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                   Else
                                        Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, ""): oCommand.Parameters.Append oParameter
                                        Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, Null): oParameter.Attributes = adParamNullable: oCommand.Parameters.Append oParameter
                                   End If
                                Else
                                    Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, ""): oCommand.Parameters.Append oParameter
                                    Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, Null): oParameter.Attributes = adParamNullable: oCommand.Parameters.Append oParameter
                                End If
                                Set oParameter = oCommand.CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!nomPa), Left(.Fields!nomPa, 20), "")): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!nomMa), Left(.Fields!nomMa, 20), "")): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!fichaf), Left(.Fields!fichaf, 20), "")): oCommand.Parameters.Append oParameter
                                oCommand.Execute
                                lnIdPaciente = oCommand.Parameters("@IdPaciente")
                          Set oCommand = Nothing
                          Set oParameter = Nothing
                          
                          'Graba HistoriasClinicas

                        
                                oCommand.CommandType = adCmdStoredProc
                                Set oCommand.ActiveConnection = wxConexionRed
                                oCommand.CommandTimeout = 150
                                oCommand.CommandText = "HistoriasClinicasAgregarPorIdPaciente"
                                Set oParameter = oCommand.CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, CDate(lcFechaAnt)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnTipoNumeracion): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@HistoriaSistemaAnterior", adVarChar, adParamInput, 50, .Fields!nro_his): oCommand.Parameters.Append oParameter
                                oCommand.Execute
                          Set oCommand = Nothing
                          Set oParameter = Nothing
                         
                          wxConexionRed.CommitTrans
                  End If
              End If
              .MoveNext
           Loop
       End With
       wxConexionJAMO.Close
    End If
    cmdCargaAtencionesHist_Click
    
    Unload Me
    Exit Sub
err_proceso:
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!HC & "     Paciente:" & wrs_LolCli.Fields!Paterno & " " & wrs_LolCli.Fields!Materno & " " & wrs_LolCli.Fields!Pnombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    'Resume
    Unload Me

End Sub


Private Sub cmdGrabaDistritoProvinciaDepartamento_Click()
    If MsgBox("¿Esta seguro de actualizar Distritos,Provincia y Departamentos ?", vbYesNo, Me.Caption) = vbNo Then Exit Sub
     
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim lnIdDepartamento As Long
    Dim lcDescrDepartamento As String
    Dim lnIdProvincia As Long
    Dim lcDescrProvincia As String
    Dim lnIdDistrito As Long
    Dim lcDescrDistrito As String
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtExcelrptUbegeo.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("ubicacionGeografica")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, lcCodigo As String
    Dim lcDepartamento As String, lcProvincia As String, lcDistrito As String
    lnFila = 2
    lnFilaFinal = 20000
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Dim oDoDepartamento As New DODepartamento
    Dim oDoProvincia As New DOProvincia
    Dim oDoDistrito As New DODistrito
    Dim Departamentos As New SIGHDatos.Departamentos
    Dim Provincia As New SIGHDatos.Provincias
    Dim Distrito As New SIGHDatos.Distritos
    
    Set Departamentos.Conexion = oConexion
    Set Provincia.Conexion = oConexion
    Set Distrito.Conexion = oConexion
    For lnFor = lnFila To lnFilaFinal
        lcRango = "B" + Trim(Str(lnFor))
        If Trim(Trim(s.Range(lcRango).Value)) = "" Then Exit For
        
        lnIdDepartamento = CLng(Mid(Trim(s.Range(lcRango).Value), 1, 2))
        lcDescrDepartamento = Mid(Trim(s.Range(lcRango).Value), 4, Len(Trim(s.Range(lcRango).Value)) - 3)

        oDoDepartamento.IdDepartamento = lnIdDepartamento
        oDoDepartamento.nombre = lcDescrDepartamento
        If Departamentos.SeleccionarPorId(oDoDepartamento) = True Then
        Else
            If Departamentos.Insertar(oDoDepartamento) = False Then
                Exit For
            End If
        End If
           
        lcRango = "E" + Trim(Str(lnFor))
        If Not (Trim(s.Range(lcRango).Value) = "") Then
'           lnIdProvincia = CLng(Mid(Trim(Val(s.Range(lcRango).Value)), 0, 2))
           lcDescrProvincia = Mid(Trim(s.Range(lcRango).Value), 4, Len(Trim(s.Range(lcRango).Value)) - 3)
           
           lcRango = "K" + Trim(Str(lnFor))
           lcCodigo = Mid(Trim(s.Range(lcRango).Value), 1, 4)
                    
           oDoProvincia.IdDepartamento = lnIdDepartamento
           oDoProvincia.IdProvincia = CLng(lcCodigo)
           oDoProvincia.nombre = lcDescrProvincia
           If Provincia.SeleccionarPorId(oDoProvincia) = True Then
           Else
               If Provincia.Insertar(oDoProvincia) = False Then
                   Exit For
               End If
           End If
            
           lcRango = "F" + Trim(Str(lnFor))
           If Not (Trim(s.Range(lcRango).Value) = "") Then
'                lnIdDistrito = CLng(Mid(Trim(Val(s.Range(lcRango).Value)), 0, 2))
                lcDescrDistrito = Mid(Trim(s.Range(lcRango).Value), 4, Len(Trim(s.Range(lcRango).Value)) - 3)
                lcRango = "K" + Trim(Str(lnFor))
                lcCodigo = CLng(Trim(s.Range(lcRango).Value))
                
                oDoDistrito.IdProvincia = oDoProvincia.IdProvincia
                oDoDistrito.IdDistrito = lcCodigo
                oDoDistrito.nombre = lcDescrDistrito
                If Distrito.SeleccionarPorId(oDoDistrito) = True Then
                Else
                    If Distrito.Insertar(oDoDistrito) = False Then
                        Exit For
                    End If
                End If

            End If
        End If
    Next
    MsgBox "Terminó exitosamente el proceso de actualización de las tablas Departamentos, Provincias y Distritos", vbInformation, Me.Caption
    Set s = Nothing
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    Unload Me

End Sub

Private Sub cmdHCproblemas_Click()
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        BuscarProblemasHCLolcli
     End If
       
End Sub

Sub BuscarProblemasHCLolcli()
       Dim wrs_LolCli As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim lnNroHistoriaClinica As Long
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim lcSql As String
       Dim lntotReg As Long
       Dim lcHC As String
        Dim lcpacHis As String
        Dim lcpacPat As String
        Dim lcpacMat As String
        Dim lcpacNam As String
        Dim ldpacFin As Date
        Dim lcpacHis1 As String
        Dim lcpacPat1 As String
        Dim lcpacMat1 As String
        Dim lcpacNam1 As String
        Dim ldpacFin1 As Date
        Dim lbNuevo As Boolean
        Dim lcAutogenerado As String
        Dim lnCant As Long
              Me.MousePointer = 11
 
       wrs_GalenHos1.Open "delete  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
       wrs_GalenHos1.Open "select *  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
       With wrs_GalenHos
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, " & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   " dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   "  WHERE len(ltrim(dbo.HistoriasClinicas.NroHistoriaClinica))=8    "
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
                .MoveFirst
                Do While Not .EOF
                   'lnNroHistoriaClinica = Val(.Fields!pacHis)
                   If Year(.Fields!FechaCreacion) = 2009 Then
                      wrs_GalenHos1.AddNew
                      wrs_GalenHos1.Fields!pacHis = .Fields!HistoriaSistemaAnterior
                      wrs_GalenHos1.Fields!pacPat = .Fields!ApellidoPaterno
                      wrs_GalenHos1.Fields!pacMat = .Fields!ApellidoMaterno
                      wrs_GalenHos1.Fields!pacNam = .Fields!PrimerNombre
                      wrs_GalenHos1.Fields!pacFin = .Fields!FechaCreacion
                      'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                      wrs_GalenHos1.Update
                      If Left(.Fields!HistoriaSistemaAnterior, 1) = "0" Then
                          lcHC = Trim(Val(.Fields!HistoriaSistemaAnterior))
                      Else
                          lcHC = "0" & Trim(.Fields!HistoriaSistemaAnterior)
                      End If
                      lcSql = "SELECT     dbo.Pacientes.IdPaciente, dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, " & _
                                "          dbo.Pacientes.SegundoNombre , dbo.Pacientes.NroHistoriaClinica, dbo.HistoriasClinicas.HistoriaSistemaAnterior," & _
                                "          dbo.HistoriasClinicas.fechaCreacion, dbo.HistoriasClinicas.HistoriaSistemaAnterior" & _
                                " FROM         dbo.Pacientes INNER JOIN" & _
                                "                      dbo.HistoriasClinicas ON dbo.Pacientes.IdPaciente = dbo.HistoriasClinicas.IdPaciente" & _
                                " WHERE     (dbo.HistoriasClinicas.HistoriaSistemaAnterior = '" & lcHC & "')" & _
                                " ORDER BY dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre"
                      wrs_GalenHos2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                      If wrs_GalenHos2.RecordCount > 0 Then
                         wrs_GalenHos2.MoveFirst
                         Do While Not wrs_GalenHos2.EOF
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = wrs_GalenHos2.Fields!HistoriaSistemaAnterior
                            wrs_GalenHos1.Fields!pacPat = wrs_GalenHos2.Fields!ApellidoPaterno
                            wrs_GalenHos1.Fields!pacMat = wrs_GalenHos2.Fields!ApellidoMaterno
                            wrs_GalenHos1.Fields!pacNam = wrs_GalenHos2.Fields!PrimerNombre
                            wrs_GalenHos1.Fields!pacFin = wrs_GalenHos2.Fields!FechaCreacion
                            'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                            wrs_GalenHos1.Update
                            wrs_GalenHos2.MoveNext
                         Loop
                      End If
                      wrs_GalenHos2.Close
                   End If
                   .MoveNext
                Loop
           End If
           .Close
           'Historias Clinicas Anteriores NULL o VACIAS
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   "  where (dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or dbo.HistoriasClinicas.HistoriaSistemaAnterior=''"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
              Do While Not .EOF
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = .Fields!PrimerNombre
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = .Fields!Autogenerado
                    wrs_GalenHos1.AddNew
                    wrs_GalenHos1.Fields!pacHis = ""
                    wrs_GalenHos1.Fields!pacPat = lcpacPat
                    wrs_GalenHos1.Fields!pacMat = lcpacMat
                    wrs_GalenHos1.Fields!pacNam = lcpacNam
                    wrs_GalenHos1.Fields!pacFin = ldpacFin
                    'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                    wrs_GalenHos1.Update
                    .Update
                    .MoveNext
              Loop
           End If
           .Close
           'Autogenerado Repetidos
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   " Where  not ((dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or (dbo.HistoriasClinicas.HistoriaSistemaAnterior='')) " & _
                   "  order by dbo.Pacientes.autogenerado"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
              Do While Not .EOF
                    lcpacHis = .Fields!HistoriaSistemaAnterior
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = .Fields!PrimerNombre
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = .Fields!Autogenerado
                    lnCant = 0
                    Do While Not .EOF And lcAutogenerado = .Fields!Autogenerado
                        lnCant = lnCant + 1
                        If lnCant > 1 Then
                            lcpacHis1 = .Fields!HistoriaSistemaAnterior
                            lcpacPat1 = .Fields!ApellidoPaterno
                            lcpacMat1 = .Fields!ApellidoMaterno
                            lcpacNam1 = .Fields!PrimerNombre
                            ldpacFin1 = .Fields!FechaCreacion
                        End If
                        .MoveNext
                        If .EOF Then
                           Exit Do
                        End If
                    Loop
                    If lnCant > 1 Then
                       lbNuevo = True
                       If wrs_GalenHos1.RecordCount > 0 Then
                          wrs_GalenHos1.MoveFirst
                          wrs_GalenHos1.Find "pacHis='" & lcpacHis & "'"
                          If Not wrs_GalenHos1.EOF Then
                             lbNuevo = False
                          End If
                       End If
                       If lbNuevo = True Then
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis
                            wrs_GalenHos1.Fields!pacPat = lcpacPat
                            wrs_GalenHos1.Fields!pacMat = lcpacMat
                            wrs_GalenHos1.Fields!pacNam = lcpacNam
                            wrs_GalenHos1.Fields!pacFin = ldpacFin
                            wrs_GalenHos1.Fields!autogeneradoGalenHos = lcAutogenerado
                            'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                            wrs_GalenHos1.Update
                            '
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis1
                            wrs_GalenHos1.Fields!pacPat = lcpacPat1
                            wrs_GalenHos1.Fields!pacMat = lcpacMat1
                            wrs_GalenHos1.Fields!pacNam = lcpacNam1
                            wrs_GalenHos1.Fields!pacFin = ldpacFin1
                            wrs_GalenHos1.Fields!autogeneradoGalenHos = lcAutogenerado
                            'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                            wrs_GalenHos1.Update
                       End If
                    End If
              Loop
           End If
           .Close
           'Historias  Repetidas
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   " Where  not ((dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or (dbo.HistoriasClinicas.HistoriaSistemaAnterior='')) " & _
                   "  order by dbo.HistoriasClinicas.HistoriaSistemaAnterior"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
              Do While Not .EOF
                    lcpacHis = .Fields!HistoriaSistemaAnterior
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = .Fields!PrimerNombre
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = Trim(.Fields!HistoriaSistemaAnterior)
                    lnCant = 0
                    Do While Not .EOF And lcAutogenerado = Trim(.Fields!HistoriaSistemaAnterior)
                        lnCant = lnCant + 1
                        If lnCant > 1 Then
                            lcpacHis1 = .Fields!HistoriaSistemaAnterior
                            lcpacPat1 = .Fields!ApellidoPaterno
                            lcpacMat1 = .Fields!ApellidoMaterno
                            lcpacNam1 = .Fields!PrimerNombre
                            ldpacFin1 = .Fields!FechaCreacion
                        End If
                        .MoveNext
                        If .EOF Then
                           Exit Do
                        End If
                    Loop
                    If lnCant > 1 Then
                       lbNuevo = True
                       If wrs_GalenHos1.RecordCount > 0 Then
                          wrs_GalenHos1.MoveFirst
                          wrs_GalenHos1.Find "pacHis='" & lcpacHis & "'"
                          If Not wrs_GalenHos1.EOF Then
                             lbNuevo = False
                          End If
                       End If
                       If lbNuevo = True Then
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis
                            wrs_GalenHos1.Fields!pacPat = lcpacPat
                            wrs_GalenHos1.Fields!pacMat = lcpacMat
                            wrs_GalenHos1.Fields!pacNam = lcpacNam
                            wrs_GalenHos1.Fields!pacFin = ldpacFin
                            'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                            wrs_GalenHos1.Update
                            '
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis1
                            wrs_GalenHos1.Fields!pacPat = lcpacPat1
                            wrs_GalenHos1.Fields!pacMat = lcpacMat1
                            wrs_GalenHos1.Fields!pacNam = lcpacNam1
                            wrs_GalenHos1.Fields!pacFin = ldpacFin1
                            'wrs_GalenHos1.Fields!nroHistoriaGalenhos = .Fields!NroHistoriaClinica
                            wrs_GalenHos1.Update
                       End If
                    End If
              Loop
           End If
           .Close
           
       End With
       
      Me.MousePointer = 1

       Unload Me

End Sub


Private Sub cmdHISvsGalenhos_Click()
        If oRsGrdSIS.RecordCount > 0 Then
           If Me.txtSisCatEESS.Text = "" Or Me.txtSisDisa.Text = "" Or Me.txtSISptoDigitacion.Text = "" Or Me.txtSISCodigoUDR.Text = "" Then
              MsgBox "Tiene que ingresar los 4 datos para el SIS", vbCritical
              Exit Sub
           End If
        End If
        Dim EXL As Excel.Application
        Set EXL = New Excel.Application
        Dim W As Excel.Workbook
        Dim s As Excel.Worksheet
        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        Dim oRsFox1 As New Recordset
        Dim oConexionFox As New Connection
        Dim oConexion As New Connection
        Dim oConexionSIS As New Connection
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim oDOhis_historicoAten As New DOhis_historicoAten, oHIS_historicoAten As New HIS_historicoAten
        Dim oDOPaciente As New DOPaciente, oPacientes As New Pacientes
        Dim oDOHistoriaClinica As New DOHistoriaClinica, oHistoriasClinicas As New HistoriasClinicas
        Dim oDOPArametro As New DOPArametro, oParametros As New Parametros
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim oSisConsumoWeb As New SIGHNegocios.SisConsumoWeb
        Dim lnIdUsuario As Long, lbContinuar As Boolean, ldFechaNac As Date
        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
        Dim lcSegundoNombre As String, lnTipoSexo As Long, lnIdPaciente As Long
        Dim lcAutogenerado As String, lcDNI As String, lnNroHistoriaClinica As Long
        Dim lcSql As String, lcCodDx As String, lcCod2000 As String, lbEsNuevaHC As Boolean
        Dim oFila As Long, lcFichaFam As String
        Const lnIdTipoNumeracion As Long = 2
        On Error GoTo ErrProAtHis
        '
        Set W = EXL.Workbooks.Open("c:\excel.xls")
        Set s = W.Sheets("hoja1")
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        lcSql = "select * from histCab"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        
        If oRsFox.RecordCount > 0 Then
            ProgressBar2.Max = oRsFox.RecordCount
        End If
        If oRsFox.RecordCount > 0 Then
           ProgressBar2.Min = 0
           ProgressBar2.Value = 0
           
           oFila = 1
           s.Cells(oFila, 1).Value = "Año Filiación"
           s.Cells(oFila, 3).Value = "N° Historia"
           s.Cells(oFila, 2).Value = "DNI"
           s.Cells(oFila, 4).Value = "Autogenerado SISGALENPLUS"
           s.Cells(oFila, 5).Value = "Autogenerado HIS"
           oFila = oFila + 2
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
               DoEvents: ProgressBar2.Value = ProgressBar2.Value + 1: Me.Refresh
               lcDNI = Right("        " & Trim(oRsFox.Fields!DNI), 8)
               Set oRsTmp = mo_ReglasAdmision.PacientesXdni(lcDNI, oConexion)
               If oRsTmp.RecordCount > 0 Then
                    Set oRsFox1 = mo_ReglasArchivoClinico.HistoriasClinicasXIdPaciente(oRsTmp!idPaciente, oConexion)
                    If oRsFox1.RecordCount > 0 Then
                          lcFichaFam = Left(UCase(oRsTmp!ApellidoPaterno), 2) & _
                                                    Left(UCase(oRsTmp!ApellidoMaterno), 2) & _
                                                    Left(UCase(oRsTmp!PrimerNombre), 2) & _
                                                    Format(oRsTmp!FechaNacimiento, "yyyymmdd")
                          lcSql = Left(UCase(oRsFox!pApellido), 2) & _
                                                    Left(UCase(oRsFox!sApellido), 2) & _
                                                    Left(UCase(oRsFox!Nombres), 2) & _
                                                    Format(oRsFox!fnac, "yyyymmdd")
                          s.Cells(oFila, 1).Value = Year(oRsFox1!FechaCreacion)
                          s.Cells(oFila, 3).Value = oRsFox1!NroHistoriaClinica
                          s.Cells(oFila, 2).Value = oRsTmp!NroDocumento
                          s.Cells(oFila, 4).Value = lcFichaFam
                          s.Cells(oFila, 5).Value = lcSql
                          s.Cells(oFila, 6).Value = IIf(lcFichaFam = lcSql, "", "Autog.Diferentes")
                          
                        
                          oFila = oFila + 1
                    End If
                    oRsFox1.Close
                  
               End If
               oRsTmp.Close
               oRsFox.MoveNext
           Loop
        End If
        
        '
        Me.MousePointer = 11
        oConexionFox.Close
        oConexion.Close
        
        EXL.Visible = True
        W.PrintPreview
        Set s = Nothing
        Set W = Nothing
        Set EXL = Nothing

        Set oRsTmp = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Set oConexionFox = Nothing
        Set oConexion = Nothing
        Set oDOhis_historicoAten = Nothing
        Set oHIS_historicoAten = Nothing
        Set oDOPaciente = Nothing
        Set oPacientes = Nothing
        Set oDOHistoriaClinica = Nothing
        Set oHistoriasClinicas = Nothing
        Set lcBuscaParametro = Nothing
        Set oDOPArametro = Nothing
        Set oParametros = Nothing
        Set oConexionSIS = Nothing
        Unload Me
        Exit Sub
ErrProAtHis:
        MsgBox Err.Description
Resume
        Set oRsTmp = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Set oConexionFox = Nothing
        Set oConexion = Nothing
        Set oDOhis_historicoAten = Nothing
        Set oHIS_historicoAten = Nothing
        Set oDOPaciente = Nothing
        Set oPacientes = Nothing
        Set oDOHistoriaClinica = Nothing
        Set oHistoriasClinicas = Nothing
        Set lcBuscaParametro = Nothing
        Set oConexionSIS = Nothing
        Unload Me

End Sub

Private Sub cmdKimbiri_Click()
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=HIS"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
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
       Dim lbContinuarProceso As Boolean
       Dim wFec1 As String, wFec2 As String, wFFecha As Date
       With wrs_LolCli
           wFFecha = CDate(Me.txtCuzcoF1.Text)
           wFec1 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           wFFecha = CDate(Me.txtCuzcoF2.Text)
           wFec2 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           
           Me.txtIni1.Text = "01/01/2011"
           Me.txtFin1.Text = "01/01/2011"
           wRsProblemas.Open "delete from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           'elimina historias GalenHos de esas Fechas
           lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
           lcSql = "select * from HistoriasClinicas where fechaCreacion Between (CONVERT(DATETIME,'" & Me.txtIni1.Text & " 00:00:00',103)) and (CONVERT(DATETIME,'" & Me.txtFin1.Text & " 23:59:59',103))"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar2.Min = 0
                ProgressBar2.Max = lntotReg
                lnRegAct = 0
                Do While Not wrs_GalenHos.EOF
                   lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                   wxConexionRed.BeginTrans
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   wrs_GalenHos.Delete
                   wrs_GalenHos.Update
                   wrs_GalenHos1.Open "delete from Pacientes where idpaciente=" & lnIdPaciente, wxConexionRed, adOpenKeyset, adLockOptimistic
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
                
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando11.Caption = "Insertando HC en GalenHos"
           lcSql = "select * from HistoriasClinicas"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcSql = "select * from Pacientes"
           wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcSql = "select * from historias "
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en JAMO"
              wxConexionJAMO.Close
              Exit Sub
           End If
           wRsProblemas.Open "select *  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
'If lnRegAct > 1000 Then
'   Exit Do
'End If
              ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
'              lbContinuarProceso = False
'              If IsNull(.Fields!FechaIngreso) Or .Fields!FechaIngreso = "" Then
'                 lbContinuarProceso = True
'              ElseIf CDate(.Fields!FechaIngreso) >= CDate(Me.txtIni1.Text) And CDate(.Fields!FechaIngreso) <= CDate(Me.txtFin1.Text) Then
'                 lbContinuarProceso = True
'              End If
              lbContinuarProceso = True
              If lbContinuarProceso = True Then
                  lbNuevoHC = False
If .Fields!HC = 260 Then
lbNuevoHC = False
End If
                  If Not sighentidades.EsFecha(Trim(.Fields!nacimiento), "DD/MM/AAAA") Then
                        lcFechaNac = "01/01/1990"
                  Else
                        lcFechaNac = .Fields!nacimiento
                  End If
                  lnNroHistoriaClinica = 0
                  lnNroHistoriaClinica = .Fields!HC
                  If lnNroHistoriaClinica = 0 Then
                     'Historia clinica con problemas
                     lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!HC)
    '                 lbNuevoHC = True
                  End If
                  If IsNull(.Fields!Pnombre) Then
                      lcPrimerNombre = "NN"
                      lcSegundoNombre = ""
                  Else
                      lcPrimerNombre = Trim(.Fields!Pnombre)
                      lcSegundoNombre = Trim(.Fields!sNombre)
                  End If
                  lnTipoSexo = IIf(Left(UCase(Trim(.Fields!M)), 1) = "X", 1, 2)
                  If lnNroHistoriaClinica = 0 Then
                     lbNuevoHC = True
                  Else
                        'Busca si ya existe Nro Historia en GalenHos
                        wrs_GalenHos2.Open "select * from Pacientes where nroHistoriaClinica=" & lnNroHistoriaClinica, wxConexionRed, adOpenKeyset, adLockOptimistic
                        If wrs_GalenHos2.RecordCount > 0 Then
                           lbNuevoHC = True
                        End If
                  End If
                  If lbNuevoHC = True Then
                          'HC repetida
                          'solo graba como problemas
                          wRsProblemas.AddNew
                          wRsProblemas.Fields!pacHis = .Fields!HC
                          wRsProblemas.Fields!pacPat = .Fields!Paterno
                          wRsProblemas.Fields!pacMat = .Fields!Materno
                          wRsProblemas.Fields!pacNam = .Fields!Pnombre
                          wRsProblemas.Fields!pacFin = CDate(Me.txtIni1.Text)
                          wRsProblemas.Fields!nroHistoriaGalenHos = lnNroHistoriaClinica
                          wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "HcRep-Kimbiri"
                          wRsProblemas.Update
                          '
                          If lnNroHistoriaClinica <> 0 Then
                            wRsProblemas.AddNew
                            wRsProblemas.Fields!pacHis = wrs_GalenHos2.Fields!NroHistoriaClinica
                            wRsProblemas.Fields!pacPat = wrs_GalenHos2.Fields!ApellidoPaterno
                            wRsProblemas.Fields!pacMat = wrs_GalenHos2.Fields!ApellidoMaterno
                            wRsProblemas.Fields!pacNam = Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)
                            wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"
                            wRsProblemas.Update
                            '
                            wrs_GalenHos2.Close
                          End If
                  Else
                          wrs_GalenHos2.Close
                          lcApellidoPaterno = UCase(Left(Trim(.Fields!Paterno), 20))
                          lcApellidoMaterno = UCase(Left(Trim(.Fields!Materno), 20))
                          lcPrimerNombre = UCase(lcPrimerNombre)
                          lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                          lcFechaAnt = Me.txtIni1.Text
                          'Busca en Tabla xx Equivalencia LolCli
                          lntipoOcupacion = 0
                          lnIdDepartamentoDomicilio = 0
                          LnIdProvinciaDomicilio = 0
                          lnIdDistritoDomicilio = 0
'                          If Not IsNull(.Fields!codgeo) Then
'                            lnIdDepartamentoDomicilio = Val(Left(.Fields!codgeo, 2))
'                            LnIdProvinciaDomicilio = Val(Left(.Fields!codgeo, 4))
'                            lnIdDistritoDomicilio = Val(.Fields!codgeo)
'                          End If
                          lnIdDepartamentoNacimiento = 0
                          LnIdProvinciaNacimiento = 0
                          lnIdDistritoNacimiento = 0
                          lnIdEstadoCivil = 0
'                          If Not IsNull(.Fields!estCivil) Then
'                             Select Case UCase(Left(.Fields!estCivil, 1))
'                             Case "S"   'soltero
'                                  lnIdEstadoCivil = 2
'                             Case "C"   'casado
'                                  lnIdEstadoCivil = 1
'                             End Select
'                          End If
                          'Graba Pacientes
                          wxConexionRed.BeginTrans
                          wrs_GalenHos1.AddNew
                          wrs_GalenHos1.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                          wrs_GalenHos1.Fields!ApellidoPaterno = lcApellidoPaterno
                          wrs_GalenHos1.Fields!ApellidoMaterno = lcApellidoMaterno
                          wrs_GalenHos1.Fields!PrimerNombre = lcPrimerNombre
                          wrs_GalenHos1.Fields!SegundoNombre = lcSegundoNombre
                          wrs_GalenHos1.Fields!idTipoSexo = lnTipoSexo
                          wrs_GalenHos1.Fields!FechaNacimiento = CDate(lcFechaNac)
                          wrs_GalenHos1.Fields!IdTipoNumeracion = lnTipoNumeracion
                          wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado
                          wrs_GalenHos1.Fields!IdDistritoDomicilio = lnIdDistritoDomicilio
                          If lnIdEstadoCivil > 0 Then
                             wrs_GalenHos1.Fields!IdEstadoCivil = lnIdEstadoCivil
                          End If
                          If Not IsNull(.Fields!domicilio) Then
                             wrs_GalenHos1.Fields!DireccionDomicilio = Left(.Fields!domicilio, 50)
                          End If
'                          If Not IsNull(.Fields!padreNomb) Then
'                             wrs_GalenHos1.Fields!NombrePadre = Left(.Fields!padreNomb, 20)
'                          End If
'                          If Not IsNull(.Fields!madreNomb) Then
'                             wrs_GalenHos1.Fields!NombreMadre = Left(.Fields!madreNomb, 20)
'                          End If
'                          If Not IsNull(.Fields!le) Then
'                             If Len(Trim(.Fields!le)) = 8 Then
'                                wrs_GalenHos1.Fields!NroDocumento = Left(.Fields!le, 8)
'                                wrs_GalenHos1.Fields!IdDocIdentidad = 1
'                             End If
'                          End If
                          'wrs_GalenHos1.Fields!IdEtnia = "80"
                          wrs_GalenHos1.Fields!IdPaisDomicilio = 166
                          wrs_GalenHos1.Fields!IdPaisProcedencia = 166
                          wrs_GalenHos1.Fields!IdPaisNacimiento = 166
                          wrs_GalenHos1.Update
                          lnIdPaciente = wrs_GalenHos1.Fields!idPaciente
                          'Graba HistoriasClinicas
                          wrs_GalenHos.AddNew
                          wrs_GalenHos.Fields!idPaciente = lnIdPaciente
                          wrs_GalenHos.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                          wrs_GalenHos.Fields!FechaCreacion = lcFechaAnt
                          wrs_GalenHos.Fields!IdTipoNumeracion = lnTipoNumeracion
                          wrs_GalenHos.Fields!IdEstadoHistoria = 1
                          wrs_GalenHos.Fields!IdTipoHistoria = 1
                          wrs_GalenHos.Fields!HistoriaSistemaAnterior = .Fields!HC
                          wrs_GalenHos.Update
                          wxConexionRed.CommitTrans
                  End If
              End If
              .MoveNext
           Loop
       End With
    End If
    wxConexionJAMO.Close
    Unload Me
    Exit Sub
err_proceso:
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!HC & "     Paciente:" & wrs_LolCli.Fields!Paterno & " " & wrs_LolCli.Fields!Materno & " " & wrs_LolCli.Fields!Pnombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    Resume
    Unload Me

End Sub

Private Sub cmdMigraAtencionesJamo_Click()
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       Dim lcSql As String
       Dim lnCant As Long
       Dim lntotReg As Long
       Dim oRsTmpNew1 As New Recordset
       Dim oRsTmpJamo1 As New Recordset
       Dim oRsTmpJamo2 As New Recordset
       Dim oRsTmpJamo3 As New Recordset
       Dim lcBuscaParametro As New SIGHDatos.Parametros
       Dim oConexionBDnew As New Connection
       Dim lcHistoriaProcesada As String
       '
       lcSql = lcBuscaParametro.SeleccionaFilaParametro(273)
       oConexionBDnew.Open lcSql
       oConexionBDnew.CursorLocation = adUseClient
       oConexionBDnew.CommandTimeout = 150
       
       '
       wxConexionJAMO.Open "dsn=Jamo"
       wxConexionJAMO.CursorLocation = adUseClient
       wxConexionJAMO.CommandTimeout = 150
       '
       Dim wrs_GalenHos As New ADODB.Recordset
       'Servidor Nuevo
       lcSql = "delete from atencionesCE where CitaFecha Between (CONVERT(DATETIME,'" & Me.txtFCita1.Text & " 00:00:01',103)) and (CONVERT(DATETIME,'" & Me.txtFCita2.Text & " 23:59:59',103))"
       oRsTmpNew1.Open lcSql, oConexionBDnew, adOpenKeyset, adLockOptimistic
       lcSql = "select * from atencionesCE"
       oRsTmpNew1.Open lcSql, oConexionBDnew, adOpenKeyset, adLockOptimistic
       'Filtro de Todas las CITAS
       lcSql = "SELECT     dbo.Citas.IdCitas, dbo.Citas.IdServicio, dbo.Citas.FechadeCita, dbo.Medico.Nombre AS dMedico, dbo.Servicio.Servicio, dbo.DetalleHc.NumHc, " & _
                "                      dbo.DetalleHc.Nombre , dbo.DetalleHc.ApellidoP, dbo.DetalleHc.ApellidoM,dbo.Citas.Dni" & _
                " FROM         dbo.Citas LEFT OUTER JOIN" & _
                "                      dbo.DetalleHc ON dbo.Citas.NumHc = dbo.DetalleHc.NumHc LEFT OUTER JOIN" & _
                "                      dbo.Medico ON dbo.Citas.Dni = dbo.Medico.Dni LEFT OUTER JOIN" & _
                "                      dbo.Servicio ON dbo.Medico.IdServicio = dbo.Servicio.IdServicio"
       oRsTmpJamo1.Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
'oRsTmpJamo1.Filter = "NumHc='0100033'"
       lntotReg = oRsTmpJamo1.RecordCount
       If lntotReg > 0 Then
             ProgressBar3.Min = 0
             ProgressBar3.Max = lntotReg
             lnCant = 0
             oRsTmpJamo1.MoveFirst
             Do While Not oRsTmpJamo1.EOF
                DoEvents
                ProgressBar3.Value = lnCant
                Me.Refresh
                lnCant = lnCant + 1
                If CDate(oRsTmpJamo1.Fields!fechaDeCita) >= CDate(Me.txtFCita1.Text) And CDate(oRsTmpJamo1.Fields!fechaDeCita) <= CDate(Me.txtFCita2.Text) And Trim(oRsTmpJamo1.Fields!numHc) <> "" Then
                    lcHistoriaProcesada = oRsTmpJamo1.Fields!numHc
                    'Filtro para Triaje
                    lcSql = "SELECT     IdCitas, NumHc, FechaTriaje, PresionA, Talla, Temperatura, Peso, Edad, HoraTriaje" & _
                         " From dbo.Triaje" & _
                         " WHERE     IdCitas = '" & oRsTmpJamo1.Fields!idCitas & "' and NumHc='" & oRsTmpJamo1.Fields!numHc & "'"
                    oRsTmpJamo2.Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
                    'Filtro para buscar en "bd JAMO" las 'atenciones registradas' de las Citas
                    lcSql = "SELECT     dbo.Diagnostico.Fecha, dbo.Medico.Nombre, dbo.Servicio.Servicio, dbo.Diagnostico.NumHc, dbo.Diagnostico.Motivo, " & _
                             "                      dbo.Diagnostico.Anamnesis AS ExamenClinico, dbo.Diagnostico.DiagMed, dbo.Diagnostico.Tratamiento, dbo.Diagnostico.ExClinicos," & _
                             "                      dbo.Diagnostico.Observaciones , dbo.Diagnostico.idServicio, dbo.Diagnostico.DNI" & _
                             " FROM         dbo.Diagnostico LEFT OUTER JOIN" & _
                             "                      dbo.Servicio ON dbo.Diagnostico.IdServicio = dbo.Servicio.IdServicio LEFT OUTER JOIN" & _
                             "                      dbo.Medico ON dbo.Diagnostico.Dni = dbo.Medico.Dni" & _
                             " WHERE     dbo.Diagnostico.Fecha='" & oRsTmpJamo1.Fields!fechaDeCita & "' and  dbo.Diagnostico.DNI='" & oRsTmpJamo1.Fields!DNI & "' and dbo.Diagnostico.NumHc='" & oRsTmpJamo1.Fields!numHc & "'"
                    oRsTmpJamo3.Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
                    '
                    oRsTmpNew1.AddNew
                    oRsTmpNew1.Fields!NroHistoriaClinica = oRsTmpJamo1.Fields!numHc
                    oRsTmpNew1.Fields!CitaDniMedicoJamo = oRsTmpJamo1.Fields!DNI
                    oRsTmpNew1.Fields!CitaFecha = oRsTmpJamo1.Fields!fechaDeCita
                    oRsTmpNew1.Fields!CitaMedico = Left(oRsTmpJamo1.Fields!dMedico, 100)
                    If oRsTmpJamo3.RecordCount > 0 Then
                        oRsTmpNew1.Fields!CitaIdServicio = oRsTmpJamo3.Fields!IdServicio
                        oRsTmpNew1.Fields!CitaServicioJamo = oRsTmpJamo3.Fields!Servicio
                        oRsTmpNew1.Fields!CitaMotivo = oRsTmpJamo3.Fields!motivo
                        oRsTmpNew1.Fields!CitaExamenClinico = oRsTmpJamo3.Fields!ExamenClinico
                        oRsTmpNew1.Fields!CitaDiagMed = oRsTmpJamo3.Fields!DiagMed
                        oRsTmpNew1.Fields!CitaTratamiento = oRsTmpJamo3.Fields!Tratamiento
                        oRsTmpNew1.Fields!CitaExClinicos = oRsTmpJamo3.Fields!ExClinicos
                        oRsTmpNew1.Fields!CitaObservaciones = oRsTmpJamo3.Fields!Observaciones
                    Else
                        oRsTmpNew1.Fields!CitaServicioJamo = Left(oRsTmpJamo1.Fields!Servicio, 100)
                        oRsTmpNew1.Fields!CitaIdServicio = oRsTmpJamo1.Fields!IdServicio
                    End If
                    If oRsTmpJamo2.RecordCount > 0 Then
                        oRsTmpNew1.Fields!TriajeEdad = Left(oRsTmpJamo2.Fields!Edad, 6)
                        oRsTmpNew1.Fields!TriajeFecha = CDate(oRsTmpJamo2.Fields!FechaTriaje & " " & oRsTmpJamo2.Fields!HoraTriaje)
                        oRsTmpNew1.Fields!TriajePresion = Left(oRsTmpJamo2.Fields!presionA, 13)
                        oRsTmpNew1.Fields!TriajeTalla = Left(oRsTmpJamo2.Fields!Talla, 7)
                        oRsTmpNew1.Fields!TriajeTemperatura = Left(oRsTmpJamo2.Fields!Temperatura, 6)
                        oRsTmpNew1.Fields!TriajePeso = Left(oRsTmpJamo2.Fields!Peso, 7)
                    End If
                    oRsTmpNew1.Update
                    '
                    oRsTmpJamo2.Close
                    oRsTmpJamo3.Close
                End If
                oRsTmpJamo1.MoveNext
             Loop
          
       End If
       oRsTmpJamo1.Close
       wxConexionJAMO.Close
       oConexionBDnew.Close
       Unload Me
       Exit Sub
    End If
err_proceso:
    MsgBox Err.Description & Chr(13) & "Historia con problemas JAMO: " & lcHistoriaProcesada
    Resume
End Sub

Private Sub cmdMigraHRC_Click()
    lcMensajeError = "todavía no se conecta a la BD dbHospi"
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=dbhospi"
       lcMensajeError = "Ya se conecto a la BD dbhospi por ODBC"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
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
       Dim lbContinuarProceso As Boolean
       Dim wFec1 As String, wFec2 As String, wFFecha As Date
       Dim lnTipoNumeracion1 As Long, ldFechaO As Date
       Dim lnFechasOk As Long, lnFechasNoOk As Long, lnNulos As Long, PacienteOK As Long, PacienteNoOK As Long
       lnTipoNumeracion1 = 1
       lnFechasOk = 0: lnFechasNoOk = 0: lnNulos = 0: PacienteOK = 0: PacienteNoOK = 0
       With wrs_LolCli
           wFFecha = CDate(Me.Text1.Text)
           wFec1 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           wFFecha = CDate(Me.Text2.Text)
           wFec2 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           
           
           Me.txtIni1.Text = wFec1
           Me.txtFin1.Text = wFec2
           'wRsProblemas.Open "delete from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           'elimina historias GalenHos de esas Fechas
           If chkkPacientes1.Value = 0 Then
                lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
                lcMensajeError = "antes de HistoriasClinicas"
                'lcSql = "select * from HistoriasClinicas where fechaCreacion >=" & Me.txtIni1.Text & " and fechacreacion<=" & Me.txtFin1.Text
                lcSql = "select * from HistoriasClinicas"
                wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcMensajeError = "Cargo Historias en GalenHos para eliminarlas y volver a Procesar"
                lntotReg = wrs_GalenHos.RecordCount
                If lntotReg > 0 Then
                     wrs_GalenHos.MoveFirst
                     ProgressBar2.Min = 0
                     ProgressBar2.Max = lntotReg
                     lnRegAct = 0
                     Do While Not wrs_GalenHos.EOF
                        lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                        ldFechaO = CDate(Format(wrs_GalenHos.Fields!FechaCreacion, "dd/mm/yyyy"))
                        If ldFechaO >= CDate(Me.Text1.Text) And ldFechaO <= CDate(Me.Text2.Text) Then
                             wxConexionRed.BeginTrans
                             lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                             wrs_GalenHos.Delete
                             wrs_GalenHos.Update
                             wrs_GalenHos1.Open "delete from Pacientes where idpaciente=" & lnIdPaciente, wxConexionRed, adOpenKeyset, adLockOptimistic
                             wxConexionRed.CommitTrans
                        End If
                        wrs_GalenHos.MoveNext
                     Loop
                     
                End If
                wrs_GalenHos.Close
           End If
           If chkActualizaFechaREg.Value = 1 Then
              lcSql = "update paciente set fec_reg=today() where fec_reg is null"
              .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           End If
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando11.Caption = "Insertando HC en GalenHos"
           lcSql = "select * from HistoriasClinicas"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcSql = "select * from Pacientes"
           wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcMensajeError = "Abre las tablas Pacientes,Historias para agregar datos"
           If chkkPacientes1.Value = 0 Then
              lcSql = "select * from Paciente "
           Else
              lcSql = "select * from Paciente2 "
           End If
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lcMensajeError = "Lee la tabla Paciente de DBHOSPI para barrer c/u"
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en CS Nazareno"
              wxConexionJAMO.Close
              Unload Me
              Exit Sub
           End If
           lcMensajeError = "Lee la tabla LolcliProblemasHC"
           wRsProblemas.Open "select *  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
               lcMensajeError = "Empieza a barrer uno por uno - Antes de Fecha"
               ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
               If Not IsNull(.Fields!fec_reg) Then
               If sighentidades.EsFecha(.Fields!fec_reg, "DD/MM/AAAA") Then
               ldFechaO = CDate(Format(.Fields!fec_reg, "dd/mm/yyyy"))
               lcMensajeError = "Fecha  OK"
               If ldFechaO >= CDate(Me.Text1.Text) And ldFechaO <= CDate(Me.Text2.Text) Then
                  lcMensajeError = "Empieza a barrer uno por uno - Paso Fecha"
                  lbContinuarProceso = True
                  If IsNull(.Fields!ap_pac) Or Trim(.Fields!ap_pac) = "" Or IsNull(.Fields!am_pac) Or Trim(.Fields!am_pac) = "" Or IsNull(.Fields!hc_paciente) Or Trim(.Fields!hc_paciente) = "" Then
                     lbContinuarProceso = False
                  End If
    '              ElseIf CDate(.Fields!FechaIngreso) >= CDate(Me.txtIni1.Text) And CDate(.Fields!FechaIngreso) <= CDate(Me.txtFin1.Text) Then
    '                 lbContinuarProceso = True
    '              End If
                  If lbContinuarProceso = True Then
                      lbNuevoHC = False
    If Val(.Fields!hc_paciente) = 41403 Then
    lbNuevoHC = False
    End If
                      lcMensajeError = "Paso continuarProceso"
                      If IsNull(.Fields!fec_nac_pac) Then
                            lcFechaNac = "01/01/1990"
                      ElseIf (Not sighentidades.EsFecha(Trim(.Fields!fec_nac_pac), "DD/MM/AAAA")) Then
                            lcFechaNac = "01/01/1990"
                      Else
                            lcFechaNac = .Fields!fec_nac_pac
                      End If
                      lcMensajeError = "Paso fecha de Nacimiento"
                      lnNroHistoriaClinica = 0
                      If UCase(Left(.Fields!hc_paciente, 2)) = "T-" Then
                         lnNroHistoriaClinica = "999" + Trim(Str(Val(Mid(.Fields!hc_paciente, 3, 100))))
                      Else
                         lnNroHistoriaClinica = Val(.Fields!hc_paciente)
                      End If
                      If lnNroHistoriaClinica = 0 Then
                         'Historia clinica con problemas
                         lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!hc_paciente)
        '                 lbNuevoHC = True
                      End If
                      If IsNull(.Fields!nom_pac) Then
                          lcPrimerNombre = "NN"
                          lcSegundoNombre = ""
                      Else
                          lcPrimerNombre = Trim(.Fields!nom_pac)
                          lcPrimerNombre = Left(RetornaPrimerNombre(lcPrimerNombre), 20)
                          lcSegundoNombre = Trim(.Fields!nom_pac)
                          lcSegundoNombre = Left(RetornaSegundoNombre(lcSegundoNombre), 20)
                      End If
                      
                      lnTipoSexo = IIf(UCase(.Fields!sexo_pac) = "M", 1, 2)
                      If lnNroHistoriaClinica = 0 Then
                         lbNuevoHC = True
                      Else
                            'Busca si ya existe Nro Historia en GalenHos
                            wrs_GalenHos2.Open "select * from Pacientes where nroHistoriaClinica=" & lnNroHistoriaClinica, wxConexionRed, adOpenKeyset, adLockOptimistic
                            If wrs_GalenHos2.RecordCount > 0 Then
                               lbNuevoHC = True
                            End If
                      End If
                      If lbNuevoHC = True Then
                              lcMensajeError = "Historia con problemas"
                              'HC repetida
                              'solo graba como problemas
                              wRsProblemas.AddNew
                              wRsProblemas.Fields!pacHis = .Fields!hc_paciente
                              wRsProblemas.Fields!pacPat = .Fields!ap_pac
                              wRsProblemas.Fields!pacMat = .Fields!am_pac
                              wRsProblemas.Fields!pacNam = .Fields!nom_pac
                              wRsProblemas.Fields!pacFin = CDate(Me.txtFIniCSN.Text)
                              wRsProblemas.Fields!nroHistoriaGalenHos = lnNroHistoriaClinica
                              wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "HcRep-HRC"
                              wRsProblemas.Update
                              '
                              If lnNroHistoriaClinica <> 0 Then
                                wRsProblemas.AddNew
                                wRsProblemas.Fields!pacHis = wrs_GalenHos2.Fields!NroHistoriaClinica
                                wRsProblemas.Fields!pacPat = wrs_GalenHos2.Fields!ApellidoPaterno
                                wRsProblemas.Fields!pacMat = wrs_GalenHos2.Fields!ApellidoMaterno
                                wRsProblemas.Fields!pacNam = Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)
                                wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"
                                wRsProblemas.Update
                                '
                                wrs_GalenHos2.Close
                              End If
                              PacienteNoOK = PacienteNoOK + 1
                      Else
                              lcMensajeError = "Historia OK"
                              wrs_GalenHos2.Close
                              lcApellidoPaterno = UCase(Left(Trim(.Fields!ap_pac), 20))
                              lcApellidoMaterno = UCase(Left(Trim(.Fields!am_pac), 20))
                              lcPrimerNombre = UCase(lcPrimerNombre)
                              lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                              lcFechaAnt = .Fields!fec_reg
                              'Busca en Tabla xx Equivalencia LolCli
                              lntipoOcupacion = 0
                              lnIdDepartamentoDomicilio = 0
                              LnIdProvinciaDomicilio = 0
                              lnIdDistritoDomicilio = 0
    '                          If Not IsNull(.Fields!codgeo) Then
    '                            lnIdDepartamentoDomicilio = Val(Left(.Fields!codgeo, 2))
    '                            LnIdProvinciaDomicilio = Val(Left(.Fields!codgeo, 4))
    '                            lnIdDistritoDomicilio = Val(.Fields!codgeo)
    '                          End If
                              lnIdDepartamentoNacimiento = 0
                              LnIdProvinciaNacimiento = 0
                              lnIdDistritoNacimiento = 0
                              lnIdEstadoCivil = 0
    '                          If Not IsNull(.Fields!estCivil) Then
    '                             Select Case UCase(Left(.Fields!estCivil, 1))
    '                             Case "S"   'soltero
    '                                  lnIdEstadoCivil = 2
    '                             Case "C"   'casado
    '                                  lnIdEstadoCivil = 1
    '                             End Select
    '                          End If
                              'Graba Pacientes
                              wxConexionRed.BeginTrans
                              wrs_GalenHos1.AddNew
                              wrs_GalenHos1.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                              wrs_GalenHos1.Fields!ApellidoPaterno = lcApellidoPaterno
                              wrs_GalenHos1.Fields!ApellidoMaterno = lcApellidoMaterno
                              wrs_GalenHos1.Fields!PrimerNombre = lcPrimerNombre
                              wrs_GalenHos1.Fields!SegundoNombre = lcSegundoNombre
                              wrs_GalenHos1.Fields!idTipoSexo = lnTipoSexo
                              wrs_GalenHos1.Fields!FechaNacimiento = CDate(lcFechaNac)
                              wrs_GalenHos1.Fields!IdTipoNumeracion = lnTipoNumeracion1
                              wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado
                              wrs_GalenHos1.Fields!IdDistritoDomicilio = lnIdDistritoDomicilio
    '                          If lnIdEstadoCivil > 0 Then
    '                             wrs_GalenHos1.Fields!IdEstadoCivil = lnIdEstadoCivil
    '                          End If
                              If Not IsNull(.Fields!domicilio) Then
                                 wrs_GalenHos1.Fields!DireccionDomicilio = Left(.Fields!domicilio, 50)
                              End If
                              If Not IsNull(.Fields!nom_padre) Then
                                 wrs_GalenHos1.Fields!NombrePadre = Left(.Fields!nom_padre, 20)
                              End If
                              If Not IsNull(.Fields!nom_madre) Then
                                 wrs_GalenHos1.Fields!NombreMadre = Left(.Fields!nom_madre, 20)
                              End If
                              If Not IsNull(.Fields!DNI) Then
                                 If Len(Trim(.Fields!DNI)) = 8 Then
                                    wrs_GalenHos1.Fields!NroDocumento = Left(.Fields!DNI, 8)
                                    wrs_GalenHos1.Fields!IdDocIdentidad = 1
                                 End If
                              End If
                              'wrs_GalenHos1.Fields!IdEtnia = "80"
                              wrs_GalenHos1.Fields!IdPaisDomicilio = 166
                              wrs_GalenHos1.Fields!IdPaisProcedencia = 166
                              wrs_GalenHos1.Fields!IdPaisNacimiento = 166
                              wrs_GalenHos1.Update
                              lnIdPaciente = wrs_GalenHos1.Fields!idPaciente
                              'Graba HistoriasClinicas
                              wrs_GalenHos.AddNew
                              wrs_GalenHos.Fields!idPaciente = lnIdPaciente
                              wrs_GalenHos.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                              wrs_GalenHos.Fields!FechaCreacion = lcFechaAnt
                              wrs_GalenHos.Fields!IdTipoNumeracion = lnTipoNumeracion1
                              wrs_GalenHos.Fields!IdEstadoHistoria = 1
                              wrs_GalenHos.Fields!IdTipoHistoria = 1
                              wrs_GalenHos.Fields!HistoriaSistemaAnterior = .Fields!hc_paciente
                              wrs_GalenHos.Update
                              wxConexionRed.CommitTrans
                              PacienteOK = PacienteOK + 1
                      End If
                  End If
               Else
                  lnFechasOk = lnFechasOk + 1
               End If
               Else
                    lnFechasNoOk = lnFechasNoOk + 1
               End If
               Else
                    lnNulos = lnNulos + 1
               End If
              .MoveNext
           Loop
       End With
       wxConexionJAMO.Close
       On Error Resume Next
       wrs_GalenHos2.Close
    End If
    MsgBox "Fechas OK, pero fuera del rango: " & Chr(13) & lnFechasOk & "       Fechas con datos (pero no son fechas): " & Chr(13) & lnFechasNoOk & "Fechas NULAS: " & lnNulos
    
    wrs_GalenHos2.Open "select top 1 nroHistoriaClinica from Pacientes order by nroHistoriaClinica desc", wxConexionRed, adOpenKeyset, adLockOptimistic
    lnNroHistoriaClinica = wrs_GalenHos2.Fields!NroHistoriaClinica
    wrs_GalenHos2.Close
    wrs_GalenHos2.Open "update generadorNroHistoriaClinica set nroHistoriaClinica=" & lnNroHistoriaClinica & " where idNumerador=17", wxConexionRed, adOpenKeyset, adLockOptimistic
    Unload Me
    Exit Sub
err_proceso:
    If MsgBox(Err.Description & Chr(13) & lcMensajeError & Chr(13) & "Vuelve a probar?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Resume
    End If
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!HC & "     Paciente:" & wrs_LolCli.Fields!Paterno & " " & wrs_LolCli.Fields!Materno & " " & wrs_LolCli.Fields!Pnombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    Unload Me
End Sub

Private Sub cmdMigraJamo_Click()
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=Jamo"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
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
       Dim lbContinuarProceso As Boolean
       With wrs_LolCli
           If chkLimpiaLolcli.Value = 1 Then
               wRsProblemas.Open "delete from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           End If
           'elimina historias GalenHos de esas Fechas
           lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
           lcSql = "select * from HistoriasClinicas where fechaCreacion Between (CONVERT(DATETIME,'" & Me.txtIni1.Text & " 00:00:00',103)) and (CONVERT(DATETIME,'" & Me.txtFin1.Text & " 23:59:59',103))"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar2.Min = 0
                ProgressBar2.Max = lntotReg
                lnRegAct = 0
                Do While Not wrs_GalenHos.EOF
                   lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                   wxConexionRed.BeginTrans
If wrs_GalenHos.Fields!idPaciente = 613609 Then
lnIdPaciente = 0
End If
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   wrs_GalenHos.Delete
                   wrs_GalenHos.Update
                   
                   lcSql = "delete from SunasaPacientesHistoricos where idpaciente=" & lnIdPaciente
                   wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                   
                   lcSql = "delete from Pacientes where idpaciente=" & lnIdPaciente
                   wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
                
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando1.Caption = "Insertando HC en GalenHos"
           lcSql = "select * from HistoriasClinicas"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcSql = "select * from Pacientes"
           wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
            lcSql = "select * from DetalleHC " & _
                     " order by fechaIngreso"
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en JAMO"
              wxConexionJAMO.Close
              Exit Sub
           End If
           wRsProblemas.Open "select *  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
              ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
              lbContinuarProceso = False
              If sighentidades.EsFecha(.Fields!FechaIngreso, "DD/MM/AAAA") = False Then
                 lbContinuarProceso = True
              ElseIf CDate(.Fields!FechaIngreso) >= CDate(Me.txtIni1.Text) And CDate(.Fields!FechaIngreso) <= CDate(Me.txtFin1.Text) Then
                 lbContinuarProceso = True
              End If
              If lbContinuarProceso = True Then
                  lbNuevoHC = False
                  
                  If sighentidades.EsFecha(.Fields!FechaNac, "DD/MM/AAAA") Then
                     lcFechaNac = .Fields!FechaNac
                  Else
                     lcFechaNac = "01/01/1990"
                  End If
                  lnNroHistoriaClinica = 0
                  lnNroHistoriaClinica = Val(.Fields!numHc)
                  If lnNroHistoriaClinica = 0 Then
                     'Historia clinica con problemas
                     lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!numHc)
    '                 lbNuevoHC = True
                  End If
                  If IsNull(.Fields!nombre) Then
                      lcPrimerNombre = "NN"
                      lcSegundoNombre = ""
                  Else
                      lcPrimerNombre = Left(RetornaPrimerNombre(.Fields!nombre), 20)
                      lcSegundoNombre = Left(RetornaSegundoNombre(.Fields!nombre), 20)
                  End If
                  lnTipoSexo = IIf(Left(UCase(.Fields!sexo), 1) = "F", 2, 1)
                  'Busca si ya existe Nro Historia en GalenHos
                  wrs_GalenHos2.Open "select * from Pacientes where nroHistoriaClinica=" & lnNroHistoriaClinica, wxConexionRed, adOpenKeyset, adLockOptimistic
                  If wrs_GalenHos2.RecordCount > 0 Then
                     lbNuevoHC = True
                  End If
                  If lbNuevoHC = True Then
                          'HC repetida
                          'solo graba como problemas
                          wRsProblemas.AddNew
                          wRsProblemas.Fields!pacHis = .Fields!numHc
                          wRsProblemas.Fields!pacPat = .Fields!ApellidoP
                          wRsProblemas.Fields!pacMat = .Fields!ApellidoM
                          wRsProblemas.Fields!pacNam = .Fields!nombre
                          wRsProblemas.Fields!pacFin = .Fields!FechaIngreso
                          wRsProblemas.Fields!nroHistoriaGalenHos = lnNroHistoriaClinica
                          wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "HcRep-Jamo"
                          wRsProblemas.Update
                          '
                          wRsProblemas.AddNew
                          wRsProblemas.Fields!pacHis = wrs_GalenHos2.Fields!NroHistoriaClinica
                          wRsProblemas.Fields!pacPat = wrs_GalenHos2.Fields!ApellidoPaterno
                          wRsProblemas.Fields!pacMat = wrs_GalenHos2.Fields!ApellidoMaterno
                          wRsProblemas.Fields!pacNam = Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)
                          wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"
                          wRsProblemas.Update
                          '
                          wrs_GalenHos2.Close
                  Else
                          wrs_GalenHos2.Close
                          If lbNuevoHC Then
                             lnNroHistoriaClinica = generaNuevaNroHistoria(lnNroHistoriaClinica)
                             lcSegundoNombre = "." & LCase(Trim(lcSegundoNombre))
                          End If
                          lcApellidoPaterno = Left(.Fields!ApellidoP, 20)
                          lcApellidoMaterno = Left(.Fields!ApellidoM, 20)
                          lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                          lcFechaAnt = IIf(sighentidades.EsFecha(.Fields!FechaIngreso, "DD/MM/AAAA") = False, "19/04/2005", .Fields!FechaIngreso)
                          'Busca en Tabla xx Equivalencia LolCli
                          lntipoOcupacion = 0
                          lnIdDepartamentoDomicilio = 0
                          LnIdProvinciaDomicilio = 0
                          lnIdDistritoDomicilio = 0
                          lnIdDepartamentoNacimiento = 0
                          LnIdProvinciaNacimiento = 0
                          lnIdDistritoNacimiento = 0
                          lnIdEstadoCivil = 0
                          If Not IsNull(.Fields!estadoCivil) Then
                             Select Case UCase(Left(.Fields!estadoCivil, 1))
                             Case "S"   'soltero
                                  lnIdEstadoCivil = 2
                             Case "C"   'casado
                                  lnIdEstadoCivil = 1
                             Case "O"   'otros
                                  lnIdEstadoCivil = 9
                             End Select
                          End If
                          'Graba Pacientes
                          wxConexionRed.BeginTrans
                          wrs_GalenHos1.AddNew
                          wrs_GalenHos1.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                          wrs_GalenHos1.Fields!ApellidoPaterno = lcApellidoPaterno
                          wrs_GalenHos1.Fields!ApellidoMaterno = lcApellidoMaterno
                          wrs_GalenHos1.Fields!PrimerNombre = lcPrimerNombre
                          wrs_GalenHos1.Fields!SegundoNombre = lcSegundoNombre
                          wrs_GalenHos1.Fields!idTipoSexo = lnTipoSexo
                          wrs_GalenHos1.Fields!FechaNacimiento = CDate(lcFechaNac)
                          wrs_GalenHos1.Fields!IdTipoNumeracion = lnTipoNumeracion
                          wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado
                          If lnIdEstadoCivil > 0 Then
                             wrs_GalenHos1.Fields!IdEstadoCivil = lnIdEstadoCivil
                          End If
                          If Not IsNull(.Fields!Direccion) Then
                             wrs_GalenHos1.Fields!DireccionDomicilio = Left(.Fields!Direccion, 50)
                          End If
                          'wrs_GalenHos1.Fields!IdEtnia = "80"
                          wrs_GalenHos1.Fields!IdPaisDomicilio = 166
                          wrs_GalenHos1.Fields!IdPaisProcedencia = 166
                          wrs_GalenHos1.Fields!IdPaisNacimiento = 166
                          wrs_GalenHos1.Update
                          lnIdPaciente = wrs_GalenHos1.Fields!idPaciente
                          'Graba HistoriasClinicas
                          wrs_GalenHos.AddNew
                          wrs_GalenHos.Fields!idPaciente = lnIdPaciente
                          wrs_GalenHos.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                          wrs_GalenHos.Fields!FechaCreacion = lcFechaAnt
                          wrs_GalenHos.Fields!IdTipoNumeracion = lnTipoNumeracion
                          wrs_GalenHos.Fields!IdEstadoHistoria = 1
                          wrs_GalenHos.Fields!IdTipoHistoria = 1
                          wrs_GalenHos.Fields!HistoriaSistemaAnterior = .Fields!numHc
                          wrs_GalenHos.Update
                          wxConexionRed.CommitTrans
                  End If
              End If
              .MoveNext
           Loop
       End With
    End If
    wxConexionJAMO.Close
    Unload Me
    Exit Sub
err_proceso:
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!numHc & "     F.Registro: " & wrs_LolCli.Fields!FechaIngreso & "     Paciente:" & wrs_LolCli.Fields!ApellidoP & " " & wrs_LolCli.Fields!ApellidoM & " " & wrs_LolCli.Fields!nombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    Resume
    Unload Me
End Sub

Private Sub cmdMigraSicuani_Click()
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=HIS"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
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
       Dim lbContinuarProceso As Boolean
       Dim wFec1 As String, wFec2 As String, wFFecha As Date
       With wrs_LolCli
           wFFecha = CDate(Me.txtCuzcoF1.Text)
           wFec1 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           wFFecha = CDate(Me.txtCuzcoF2.Text)
           wFec2 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
       
           If chkLimpiaLolcli1.Value = 1 Then
               wRsProblemas.Open "delete from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           End If
           'elimina historias GalenHos de esas Fechas
           lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
           lcSql = "select * from HistoriasClinicas where fechaCreacion Between (CONVERT(DATETIME,'" & Me.txtIni1.Text & " 00:00:00',103)) and (CONVERT(DATETIME,'" & Me.txtFin1.Text & " 23:59:59',103))"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar2.Min = 0
                ProgressBar2.Max = lntotReg
                lnRegAct = 0
                Do While Not wrs_GalenHos.EOF
                   lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                   wxConexionRed.BeginTrans
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   wrs_GalenHos.Delete
                   wrs_GalenHos.Update
                   wrs_GalenHos1.Open "delete from Pacientes where idpaciente=" & lnIdPaciente, wxConexionRed, adOpenKeyset, adLockOptimistic
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
                
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando11.Caption = "Insertando HC en GalenHos"
           lcSql = "select * from HistoriasClinicas"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcSql = "select * from Pacientes"
           wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcSql = "select * from histClinic " & _
                        "  where  f_admis>=" & wFec1 & " and f_admis<=" & wFec2 & _
                         " order by f_admis"
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en JAMO"
              wxConexionJAMO.Close
              Exit Sub
           End If
           wRsProblemas.Open "select *  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
'If lnRegAct > 1000 Then
'   Exit Do
'End If
              ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
'              lbContinuarProceso = False
'              If IsNull(.Fields!FechaIngreso) Or .Fields!FechaIngreso = "" Then
'                 lbContinuarProceso = True
'              ElseIf CDate(.Fields!FechaIngreso) >= CDate(Me.txtIni1.Text) And CDate(.Fields!FechaIngreso) <= CDate(Me.txtFin1.Text) Then
'                 lbContinuarProceso = True
'              End If
              lbContinuarProceso = True
              If lbContinuarProceso = True Then
                  lbNuevoHC = False
                  
                  If Not IsNull(.Fields!f_nacim) Then
                     If .Fields!f_nacim < 1900 Then
                        lcFechaNac = "01/01/1900"
                     Else
                        lcFechaNac = .Fields!f_nacim
                     End If
                  Else
                      
                     lcFechaNac = "01/01/1990"
                  End If
                  lnNroHistoriaClinica = 0
                  lnNroHistoriaClinica = Val(.Fields!Codigo)
                  If lnNroHistoriaClinica = 0 Then
                     'Historia clinica con problemas
                     lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!Codigo)
    '                 lbNuevoHC = True
                  End If
                  If IsNull(.Fields!Nombres) Then
                      lcPrimerNombre = "NN"
                      lcSegundoNombre = ""
                  Else
                      lcPrimerNombre = Trim(.Fields!Nombres)
                      lcPrimerNombre = Left(RetornaPrimerNombre(lcPrimerNombre), 20)
                      lcSegundoNombre = Trim(.Fields!Nombres)
                      lcSegundoNombre = Left(RetornaSegundoNombre(lcSegundoNombre), 20)
                  End If
                  lnTipoSexo = IIf(Left(UCase(.Fields!sexo), 1) = "F", 2, 1)
                  If lnNroHistoriaClinica = 0 Then
                     lbNuevoHC = True
                  Else
                        'Busca si ya existe Nro Historia en GalenHos
                        wrs_GalenHos2.Open "select * from Pacientes where nroHistoriaClinica=" & lnNroHistoriaClinica, wxConexionRed, adOpenKeyset, adLockOptimistic
                        If wrs_GalenHos2.RecordCount > 0 Then
                           lbNuevoHC = True
                        End If
                  End If
                  If lbNuevoHC = True Then
                          'HC repetida
                          'solo graba como problemas
                          wRsProblemas.AddNew
                          wRsProblemas.Fields!pacHis = .Fields!Codigo
                          wRsProblemas.Fields!pacPat = .Fields!ApPaterno
                          wRsProblemas.Fields!pacMat = .Fields!apMaterno
                          wRsProblemas.Fields!pacNam = .Fields!Nombres
                          wRsProblemas.Fields!pacFin = .Fields!f_admis
                          wRsProblemas.Fields!nroHistoriaGalenHos = lnNroHistoriaClinica
                          wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "HcRep-Cuzco"
                          wRsProblemas.Update
                          '
                          If lnNroHistoriaClinica <> 0 Then
                            wRsProblemas.AddNew
                            wRsProblemas.Fields!pacHis = wrs_GalenHos2.Fields!NroHistoriaClinica
                            wRsProblemas.Fields!pacPat = wrs_GalenHos2.Fields!ApellidoPaterno
                            wRsProblemas.Fields!pacMat = wrs_GalenHos2.Fields!ApellidoMaterno
                            wRsProblemas.Fields!pacNam = Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)
                            wRsProblemas.Fields!autogeneradoGalenHos = "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"
                            wRsProblemas.Update
                            '
                            wrs_GalenHos2.Close
                          End If
                  Else
                          wrs_GalenHos2.Close
                          lcApellidoPaterno = Left(Trim(.Fields!ApPaterno), 20)
                          lcApellidoMaterno = Left(Trim(.Fields!apMaterno), 20)
                          lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                          lcFechaAnt = IIf(IsNull(.Fields!f_admis), "19/04/2005", .Fields!f_admis)
                          'Busca en Tabla xx Equivalencia LolCli
                          lntipoOcupacion = 0
                          lnIdDepartamentoDomicilio = 0
                          LnIdProvinciaDomicilio = 0
                          lnIdDistritoDomicilio = 0
                          If Not IsNull(.Fields!codgeo) Then
                            lnIdDepartamentoDomicilio = Val(Left(.Fields!codgeo, 2))
                            LnIdProvinciaDomicilio = Val(Left(.Fields!codgeo, 4))
                            lnIdDistritoDomicilio = Val(.Fields!codgeo)
                          End If
                          lnIdDepartamentoNacimiento = 0
                          LnIdProvinciaNacimiento = 0
                          lnIdDistritoNacimiento = 0
                          lnIdEstadoCivil = 0
                          If Not IsNull(.Fields!estCivil) Then
                             Select Case UCase(Left(.Fields!estCivil, 1))
                             Case "S"   'soltero
                                  lnIdEstadoCivil = 2
                             Case "C"   'casado
                                  lnIdEstadoCivil = 1
                             End Select
                          End If
                          'Graba Pacientes
                          wxConexionRed.BeginTrans
                          wrs_GalenHos1.AddNew
                          wrs_GalenHos1.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                          wrs_GalenHos1.Fields!ApellidoPaterno = lcApellidoPaterno
                          wrs_GalenHos1.Fields!ApellidoMaterno = lcApellidoMaterno
                          wrs_GalenHos1.Fields!PrimerNombre = lcPrimerNombre
                          wrs_GalenHos1.Fields!SegundoNombre = lcSegundoNombre
                          wrs_GalenHos1.Fields!idTipoSexo = lnTipoSexo
                          wrs_GalenHos1.Fields!FechaNacimiento = CDate(lcFechaNac)
                          wrs_GalenHos1.Fields!IdTipoNumeracion = lnTipoNumeracion
                          wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado
                          wrs_GalenHos1.Fields!IdDistritoDomicilio = lnIdDistritoDomicilio
                          If lnIdEstadoCivil > 0 Then
                             wrs_GalenHos1.Fields!IdEstadoCivil = lnIdEstadoCivil
                          End If
                          If Not IsNull(.Fields!direccio1) Then
                             wrs_GalenHos1.Fields!DireccionDomicilio = Left(.Fields!direccio1, 50)
                          End If
                          If Not IsNull(.Fields!padreNomb) Then
                             wrs_GalenHos1.Fields!NombrePadre = Left(.Fields!padreNomb, 20)
                          End If
                          If Not IsNull(.Fields!madreNomb) Then
                             wrs_GalenHos1.Fields!NombreMadre = Left(.Fields!madreNomb, 20)
                          End If
                          If Not IsNull(.Fields!le) Then
                             If Len(Trim(.Fields!le)) = 8 Then
                                wrs_GalenHos1.Fields!NroDocumento = Left(.Fields!le, 8)
                                wrs_GalenHos1.Fields!IdDocIdentidad = 1
                             End If
                          End If
                          'wrs_GalenHos1.Fields!IdEtnia = "80"
                          wrs_GalenHos1.Fields!IdPaisDomicilio = 166
                          wrs_GalenHos1.Fields!IdPaisProcedencia = 166
                          wrs_GalenHos1.Fields!IdPaisNacimiento = 166
                          wrs_GalenHos1.Update
                          lnIdPaciente = wrs_GalenHos1.Fields!idPaciente
                          'Graba HistoriasClinicas
                          wrs_GalenHos.AddNew
                          wrs_GalenHos.Fields!idPaciente = lnIdPaciente
                          wrs_GalenHos.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                          wrs_GalenHos.Fields!FechaCreacion = lcFechaAnt
                          wrs_GalenHos.Fields!IdTipoNumeracion = lnTipoNumeracion
                          wrs_GalenHos.Fields!IdEstadoHistoria = 1
                          wrs_GalenHos.Fields!IdTipoHistoria = 1
                          wrs_GalenHos.Fields!HistoriaSistemaAnterior = .Fields!Codigo
                          wrs_GalenHos.Update
                          wxConexionRed.CommitTrans
                  End If
              End If
              .MoveNext
           Loop
       End With
    End If
    wxConexionJAMO.Close
    Unload Me
    Exit Sub
err_proceso:
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!numHc & "     F.Registro: " & wrs_LolCli.Fields!FechaIngreso & "     Paciente:" & wrs_LolCli.Fields!ApellidoP & " " & wrs_LolCli.Fields!ApellidoM & " " & wrs_LolCli.Fields!nombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    Resume
    Unload Me

End Sub

Private Sub cmdNuevosCptVsGrupo_Click()
       Me.MousePointer = 11
    
        Dim EXL As Excel.Application
        Set EXL = New Excel.Application
        Dim W As Excel.Workbook
        Set W = EXL.Workbooks.Open("c:\excel.xls")
        Dim s As Excel.Worksheet
        Set s = W.Sheets("libro1")
        Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, lcCodigo As String
        Dim lcDepartamento As String, lcProvincia As String, lcDistrito As String
        lnFila = 1
        lnFilaFinal = 20000
       
       
       Dim oConexODBC As New Connection
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsProductos As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim lbEsNuevo As Boolean, lcSql As String, lnPqte1 As Long
       oConexODBC.Open "dsn=GALENHOS"
       lcSql = "select  * from labPruebas"
       oRsProductos.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
       
       With oRsTmp1
        .Fields.Append "codigoPrueba", adVarChar, 6
        .Fields.Append "codigoCPT", adVarChar, 20
        .LockType = adLockOptimistic
        .Open
       End With
       Dim lcCodigoCPT As String, lcCodigoPrueba As String
       For lnFor = lnFila To lnFilaFinal
            
           lcCodigoCPT = "A" + Trim(Str(lnFor))
           lcCodigoCPT = s.Range(lcCodigoCPT).Value
           If lcCodigoCPT = "" Then
              Exit For
           End If
           lcCodigoPrueba = "B" + Trim(Str(lnFor))
           lcCodigoPrueba = s.Range(lcCodigoPrueba).Value
           oRsTmp1.AddNew
           oRsTmp1.Fields!CodigoPrueba = lcCodigoPrueba
           oRsTmp1.Fields!CodigoCPT = Right("0000" & lcCodigoCPT, 5)
           oRsTmp1.Update
       
       Next
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lcSql = "select * from labPruebas where codigoPrueba='" & oRsTmp1!CodigoPrueba & _
                "' and codigoCPT='" & oRsTmp1!CodigoCPT & "'"
          oRsTmp2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
          If oRsTmp2.RecordCount = 0 Then
             oRsProductos.AddNew
             oRsProductos.Fields!CodigoPrueba = oRsTmp1!CodigoPrueba
             oRsProductos.Fields!CodigoCPT = oRsTmp1!CodigoCPT
             oRsProductos.Update
          End If
          oRsTmp2.Close
          oRsTmp1.MoveNext
       Loop
       oRsProductos.Close
       oRsTmp1.Close
       Me.MousePointer = 1
       oConexODBC.Close
       Set oConexODBC = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set oRsProductos = Nothing
       
        Set s = Nothing
        W.Close
        Set W = Nothing
        Set EXL = Nothing
    
       Unload Me
    
End Sub

Private Sub cmdProblemasJamo_Click()
       Dim wrs_LolCli As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim lnNroHistoriaClinica As Long
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim lcSql As String
       Dim lntotReg As Long
       Dim lcHC As String
        Dim lcpacHis As String
        Dim lcpacPat As String
        Dim lcpacMat As String
        Dim lcpacNam As String
        Dim ldpacFin As Date
        Dim lcpacHis1 As String
        Dim lcpacPat1 As String
        Dim lcpacMat1 As String
        Dim lcpacNam1 As String
        Dim ldpacFin1 As Date
        Dim lbNuevo As Boolean
        Dim lcAutogenerado As String
        Dim lnCant As Long
        Dim lnRepetidos As Long
       Me.MousePointer = 11
       On Error GoTo ErrProblJamo
 
       wrs_GalenHos1.Open "delete  from lolcliProblemasHC where left(autogeneradoGalenHos,1) <> '*' ", wxConexionRed, adOpenKeyset, adLockOptimistic
       wrs_GalenHos1.Open "select *  from lolcliProblemasHC", wxConexionRed, adOpenKeyset, adLockOptimistic
       With wrs_GalenHos
           'Historias Clinicas Anteriores NULL o VACIAS
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.Pacientes.SegundoNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   "  where (dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or dbo.HistoriasClinicas.HistoriaSistemaAnterior=''"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           lnCant = 1
           If lntotReg > 0 Then
              Do While Not .EOF
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = Left(Trim(.Fields!PrimerNombre) & " " & Trim(.Fields!SegundoNombre), 50)
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = .Fields!Autogenerado
                    wrs_GalenHos1.AddNew
                    wrs_GalenHos1.Fields!pacHis = ""
                    wrs_GalenHos1.Fields!pacPat = lcpacPat
                    wrs_GalenHos1.Fields!pacMat = lcpacMat
                    wrs_GalenHos1.Fields!pacNam = lcpacNam
                    wrs_GalenHos1.Fields!pacFin = ldpacFin
                    wrs_GalenHos1.Fields!autogeneradoGalenHos = Trim(Str(lnCant)) & "HC null vacias"
                    wrs_GalenHos1.Update
                    lnCant = lnCant + 1
                    .Update
                    .MoveNext
              Loop
           End If
           .Close
           'Autogenerado Repetidos
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.Pacientes.SegundoNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   " Where  not ((dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or (dbo.HistoriasClinicas.HistoriaSistemaAnterior='')) " & _
                   "  order by dbo.Pacientes.autogenerado"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
              lnRepetidos = 1
              Do While Not .EOF
                    lcpacHis = .Fields!HistoriaSistemaAnterior
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = Left(Trim(.Fields!PrimerNombre) & " " & Trim(.Fields!SegundoNombre), 50)
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = .Fields!Autogenerado
                    lnCant = 0
                    Do While Not .EOF And lcAutogenerado = .Fields!Autogenerado
                        lnCant = lnCant + 1
                        If lnCant > 1 Then
                            lcpacHis1 = .Fields!HistoriaSistemaAnterior
                            lcpacPat1 = .Fields!ApellidoPaterno
                            lcpacMat1 = .Fields!ApellidoMaterno
                            lcpacNam1 = Left(Trim(.Fields!PrimerNombre) & " " & Trim(.Fields!SegundoNombre), 50)
                            ldpacFin1 = .Fields!FechaCreacion
                        End If
                        .MoveNext
                        If .EOF Then
                           Exit Do
                        End If
                    Loop
                    If lnCant > 1 Then
                       lbNuevo = True
                       If wrs_GalenHos1.RecordCount > 0 Then
                          wrs_GalenHos1.MoveFirst
                          wrs_GalenHos1.Find "pacHis='" & lcpacHis & "'"
                          If Not wrs_GalenHos1.EOF Then
                             lbNuevo = False
                          End If
                       End If
                       If lbNuevo = True Then
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis
                            wrs_GalenHos1.Fields!pacPat = lcpacPat
                            wrs_GalenHos1.Fields!pacMat = lcpacMat
                            wrs_GalenHos1.Fields!pacNam = lcpacNam
                            wrs_GalenHos1.Fields!pacFin = ldpacFin
                            wrs_GalenHos1.Fields!autogeneradoGalenHos = Trim(Str(lnRepetidos)) & "Autog-Repet"
                            wrs_GalenHos1.Update
                            '
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis1
                            wrs_GalenHos1.Fields!pacPat = lcpacPat1
                            wrs_GalenHos1.Fields!pacMat = lcpacMat1
                            wrs_GalenHos1.Fields!pacNam = lcpacNam1
                            wrs_GalenHos1.Fields!pacFin = ldpacFin1
                            wrs_GalenHos1.Fields!autogeneradoGalenHos = Trim(Str(lnRepetidos)) & "Autog-Repet"
                            wrs_GalenHos1.Update
                            lnRepetidos = lnRepetidos + 1
                       End If
                    End If
              Loop
           End If
           .Close
           'Historias  Repetidas
           lcSql = "SELECT dbo.HistoriasClinicas.FechaCreacion, dbo.HistoriasClinicas.NroHistoriaClinica, dbo.Pacientes.Autogenerado," & _
                   "  dbo.HistoriasClinicas.HistoriaSistemaAnterior , dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.Pacientes.SegundoNombre" & _
                   " FROM         dbo.HistoriasClinicas INNER JOIN" & _
                   "    dbo.Pacientes ON dbo.HistoriasClinicas.IdPaciente = dbo.Pacientes.IdPaciente" & _
                   " Where  not ((dbo.HistoriasClinicas.HistoriaSistemaAnterior is null) or (dbo.HistoriasClinicas.HistoriaSistemaAnterior='')) " & _
                   "  order by dbo.HistoriasClinicas.HistoriaSistemaAnterior"
           .Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg > 0 Then
              lnRepetidos = 1
              Do While Not .EOF
                    lcpacHis = .Fields!HistoriaSistemaAnterior
                    lcpacPat = .Fields!ApellidoPaterno
                    lcpacMat = .Fields!ApellidoMaterno
                    lcpacNam = Left(Trim(.Fields!PrimerNombre) & " " & Trim(.Fields!SegundoNombre), 50)
                    ldpacFin = .Fields!FechaCreacion
                    lcAutogenerado = Trim(.Fields!HistoriaSistemaAnterior)
                    lnCant = 0
                    Do While Not .EOF And lcAutogenerado = Trim(.Fields!HistoriaSistemaAnterior)
                        lnCant = lnCant + 1
                        If lnCant > 1 Then
                            lcpacHis1 = .Fields!HistoriaSistemaAnterior
                            lcpacPat1 = .Fields!ApellidoPaterno
                            lcpacMat1 = .Fields!ApellidoMaterno
                            lcpacNam1 = Left(Trim(.Fields!PrimerNombre) & " " & Trim(.Fields!SegundoNombre), 50)
                            ldpacFin1 = .Fields!FechaCreacion
                        End If
                        .MoveNext
                        If .EOF Then
                           Exit Do
                        End If
                    Loop
                    If lnCant > 1 Then
                       lbNuevo = True
                       If wrs_GalenHos1.RecordCount > 0 Then
                          wrs_GalenHos1.MoveFirst
                          wrs_GalenHos1.Find "pacHis='" & lcpacHis & "'"
                          If Not wrs_GalenHos1.EOF Then
                             lbNuevo = False
                          End If
                       End If
                       If lbNuevo = True Then
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis
                            wrs_GalenHos1.Fields!pacPat = lcpacPat
                            wrs_GalenHos1.Fields!pacMat = lcpacMat
                            wrs_GalenHos1.Fields!pacNam = lcpacNam
                            wrs_GalenHos1.Fields!pacFin = ldpacFin
                            wrs_GalenHos1.Fields!autogeneradoGalenHos = Trim(Str(lnRepetidos)) & "HC-Repet-Gal"
                            wrs_GalenHos1.Update
                            '
                            wrs_GalenHos1.AddNew
                            wrs_GalenHos1.Fields!pacHis = lcpacHis1
                            wrs_GalenHos1.Fields!pacPat = lcpacPat1
                            wrs_GalenHos1.Fields!pacMat = lcpacMat1
                            wrs_GalenHos1.Fields!pacNam = lcpacNam1
                            wrs_GalenHos1.Fields!pacFin = ldpacFin1
                            wrs_GalenHos1.Fields!autogeneradoGalenHos = Trim(Str(lnRepetidos)) & "HC-Repet-Gal"
                            wrs_GalenHos1.Update
                            lnRepetidos = lnRepetidos + 1
                       End If
                    End If
              Loop
           End If
           .Close
           
       End With
      Me.MousePointer = 1
       Unload Me
       Exit Sub
ErrProblJamo:
    MsgBox Err.Description
    Resume
End Sub

Private Sub cmdProcesa_Click()
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
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
       With wrs_LolCli
           'elimina historias GalenHos de esas Fechas
           lblProcesando.Caption = "Eliminando HC ya migradas, en GalenHos"
           lcSql = "select * from HistoriasClinicas where fechaCreacion Between (CONVERT(DATETIME,'" & txtFechaIni.Text & " 00:00:00',103)) and (CONVERT(DATETIME,'" & txtFechaFin.Text & " 23:59:59',103))"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar1.Min = 0
                ProgressBar1.Max = lntotReg
                lnRegAct = 0
                Do While Not wrs_GalenHos.EOF
                   lnRegAct = lnRegAct + 1: ProgressBar1.Value = lnRegAct
                   wxConexionRed.BeginTrans
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   wrs_GalenHos.Delete
                   wrs_GalenHos.Update
                   wrs_GalenHos1.Open "delete from Pacientes where idpaciente=" & lnIdPaciente, wxConexionRed, adOpenKeyset, adLockOptimistic
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
                
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando.Caption = "Insertando HC en GalenHos"
           lcSql = "select * from HistoriasClinicas"
           wrs_GalenHos.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcSql = "select * from Pacientes"
           wrs_GalenHos1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
           lcSql = "select * from Pacientes where pacfin Between (CONVERT(DATETIME,'" & txtFechaIni.Text & " 00:00:00',103)) and (CONVERT(DATETIME,'" & txtFechaFin.Text & " 23:59:59',103))" & _
                   " order by pacfin"
           .Open lcSql, wxConexion, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en LOLCLI"
              Exit Sub
           End If
           ProgressBar1.Min = 0
           ProgressBar1.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
              lbNuevoHC = False
              ProgressBar1.Value = lnRegAct: lnRegAct = lnRegAct + 1
              If Not IsNull(.Fields!pacFen) Then
                 lcFechaNac = .Fields!pacFen
              Else
                 lcFechaNac = "  /  /    "
              End If
              lnNroHistoriaClinica = Val(.Fields!pacHis)
              If lnNroHistoriaClinica = 0 Then
                 'Historia clinica con problemas
                 lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!pacHis)
'                 lbNuevoHC = True
              End If
              lcPrimerNombre = Left(RetornaPrimerNombre(.Fields!pacNam), 20)
              lcSegundoNombre = Left(RetornaSegundoNombre(.Fields!pacNam), 20)
              lnTipoSexo = IIf(UCase(.Fields!sexCod) = "FE", 2, 1)
              'Busca si ya existe Nro Historia en GalenHos
              wrs_GalenHos2.Open "select nroHistoriaClinica from historiasClinicas where nroHistoriaClinica=" & lnNroHistoriaClinica, wxConexionRed, adOpenKeyset, adLockOptimistic
              If wrs_GalenHos2.RecordCount > 0 Then
                 lbNuevoHC = True
              End If
              wrs_GalenHos2.Close
              If lbNuevoHC = False Then
                      If lbNuevoHC Then
                         lnNroHistoriaClinica = generaNuevaNroHistoria(lnNroHistoriaClinica)
                         lcSegundoNombre = "." & LCase(Trim(lcSegundoNombre))
                      End If
                      
                      lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, .Fields!pacPat, .Fields!pacMat, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                      lcFechaAnt = .Fields!pacFin
                      'Busca en Tabla xx Equivalencia LolCli
                      lntipoOcupacion = 0
                      lcSql = "select * from TiposOcupacion where lolcli='" & .Fields!ocuCod & "'"
                      wrs_GalenHos2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                      If wrs_GalenHos2.RecordCount > 0 Then
                         lntipoOcupacion = wrs_GalenHos2.Fields!IdTipoOcupacion
                      End If
                      wrs_GalenHos2.Close
                      If lnNroHistoriaClinica = 501262 Then
                        lnIdDepartamentoDomicilio = 0
                      End If
                      lnIdDepartamentoDomicilio = 0
                      LnIdProvinciaDomicilio = 0
                      lnIdDistritoDomicilio = 0
                      If Not IsNull(.Fields!ubicod) Then
                        wrs_GalenHos2.Open "select * from lolCliUbigeo where ubigeoLolcli='" & .Fields!ubicod & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
                        If wrs_GalenHos2.RecordCount > 0 Then
                              lnIdDepartamentoDomicilio = wrs_GalenHos2.Fields!IdDepartamento
                              LnIdProvinciaDomicilio = wrs_GalenHos2.Fields!IdProvincia
                              lnIdDistritoDomicilio = wrs_GalenHos2.Fields!IdDistrito
                        Else
                              If .Fields!ubicod = "050000" Then
                                    lnIdDepartamentoDomicilio = "05"
                                    LnIdProvinciaDomicilio = "0501"
                                    lnIdDistritoDomicilio = "050101"
                              Else
                                    lnIdDepartamentoDomicilio = Val(Left(.Fields!ubicod, 2))
                                    LnIdProvinciaDomicilio = Val(Left(.Fields!ubicod, 4))
                                    lnIdDistritoDomicilio = Val(.Fields!ubicod)
                              End If
                        End If
                        wrs_GalenHos2.Close
                      End If
                      lnIdDepartamentoNacimiento = 0
                      LnIdProvinciaNacimiento = 0
                      lnIdDistritoNacimiento = 0
                      If Not IsNull(.Fields!pacLun) Then
                            wrs_GalenHos2.Open "select * from lolCliUbigeo where ubigeoLolcli='" & .Fields!pacLun & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
                            If wrs_GalenHos2.RecordCount > 0 Then
                                  lnIdDepartamentoNacimiento = wrs_GalenHos2.Fields!IdDepartamento
                                  LnIdProvinciaNacimiento = wrs_GalenHos2.Fields!IdProvincia
                                  lnIdDistritoNacimiento = wrs_GalenHos2.Fields!IdDistrito
                            End If
                            wrs_GalenHos2.Close
                      End If
                      lnIdEstadoCivil = 0
                      wrs_GalenHos2.Open "select * from TiposEstadoCivil where lolcli='" & .Fields!eciCod & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
                      If wrs_GalenHos2.RecordCount > 0 Then
                         lnIdEstadoCivil = wrs_GalenHos2.Fields!IdEstadoCivil
                      End If
                      wrs_GalenHos2.Close
                      'Graba Pacientes
                      wxConexionRed.BeginTrans
                      wrs_GalenHos1.AddNew
                      wrs_GalenHos1.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                      wrs_GalenHos1.Fields!ApellidoPaterno = Left(.Fields!pacPat, 20)
                      wrs_GalenHos1.Fields!ApellidoMaterno = Left(.Fields!pacMat, 20)
                      wrs_GalenHos1.Fields!PrimerNombre = lcPrimerNombre
                      wrs_GalenHos1.Fields!SegundoNombre = lcSegundoNombre
                      wrs_GalenHos1.Fields!idTipoSexo = lnTipoSexo
                      If Not IsNull(.Fields!pacFen) Then
                         wrs_GalenHos1.Fields!FechaNacimiento = .Fields!pacFen
                      End If
                      If Not IsNull(.Fields!pacDoc) Then
                         wrs_GalenHos1.Fields!IdDocIdentidad = 1
                         wrs_GalenHos1.Fields!NroDocumento = Left(.Fields!pacDoc, 8)
                      End If
                      wrs_GalenHos1.Fields!IdTipoNumeracion = lnTipoNumeracion
                      wrs_GalenHos1.Fields!Autogenerado = lcAutogenerado
                      If lntipoOcupacion > 0 Then
                         wrs_GalenHos1.Fields!IdTipoOcupacion = lntipoOcupacion
                      End If
                      wrs_GalenHos1.Fields!DireccionDomicilio = Left(.Fields!pacDir, 50)
                      If lnIdDepartamentoDomicilio > 0 Then
                         'wrs_GalenHos1.Fields!IdDepartamentoDomicilio = lnIdDepartamentoDomicilio
                      End If
                      If LnIdProvinciaDomicilio > 0 Then
                         'wrs_GalenHos1.Fields!IdProvinciaDomicilio = LnIdProvinciaDomicilio
                      End If
                      If lnIdDistritoDomicilio > 0 Then
                         wrs_GalenHos1.Fields!IdDistritoDomicilio = lnIdDistritoDomicilio
                      End If
                      If lnIdDepartamentoNacimiento > 0 Then
                         'wrs_GalenHos1.Fields!IdDepartamentoNacimiento = lnIdDepartamentoNacimiento
                      End If
                      If LnIdProvinciaNacimiento > 0 Then
                         'wrs_GalenHos1.Fields!IdProvinciaNacimiento = LnIdProvinciaNacimiento
                      End If
                      If lnIdDistritoNacimiento > 0 Then
                         wrs_GalenHos1.Fields!IdDistritoNacimiento = lnIdDistritoNacimiento
                      End If
                      If Not IsNull(.Fields!pacTel) Then
                         wrs_GalenHos1.Fields!Telefono = Left(.Fields!pacTel, 10)
                      End If
                      wrs_GalenHos1.Fields!IdEstadoCivil = lnIdEstadoCivil
                      wrs_GalenHos1.Update
                      lnIdPaciente = wrs_GalenHos1.Fields!idPaciente
                      'Graba HistoriasClinicas
                      wrs_GalenHos.AddNew
                      wrs_GalenHos.Fields!idPaciente = lnIdPaciente
                      wrs_GalenHos.Fields!NroHistoriaClinica = lnNroHistoriaClinica
                      wrs_GalenHos.Fields!FechaCreacion = .Fields!pacFin
                      wrs_GalenHos.Fields!IdTipoNumeracion = lnTipoNumeracion
                      wrs_GalenHos.Fields!IdEstadoHistoria = 1
                      wrs_GalenHos.Fields!IdTipoHistoria = 1
                      wrs_GalenHos.Fields!HistoriaSistemaAnterior = .Fields!pacHis
                      wrs_GalenHos.Update
                      wxConexionRed.CommitTrans
              End If
              .MoveNext
           Loop
       End With
    End If
    Unload Me
    Exit Sub
err_proceso:
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!pacHis & "     F.Registro: " & wrs_LolCli.Fields!pacFin & "     Paciente:" & wrs_LolCli.Fields!pacPat & " " & wrs_LolCli.Fields!pacMat & " " & wrs_LolCli.Fields!pacNam & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    'Resume
End Sub

Private Sub cmdProcesaCuentas_Click()
       Dim mo_ReglasFacturacion  As New SIGHNegocios.ReglasFacturacion
       Dim lnFor As Long, lcHoraInicio As String, lcHoraActual As String, lcSql As String
       Me.MousePointer = 11
       '
       If Me.optActImpUnaCuenta.Value = True Then
           mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar Val(txtCtaInicial.Text), True, Val(txtTiempoProcesaCtas.Text)
       Else
           lcHoraInicio = Time
           For lnFor = Val(Me.txtCta1.Text) To Val(Me.txtCta2.Text)
                lcHoraActual = Time
                lcSql = DateDiff("n", CDate(lcHoraInicio), CDate(lcHoraActual))
                If Val(lcSql) > Val(txtTiempoProcesaCtas.Text) Then
                   Exit For
                End If
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar lnFor, False, 0
           Next
       End If
       Me.MousePointer = 1
       Unload Me
End Sub

Private Sub cmdProcesaHISlluy_Click()
        Dim oRsFox As New Recordset
        Dim oRsFox1 As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        Dim oConexion1 As New Connection
        Dim oConexionFox As New Connection
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim lnIdUsuario As Long, lbContinuar As Boolean, ldFechaNac As Date
        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
        Dim lcSegundoNombre As String, lnTipoSexo As Long, lnIdPaciente As Long
        Dim lcAutogenerado As String, lcDNI As String, lnNroHistoriaClinica As Long
        Dim lcSql As String, lcCodDx As String, lcCod2000 As String, lbEsNuevaHC As Boolean
        Const lnIdTipoNumeracion As Long = 2
        Dim lbEncontroDx As Boolean, lcFechaAtencion As String, lcHistoria As String
        Dim lbEncontroCitaEnGalenhos As Boolean, lbDxDiferentes As Boolean, lcDxGalenhos As String
        Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
        On Error GoTo ErrProAtHis
        '
        oConexion1.CommandTimeout = 300
        oConexion1.CursorLocation = adUseClient
        oConexion1.Open sighentidades.CadenaConexion
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        lcSql = "update LluyDet set iguales='',dxgalenhos='',dxhis=''"
        If oRsFox1.State = 1 Then oRsFox1.Close
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        
        lcSql = "select * from LluyCab"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           ProgressBar2.Max = oRsFox.RecordCount
           ProgressBar2.Min = 0
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
              Me.Refresh: ProgressBar2.Value = ProgressBar2.Value + 1
              DoEvents
              '
              lcDNI = Mid(oRsFox!DNI, 5, 8)
              Text3.Text = lcDNI
              lcSql = "select * from LluyDet where left(DNI,12)='" & Left(oRsFox!DNI, 12) & "'"
              If oRsFox1.State = 1 Then oRsFox1.Close
              oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
              '
              lcHistoria = Trim(Str(oRsFox!fichafam))
              lcSql = "select * from LluyDet where left(DNI,12)='" & Left(oRsFox!DNI, 12) & "'"
              If oRsFox1.State = 1 Then oRsFox1.Close
              oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
              If oRsFox1.RecordCount > 0 Then
                 oRsFox1.MoveFirst
                 Do While Not oRsFox1.EOF
                    lbEncontroCitaEnGalenhos = False
                    lcFechaAtencion = Right("0" & Trim(Str(oRsFox1!dia)), 2) & "/" & Right("0" & Trim(Str(oRsFox1!Mes)), 2) & "/" & Trim(Str(oRsFox1!ano))
                    lcSql = "select * from atencionesCE where NroHistoriaClinica=" & lcHistoria & " and CitaFecha='" & lcFechaAtencion & "'"
                    If oRsTmp1.State = 1 Then oRsTmp1.Close
                    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    If oRsTmp1.RecordCount = 0 Then
                        lcSql = "select * from Pacientes where nroDocumento='" & lcDNI & "' and idDocIdentidad='1'"
                        If oRsTmp1.State = 1 Then oRsTmp1.Close
                        oRsTmp1.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
                        If oRsTmp1.RecordCount > 0 Then
                           lcHistoria = Trim(Str(oRsTmp1.Fields!NroHistoriaClinica))
                           lcSql = "select * from atencionesCE where NroHistoriaClinica=" & lcHistoria & " and CitaFecha='" & lcFechaAtencion & "'"
                           lcSql = "select * from atencionesCE where NroHistoriaClinica=" & lcHistoria
                           If oRsTmp1.State = 1 Then oRsTmp1.Close
                           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                           If oRsTmp1.RecordCount > 0 Then
                              lbEncontroCitaEnGalenhos = True
                           End If
                        End If
                    Else
                        lbEncontroCitaEnGalenhos = True
                    End If
                    If lbEncontroCitaEnGalenhos = True Then
                       oRsTmp1.MoveFirst
                       lcDx1 = ""
                       If soloNumeros(oRsFox1!codigo1) = False Then
                          lcDx1 = Left(DxAsignaPunto(oRsFox1!codigo1), 5)
                       End If
                       lcDx2 = ""
                       If soloNumeros(oRsFox1!codigo2) = False Then
                          lcDx2 = Left(DxAsignaPunto(oRsFox1!codigo2), 5)
                       End If
                       lcDx3 = ""
                       If soloNumeros(oRsFox1!codigo3) = False Then
                          lcDx3 = Left(DxAsignaPunto(oRsFox1!codigo3), 5)
                       End If
                       lcDx4 = ""
                       If soloNumeros(oRsFox1!codigo4) = False Then
                          lcDx4 = Left(DxAsignaPunto(oRsFox1!codigo4), 5)
                       End If
                       lcDx5 = ""
                       If soloNumeros(oRsFox1!codigo5) = False Then
                          lcDx5 = Left(DxAsignaPunto(oRsFox1!codigo5), 5)
                       End If
                       lcDx6 = ""
                       If soloNumeros(oRsFox1!codigo6) = False Then
                          lcDx6 = Left(DxAsignaPunto(oRsFox1!codigo6), 5)
                       End If
                       lbDxDiferentes = True
                       lcDxGalenhos = ""
                       Do While Not oRsTmp1.EOF
                          If Year(oRsTmp1.Fields!CitaFecha) = oRsFox1!ano And Month(oRsTmp1.Fields!CitaFecha) = oRsFox1!Mes And Day(oRsTmp1.Fields!CitaFecha) = oRsFox1!dia Then
                            lcDxGalenhos = Left(oRsTmp1.Fields!CitaDiagMed, 250)
                            If lcDx1 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx1) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx2 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx2) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx3 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx3) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx4 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx4) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx5 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx5) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx6 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx6) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5 & ", Dx6=" & lcDx6
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            End If
                          End If
                          oRsTmp1.MoveNext
                       Loop
                       If lcDxGalenhos = "" Then
                            lbEncontroCitaEnGalenhos = False
                       ElseIf lbDxDiferentes = True Then
                            oRsFox1!iguales = "DxDIFERENTES"
                            oRsFox1!DxGalenhos = lcDxGalenhos
                            oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5 & ", Dx6=" & lcDx6
                            oRsFox1.Update
                       End If
                    End If
                    If lbEncontroCitaEnGalenhos = False Then
                       oRsFox1!iguales = "SinCitaEnGalenhos"
                       oRsFox1.Update
                    End If
                    oRsFox1.MoveNext
                 Loop
              Else
                 oRsFox1!iguales = "SinDetalleDBF"
                 oRsFox1.Update
              End If
              oRsFox1.Close
              
              oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        oConexionFox.Close
        Set oConexionFox = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Me.MousePointer = 11
        Unload Me
        Exit Sub
ErrProAtHis:
    MsgBox Err.Description
    Resume
End Sub

Function DxAsignaPunto(lcDx As String) As String
     DxAsignaPunto = Trim(Left(lcDx, 3) & "." & Mid(lcDx, 4))
End Function
Function soloNumeros(lcTexto As String) As Boolean
    Dim lnFor As Integer
    soloNumeros = True
    For lnFor = 1 To Len(lcTexto)
        If InStr(Mid(lcTexto, lnFor, 1), "0123456789") = 0 Then
           soloNumeros = False
           Exit Function
        End If
    Next
    
End Function


Private Sub cmdProcesaHISlluy1_Click()
        Dim oRsFox As New Recordset
        Dim oRsFox1 As New Recordset
        Dim oRsTmp As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        Dim oConexion1 As New Connection
        Dim oConexionFox As New Connection
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim lnIdUsuario As Long, lbContinuar As Boolean, ldFechaNac As Date
        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
        Dim lcSegundoNombre As String, lnTipoSexo As Long, lnIdPaciente As Long
        Dim lcAutogenerado As String, lcDNI As String, lnNroHistoriaClinica As Long
        Dim lcSql As String, lcCodDx As String, lcCod2000 As String, lbEsNuevaHC As Boolean
        Const lnIdTipoNumeracion As Long = 2
        Dim lbEncontroDx As Boolean, lcFechaAtencion As String, lcHistoria As String
        Dim lbEncontroCitaEnGalenhos As Boolean, lbDxDiferentes As Boolean, lcDxGalenhos As String
        Dim lbEncontroCitaHis As Boolean
        Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
        On Error GoTo ErrProAtHis
        '
        oConexion1.CommandTimeout = 300
        oConexion1.CursorLocation = adUseClient
        oConexion1.Open sighentidades.CadenaConexion
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        lcSql = "update LluyDet set iguales='',dxgalenhos='',dxhis=''"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        lcSql = "select * from atencionesCE where month(CitaFecha)=4 and year(citaFecha)=2014"
        oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp1.RecordCount > 0 Then
           ProgressBar2.Max = oRsTmp1.RecordCount
           ProgressBar2.Min = 0
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
If oRsTmp1!NroHistoriaClinica = 11598 Or oRsTmp1!NroHistoriaClinica = 1165 Then
lcSql = ""
End If

              Me.Refresh: ProgressBar2.Value = ProgressBar2.Value + 1
              DoEvents
              '
              lcHistoria = Trim(Str(oRsTmp1!NroHistoriaClinica))
              lcDNI = ""
              lcSql = "select * from Pacientes where NroHistoriaClinica=" & lcHistoria
              If oRsTmp.State = 1 Then oRsTmp.Close
              oRsTmp.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
              If oRsTmp.RecordCount > 0 Then
                 If Not IsNull(oRsTmp.Fields!NroDocumento) Then
                    lcDNI = oRsTmp.Fields!NroDocumento
                 End If
              End If
              '
              lcSql = "select * from LluyDet where left(dni,12)='" & "PER1" & lcDNI & "'"
              If oRsFox1.State = 1 Then oRsFox1.Close
              oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
              lbEncontroCitaHis = False
              If oRsFox1.RecordCount = 0 Then
                 lcSql = "select * from LluyDet where fichaFam=" & lcHistoria
                 If oRsFox1.State = 1 Then oRsFox1.Close
                 oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
                 If oRsFox1.RecordCount > 0 Then
                    lbEncontroCitaHis = True
                 End If
              Else
                 lbEncontroCitaHis = True
              End If
              If lbEncontroCitaHis = True Then
                 lcDxGalenhos = ""
                 lbDxDiferentes = True
                 oRsFox1.MoveFirst
                 Do While Not oRsFox1.EOF
                       If Year(oRsTmp1.Fields!CitaFecha) = oRsFox1!ano And _
                                        Month(oRsTmp1.Fields!CitaFecha) = oRsFox1!Mes And Day(oRsTmp1.Fields!CitaFecha) = oRsFox1!dia Then
                            lcDx1 = ""
                            If soloNumeros(oRsFox1!codigo1) = False Then
                               lcDx1 = Left(DxAsignaPunto(oRsFox1!codigo1), 5)
                            End If
                            lcDx2 = ""
                            If soloNumeros(oRsFox1!codigo2) = False Then
                               lcDx2 = Left(DxAsignaPunto(oRsFox1!codigo2), 5)
                            End If
                            lcDx3 = ""
                            If soloNumeros(oRsFox1!codigo3) = False Then
                               lcDx3 = Left(DxAsignaPunto(oRsFox1!codigo3), 5)
                            End If
                            lcDx4 = ""
                            If soloNumeros(oRsFox1!codigo4) = False Then
                               lcDx4 = Left(DxAsignaPunto(oRsFox1!codigo4), 5)
                            End If
                            lcDx5 = ""
                            If soloNumeros(oRsFox1!codigo5) = False Then
                               lcDx5 = Left(DxAsignaPunto(oRsFox1!codigo5), 5)
                            End If
                            lcDx6 = ""
                            If soloNumeros(oRsFox1!codigo6) = False Then
                               lcDx6 = Left(DxAsignaPunto(oRsFox1!codigo6), 5)
                            End If
                            '
                       
                            lcDxGalenhos = Left(oRsTmp1.Fields!CitaDiagMed, 250)
                            If lcDx1 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx1) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1
                               oRsFox1.Update
                               lbDxDiferentes = False
                            ElseIf lcDx2 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx2) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2
                               oRsFox1.Update
                               lbDxDiferentes = False
                            ElseIf lcDx3 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx3) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3
                               oRsFox1.Update
                               lbDxDiferentes = False
                            ElseIf lcDx4 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx4) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4
                               oRsFox1.Update
                               lbDxDiferentes = False
                            ElseIf lcDx5 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx5) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5
                               oRsFox1.Update
                               lbDxDiferentes = False
                            ElseIf lcDx6 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx6) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5 & ", Dx6=" & lcDx6
                               oRsFox1.Update
                               lbDxDiferentes = False
                            Else
                               oRsFox1!iguales = "DxDIFERENTES"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5 & ", Dx6=" & lcDx6
                               oRsFox1.Update
                            
                            End If
                       End If
                       
                    
                       oRsFox1.MoveNext
                 Loop
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        Set oConexionFox = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Me.MousePointer = 11
        Unload Me
        Exit Sub
ErrProAtHis:
    MsgBox Err.Description
    Resume

End Sub

Private Sub cmdProcesaHISlluyDET_Click()
        Dim oRsFox As New Recordset
        Dim oRsFox1 As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        Dim oConexion1 As New Connection
        Dim oConexionFox As New Connection
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim lnIdUsuario As Long, lbContinuar As Boolean, ldFechaNac As Date
        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
        Dim lcSegundoNombre As String, lnTipoSexo As Long, lnIdPaciente As Long
        Dim lcAutogenerado As String, lcDNI As String, lnNroHistoriaClinica As Long
        Dim lcSql As String, lcCodDx As String, lcCod2000 As String, lbEsNuevaHC As Boolean
        Const lnIdTipoNumeracion As Long = 2
        Dim lbEncontroDx As Boolean, lcFechaAtencion As String, lcHistoria As String
        Dim lbEncontroCitaEnGalenhos As Boolean, lbDxDiferentes As Boolean, lcDxGalenhos As String
        Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
        On Error GoTo ErrProAtHis
        '
        oConexion1.CommandTimeout = 300
        oConexion1.CursorLocation = adUseClient
        oConexion1.Open sighentidades.CadenaConexion
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
              
              lcSql = "update LluyDet set iguales='',dxgalenhos='',dxhis=''"
              If oRsFox1.State = 1 Then oRsFox1.Close
              oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
              
              lcSql = "select * from LluyDet where ano=2014 and mes=4"
              If oRsFox1.State = 1 Then oRsFox1.Close
              oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
              If oRsFox1.RecordCount > 0 Then
                 ProgressBar2.Max = oRsFox1.RecordCount
                 ProgressBar2.Min = 0
                 oRsFox1.MoveFirst
                 Do While Not oRsFox1.EOF
                    Me.Refresh: ProgressBar2.Value = ProgressBar2.Value + 1
                    DoEvents
                    '
                    lcDNI = Mid(oRsFox1!DNI, 5, 8)
                    Text3.Text = lcDNI
                    lcHistoria = Trim(Str(oRsFox1!fichafam))
                    '
                    lbEncontroCitaEnGalenhos = False
                    lcFechaAtencion = Right("0" & Trim(Str(oRsFox1!dia)), 2) & "/" & Right("0" & Trim(Str(oRsFox1!Mes)), 2) & "/" & Trim(Str(oRsFox1!ano))
                    lcSql = "select * from atencionesCE where NroHistoriaClinica=" & lcHistoria
                    If oRsTmp1.State = 1 Then oRsTmp1.Close
                    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    If oRsTmp1.RecordCount = 0 Then
                        lcSql = "select * from Pacientes where nroDocumento='" & lcDNI & "' and idDocIdentidad='1'"
                        If oRsTmp1.State = 1 Then oRsTmp1.Close
                        oRsTmp1.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
                        If oRsTmp1.RecordCount > 0 Then
                           lcHistoria = Trim(Str(oRsTmp1.Fields!NroHistoriaClinica))
                           lcSql = "select * from atencionesCE where NroHistoriaClinica=" & lcHistoria
                           If oRsTmp1.State = 1 Then oRsTmp1.Close
                           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                           If oRsTmp1.RecordCount > 0 Then
                              lbEncontroCitaEnGalenhos = True
                           End If
                        End If
                    Else
                        lbEncontroCitaEnGalenhos = True
                    End If
                    If lbEncontroCitaEnGalenhos = True Then
                       oRsTmp1.MoveFirst
                       lcDx1 = ""
                       If soloNumeros(oRsFox1!codigo1) = False Then
                          lcDx1 = Left(DxAsignaPunto(oRsFox1!codigo1), 5)
                       End If
                       lcDx2 = ""
                       If soloNumeros(oRsFox1!codigo2) = False Then
                          lcDx2 = Left(DxAsignaPunto(oRsFox1!codigo2), 5)
                       End If
                       lcDx3 = ""
                       If soloNumeros(oRsFox1!codigo3) = False Then
                          lcDx3 = Left(DxAsignaPunto(oRsFox1!codigo3), 5)
                       End If
                       lcDx4 = ""
                       If soloNumeros(oRsFox1!codigo4) = False Then
                          lcDx4 = Left(DxAsignaPunto(oRsFox1!codigo4), 5)
                       End If
                       lcDx5 = ""
                       If soloNumeros(oRsFox1!codigo5) = False Then
                          lcDx5 = Left(DxAsignaPunto(oRsFox1!codigo5), 5)
                       End If
                       lcDx6 = ""
                       If soloNumeros(oRsFox1!codigo6) = False Then
                          lcDx6 = Left(DxAsignaPunto(oRsFox1!codigo6), 5)
                       End If
                       lbDxDiferentes = True
                       lcDxGalenhos = ""
                       Do While Not oRsTmp1.EOF
                          If IsNull(oRsTmp1.Fields!CitaFecha) Then
                          Else
                          If Year(oRsTmp1.Fields!CitaFecha) = oRsFox1!ano And Month(oRsTmp1.Fields!CitaFecha) = oRsFox1!Mes And Day(oRsTmp1.Fields!CitaFecha) = oRsFox1!dia Then
                            lcDxGalenhos = Left(oRsTmp1.Fields!CitaDiagMed, 250)
                            If lcDx1 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx1) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx2 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx2) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx3 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx3) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx4 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx4) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx5 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx5) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            ElseIf lcDx6 <> "" And InStr(oRsTmp1.Fields!CitaDiagMed, lcDx6) > 0 Then
                               oRsFox1!iguales = "DxIGUAL"
                               oRsFox1!DxGalenhos = lcDxGalenhos
                               oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5 & ", Dx6=" & lcDx6
                               oRsFox1.Update
                               lbDxDiferentes = False
                               Exit Do
                            End If
                          End If
                          End If
                          oRsTmp1.MoveNext
                       Loop
                       If lcDxGalenhos = "" Then
                            lbEncontroCitaEnGalenhos = False
                       ElseIf lbDxDiferentes = True Then
                            oRsFox1!iguales = "DxDIFERENTES"
                            oRsFox1!DxGalenhos = lcDxGalenhos
                            oRsFox1!DxHis = "Dx1=" & lcDx1 & ", Dx2=" & lcDx2 & ", Dx3=" & lcDx3 & ", Dx4=" & lcDx4 & ", Dx5=" & lcDx5 & ", Dx6=" & lcDx6
                            oRsFox1.Update
                       End If
                    End If
                    If lbEncontroCitaEnGalenhos = False Then
                       oRsFox1!iguales = "SinCitaEnGalenhos"
                       oRsFox1!DxGalenhos = ""
                       oRsFox1!DxHis = ""
                       oRsFox1.Update
                    End If
                    oRsFox1.MoveNext
                 Loop
              End If
        oRsFox1.Close
        oConexionFox.Close
        Set oConexionFox = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Me.MousePointer = 11
        Unload Me
        Exit Sub
ErrProAtHis:
    MsgBox Err.Description
    Resume

End Sub

Private Sub cmdProcesaPartidas_Click()
       Dim mo_ReglasCaja  As New SIGHNegocios.ReglasCaja
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim mrs_Tmp As New Recordset
       Dim oConexion As New Connection
       Dim oPartidasPresupuestales As New PartidasPresupuestales
       Dim lnFor As Long, lcHoraInicio As String, lcHoraActual As String, lcSql As String
       Dim ldFechaInicio As Date, ldFechaFinal As Date, ldFechaMaxima As Date
       Dim lnCorrelativo As Long
       On Error GoTo errDesc
       Me.MousePointer = 11
       ProgressBar2.Min = 0
       ProgressBar2.Max = 367
       lcHoraInicio = Time
       '
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       'chequea si llegó al final del año
       ldFechaMaxima = CDate("31/12/" & Me.txtAnioProc.Text)
       Set oRsTmp1 = mo_ReglasCaja.FactPartidasPresupuestalesXMesSelecionaUltimoProceso(ldFechaMaxima, oConexion)
       ldFechaInicio = CDate("01/01/" & Me.txtAnioProc.Text)
       If oRsTmp1.RecordCount > 0 Then
           If oRsTmp1!Fecha = ldFechaMaxima Then
              MsgBox "Ese año ya se terminó de procesar"
              Unload Me
           Else
              If Year(oRsTmp1!Fecha) = Val(Me.txtAnioProc.Text) Then
                 ldFechaInicio = oRsTmp1!Fecha + 1
              End If
              oRsTmp1.Close
              Set oRsTmp1 = mo_ReglasCaja.CajaComprobantePagoUltimaBoletaDelAnio(ldFechaMaxima, oConexion)
              If oRsTmp1.RecordCount > 0 Then
                 ldFechaMaxima = CDate(Format(oRsTmp1!FechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY))
              End If
              If ldFechaInicio > ldFechaMaxima Then
                   MsgBox "Ese año ya se terminó de procesar"
                   Unload Me
              End If
           End If
       End If
       lnCorrelativo = 1
       oRsTmp1.Close
       Do While True
             '
             DoEvents
             ProgressBar2.Value = DateDiff("d", CDate("01/01/" & Me.txtAnioProc.Text), ldFechaInicio)
             Me.Refresh
             '
             
             lcHoraActual = Time
             lcSql = DateDiff("n", CDate(lcHoraInicio), CDate(lcHoraActual))
             If Val(lcSql) > Val(txtMinutosProc.Text) Or ldFechaMaxima < ldFechaInicio Then
                Exit Do
             End If
             'procesa Partidas por día
             mo_ReglasCaja.FactPartidasPresupuestalesXMesEliminar ldFechaInicio, oConexion
             ldFechaFinal = CDate(Format(ldFechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY) & " 23:59:59")
             mo_SIGHProxies.ReportePartidaREsumen mrs_Tmp, ldFechaInicio, ldFechaFinal, _
                                                 99999, sghProcesaYgraba, oConexion, True, 0, False
             lnCorrelativo = lnCorrelativo + 1
             If mrs_Tmp.RecordCount > 0 Then
                mrs_Tmp.MoveFirst
                Do While Not mrs_Tmp.EOF
                      mo_ReglasCaja.FactPartidasPresupuestalesXMesAgregar ldFechaInicio, _
                                                                          mrs_Tmp!IdPartida, _
                                                                          mrs_Tmp!identificador, _
                                                                          mrs_Tmp!ImpAnulado, _
                                                                          mrs_Tmp!ImpExonerado, _
                                                                          mrs_Tmp!ImpNormal, _
                                                                          mrs_Tmp!ImpCancelado, _
                                                                          oConexion
                      mrs_Tmp.MoveNext
                Loop
             End If
             Set mrs_Tmp = Nothing
             'pasa al siguiente día
             ldFechaInicio = ldFechaInicio + 1
       Loop
       Me.MousePointer = 1
       Set mo_ReglasCaja = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set mrs_Tmp = Nothing
       Set oConexion = Nothing
       Set oPartidasPresupuestales = Nothing
       
       Unload Me
       Exit Sub
errDesc:
       MsgBox Err.Description
       Resume
End Sub

Private Sub cmdProcesaStaElena_Click()

    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=HIS"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
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
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim lcApellidoPaterno As String, lcApellidoMaterno As String
       Dim lbContinuarProceso As Boolean
       Dim wFec1 As String, wFec2 As String, wFFecha As Date
       Dim lnTipoNumeracion1 As Long
       lnTipoNumeracion1 = 1
       With wrs_LolCli
           wFFecha = CDate(Me.txtINIse.Text)
           wFec1 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           wFFecha = CDate(Me.txtFINse.Text)
           wFec2 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           
           Me.txtIni1.Text = Me.txtINIse.Text
           Me.txtFin1.Text = Me.txtFINse.Text
           lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "HistoriasClinicasSeleccionarPorFechaCreacion"
                Set oParameter = .CreateParameter("@FechaInicio", adVarChar, adParamInput, 20, Me.txtIni1.Text): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaFin", adVarChar, adParamInput, 20, Me.txtFin1.Text): .Parameters.Append oParameter
                Set wrs_GalenHos = .Execute
                Set wrs_GalenHos.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           Set oParameter = Nothing
                      
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar2.Min = 0
                ProgressBar2.Max = lntotReg
                lnRegAct = 0
                Do While Not wrs_GalenHos.EOF
                   lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                   wxConexionRed.BeginTrans
                   
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   
                   With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "HistoriasClinicasEliminar"
                         Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, wrs_GalenHos.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                         Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                         .Execute
                   End With
                   Set oCommand = Nothing
                   Set oParameter = Nothing
                   
                   
                   With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "PacientesEliminarPorIdPaciente"
                         Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
                         Set wrs_GalenHos1 = .Execute
                         Set wrs_GalenHos1.ActiveConnection = Nothing
                   End With
                   Set oCommand = Nothing
                   
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
                
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando11.Caption = "Insertando HC en GalenHos"
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "HistoriasClinicasSeleccionarTodos"
                Set wrs_GalenHos = .Execute
                Set wrs_GalenHos.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "PacientesConsultarTodos"
                Set wrs_GalenHos1 = .Execute
                Set wrs_GalenHos1.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           lcSql = "select * from t_hiscli " & _
                    "  where  fec_exp>=" & wFec1 & " and fec_exp<=" & wFec2 & _
                     " order by fec_exp"
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en CS Santa Elena"
              wxConexionJAMO.Close
              Unload Me
              Exit Sub
           End If
           
           
           With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = wxConexionRed
                  .CommandTimeout = 150
                  .CommandText = "lolcliProblemasHCSeleccionarTodos"
                  Set wRsProblemas = .Execute
                  Set wRsProblemas.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
                      
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
              ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
              lbContinuarProceso = True
              If IsNull(.Fields!apepat) Or Trim(.Fields!apepat) = "" Or IsNull(.Fields!apemat) Or Trim(.Fields!apemat) = "" Or IsNull(.Fields!codHis) Or Trim(.Fields!codHis) = "" Then
                 lbContinuarProceso = False
              End If
              If lbContinuarProceso = True Then
                  lbNuevoHC = False
If Val(.Fields!codHis) = 41403 Then
lbNuevoHC = False
End If
                  If Not sighentidades.EsFecha(Trim(.Fields!fecnac), "DD/MM/AAAA") Then
                        lcFechaNac = "01/01/1990"
                  Else
                        lcFechaNac = .Fields!fecnac
                  End If
                  lnNroHistoriaClinica = 0
                  If UCase(Left(.Fields!codHis, 2)) = "T-" Then
                     lnNroHistoriaClinica = "999" + Trim(Str(Val(Mid(.Fields!codHis, 3, 100))))
                  Else
                     lnNroHistoriaClinica = Val(.Fields!codHis)
                  End If
                  If lnNroHistoriaClinica = 0 Then
                     'Historia clinica con problemas
                     lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!codHis)
    '                 lbNuevoHC = True
                  End If
                  If IsNull(.Fields!nombre) Then
                      lcPrimerNombre = "NN"
                      lcSegundoNombre = ""
                  Else
                      lcPrimerNombre = Trim(.Fields!nombre)
                      lcPrimerNombre = Left(RetornaPrimerNombre(lcPrimerNombre), 20)
                      lcSegundoNombre = Trim(.Fields!nombre)
                      lcSegundoNombre = Left(RetornaSegundoNombre(lcSegundoNombre), 20)
                  End If
                  
                  lnTipoSexo = IIf(UCase(.Fields!sexo) = "M", 1, 2)
                  If lnNroHistoriaClinica = 0 Then
                     lbNuevoHC = True
                  Else
                        'Busca si ya existe Nro Historia en GalenHos
                       
                        
                        With oCommand
                               .CommandType = adCmdStoredProc
                               Set .ActiveConnection = wxConexionRed
                               .CommandTimeout = 150
                               .CommandText = "PacientesSeleccionarPorNroHistoriaClinica"
                               Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                               Set wrs_GalenHos2 = .Execute
                               Set wrs_GalenHos2.ActiveConnection = Nothing
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        
                        If wrs_GalenHos2.RecordCount > 0 Then
                           lbNuevoHC = True
                        End If
                        
                  End If
                  If lbNuevoHC = True Then
                          'HC repetida
                          'solo graba como problemas
                          
                          '
                              oCommand.CommandType = adCmdStoredProc
                              Set oCommand.ActiveConnection = wxConexionRed
                              oCommand.CommandTimeout = 150
                              oCommand.CommandText = "lolcliProblemasHCAgregar"
                              Set oParameter = oCommand.CreateParameter("@pacHis", adChar, adParamInput, 50, .Fields!codHis): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacPat", adChar, adParamInput, 30, .Fields!ape_pat): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacMat", adChar, adParamInput, 30, .Fields!ape_mat): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacNam", adChar, adParamInput, 50, .Fields!nombre): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacFin", adDBTimeStamp, adParamInput, 0, CDate(Me.txtFIniCSN.Text)): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@nroHistoriaGalenHos", adVarChar, adParamInput, 50, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@autogeneradoGalenHos", adVarChar, adParamInput, 50, "*" & Trim(Str(lnRegAct)) & "HcRep-StaElena"): oCommand.Parameters.Append oParameter
                              oCommand.Execute
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                          '
                          If lnNroHistoriaClinica <> 0 Then
                          
                            
                            With oCommand
                                .CommandType = adCmdStoredProc
                                Set .ActiveConnection = wxConexionRed
                                .CommandTimeout = 150
                                .CommandText = "lolcliProblemasHCAgregarSinNroHistoriaClinica"
                                Set oParameter = .CreateParameter("@pacHis", adChar, adParamInput, 50, wrs_GalenHos2.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacPat", adChar, adParamInput, 30, wrs_GalenHos2.Fields!ApellidoPaterno): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacMat", adChar, adParamInput, 30, wrs_GalenHos2.Fields!ApellidoMaterno): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacNam", adChar, adParamInput, 50, Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@autogeneradoGalenHos", adVarChar, adParamInput, 50, "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"): .Parameters.Append oParameter
                                .Execute
                            End With
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                            '
                            wrs_GalenHos2.Close
                          End If
                  Else
                          wrs_GalenHos2.Close
                          lcApellidoPaterno = UCase(Left(Trim(.Fields!apepat), 20))
                          lcApellidoMaterno = UCase(Left(Trim(.Fields!apemat), 20))
                          lcPrimerNombre = UCase(lcPrimerNombre)
                          lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                          lcFechaAnt = .Fields!FEC_EXP
                          'Busca en Tabla xx Equivalencia LolCli
                          lntipoOcupacion = 0
                          lnIdDepartamentoDomicilio = 0
                          LnIdProvinciaDomicilio = 0
                          lnIdDistritoDomicilio = 0
                          lnIdDepartamentoNacimiento = 0
                          LnIdProvinciaNacimiento = 0
                          lnIdDistritoNacimiento = 0
                          lnIdEstadoCivil = 0
                          'Graba Pacientes
                          wxConexionRed.BeginTrans
                          
                          
                                oCommand.CommandType = adCmdStoredProc
                                Set oCommand.ActiveConnection = wxConexionRed
                                oCommand.CommandTimeout = 150
                                oCommand.CommandText = "PacientesAgregarPorHistoriaClinica"
                                Set oParameter = oCommand.CreateParameter("@IdPaciente", adInteger, adParamOutput, 0, 0): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, UCase(lcApellidoPaterno)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, UCase(lcApellidoMaterno)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, UCase(lcPrimerNombre)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, lcSegundoNombre): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, lnTipoSexo): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, CDate(lcFechaNac)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnTipoNumeracion1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, lcAutogenerado): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, lnIdDistritoDomicilio): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(Not IsNull(.Fields!domicilio), Left(Trim(.Fields!domicilio), 50), "")): oCommand.Parameters.Append oParameter
                                If Not IsNull(.Fields!DNI) Then
                                   If Len(Trim(.Fields!DNI)) = 8 Then
                                        Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, Left(.Fields!DNI, 8)): oCommand.Parameters.Append oParameter
                                        Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                   Else
                                        Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, ""): oCommand.Parameters.Append oParameter
                                        Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, Null): oParameter.Attributes = adParamNullable: oCommand.Parameters.Append oParameter
                                   End If
                                Else
                                    Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, ""): oCommand.Parameters.Append oParameter
                                    Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, Null): oParameter.Attributes = adParamNullable: oCommand.Parameters.Append oParameter
                                End If
                                Set oParameter = oCommand.CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!nomPa), Left(.Fields!nomPa, 20), "")): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!nomMa), Left(.Fields!nomMa, 20), "")): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, ""): oCommand.Parameters.Append oParameter
                                lnIdPaciente = oCommand.Parameters("@IdPaciente")
                          Set oCommand = Nothing
                          Set oParameter = Nothing
                          
                        
                                oCommand.CommandType = adCmdStoredProc
                                Set .ActiveConnection = wxConexionRed
                                oCommand.CommandTimeout = 150
                                oCommand.CommandText = "HistoriasClinicasAgregarPorIdPaciente"
                                Set oParameter = oCommand.CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, CDate(lcFechaAnt)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnTipoNumeracion1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@HistoriaSistemaAnterior", adVarChar, adParamInput, 50, .Fields!codHis): oCommand.Parameters.Append oParameter
                                oCommand.Execute
                          Set oCommand = Nothing
                          Set oParameter = Nothing
                          
                          wxConexionRed.CommitTrans
                  End If
              End If
              .MoveNext
           Loop
       End With
       wxConexionJAMO.Close
       On Error Resume Next
       wrs_GalenHos2.Close
    End If
  
           
    With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "PacientesSeleccionarPorNroHistoriaClinicaTop1"
                Set wrs_GalenHos2 = .Execute
                Set wrs_GalenHos2.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
        
    lnNroHistoriaClinica = wrs_GalenHos2.Fields!NroHistoriaClinica
    wrs_GalenHos2.Close
        
    With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "GeneradorNroHistoriaClinicaActualizarNroHistoriaClinica"
                Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                Set wrs_GalenHos2 = .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    Unload Me
    Exit Sub
err_proceso:
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!HC & "     Paciente:" & wrs_LolCli.Fields!Paterno & " " & wrs_LolCli.Fields!Materno & " " & wrs_LolCli.Fields!Pnombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    Resume
    Unload Me

End Sub

Private Sub cmdProceVariasExcel_Click()
    Dim ml_EdadEnMeses As Long
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Dim W1 As Excel.Workbook
    Dim s1 As Excel.Worksheet
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim oRsExcelHojas As New Recordset
    Dim oFila As Long, oFila1 As Long, oCol1 As Integer, lnExcelHojas As Integer
    Dim lcExcel As String, lcHoja As String
    Dim lnColumMax As Integer
    Const lcRuta As String = "c:\barrantes\"
    'Carga a temporal
    With oRsExcelHojas
          .Fields.Append "Excel", adVarChar, 100
          .Fields.Append "Hoja", adVarChar, 100
          .LockType = adLockOptimistic
          .Open
    End With
    lcExcel = Dir(lcRuta & "*.*")
    Do While lcExcel <> ""
        Set W1 = EXL.Workbooks.Open(lcRuta & lcExcel)
        For oFila1 = 1 To W1.Sheets.Count
            lcHoja = W1.Sheets.Item(oFila1).Name
            oRsExcelHojas.AddNew
            oRsExcelHojas.Fields!Excel = lcExcel
            oRsExcelHojas.Fields!hoja = lcHoja
            oRsExcelHojas.Update
        Next
        W1.Close
        Set W1 = Nothing
        lcExcel = Dir
    Loop
'    Set W1 = EXL.Workbooks.Open("D:\excel.xls")
'    Set s1 = W1.Sheets("Hoja1")
'    oFila = 1
'    Do While True
'        If Val(s1.Cells(oFila, 1).Value) = 0 Then
'           Exit Do
'        End If
'        oRsExcelHojas.AddNew
'        oRsExcelHojas.Fields!Excel = s1.Cells(oFila, 1).Value
'        oRsExcelHojas.Fields!hoja = s1.Cells(oFila, 2).Value
'        oRsExcelHojas.Update
'        oFila = oFila + 1
'    Loop
'    W1.Close True
    'proceso
    lnExcelHojas = oRsExcelHojas.RecordCount
    If lnExcelHojas > 1 Then
       oRsExcelHojas.MoveFirst
       lcExcel = lcRuta & oRsExcelHojas.Fields!Excel
       lcHoja = oRsExcelHojas.Fields!hoja
       
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(lcExcel)
        oWorkBookPlantilla.Worksheets(lcHoja).Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        lnColumMax = oWorkSheet.Columns.Columns.Count
        For oCol1 = 1 To oWorkSheet.Columns.Columns.Count
           If oWorkSheet.Cells(1, oCol1).Value = "" Then
              lnColumMax = oCol1 - 1
              Exit For
           End If
        Next
       oFila1 = 2
       Do While True
            If oWorkSheet.Cells(oFila1, 1).Value = "" Then
               Exit Do
            End If
            oFila1 = oFila1 + 1
       Loop
'       Set W1 = EXL.Workbooks.Open(lcExcel)
'       Set s1 = W1.Sheets(lcHoja)
       '
'       lnColumMax = s1.Columns.Columns.Count
'       For oCol1 = 1 To s1.Columns.Columns.Count
'           If s1.Cells(1, oCol1).Value = "" Then
'              lnColumMax = oCol1 - 1
'              Exit For
'           End If
'       Next
       '
'       oFila1 = 2
'       Do While True
'            If s1.Cells(oFila1, 1).Value = "" Then
'               Exit Do
'            End If
'            oFila1 = oFila1 + 1
'       Loop
       oRsExcelHojas.MoveNext
       Do While Not oRsExcelHojas.EOF
            lcExcel = lcRuta & oRsExcelHojas.Fields!Excel
            lcHoja = oRsExcelHojas.Fields!hoja
            Set W = EXL.Workbooks.Open(lcExcel)
            Set s = W.Sheets(lcHoja)
            oFila = 2
            Do While True
                 If s.Cells(oFila, 1).Value = "" Then
                    Exit Do
                 End If
                 DoEvents
                 txtNroHojas.Text = Trim(Str(oRsExcelHojas.AbsolutePosition)) & "/" & _
                                    Trim(Str(lnExcelHojas)) & "/" & Trim(Str(oFila))
                 Me.Refresh
                 For oCol1 = 1 To lnColumMax
                     's1.Cells(oFila1, oCol1).Value = s.Cells(oFila, oCol1).Value
                     oWorkSheet.Cells(oFila1, oCol1).Value = s.Cells(oFila, oCol1).Value
                     
                 Next
                 s1.Cells(oFila1, lnColumMax + 1).Value = lcExcel & "-" & lcHoja
                 oFila1 = oFila1 + 1
                 oFila = oFila + 1
            Loop
            W.Close True
            oRsExcelHojas.MoveNext
       Loop
       oExcel.Visible = True
       oWorkSheet.PrintPreview
       
    Else
       MsgBox "Debe haber 2 filas al menos"
    End If
    oRsExcelHojas.Close
    W1.Close True
    '
    Set s = Nothing
    Set s1 = Nothing
    Set W = Nothing
    Set W1 = Nothing
    Set EXL = Nothing
    Set oExcel = Nothing
    Set oWorkBookPlantilla = Nothing
    Set oWorkBook = Nothing
    Set oWorkSheet = Nothing
    Unload Me
    Exit Sub
ErrRptHuelga:
    MsgBox Err.Description
'    Resume

End Sub

Private Sub cmdSanJuanAyacucho_Click()
    mo_ReglasAdmision.his_historicoAtencionesEliminarTodas wxConexionRed
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionJAMO.Open "dsn=HIS"
       Dim wrs_GalenHos As New ADODB.Recordset
       Dim wrs_GalenHos1 As New ADODB.Recordset
       Dim wrs_GalenHos2 As New ADODB.Recordset
       Dim wRsProblemas As New Recordset
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
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim lnIdEstadoCivil  As Long
       Dim lbNuevoHC As Boolean
       Dim lcApellidoPaterno As String, lcApellidoMaterno As String
       Dim lbContinuarProceso As Boolean
       Dim wFec1 As String, wFec2 As String, wFFecha As Date
       Dim lnTipoNumeracion1 As Long, lnFor As Integer, lnCantGuiones As Integer, lnHistoriaAutogenerada As Long
       lnTipoNumeracion1 = 2
       With wrs_LolCli
           wFFecha = CDate(Me.txtIniCSsb.Text)
           wFec1 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           wFFecha = CDate(Me.txtFinCSsb.Text)
           wFec2 = "date(" & Str(Year(wFFecha)) & "," & Str(Month(wFFecha)) & "," & Str(Day(wFFecha)) & ")"
           
           Me.txtIni1.Text = Me.txtIniCSsb.Text
           Me.txtFin1.Text = Me.txtFinCSsb.Text
           'elimina historias GalenHos de esas Fechas
           lblProcesando1.Caption = "Eliminando HC ya migradas, en GalenHos"
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "HistoriasClinicasSeleccionarPorFechaCreacion"
                Set oParameter = .CreateParameter("@FechaInicio", adVarChar, adParamInput, 20, Me.txtIni1.Text): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaFin", adVarChar, adParamInput, 20, Me.txtFin1.Text): .Parameters.Append oParameter
                Set wrs_GalenHos = .Execute
                Set wrs_GalenHos.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           Set oParameter = Nothing
           
           lntotReg = wrs_GalenHos.RecordCount
           If lntotReg > 0 Then
                wrs_GalenHos.MoveFirst
                ProgressBar2.Min = 0
                ProgressBar2.Max = lntotReg
                lnRegAct = 0
                Do While Not wrs_GalenHos.EOF
                   DoEvents
                   lnRegAct = lnRegAct + 1: ProgressBar2.Value = lnRegAct
                   Me.Refresh
                   wxConexionRed.BeginTrans
                   lnIdPaciente = wrs_GalenHos.Fields!idPaciente
                   
                   With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "HistoriasClinicasEliminar"
                         Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, wrs_GalenHos.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                         Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                         .Execute
                   End With
                   Set oCommand = Nothing
                   Set oParameter = Nothing
                   
                   
                   With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "PacientesEliminarPorIdPaciente"
                         Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
                         .Execute
                   End With
                   Set oCommand = Nothing
                   
                   wxConexionRed.CommitTrans
                   wrs_GalenHos.MoveNext
                Loop
                
           End If
           wrs_GalenHos.Close
           'Busca Historias LolCli de esas fechas, para añadirlos a GalenHos
           lblProcesando11.Caption = "Insertando HC en GalenHos"
            
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "HistoriasClinicasSeleccionarTodos"
                Set wrs_GalenHos = .Execute
                Set wrs_GalenHos.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "PacientesConsultarTodos"
                Set wrs_GalenHos1 = .Execute
                Set wrs_GalenHos1.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           
           lcSql = "select * from t_hiscli " & _
                    "  where  fec_exp>=" & wFec1 & " and fec_exp<=" & wFec2 & _
                     " order by fec_exp"
           .Open lcSql, wxConexionJAMO, adOpenKeyset, adLockOptimistic
           lntotReg = .RecordCount
           If lntotReg = 0 Then
              MsgBox "No hay registros en CS San Juan 2"
              wxConexionJAMO.Close
              Unload Me
              Exit Sub
           End If
           If chkFF2.Value = 1 Then
                    lnTipoNumeracion1 = 1
                    lnHistoriaAutogenerada = 1000001
                    
                    
                    With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = wxConexionRed
                         .CommandTimeout = 150
                         .CommandText = "PacientesSeleccionarPorIdTipoNumeracion1"
                         Set wRsProblemas = .Execute
                         Set wRsProblemas.ActiveConnection = Nothing
                    End With
                    Set oCommand = Nothing
                    
                    If wRsProblemas.RecordCount > 0 Then
                       lnHistoriaAutogenerada = wRsProblemas.Fields!NroHistoriaClinica + 1
                    End If
                    wRsProblemas.Close
                    .MoveFirst
                    Do While Not .EOF
                       If IsNull(.Fields!fichaf) Then
                          MsgBox "Existe FICHA FAMILIAR vacia, NO SE PODRA SEGUIR AÑADIENDO FICHAS FAMILIARES", vbCritical, "Mensaje"
                          Exit Sub
                       End If
                       lcSegundoNombre = Trim(.Fields!fichaf)
                       lnCantGuiones = 0
                       For lnFor = 1 To Len(lcSegundoNombre)
                           If Mid(lcSegundoNombre, lnFor, 1) = "-" Then
                              lnCantGuiones = lnCantGuiones + 1
                           End If
                       Next
                       If lnCantGuiones <> 2 Then
                          MsgBox "Ficha: " & .Fields!fichaf & Chr(13) & "Todas las Fichas Familiares deben tener 2 GUIONES" & Chr(13) & "El formato de la FICHA FAMILIAR es: sector-NumeroHistoria-NumeroIntegranteFamilia" & Chr(13) & " NO SE PODRA SEGUIR AÑADIENDO FICHAS FAMILIARES", vbCritical, "Mensaje"
                          Exit Sub
                       End If
                       .MoveNext
                    Loop
           End If
          
           
           With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = wxConexionRed
                  .CommandTimeout = 150
                  .CommandText = "lolcliProblemasHCSeleccionarTodos"
                  Set wRsProblemas = .Execute
                  Set wRsProblemas.ActiveConnection = Nothing
           End With
           Set oCommand = Nothing
           
           ProgressBar2.Min = 0
           ProgressBar2.Max = lntotReg
           .MoveFirst
           lnRegAct = 1
           Do While Not .EOF
              DoEvents
              ProgressBar2.Value = lnRegAct: lnRegAct = lnRegAct + 1
              Me.Refresh
              lbContinuarProceso = True
              If IsNull(.Fields!apepat) Or Trim(.Fields!apepat) = "" Or IsNull(.Fields!apemat) Or Trim(.Fields!apemat) = "" Or IsNull(.Fields!codHis) Or Trim(.Fields!codHis) = "" Then
                 lbContinuarProceso = False
              End If
              If lbContinuarProceso = True Then
                  lbNuevoHC = False
                  If Not sighentidades.EsFecha(Trim(.Fields!fecnac), "DD/MM/AAAA") Then
                        lcFechaNac = "01/01/1990"
                  Else
                        lcFechaNac = .Fields!fecnac
                  End If
                  lnNroHistoriaClinica = 0
                  If chkFF2.Value = 1 Then
                     lnNroHistoriaClinica = lnHistoriaAutogenerada
                     lnHistoriaAutogenerada = lnHistoriaAutogenerada + 1
                  Else
                      lnNroHistoriaClinica = Val(.Fields!codHis)
                  End If
                  If lnNroHistoriaClinica = 0 Then
                     'Historia clinica con problemas
                     lnNroHistoriaClinica = SoloNumerosDeHC(.Fields!codHis)
    '                 lbNuevoHC = True
                  End If
                  If IsNull(.Fields!nombre) Then
                      lcPrimerNombre = "NN"
                      lcSegundoNombre = ""
                  Else
                      lcPrimerNombre = Trim(.Fields!nombre)
                      lcPrimerNombre = Left(RetornaPrimerNombre(lcPrimerNombre), 20)
                      lcSegundoNombre = Trim(.Fields!nombre)
                      lcSegundoNombre = Left(RetornaSegundoNombre(lcSegundoNombre), 20)
                  End If
                  
                  lnTipoSexo = IIf(UCase(.Fields!sexo) = "M", 1, 2)
                  If lnNroHistoriaClinica = 0 Then
                     lbNuevoHC = True
                  Else
                        'Busca si ya existe Nro Historia en GalenHos
                       
                        
                        With oCommand
                               .CommandType = adCmdStoredProc
                               Set .ActiveConnection = wxConexionRed
                               .CommandTimeout = 150
                               .CommandText = "PacientesSeleccionarPorNroHistoriaClinica"
                               Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                               Set wrs_GalenHos2 = .Execute
                               Set wrs_GalenHos2.ActiveConnection = Nothing
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        
                        If wrs_GalenHos2.RecordCount > 0 Then
                           lbNuevoHC = True
                        End If
                  End If
                  If lbNuevoHC = True Then
                          'HC repetida
                          'solo graba como problemas
                          '
                              oCommand.CommandType = adCmdStoredProc
                              Set oCommand.ActiveConnection = wxConexionRed
                              oCommand.CommandTimeout = 150
                              oCommand.CommandText = "lolcliProblemasHCAgregar"
                              Set oParameter = oCommand.CreateParameter("@pacHis", adChar, adParamInput, 50, .Fields!codHis): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacPat", adChar, adParamInput, 30, .Fields!apepat): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacMat", adChar, adParamInput, 30, .Fields!apemat): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacNam", adChar, adParamInput, 50, .Fields!nombre): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@pacFin", adDBTimeStamp, adParamInput, 0, CDate(Me.txtFIniCSN.Text)): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@nroHistoriaGalenHos", adVarChar, adParamInput, 50, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                              Set oParameter = oCommand.CreateParameter("@autogeneradoGalenHos", adVarChar, adParamInput, 50, "*" & Trim(Str(lnRegAct)) & "HcRep-SanJuanB"): oCommand.Parameters.Append oParameter
                              oCommand.Execute
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                          
                          
                          If lnNroHistoriaClinica <> 0 Then

                            With oCommand
                                .CommandType = adCmdStoredProc
                                Set .ActiveConnection = wxConexionRed
                                .CommandTimeout = 150
                                .CommandText = "lolcliProblemasHCAgregarSinNroHistoriaClinica"
                                Set oParameter = .CreateParameter("@pacHis", adChar, adParamInput, 50, wrs_GalenHos2.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacPat", adChar, adParamInput, 30, wrs_GalenHos2.Fields!ApellidoPaterno): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacMat", adChar, adParamInput, 30, wrs_GalenHos2.Fields!ApellidoMaterno): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@pacNam", adChar, adParamInput, 50, Left(Trim(wrs_GalenHos2.Fields!PrimerNombre) & " " & wrs_GalenHos2.Fields!SegundoNombre, 50)): .Parameters.Append oParameter
                                Set oParameter = .CreateParameter("@autogeneradoGalenHos", adVarChar, adParamInput, 50, "*" & Trim(Str(lnRegAct)) & "YaMigradoEnGalenHos"): .Parameters.Append oParameter
                                .Execute
                            End With
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                            '
                            wrs_GalenHos2.Close
                          End If
                  Else
                          wrs_GalenHos2.Close
                          lcApellidoPaterno = UCase(Left(Trim(.Fields!apepat), 20))
                          lcApellidoMaterno = UCase(Left(Trim(.Fields!apemat), 20))
                          lcPrimerNombre = UCase(lcPrimerNombre)
                          lcAutogenerado = PacienteCrearNroAutogenerado1(lcFechaNac, lcApellidoPaterno, lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                          lcFechaAnt = .Fields!FEC_EXP
                          'Busca en Tabla xx Equivalencia LolCli
                          lntipoOcupacion = 0
                          lnIdDepartamentoDomicilio = 0
                          LnIdProvinciaDomicilio = 0
                          lnIdDistritoDomicilio = 0
'                          If Not IsNull(.Fields!codgeo) Then
'                            lnIdDepartamentoDomicilio = Val(Left(.Fields!codgeo, 2))
'                            LnIdProvinciaDomicilio = Val(Left(.Fields!codgeo, 4))
'                            lnIdDistritoDomicilio = Val(.Fields!codgeo)
'                          End If
                          lnIdDepartamentoNacimiento = 0
                          LnIdProvinciaNacimiento = 0
                          lnIdDistritoNacimiento = 0
                          lnIdEstadoCivil = 0
'                          If Not IsNull(.Fields!estCivil) Then
'                             Select Case UCase(Left(.Fields!estCivil, 1))
'                             Case "S"   'soltero
'                                  lnIdEstadoCivil = 2
'                             Case "C"   'casado
'                                  lnIdEstadoCivil = 1
'                             End Select
'                          End If
                          'Graba Pacientes
                          wxConexionRed.BeginTrans
                          
                          
                                oCommand.CommandType = adCmdStoredProc
                                Set oCommand.ActiveConnection = wxConexionRed
                                oCommand.CommandTimeout = 150
                                oCommand.CommandText = "PacientesAgregarPorHistoriaClinica"
                                Set oParameter = oCommand.CreateParameter("@IdPaciente", adInteger, adParamOutput, 0, 0): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, lcApellidoPaterno): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, lcApellidoMaterno): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, lcPrimerNombre): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, lcSegundoNombre): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, lnTipoSexo): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, CDate(lcFechaNac)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnTipoNumeracion): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, lcAutogenerado): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, lnIdDistritoDomicilio): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(Not IsNull(.Fields!domicilio), Left(Trim(.Fields!domicilio), 50), "")): oCommand.Parameters.Append oParameter
                                If Not IsNull(.Fields!DNI) Then
                                   If Len(Trim(.Fields!DNI)) = 8 Then
                                        Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, Left(.Fields!DNI, 8)): oCommand.Parameters.Append oParameter
                                        Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                   Else
                                        Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, ""): oCommand.Parameters.Append oParameter
                                        Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, Null): oParameter.Attributes = adParamNullable: oCommand.Parameters.Append oParameter
                                   End If
                                Else
                                    Set oParameter = oCommand.CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, ""): oCommand.Parameters.Append oParameter
                                    Set oParameter = oCommand.CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, Null): oParameter.Attributes = adParamNullable: oCommand.Parameters.Append oParameter
                                End If
                                Set oParameter = oCommand.CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!nomPa), Left(.Fields!nomPa, 20), "")): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!nomMa), Left(.Fields!nomMa, 20), "")): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, 166): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, IIf(Not IsNull(.Fields!fichaf), Left(.Fields!fichaf, 20), "")): oCommand.Parameters.Append oParameter
                                oCommand.Execute
                                lnIdPaciente = oCommand.Parameters("@IdPaciente")
                          Set oCommand = Nothing
                          Set oParameter = Nothing
                          
                          'Graba HistoriasClinicas
                          
                        
                                oCommand.CommandType = adCmdStoredProc
                                Set oCommand.ActiveConnection = wxConexionRed
                                oCommand.CommandTimeout = 150
                                oCommand.CommandText = "HistoriasClinicasAgregarPorIdPaciente"
                                Set oParameter = oCommand.CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, CDate(lcFechaAnt)): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lnTipoNumeracion1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 0, 1): oCommand.Parameters.Append oParameter
                                Set oParameter = oCommand.CreateParameter("@HistoriaSistemaAnterior", adVarChar, adParamInput, 50, .Fields!codHis): oCommand.Parameters.Append oParameter
                                oCommand.Execute
                          Set oCommand = Nothing
                          Set oParameter = Nothing
                          wxConexionRed.CommitTrans
                  End If
              End If
              .MoveNext
           Loop
       End With
       wxConexionJAMO.Close
       On Error Resume Next
       wrs_GalenHos2.Close
    End If
      
           
    With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "PacientesSeleccionarPorNroHistoriaClinicaTop1"
                Set wrs_GalenHos2 = .Execute
                Set wrs_GalenHos2.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
        
    lnNroHistoriaClinica = wrs_GalenHos2.Fields!NroHistoriaClinica
    wrs_GalenHos2.Close
        
    With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "GeneradorNroHistoriaClinicaActualizarNroHistoriaClinica"
                Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                Set wrs_GalenHos2 = .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    cmdCargaAtencionesHist_Click
    
    Unload Me
    Exit Sub
err_proceso:
    MsgBox "         Procesó hasta " & lcFechaAnt & Chr(13) & " " & Chr(13) & " " & Chr(13) & " " & Chr(13) & "Fallo en HC: " & wrs_LolCli.Fields!HC & "     Paciente:" & wrs_LolCli.Fields!Paterno & " " & wrs_LolCli.Fields!Materno & " " & wrs_LolCli.Fields!Pnombre & Chr(13) & " " & Chr(13) & " " & Chr(13) & Err.Description
    lcFechaAnt = lcFechaAnt - 1
    wxConexionRed.RollbackTrans
    Resume
    Unload Me

End Sub

Private Sub Command1_Click()
   cmdProblemasJamo_Click
End Sub



Private Sub Command2_Click()
        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long
        Dim lnId As Long
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=Sicuani"
        '
        Me.MousePointer = 1
        lcSql = "select * from tformatodet"
        oRsTmpOpc.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsTmpOpc.MoveFirst
        Do While Not oRsTmpOpc.EOF
           If Not IsNull(oRsTmpOpc.Fields!FEC_EXP) Then
           If Val(Right(oRsTmpOpc.Fields!FEC_EXP, 2)) > 28 And Val(Mid(oRsTmpOpc.Fields!FEC_EXP, 5, 2)) = 2 Then
              lcSql = "28/" & Mid(oRsTmpOpc.Fields!FEC_EXP, 5, 2) & "/" & Left(oRsTmpOpc.Fields!FEC_EXP, 4)
           Else
              lcSql = Right(oRsTmpOpc.Fields!FEC_EXP, 2) & "/" & Mid(oRsTmpOpc.Fields!FEC_EXP, 5, 2) & "/" & Left(oRsTmpOpc.Fields!FEC_EXP, 4)
           End If
           
           oRsTmpOpc.Fields!fec_Exp1 = CDate(lcSql)
           oRsTmpOpc.Update
           End If
           oRsTmpOpc.MoveNext
        Loop
        oRsTmpOpc.Close
        MsgBox "Termino el proceso....vea el GRID con errores"
 Unload Me

End Sub

Private Sub Command3_Click()

     
    'mo_ReglasAdmision.his_historicoAtencionesEliminarTodas wxConexionRed
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    

    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       wxConexionRed.Open "dsn=GALENHOS"

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
             ProgressBar2.Max = wrs_GalenHos2.RecordCount + 1
             ProgressBar2.Min = 1
             lnFor = 1
             wrs_GalenHos2.MoveFirst
             Do While Not wrs_GalenHos2.EOF
                DoEvents: ProgressBar2.Value = ProgressBar2.Value + 1: Me.Refresh
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

Private Sub cmdTodosPuntosCarga_Click()
    If Val(txtTipoServicio.Text) > 0 And Val(txtTipoServicio.Text) < 4 Then
       Me.MousePointer = 11
       Dim oConexODBC As New Connection
       Dim oRsTmp1 As New Recordset
       Dim oRsPtos As New Recordset
       Dim oRsProductos As New Recordset
       Dim oCommand As New ADODB.Command
       Dim oParameter As ADODB.Parameter
       Dim lbEsNuevo As Boolean, lcSql As String, lnPqte1 As Long
       oConexODBC.CommandTimeout = 900
       oConexODBC.CursorLocation = adUseClient
       oConexODBC.Open sighentidades.CadenaConexion   ' "dsn=GALENHOS"
       lcSql = "select  idproducto from FactCatalogoServiciosHosp where idtipoFinanciamiento=1"
       oRsProductos.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
       
'       lcSql = "select * from FactCatalogoServiciosPtos"
'       oRsPtos.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
       
       lcSql = "SELECT     dbo.Servicios.IdServicio, dbo.Servicios.Nombre, dbo.FactPuntosCarga.IdPuntoCarga, " & _
               "            dbo.Servicios.IdTipoServicio " & _
               "  FROM         dbo.Servicios LEFT OUTER JOIN " & _
               "            dbo.FactPuntosCarga ON dbo.Servicios.IdServicio = dbo.FactPuntosCarga.idServicio" & _
               "  WHERE dbo.Servicios.IdTipoServicio=" & Trim(txtTipoServicio.Text)
       'lcSql = "select * from factPuntosCarga where idPuntoCarga=1"
       oRsTmp1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 And oRsProductos.RecordCount > 0 Then
          lblTotPtos.Caption = Trim(Str(oRsTmp1.RecordCount)) & "/"
          lnPqte1 = 1
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             DoEvents
             lblPto.Caption = lnPqte1
             Me.Refresh
             lnPqte1 = lnPqte1 + 1
             oRsProductos.MoveFirst
             Do While Not oRsProductos.EOF
                If Not IsNull(oRsTmp1!idPuntoCarga) And Not IsNull(oRsTmp1!idProducto) Then
                    lbEsNuevo = True
                    lcSql = "select * from FactCatalogoServiciosPtos where idPuntoCarga=" & oRsTmp1!idPuntoCarga & _
                            " and idProducto=" & oRsProductos!idProducto
                    oRsPtos.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                    If oRsPtos.RecordCount > 0 Then
                       lbEsNuevo = False
                    End If
                    oRsPtos.Close
                    If lbEsNuevo = True Then
                        With oCommand
                           .CommandType = adCmdStoredProc
                           Set .ActiveConnection = oConexODBC
                           .CommandTimeout = 150
                           .CommandText = "FactCatalogoServiciosPtosAgregar"
                           Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsTmp1!idPuntoCarga): .Parameters.Append oParameter
                           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsProductos!idProducto): .Parameters.Append oParameter
                           Set oParameter = .CreateParameter("@EsPreVenta", adBoolean, adParamInput, 0, 0): .Parameters.Append oParameter
                           .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                    End If
                End If
              
'                lbEsNuevo = True
'                If oRsTmp2.RecordCount > 0 Then
'
'                    oRsTmp2.MoveFirst
'                    Do While Not oRsTmp2.EOF
'                       If oRsTmp2!IdPuntoCarga = oRsTmp1!IdPuntoCarga And oRsTmp2!idProducto = oRsTmp3!idProducto Then
'                          lbEsNuevo = False
'                       End If
'                       oRsTmp2.MoveNext
'                    Loop
'                    If lbEsNuevo = True Then
'                        With oCommand
'                           .CommandType = adCmdStoredProc
'                           Set .ActiveConnection = oConexODBC
'                           .CommandTimeout = 150
'                           .CommandText = "FactCatalogoServiciosPtosAgregar"
'                           Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsTmp1!IdPuntoCarga): .Parameters.Append oParameter
'                           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmp3!idProducto): .Parameters.Append oParameter
'                           Set oParameter = .CreateParameter("@EsPreVenta", adBoolean, adParamInput, 0, 0): .Parameters.Append oParameter
'                           .Execute
'                        End With
'                        Set oCommand = Nothing
'                        Set oParameter = Nothing
'
'                    End If
'                End If
                oRsProductos.MoveNext
             Loop
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       Me.MousePointer = 1
       oConexODBC.Close
       Set oConexODBC = Nothing
       Set oRsTmp1 = Nothing
       Set oRsPtos = Nothing
       Set oRsProductos = Nothing
       Unload Me
    End If
     
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    txtCtaInicial.Text = 0 'DevuelveUltimaCuentaEnParametros
    
    lnTipoNumeracion = 2
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    'Ayacucho
    List1.AddItem "1-Tablas de GalenHos, que deben tener CODIGO de LolCli:"
    List1.AddItem "  LolCliUbigeo, TiposOcupacion, TiposEstadoCivil       "
    List1.AddItem ""
    List1.AddItem "2-La Tabla LolCliUbigeo debe contener todos los UBIGEOS"
    List1.AddItem "  usados en la Filiación del Sistema LolCli (Lugar de  "
    List1.AddItem "  de Nacimiento y Lugar de Domicilio)                  "
    List1.AddItem "3-Se podrá ejecutar el PROCESO DE MIGRACION DE PACIENTES"
    List1.AddItem ""
    List1.AddItem "4-Cuando se esta Procesando y LolCli tiene Nro Historias"
    List1.AddItem "  repetidas:"
    List1.AddItem "  4.1-El Sistema grabará como esta la Primera, con todos"
    List1.AddItem "      sus demas datos."
    List1.AddItem "  4.2-Para las siguientes, el sistema grabará:"
    List1.AddItem "      -El 'Nro Historia' Clinica tendrá una longitud de 8"
    List1.AddItem "       numeros, comenzando con '9' y el mismo Nro Historia"
    List1.AddItem "       que grabaron en LolCli."
    List1.AddItem "      -El 'Segundo Nombre' se antepondrá un PUNTO, seguido"
    List1.AddItem "       del resto del segundo Nombre (por el autogenerado, dos"
    List1.AddItem "       pacientes con los mismos Apellidos y Nombres)."
    List1.AddItem "      -Los demas datos se grabarán iguales (sin cambios)"
    List1.AddItem ""
    List1.AddItem "5-Existe una columna 'HistoriasClinicas.HistoriaSistemaAnterior'"
    List1.AddItem "  donde se grabá 'Nro Historia' tal y como esta grabada en el LolCli"
    txtFechaIni.Text = Date
    txtFechaFin.Text = Date
    'Tumbes
    List2.AddItem "Consideraciones:                                 "
    List2.AddItem "                                                 "
    List2.AddItem "1-Crear ODBC llamado 'JAMO' que apunte a la Base "
    List2.AddItem "  de datos BDHcHospitalJAMO.                     "
    List2.AddItem "2-Las tablas: Pacientes, HistoriasClinicas deberán"
    List2.AddItem "  estar VACIAS (solo al ejecutar este proceso por"
    List2.AddItem "  primera vez).                                  "
    txtIni1.Text = Date
    txtFin1.Text = Date
    List3.AddItem "Consideraciones:                                 "
    List3.AddItem "                                                 "
    List3.AddItem "1-Crear ODBC llamado 'JAMO' que apunte a la Base "
    List3.AddItem "  de datos BDHcHospitalJAMO.                     "
    List3.AddItem "2-Debe existir el IdParametro=273 (Servidor externo,"
    List3.AddItem "  BD externa)                                       "
    List3.AddItem "3-Migrar antes los Datos Personales de cada Paciente"
    List3.AddItem "  (dicha tabla deberá estar pulida).              "
    Me.txtFCita1.Text = Date
    Me.txtFCita2.Text = Date
    'Cuzco-Sicuani
    List4.AddItem "Consideraciones:                                 "
    List4.AddItem "                                                 "
    List4.AddItem "1-Crear ODBC llamado 'HIS' que apunte a la Base "
    List4.AddItem "  de datos 'HistClinic.dbf'.                     "
    List4.AddItem "   *  Microsoft Visual Foxpro Driver"
    List4.AddItem "   *  Directorio de Tabla Libre                                         "
    List4.AddItem "   *  c:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
    List4.AddItem "      histclinic.dbf         "
    Me.txtCuzcoF1.Text = Date
    Me.txtCuzcoF2.Text = Date

    'CS Nazareno -Ayacucho
    If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3575 Then
       Me.SSTab1.Tab = 4
    End If
    List6.AddItem "Consideraciones:                                 "
    List6.AddItem "                                                 "
    List6.AddItem "1-Crear ODBC llamado 'HIS' que apunte a la Base "
    List6.AddItem "  de datos 'tarjeta.dbf'.                     "
    List6.AddItem "   *  Microsoft Visual Foxpro Driver"
    List6.AddItem "   *  Directorio de Tabla Libre                                         "
    List6.AddItem "   *  c:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
    List6.AddItem "      tarjeta.dbf         "
    List6.AddItem "   *Campos:"
    List6.AddItem " fec_ins, date        (fecha que se registro el Paciente)         "
    List6.AddItem " ape_pat, string,20       (Apellido Paterno)         "
    List6.AddItem " ape_mat, string,20       (Apellido Materno)           "
    List6.AddItem " nro_his, string,9       (Numero de Historia)         "
    List6.AddItem " fec_Nac, date         (Fecha de Nacimiento)           "
    List6.AddItem " Nombres1, string,20       (Primer Nombre)         "
    List6.AddItem " Nombres2, string,20       (Segundo Nombre)         "
    List6.AddItem " Sexo, string          (M->masculino)        "
    List6.AddItem " dom_act, string,50       (Domicilio actual,dato no obligatorio )        "
    List6.AddItem " DNI, string,8            (DNI,dato no obligatorio )        "
    List6.AddItem " edad, string             (edad,dato no obligatorio )        "
    List6.AddItem " nomPa, string,20         (Nombre del Padre,dato no obligatorio)         "
    List6.AddItem " nomMa, string,20    (Nombre de la Madre,dato no obligatorio)         "
    List6.AddItem " FichaF,string,20    (Ficha Familiar: Sector-N°Historia-N°EnFamilia )(no obligatorio)(el campo 'nro_his' debe estar VACIO)(pulsar clic en CHECK)"
    txtFIniCSN.Text = Date
    txtFFinCSN.Text = Date
    'CS San Juan Bautista -Ayacucho
    If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3598 Then
       Me.SSTab1.Tab = 5
    End If
    List7.AddItem "Consideraciones:                                 "
    List7.AddItem "                                                 "
    List7.AddItem "1-Crear ODBC llamado 'HIS' que apunte a la Base "
    List7.AddItem "  de datos 't_hiscli.dbf'.                     "
    List7.AddItem "   *  Microsoft Visual Foxpro Driver"
    List7.AddItem "   *  Directorio de Tabla Libre                                         "
    List7.AddItem "   *  c:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
    List7.AddItem "      t_hiscli.dbf         "
    List7.AddItem "******Campos:"
    List7.AddItem " Fec_exp, date        (fecha que se registro el Paciente)         "
    List7.AddItem " apepat, string,20       (Apellido Paterno)         "
    List7.AddItem " apemat, string,20       (Apellido Materno)           "
    List7.AddItem " codhis, string,9       (Numero de Historia)         "
    List7.AddItem " fecNac, date         (Fecha de Nacimiento)           "
    List7.AddItem " Nombre, string,40       (Primer y segundo Nombre..separados por BLANCOS)         "
    List7.AddItem " Sexo, string         (M->masculino)        "
    List7.AddItem " Domicilio, string,50    (Donde vive el Paciente, dato no obligatorio)         "
    List7.AddItem " DNI, string,8          (Numero de DNI, dato no obligatorio)          "
    List7.AddItem " nomPa, string,20       (Nombre del Padre,dato no obligatorio)         "
    List7.AddItem " nomMa, string,20  (Nombre de la Madre,dato no obligatorio)         "
    List7.AddItem " FichaF,string,20  (Ficha Familiar: Sector-N°Historia-N°EnFamilia )(no obligatorio)(el campo 'codHis' debe estar VACIO)(pulsar clic en CHECK)"
    Me.txtIniCSsb.Text = Date
    Me.txtFinCSsb.Text = Date
    'CS santa elena -Ayacucho
    If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3602 Then
       Me.SSTab1.Tab = 6
    End If
    List8.AddItem "Consideraciones:                                 "
    List8.AddItem "                                                 "
    List8.AddItem "1-Crear ODBC llamado 'HIS' que apunte a la Base "
    List8.AddItem "  de datos 't_hiscli.dbf'.                     "
    List8.AddItem "   *  Microsoft Visual Foxpro Driver"
    List8.AddItem "   *  Directorio de Tabla Libre                                         "
    List8.AddItem "   *  c:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
    List8.AddItem "      t_hiscli.dbf         "
    List8.AddItem "   *Campos:"
    List8.AddItem "      Fec_exp, date        (fecha que se registro el Paciente)         "
    List8.AddItem "      apepat, string,20       (Apellido Paterno)         "
    List8.AddItem "      apemat, string,20       (Apellido Materno)           "
    List8.AddItem "      codhis, string,9       (Numero de Historia)         "
    List8.AddItem "      fecNac, date         (Fecha de Nacimiento)           "
    List8.AddItem "      Nombre, string,40       (Primer y segundo Nombre..separados por BLANCOS)         "
    List8.AddItem "      Sexo, string         (M->masculino)        "
    List8.AddItem "      Domicilio, string,50    (Donde vive el Paciente, dato no obligatorio)         "
    List8.AddItem "      DNI, string,8          (Numero de DNI), dato no obligatorio           "
    List8.AddItem "      nomPa, string,20       (Nombre del Padre,dato no obligatorio)         "
    List8.AddItem "      nomMa, string,20       (Nombre de la Madre,dato no obligatorio)         "
    List8.AddItem "      "
    Me.txtINIse.Text = Date
    Me.txtFINse.Text = Date
    'HRC -Cajamarca
    If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3602 Then
       Me.SSTab1.Tab = 6
    End If
    List9.AddItem "Consideraciones:                                 "
    List9.AddItem "                                                 "
    List9.AddItem "1-Crear ODBC llamado 'dbhospi' que apunte a la Base "
    List9.AddItem "  de datos 'dbhospi'.                     "
    List9.AddItem "   *  Sql Server"
    List9.AddItem "   *Campos:"
    List9.AddItem "      Fec_exp, date        (fecha que se registro el Paciente)         "
    List9.AddItem "      apepat, string,20       (Apellido Paterno)         "
    List9.AddItem "      apemat, string,20       (Apellido Materno)           "
    List9.AddItem "      codhis, string,9       (Numero de Historia)         "
    List9.AddItem "      fecNac, date         (Fecha de Nacimiento)           "
    List9.AddItem "      Nombre, string,40       (Primer y segundo Nombre..separados por BLANCOS)         "
    List9.AddItem "      Sexo, string         (M->masculino)        "
    List9.AddItem "      Domicilio, string,50    (Donde vive el Paciente, dato no obligatorio)         "
    List9.AddItem "      DNI, string,8          (Numero de DNI), dato no obligatorio           "
    List9.AddItem "2-Salir del Sistema de Admisión actual "
    List9.AddItem "3-Ejecutar icono windows 'CopiaBD' antes de ejecutar este proceso   "
    List9.AddItem "      "
    Me.Text1.Text = Date
    Me.Text2.Text = Date
    'Carga Datos del EESS
    txtnombre.Text = lcBuscaParametro.SeleccionaFilaParametro(205)
    txtdireccion.Text = lcBuscaParametro.SeleccionaFilaParametro(206)
    txttelefono.Text = lcBuscaParametro.SeleccionaFilaParametro(207)
    txtubigeo.Text = lcBuscaParametro.SeleccionaFilaParametro(242)
    txtcodminsa.Text = lcBuscaParametro.SeleccionaFilaParametro(208)
    txtcoddisa.Text = lcBuscaParametro.SeleccionaFilaParametro(239)
    txtcodred.Text = lcBuscaParametro.SeleccionaFilaParametro(240)
    txtcodmicrored.Text = lcBuscaParametro.SeleccionaFilaParametro(241)
    txtcodHIS.Text = lcBuscaParametro.SeleccionaFilaParametro(243)
    txtcodrenaes.Text = lcBuscaParametro.SeleccionaFilaParametro(280)
    chkCentroSalud.Value = IIf(Trim(lcBuscaParametro.SeleccionaFilaParametro(282)) = "S", 1, 0)
    Me.txtSisDisa.Text = lcBuscaParametro.SeleccionaFilaParametro(310)
    Me.txtSisCatEESS.Text = lcBuscaParametro.SeleccionaFilaParametro(303)
    Me.txtSISptoDigitacion.Text = lcBuscaParametro.SeleccionaFilaParametro(304)
    Me.txtSISCodigoUDR.Text = lcBuscaParametro.SeleccionaFilaParametro(305)
    
    Dim oSisConsumoWeb As New SIGHNegocios.SisConsumoWeb
    Set oRsGrdSIS = oSisConsumoWeb.m_eessSelecionarXcodigoRenaes(Right("0000000000" & txtcodrenaes.Text, 10))
    Me.dgrMuestra.Visible = False
    Me.grdSIS.Visible = True
    Set Me.grdSIS.DataSource = oRsGrdSIS
    '
    txtAnioProc.Text = Year(Date)
    '
    generaYcargaCartillas
End Sub

Sub generaYcargaCartillas()
   If oRsCartillas.State = 1 Then oRsCartillas.Close
   With oRsCartillas
       .Fields.Append "Nro", adInteger
       .Fields.Append "ganador", adVarChar, 1
       .Fields.Append "fijo", adVarChar, 1
       .LockType = adLockOptimistic
       .Open
       .AddNew
       .Fields!nro = 1
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 2
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 3
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 4
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 5
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 6
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 7
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 8
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 9
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 10
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       
       .Fields!nro = 11
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 12
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 13
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       .AddNew
       .Fields!nro = 14
       .Fields!ganador = "L"
       .Fields!fijo = " "
       .Update
       End With
   Set grdCartillas.DataSource = oRsCartillas
End Sub


Function PacienteCrearNroAutogenerado1(lcFechaNacimiento As String, lcApellidoPaterno As String, _
                                       lcApellidoMaterno As String, lcPrimerNombre As String, _
                                       lcSegundoNombre As String, lnIdTipoSexo As Long)
    Dim oDOPaciente As New DOPaciente
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    oDOPaciente.ApellidoPaterno = lcApellidoPaterno
    oDOPaciente.ApellidoMaterno = lcApellidoMaterno
    oDOPaciente.PrimerNombre = lcPrimerNombre
    oDOPaciente.SegundoNombre = lcSegundoNombre
    oDOPaciente.idTipoSexo = lnIdTipoSexo
    oDOPaciente.FechaNacimiento = CDate(lcFechaNacimiento)
    
    PacienteCrearNroAutogenerado1 = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDOPaciente)
    
    Set oDOPaciente = Nothing
    Set mo_AdminAdmision = Nothing

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
        
End Function

Function Modulo10(sValor As String) As Integer
Dim sTemp As String
Dim I As Integer
Dim k As Integer
Dim iTotal As Integer

    sTemp = ""
    
    For I = 1 To Len(sValor)
        If IsNumeric(Mid(sValor, I, 1)) Then
            sTemp = sTemp + Mid(sValor, I, 1)
        Else
            sTemp = sTemp + DevuelveValorEnNumeros(Mid(sValor, I, 1))
        End If
    Next I

    'Acumula total de digitos
    iTotal = 0
    For I = 1 To Len(sTemp)
        If I Mod 2 <> 0 Then
            k = CInt(Mid(sTemp, I, 1)) * 2
            iTotal = iTotal + (k - (k Mod 10)) / 10 + (k Mod 10)
        Else
            iTotal = iTotal + CInt(Mid(sTemp, I, 1))
        End If
    Next I

    If (iTotal Mod 10) = 0 Then
        Modulo10 = 0
    Else
        Modulo10 = 10 - (iTotal Mod 10)
    End If



End Function
Function DevuelveValorEnNumeros(sCaracter As String) As String

    Select Case sCaracter
    Case "A" To "N"
        DevuelveValorEnNumeros = Asc(sCaracter) - 55
    Case "Ñ"
        DevuelveValorEnNumeros = 24
    Case "O" To "Z"
        DevuelveValorEnNumeros = Asc(sCaracter) - 54
    End Select

End Function

Sub DevuelvePrimeryCuartoCaracter(sPalabra As String, C1 As String, C2 As String)
Dim sTemp As String
        If sPalabra <> "" Then
            sTemp = ObtenerUltimaPalabra(EliminarConjunciones(sPalabra))
            C1 = Left(sTemp, 1)
            C2 = DevuelveCuartoCaracter(sTemp)
        Else
            C1 = "X"
            C2 = "X"
        End If
End Sub

Function DevuelveCuartoCaracter(sPalabra) As String
    If Len(sPalabra) <= 4 Then
        DevuelveCuartoCaracter = Right(sPalabra, 1)
    Else
        DevuelveCuartoCaracter = Mid(sPalabra, 4, 1)
    End If
End Function

Function ObtenerUltimaPalabra(sTexto As String) As String
Dim p As String
Dim iUltBlanco As Integer
Dim sTemp As String


    sTemp = Trim(sTexto)

    p = InStr(sTemp, " ")
    iUltBlanco = 0
    Do While p > 0
        iUltBlanco = p
        p = InStr(p + 1, sTemp, " ")
    Loop
    If iUltBlanco > 0 Then
        ObtenerUltimaPalabra = Mid(sTemp, iUltBlanco + 1)
    Else
        ObtenerUltimaPalabra = sTemp
    End If
End Function

Function EliminarConjunciones(sPalabra As String)
Dim sTemp As String

        sTemp = ReemplazarCadena(sPalabra, " DE ", " ")
        sTemp = ReemplazarCadena(sTemp, " DEL ", " ")
        sTemp = ReemplazarCadena(sTemp, " EL ", " ")
        sTemp = ReemplazarCadena(sTemp, " LA ", " ")
        sTemp = ReemplazarCadena(sTemp, " LOS ", " ")
        sTemp = ReemplazarCadena(sTemp, " LAS ", " ")

        EliminarConjunciones = sTemp

End Function

Function ReemplazarCadena(sOriginal As String, sCadenaA As String, sCadenaR As String) As String
Dim sTemp As String
Dim lLng As Long
Dim lP As Long

    sTemp = sOriginal
    lLng = Len(sCadenaA)
    lP = InStr(sTemp, sCadenaA)
    
    Do While lP <> 0
        sTemp = Left(sTemp, lP - 1) + sCadenaR + Mid(sTemp, lP + lLng)
        lP = InStr(sTemp, sCadenaA)
    Loop

    ReemplazarCadena = sTemp
End Function

Function RetornaPrimerNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer
    RetornaPrimerNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       RetornaPrimerNombre = Trim(Left(lcPrimerSegundoNombreJuntos, ln))
    Else
       RetornaPrimerNombre = lcPrimerSegundoNombreJuntos
    End If
End Function

Function RetornaSegundoNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer
    RetornaSegundoNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       RetornaSegundoNombre = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
    Else
       RetornaSegundoNombre = ""
    End If
End Function

Function generaNuevaNroHistoria(lnActualHistoria As Long) As Long
     Dim lcId As String
     lcId = "9" + Right("0" & Trim(Str(Second(Time))), 1) + "000000"
     generaNuevaNroHistoria = Val(lcId) + lnActualHistoria
End Function

Function SoloNumerosDeHC(lcHClolCli As String) As Long
    Dim SoloNumero As String
    Dim ln As Integer
    SoloNumerosDeHC = 0
    SoloNumero = ""
    For ln = 1 To Len(lcHClolCli)
        If InStr("1234567890", Mid(lcHClolCli, ln, 1)) > 0 Then
           SoloNumero = SoloNumero + Mid(lcHClolCli, ln, 1)
        End If
    Next
    SoloNumerosDeHC = Val(SoloNumero)
End Function




Private Sub cmdCargaAtencionesHist_Click()
        If oRsGrdSIS.RecordCount > 0 Then
           If Me.txtSisCatEESS.Text = "" Or Me.txtSisDisa.Text = "" Or Me.txtSISptoDigitacion.Text = "" Or Me.txtSISCodigoUDR.Text = "" Then
              MsgBox "Tiene que ingresar los 4 datos para el SIS", vbCritical
              Exit Sub
           End If
        End If
        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        Dim oRsFox1 As New Recordset
        Dim oConexionFox As New Connection
        Dim oConexion As New Connection
        Dim oConexionSIS As New Connection
        Dim oDOhis_historicoAten As New DOhis_historicoAten, oHIS_historicoAten As New HIS_historicoAten
        Dim oDOPaciente As New DOPaciente, oPacientes As New Pacientes
        Dim oDOHistoriaClinica As New DOHistoriaClinica, oHistoriasClinicas As New HistoriasClinicas
        Dim oDOPArametro As New DOPArametro, oParametros As New Parametros
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim oSisConsumoWeb As New SIGHNegocios.SisConsumoWeb
        Dim lnIdUsuario As Long, lbContinuar As Boolean, ldFechaNac As Date
        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
        Dim lcSegundoNombre As String, lnTipoSexo As Long, lnIdPaciente As Long
        Dim lcAutogenerado As String, lcDNI As String, lnNroHistoriaClinica As Long
        Dim lcSql As String, lcCodDx As String, lcCod2000 As String, lbEsNuevaHC As Boolean, lcFichaFam As String
        Const lnIdTipoNumeracion As Long = 2
        On Error GoTo ErrProAtHis
        '
        oConexionSIS.CommandTimeout = 300
        oConexionSIS.CursorLocation = adUseClient
        oConexionSIS.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        oConexion.BeginTrans
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        Me.MousePointer = 1
        mo_ReglasAdmision.his_historicoAtencionesEliminarTodas oConexion
        'lcCod2000 = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
        lcCod2000 = Right("00000000000" & txtcodrenaes.Text, 9)
        lcSql = "select * from histdet where cod_2000='" & lcCod2000 & "' order by dni"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
            ProgressBar2.Max = oRsFox.RecordCount
        End If
        If oRsFox.RecordCount > 0 Then
           ProgressBar2.Min = 0
           ProgressBar2.Value = 0

           Set oHIS_historicoAten.Conexion = oConexion
           Set oPacientes.Conexion = oConexion
           Set oHistoriasClinicas.Conexion = oConexion
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF
               DoEvents: ProgressBar2.Value = ProgressBar2.Value + 1: Me.Refresh
               lcDNI = Right("        " & Trim(oRsFox.Fields!DNI), 8)
               lcFichaFam = Trim(oRsFox.Fields!fichafam)
               lcSql = "select * from histcab where dni='" & lcDNI & "'"
               If oRsFox1.State = 1 Then oRsFox1.Close
               oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
               If oRsFox1.RecordCount > 0 And Len(lcFichaFam) <= 6 And Val(lcFichaFam) > 0 Then
                    lnNroHistoriaClinica = Val(oRsFox.Fields!fichafam)
                    lbContinuar = True
                    lbEsNuevaHC = True
                    lcApellidoPaterno = UCase(Left(Trim(oRsFox1.Fields!pApellido), 40))
                    lcApellidoMaterno = UCase(Left(Trim(oRsFox1.Fields!sApellido), 40))
                    lcPrimerNombre = UCase(RetornaPrimerNombre(oRsFox1.Fields!Nombres))
                    lcSegundoNombre = UCase(RetornaSegundoNombre(oRsFox1.Fields!Nombres))
                    lnTipoSexo = IIf(Val(oRsFox1.Fields!sexo) = 0, 1, Val(oRsFox1.Fields!sexo))
                    ldFechaNac = IIf(IsNull(oRsFox1.Fields!fnac), CDate("01/01/1970"), oRsFox1.Fields!fnac)
                    Set oRsTmp = mo_ReglasAdmision.PacientesXdni(lcDNI, oConexion)
                    If oRsTmp.RecordCount = 0 Then
                       Set oRsTmp = mo_ReglasAdmision.PacientesXnroHistoriaTipoNumeracion(lnNroHistoriaClinica, lnIdTipoNumeracion, oConexion)
                       If oRsTmp.RecordCount > 0 Then
                          lbEsNuevaHC = False
                          lnIdPaciente = oRsTmp.Fields!idPaciente
                          If AutogeneradosIGUALES(oRsTmp, oRsFox1) = False Then
                             lbContinuar = False
                          End If
                          
                       End If
                    Else
                       lbEsNuevaHC = False
                       lnIdPaciente = oRsTmp.Fields!idPaciente
                       If AutogeneradosIGUALES(oRsTmp, oRsFox1) = False Then
                           lbContinuar = False
                       End If
                    End If
                    If lbEsNuevaHC = False And lbContinuar = True Then
                        '***** modificar datos del Paciente
                        oDOPaciente.idPaciente = lnIdPaciente
                        If oPacientes.SeleccionarPorId(oDOPaciente) = True Then
                            lcAutogenerado = PacienteCrearNroAutogenerado1(Format(ldFechaNac, sighentidades.DevuelveFechaSoloFormato_DMY), lcApellidoPaterno, _
                                             lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                            oDOPaciente.ApellidoMaterno = lcApellidoMaterno
                            oDOPaciente.ApellidoPaterno = lcApellidoPaterno
                            oDOPaciente.PrimerNombre = lcPrimerNombre
                            oDOPaciente.SegundoNombre = lcSegundoNombre
                            oDOPaciente.TercerNombre = RetornaTercerNombre(oRsFox1.Fields!Nombres)
                            oDOPaciente.FechaNacimiento = ldFechaNac
                            oDOPaciente.IdDocIdentidad = 1
                            oDOPaciente.NroDocumento = lcDNI
                            If oPacientes.Modificar(oDOPaciente, False) = True Then
                            End If
                        End If
                    ElseIf lbEsNuevaHC = True Then
                        '***** Nueva Historia
                        lcAutogenerado = PacienteCrearNroAutogenerado1(Format(ldFechaNac, sighentidades.DevuelveFechaSoloFormato_DMY), lcApellidoPaterno, _
                                         lcApellidoMaterno, lcPrimerNombre, lcSegundoNombre, lnTipoSexo)
                        oPacientes.SetDefaults oDOPaciente
                        oDOPaciente.NroHistoriaClinica = lnNroHistoriaClinica
                        oDOPaciente.ApellidoMaterno = lcApellidoMaterno
                        oDOPaciente.ApellidoPaterno = lcApellidoPaterno
                        oDOPaciente.Autogenerado = lcAutogenerado
                        If IsNull(oRsFox1.Fields!Direccion) Or Trim(oRsFox1.Fields!Direccion) = "" Then
                           oDOPaciente.DireccionDomicilio = ""
                        Else
                           oDOPaciente.DireccionDomicilio = Left(oRsFox1.Fields!Direccion, 50)
                        End If
                        oDOPaciente.FechaNacimiento = ldFechaNac
                        If Val(oRsFox1.Fields!Ubigeo) > 0 Then
                           oDOPaciente.IdDistritoDomicilio = Val(oRsFox1.Fields!Ubigeo)
                        Else
                           oDOPaciente.IdDistritoDomicilio = 0
                        End If
                        oDOPaciente.IdDocIdentidad = 1
                        oDOPaciente.IdTipoNumeracion = lnIdTipoNumeracion
                        oDOPaciente.idTipoSexo = lnTipoSexo
                        oDOPaciente.NroDocumento = lcDNI
                        oDOPaciente.PrimerNombre = lcPrimerNombre
                        oDOPaciente.SegundoNombre = lcSegundoNombre
                        oDOPaciente.TercerNombre = RetornaTercerNombre(oRsFox1.Fields!Nombres)
                       If oPacientes.Insertar(oDOPaciente) = True Then
                            oDOHistoriaClinica.FechaCreacion = Date
                            oDOHistoriaClinica.IdEstadoHistoria = 1
                            oDOHistoriaClinica.idPaciente = oDOPaciente.idPaciente
                            oDOHistoriaClinica.IdTipoHistoria = 1
                            oDOHistoriaClinica.IdTipoNumeracion = lnIdTipoNumeracion
                            oDOHistoriaClinica.IdUsuarioAuditoria = lnIdUsuario
                            oDOHistoriaClinica.NroHistoriaClinica = lnNroHistoriaClinica
                            If oHistoriasClinicas.Insertar(oDOHistoriaClinica) = False Then
                               lbContinuar = False
                            Else
                               lnIdPaciente = oDOPaciente.idPaciente
                            End If
                       Else
                            lbContinuar = False
                       End If

                    End If
                    If lbContinuar = True Then
                        Do While Not oRsFox.EOF And Val(lcDNI) = Val(oRsFox.Fields!DNI)
                           If Not (IsNull(oRsFox.Fields!diagnost) Or Trim(oRsFox.Fields!diagnost) = "") _
                                              And Not (IsNull(oRsFox.Fields!ups) Or Trim(oRsFox.Fields!ups) = "") Then
                                If IsNull(oRsFox.Fields!cpt) Or Trim(oRsFox.Fields!cpt) = "" Then
                                   oDOhis_historicoAten.cpt = ""
                                Else
                                   oDOhis_historicoAten.cpt = oRsFox.Fields!cpt
                                End If
                                If IsNull(oRsFox.Fields!diagnost) Or Trim(oRsFox.Fields!diagnost) = "" Then
                                   oDOhis_historicoAten.diagnost = oRsFox.Fields!diagnost
                                Else
                                   oDOhis_historicoAten.diagnost = oRsFox.Fields!diagnost
                                End If
                                oDOhis_historicoAten.Fecha = CDate(oRsFox.Fields!dia & "/" & oRsFox.Fields!Mes & "/" & oRsFox.Fields!ano)
                                oDOhis_historicoAten.IdUsuarioAuditoria = lnIdUsuario
                                oDOhis_historicoAten.idPaciente = lnIdPaciente
                                If IsNull(oRsFox.Fields!ups) Or Trim(oRsFox.Fields!ups) = "" Then
                                   oDOhis_historicoAten.ups = ""
                                Else
                                   oDOhis_historicoAten.ups = oRsFox.Fields!ups
                                End If
                                If oHIS_historicoAten.Insertar(oDOhis_historicoAten) = True Then
                                   lcSql = ""
                                End If
                           End If
                           oRsFox.MoveNext
                           If oRsFox.EOF Then
                              Exit Do
                           End If
                        Loop
                    End If
               Else
                    lbContinuar = False
               End If
               If lbContinuar = False Then
                    Do While Not oRsFox.EOF And Val(lcDNI) = Val(oRsFox.Fields!DNI)
                        oRsFox.MoveNext
                        If oRsFox.EOF Then
                           Exit Do
                        End If
                    Loop
               End If
           Loop

        End If
        '
        Set oParametros.Conexion = oConexion
        If Trim(txtnombre.Text) <> "" Then
            oDOPArametro.IdParametro = 205
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtnombre.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtdireccion.Text) <> "" Then
            oDOPArametro.IdParametro = 206
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtdireccion.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txttelefono.Text) <> "" Then
            oDOPArametro.IdParametro = 207
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txttelefono.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtubigeo.Text) <> "" Then
            oDOPArametro.IdParametro = 242
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtubigeo.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtcodminsa.Text) <> "" Then
            oDOPArametro.IdParametro = 208
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtcodminsa.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtcoddisa.Text) <> "" Then
            oDOPArametro.IdParametro = 239
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtcoddisa.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtcodred.Text) <> "" Then
            oDOPArametro.IdParametro = 240
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtcodred.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtcodmicrored.Text) <> "" Then
            oDOPArametro.IdParametro = 241
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtcodmicrored.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtcodHIS.Text) <> "" Then
            oDOPArametro.IdParametro = 243
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtcodHIS.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        If Trim(txtcodrenaes.Text) <> "" Then
            oDOPArametro.IdParametro = 280
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = txtcodrenaes.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
        End If
        '
        oDOPArametro.IdParametro = 282
        If Not oParametros.SeleccionarPorId(oDOPArametro) Then
           MsgBox oParametros.MensajeError: GoTo ErrProAtHis
        End If
        oDOPArametro.ValorTexto = IIf(chkCentroSalud.Value = 1, "S", "n")
        If Not oParametros.Modificar(oDOPArametro) Then
           MsgBox oParametros.MensajeError: GoTo ErrProAtHis
        End If
        '

        If oRsGrdSIS.RecordCount > 0 Then
            oDOPArametro.IdParametro = 310
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = Me.txtSisDisa.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            '
            oDOPArametro.IdParametro = 303
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = Me.txtSisCatEESS.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            '
            oDOPArametro.IdParametro = 304
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = Me.txtSISptoDigitacion.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            '
            oDOPArametro.IdParametro = 305
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If
            oDOPArametro.ValorTexto = Me.txtSISCodigoUDR.Text
            If Not oParametros.Modificar(oDOPArametro) Then
               MsgBox oParametros.MensajeError: GoTo ErrProAtHis
            End If

        End If
'        oRsTmp.Close
        '
        oConexion.CommitTrans
        Me.MousePointer = 11
        oConexionFox.Close
        oConexion.Close
        oConexionSIS.Close

        Set oRsTmp = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Set oConexionFox = Nothing
        Set oConexion = Nothing
        Set oDOhis_historicoAten = Nothing
        Set oHIS_historicoAten = Nothing
        Set oDOPaciente = Nothing
        Set oPacientes = Nothing
        Set oDOHistoriaClinica = Nothing
        Set oHistoriasClinicas = Nothing
        Set lcBuscaParametro = Nothing
        Set oDOPArametro = Nothing
        Set oParametros = Nothing
        Set oConexionSIS = Nothing
        Unload Me
        Exit Sub
ErrProAtHis:
        MsgBox Err.Description
Resume
        oConexion.RollbackTrans
        Set oRsTmp = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Set oConexionFox = Nothing
        Set oConexion = Nothing
        Set oDOhis_historicoAten = Nothing
        Set oHIS_historicoAten = Nothing
        Set oDOPaciente = Nothing
        Set oPacientes = Nothing
        Set oDOHistoriaClinica = Nothing
        Set oHistoriasClinicas = Nothing
        Set lcBuscaParametro = Nothing
        Set oConexionSIS = Nothing
        Unload Me
End Sub

Function AutogeneradosIGUALES(oRsTmp As Recordset, oRsFox1 As Recordset) As Boolean
        Dim ldFechaNac As Date, lnTipoSexo As Long
        Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String, lcSegundoNombre As String
        lcApellidoPaterno = UCase(Left(Trim(oRsFox1.Fields!pApellido), 40))
        lcApellidoMaterno = UCase(Left(Trim(oRsFox1.Fields!sApellido), 40))
        lcPrimerNombre = UCase(RetornaPrimerNombre(oRsFox1.Fields!Nombres))
        lcSegundoNombre = UCase(RetornaSegundoNombre(oRsFox1.Fields!Nombres))
        lnTipoSexo = IIf(Val(oRsFox1.Fields!sexo) = 0, 1, Val(oRsFox1.Fields!sexo))
        ldFechaNac = IIf(IsNull(oRsFox1.Fields!fnac), CDate("01/01/1970"), oRsFox1.Fields!fnac)
        AutogeneradosIGUALES = False
        If Mid(oRsTmp!ApellidoPaterno, 1, 1) = Mid(lcApellidoPaterno, 1, 1) And Mid(oRsTmp!ApellidoPaterno, 4, 1) = Mid(lcApellidoPaterno, 4, 1) Then
           If Mid(oRsTmp!ApellidoMaterno, 1, 1) = Mid(lcApellidoMaterno, 1, 1) And Mid(oRsTmp!ApellidoMaterno, 4, 1) = Mid(lcApellidoMaterno, 4, 1) Then
              If Mid(oRsTmp!PrimerNombre, 1, 1) = Mid(lcPrimerNombre, 1, 1) And Mid(oRsTmp!PrimerNombre, 4, 1) = Mid(lcPrimerNombre, 4, 1) Then
                 If lnTipoSexo = oRsTmp!idTipoSexo Then
                    If DateDiff("d", oRsTmp!FechaNacimiento, ldFechaNac) >= -15 And DateDiff("d", oRsTmp!FechaNacimiento, ldFechaNac) <= 15 Then
                       AutogeneradosIGUALES = True
                    End If
                 End If
              End If
           End If
        End If
         
End Function


Function RetornaTercerNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer, lcNombre1 As String, lcNombre2 As String, lcNombre3 As String
    RetornaTercerNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       lcNombre1 = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
       ln = InStr(lcNombre1, " ")
       If ln > 0 Then
          lcNombre2 = Trim(Left(lcNombre1, ln))
          RetornaTercerNombre = Trim(Mid(lcNombre1, ln + 1, 100))
       End If
    End If
End Function



Private Sub Text_Change()

End Sub

Private Sub Text_LostFocus()
   
    Frame2(4).Visible = True
End Sub







Private Sub txtClave1_KeyPress(KeyAscii As Integer)
   If UCase(txtClave1.Text) = "DEBB" Then
       Frame2(4).Visible = True
    End If
End Sub



Private Sub txtnombre_Change()
    If chkDejarBuscar.Value Then
        dgrMuestra.Visible = False
        Exit Sub
    End If
    If optHISActual.Value = False And optHISAnterior.Value = False Then
        MsgBox "Elija Primero el tipo de Sistema HIS", vbInformation, Me.Caption
    Else
        Dim RCS As Recordset
        Dim opt As Boolean
        dgrMuestra.Visible = True
        opt = optHISAnterior.Value
        Set RCS = New Recordset
        Set dgrMuestra.DataSource = ConsultarNombreEstablecimiento(opt, txtnombre.Text)
        dgrMuestra.Columns(0).Visible = False
        dgrMuestra.Columns(1).Width = 4000
    End If

End Sub

 Function ConsultarNombreEstablecimiento(ByVal His As Boolean, ByVal nombreEstablecimiento As String) As Recordset
    Dim oConexionMDB As New Connection, lcSql As String
    Dim oRecordset As New Recordset
    Dim tabla As String
    nombreEstablecimiento = Replace(nombreEstablecimiento, "'", "")
    
    oConexionMDB.CommandTimeout = 300
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "\parametros.mdb"
    If His Then
        tabla = "ESTABLEC_ant_his"
    Else
        tabla = "ESTABLEC_ult_his"
    End If
    lcSql = "Select cod_estab as codigo, desc_estab as Nombre from " & tabla & " where desc_estab like '%" & Trim(nombreEstablecimiento) & "%'"
    oRecordset.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    Set ConsultarNombreEstablecimiento = oRecordset
End Function

 
 
Private Sub dgrMuestra_DblClick()
    SetearDatos ConsultarEstablecimiento(optHISAnterior.Value, CStr(dgrMuestra.Columns.Item(0).Value))
    dgrMuestra.Visible = False
    txtdireccion.SetFocus
End Sub
Private Sub SetearDatos(data As Recordset)
    If data.EOF = False Then
        txtnombre.Text = data!desc_estab
        txtcodminsa.Text = Right("0000" & Trim(Str(data!cod_2000)), 5)
        txtcoddisa.Text = data!COD_DISA
        txtcodred.Text = data!COD_RED
        txtcodmicrored.Text = data!COD_MIC
        txtubigeo.Text = data!Ubigeo
        txtcodHIS.Text = data!cod_estab
        txtcodrenaes.Text = txtcodminsa.Text
        Me.txtdireccion.Text = ""
        Me.txttelefono.Text = ""
        Me.txtSisCatEESS.Text = ""
        Me.txtSISCodigoUDR.Text = ""
        Me.txtSisDisa.Text = ""
        Me.txtSISptoDigitacion.Text = ""
        Dim oSisConsumoWeb As New SIGHNegocios.SisConsumoWeb
        Set oRsGrdSIS = oSisConsumoWeb.m_eessSelecionarXcodigoRenaes(Right("0000000000" & txtcodrenaes.Text, 10))
        Set Me.grdSIS.DataSource = oRsGrdSIS
        If oRsGrdSIS.RecordCount = 1 Then
           txtSisDisa.Text = oRsGrdSIS.Fields!disa
           txtSisCatEESS.Text = oRsGrdSIS.Fields!CategoriaEESS
           txtSISptoDigitacion.Text = IIf(IsNull(oRsGrdSIS.Fields!ptoDigitacion), "", oRsGrdSIS.Fields!ptoDigitacion)
           txtSISCodigoUDR.Text = oRsGrdSIS.Fields!codigoUDR
        End If
        Me.grdSIS.Visible = True
    End If
End Sub

 Function ConsultarEstablecimiento(ByVal His As Boolean, codEstab As String) As Recordset
    Dim oConexionMDB As New Connection, lcSql As String
    Dim oRecordset As New Recordset
    Dim tabla As String
    
    oConexionMDB.CommandTimeout = 300
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "\parametros.mdb"
    If His Then
        tabla = "ESTABLEC_ant_his"
    Else
        tabla = "ESTABLEC_ult_his"
    End If
    lcSql = "select cod_estab," & _
                           "desc_estab," & _
                           "cod_2000," & _
                           "(cod_dpto + cod_prov + cod_dist) as ubigeo," & _
                           "cod_disa," & _
                           "cod_red," & _
                           "cod_mic " & _
                           "from " & tabla & " " & _
                           "where cod_estab = '" + codEstab + "'"
    oRecordset.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    Set ConsultarEstablecimiento = oRecordset
End Function


Private Sub Command5_Click()
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim oRsTmp4 As New Recordset
    Dim lcSql As String
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    lcSql = "select * from atencionesDatosAdicionales"
    oRsTmp1.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          If IsNull(oRsTmp1.Fields!SisCodigo) Then
            lcSql = "select idCuentaAtencion from atenciones where idFormaPago=2 and idAtencion=" & oRsTmp1.Fields!idAtencion
            oRsTmp2.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
            If oRsTmp2.RecordCount > 0 Then
               lcSql = "select * from sisFuaAtencion where idCuentaAtencion=" & oRsTmp2.Fields!idCuentaAtencion
               oRsTmp3.Open lcSql, lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo), adOpenKeyset, adLockOptimistic
               If oRsTmp3.RecordCount > 0 Then
                  If IsNull(oRsTmp3.Fields!Codigo) Then
                     lcSql = "select idSiaSis,Codigo from sisFiliaciones where afiliacionDisa='" & oRsTmp3.Fields!AfiliacionDisa & "' and afiliacionTipoFormato='" & oRsTmp3.Fields!AfiliacionTipoFormato & "' and afiliacionNroFormato='" & oRsTmp3.Fields!AfiliacionNroFormato & "'"
                     oRsTmp4.Open lcSql, lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo), adOpenKeyset, adLockOptimistic
                     If oRsTmp4.RecordCount > 0 Then
                        oRsTmp1.Fields!idSiasis = oRsTmp4.Fields!idSiasis
                        oRsTmp1.Fields!SisCodigo = oRsTmp4.Fields!Codigo
                        oRsTmp1.Update
                        If IsNull(oRsTmp3.Fields!idSiasis) Then
                           oRsTmp3.Fields!idSiasis = oRsTmp4.Fields!idSiasis
                           oRsTmp3.Update
                        End If
                     End If
                     oRsTmp4.Close
                  Else
                     oRsTmp1.Fields!SisCodigo = oRsTmp3.Fields!Codigo
                     oRsTmp1.Update
                  End If
               End If
               oRsTmp3.Close
            End If
            oRsTmp2.Close
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    Unload Me
End Sub

Private Sub cmdEstabNewDesdeSIS_Click()
    Me.MousePointer = 11
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexionSIGH As New ADODB.Connection
    Dim oConexion As New Connection
    Dim lnUltimoId As Long
    Dim lbNuevo As Boolean, lcCodigo As String, lcSql As String, lnBar As Long
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    '
    lcSql = "select * from m_Departamentos"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    oConexionSIGH.CursorLocation = adUseClient
    oConexionSIGH.CommandTimeout = 300
    oConexionSIGH.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionSIGH
        .CommandTimeout = 150
        .CommandText = "DepartamentosSeleccionarTodosCampos"
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    
    oRsTmp1.MoveFirst
    Do While Not oRsTmp1.EOF
       oRsTmp2.MoveFirst
       oRsTmp2.Find "IdDepartamento=" & oRsTmp1.Fields!dep_IdDep
       If oRsTmp2.EOF Then
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexionSIGH
                .CommandTimeout = 150
                .CommandText = "DepartamentosAgregar"
                Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, Val(oRsTmp1.Fields!dep_IdDep)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 20, Left(oRsTmp1.Fields!dep_Descripcion, 20)): .Parameters.Append oParameter
                .Execute
            End With
            Set oCommand = Nothing
            Set oCommand = Nothing
            
       End If
       oRsTmp1.MoveNext
    Loop
    oRsTmp1.Close
    oRsTmp2.Close
    '
    lcSql = "select * from m_Provincias"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionSIGH
        .CommandTimeout = 150
        .CommandText = "ProvinciasSeleccionarTodo"
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
        
    oRsTmp1.MoveFirst
    Do While Not oRsTmp1.EOF
       oRsTmp2.MoveFirst
       oRsTmp2.Find "IdProvincia=" & oRsTmp1.Fields!prv_IdProv
       If oRsTmp2.EOF Then
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexionSIGH
                .CommandTimeout = 150
                .CommandText = "ProvinciasAgregar"
                Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, Val(oRsTmp1.Fields!prv_IdProv)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, Left(oRsTmp1.Fields!prv_Descripcion, 50)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, Val(Left(oRsTmp1.Fields!prv_IdProv, 2))): .Parameters.Append oParameter
                .Execute
            End With
            Set oCommand = Nothing
            Set oCommand = Nothing
       End If
       oRsTmp1.MoveNext
    Loop
    oRsTmp1.Close
    oRsTmp2.Close
    '
    lcSql = "select * from m_Distritos"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic

    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionSIGH
        .CommandTimeout = 150
        .CommandText = "DistritosSeleccionarTodo"
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    
    ProgressBar1.Max = oRsTmp1.RecordCount + 1: lnBar = 1: ProgressBar1.Min = lnBar
    oRsTmp1.MoveFirst
    Do While Not oRsTmp1.EOF
       DoEvents: ProgressBar1.Value = lnBar: lnBar = lnBar + 1: Me.Refresh
       oRsTmp2.MoveFirst
       oRsTmp2.Find "IdDistrito=" & oRsTmp1.Fields!dis_IdUbigeo
       If oRsTmp2.EOF Then
           With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexionSIGH
                .CommandTimeout = 150
                .CommandText = "DistritoAgregar"
                Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, Val(oRsTmp1.Fields!dis_IdUbigeo)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, Left(oRsTmp1.Fields!dis_Descripcion, 50)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdProvincia", adInteger, adParamInput, 0, Val(Left(oRsTmp1.Fields!dis_IdUbigeo, 4))): .Parameters.Append oParameter
                .Execute
            End With
            Set oCommand = Nothing
            Set oCommand = Nothing
       End If
       oRsTmp1.MoveNext
    Loop
    oRsTmp1.Close
    oRsTmp2.Close
    '
    lcSql = "select * from m_eess"
    oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    ProgressBar1.Max = oRsTmp1.RecordCount + 1: lnBar = 1: ProgressBar1.Min = lnBar
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
                DoEvents: ProgressBar1.Value = lnBar: lnBar = lnBar + 1: Me.Refresh
                lcCodigo = Right(oRsTmp1.Fields!pre_CodigoRENAES, 5)
                If Val(oRsTmp1.Fields!pre_CodigoRENAES) <> Val(lcCodigo) Then
                   lcCodigo = oRsTmp1.Fields!pre_CodigoRENAES
                End If
                If Len(lcCodigo) <= 6 Then
                    If oRsTmp3.State = 1 Then oRsTmp3.Close
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexionSIGH
                        .CommandTimeout = 150
                        .CommandText = "EstablecimientosSeleccionarTodoCamposXCodigo"
                        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, lcCodigo): .Parameters.Append oParameter
                        Set oRsTmp3 = .Execute
                        Set oRsTmp3.ActiveConnection = Nothing
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
        
                    If oRsTmp3.RecordCount = 0 Then
                          If oRsTmp2.State = 1 Then oRsTmp2.Close
                            With oCommand
                                .CommandType = adCmdStoredProc
                                Set .ActiveConnection = oConexionSIGH
                                .CommandTimeout = 150
                                .CommandText = "EstablecimientosNoMinsaSeleccionarPorCodigo"
                                Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 50, lcCodigo): .Parameters.Append oParameter
                                Set oRsTmp2 = .Execute
                                Set oRsTmp2.ActiveConnection = Nothing
                            End With
                            Set oCommand = Nothing
                            Set oParameter = Nothing
                          
                          If oRsTmp2.RecordCount = 0 Then
    
                                    With oCommand
                                         .CommandType = adCmdStoredProc
                                            Set .ActiveConnection = oConexionSIGH
                                            .CommandTimeout = 150
                                            .CommandText = "EstablecimientosNoMinsaAgregar"
                                            Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsa", adInteger, adParamOutput, 0, 1): .Parameters.Append oParameter
                                            Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, Left(oRsTmp1.Fields!pre_Nombre, 150)): .Parameters.Append oParameter
                                            Set oParameter = .CreateParameter("@IdTipoSubsector", adInteger, adParamInput, 0, 4): .Parameters.Append oParameter
                                            Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, Val(oRsTmp1.Fields!pre_IdUbigeo)): .Parameters.Append oParameter
                                            Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, lcCodigo): .Parameters.Append oParameter
                                            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                                         .Execute
                                     End With
                                     Set oCommand = Nothing
                                     Set oCommand = Nothing
                          End If
                    End If
                End If
                oRsTmp1.MoveNext
       Loop
    End If
    Me.MousePointer = 1
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTmp3 = Nothing
    Set oConexion = Nothing
    Unload Me

End Sub

Private Sub cmdEliminaHC_Click()

If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Dim oRsTmpOpc As New Recordset
        Dim oRsTmpCat As New Recordset
        Dim oRsFox As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim lbNuevo As Boolean
        Dim lnIdOpc As Long
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=SISMEDV2"
        '
        Me.MousePointer = 1
        On Error Resume Next
        lcSql = "select nroHistori from HC"
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           Do While Not oRsFox.EOF

              With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = wxConexionRed
                  .CommandTimeout = 150
                  .CommandText = "HistoriasClinicasYPacientesEliminarPorNroHistoriaClinica"
                  Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oRsFox.Fields!nroHistori): .Parameters.Append oParameter
                  Set oRsTmpOpc = .Execute
              End With
  
              oRsFox.MoveNext
           Loop
        End If
        oRsFox.Close
        Me.MousePointer = 11
        Unload Me
End If
End Sub




Private Sub ProcesaCPTtarapoto_Click()
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtExcel.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    lnFila = 4
    lnFilaFinal = 2406
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    For lnFor = lnFila To lnFilaFinal
        lcRango = "B" + Trim(Str(lnFor))
        lcCodigo = s.Range(lcRango).Value
        lcRango = "F" + Trim(Str(lnFor))
        s.Range(lcRango).Value = ""
        If Val(lcCodigo) > 0 Then
           
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
                      
           If oRsTmp.RecordCount > 0 Then
              
              s.Range(lcRango).Value = oRsTmp.Fields!nombre
           End If
           oRsTmp.Close
        End If
    Next
    Set s = Nothing
    W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    Unload Me
End Sub


Private Sub cmdProcesaCPT_Click()
    On Error GoTo ErrorExcel
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim mo_ReporteUtil As New ReporteUtil
    Dim iFila As Integer
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    '
    Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    Set oWorkBook = oExcel.Workbooks.Add
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\HojaLibre.xls")
    oWorkBookPlantilla.Worksheets("Hoja_libre").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    Set oWorkSheet = oWorkBook.Sheets(1)
    iFila = 5
    '
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Dim oConexionFox As New ADODB.Connection
    Dim oRsFox1 As New Recordset
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    lcSql = "select * from cpt_sis"
    oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
    '
    oWorkSheet.Cells(iFila, 1).Value = "Cpt SIS"
    oWorkSheet.Cells(iFila, 2).Value = "Codigo SIS"
    oWorkSheet.Cells(iFila, 3).Value = "Descripción SIS"
    oWorkSheet.Cells(iFila, 4).Value = "Descripción GalenHos"
    If oRsFox1.RecordCount > 0 Then
       
       iFila = iFila + 2
       oRsFox1.MoveFirst
       Do While Not oRsFox1.EOF
          oWorkSheet.Cells(iFila, 1).Value = "'" & oRsFox1.Fields!a
          oWorkSheet.Cells(iFila, 2).Value = "'" & oRsFox1.Fields!b
          oWorkSheet.Cells(iFila, 3).Value = Trim(oRsFox1.Fields!C)
          
          
          With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "FactCatalogoServiciosXcodigo"
              Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 20, Trim(oRsFox1.Fields!a)): .Parameters.Append oParameter
              Set oRsTmp1 = .Execute
              Set oRsTmp1.ActiveConnection = Nothing
          End With
          Set oCommand = Nothing
          Set oParameter = Nothing
          
          If oRsTmp1.RecordCount > 0 Then
             oWorkSheet.Cells(iFila, 4).Value = Trim(oRsTmp1.Fields!nombre)
          End If
          oRsTmp1.Close
          iFila = iFila + 1
          oRsFox1.MoveNext
       Loop
    End If
    oRsFox1.Close
    '
    oExcel.Visible = True
    oWorkSheet.PrintPreview
    Exit Sub
ErrorExcel:
    MsgBox Err.Description
    
End Sub



Private Sub cmdCuentasYtarifas_Click()

     Dim oRsTmp1 As New Recordset
     Dim oRsTmp2 As New Recordset
     Dim oRsGrid As New Recordset
     Dim oConexion As New Connection
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
  
     Me.MousePointer = 11
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     'crea tmp para Errores
     With oRsGrid
      .Fields.Append "NroCuenta", adVarChar, 10, adFldIsNullable
      .Fields.Append "Plan", adVarChar, 50, adFldIsNullable
      .Fields.Append "Tarifa", adVarChar, 50, adFldIsNullable
      .Fields.Append "Fingreso", adDate
      .Fields.Append "EstadoCuenta", adVarChar, 50, adFldIsNullable
      .Fields.Append "EstadoAtencion", adVarChar, 50, adFldIsNullable
      .LockType = adLockOptimistic
      .Open
     End With
     Set grdCuentasYtarifas.DataSource = oRsGrid
     '
'     lcSql = "SELECT     dbo.Atenciones.IdFormaPago, dbo.Atenciones.idFuenteFinanciamiento, dbo.TiposFinanciamiento.Descripcion AS dTarifario, " & _
'            "                      dbo.FuentesFinanciamiento.Descripcion AS dPlan, dbo.Atenciones.IdCuentaAtencion, dbo.Atenciones.FechaIngreso," & _
'            "                      dbo.EstadosCuenta.Descripcion AS EstadoCuenta, dbo.EstadosAtencion.Descripcion AS EstadoAtencion" & _
'            " FROM         dbo.Atenciones INNER JOIN" & _
'            "                      dbo.FacturacionCuentasAtencion ON dbo.Atenciones.IdCuentaAtencion = dbo.FacturacionCuentasAtencion.IdCuentaAtencion INNER JOIN" & _
'            "                      dbo.EstadosAtencion ON dbo.Atenciones.idEstadoAtencion = dbo.EstadosAtencion.IdEstadoAtencion LEFT OUTER JOIN" & _
'            "                      dbo.EstadosCuenta ON dbo.FacturacionCuentasAtencion.IdEstado = dbo.EstadosCuenta.IdEstado LEFT OUTER JOIN" & _
'            "                      dbo.TiposFinanciamiento ON dbo.Atenciones.IdFormaPago = dbo.TiposFinanciamiento.IdTipoFinanciamiento LEFT OUTER JOIN" & _
'            "                      dbo.FuentesFinanciamiento ON dbo.Atenciones.idFuenteFinanciamiento = dbo.FuentesFinanciamiento.IdFuenteFinanciamiento" & _
'            " ORDER BY dbo.atenciones.idAtencion"
'     oRstmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        
      With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "CuentasYtarifasSeleccionar"
                Set oRsTmp1 = .Execute
      End With
      Set oCommand = Nothing
     
      ProgressBar1.Min = 0: ProgressBar1.Max = oRsTmp1.RecordCount + 2: ProgressBar1.Value = 0
       
    If oRsTmp1.RecordCount > 0 Then
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
             DoEvents:            ProgressBar1.Value = ProgressBar1.Value + 1:          Me.Refresh
            If IsNull(oRsTmp1.Fields!idFuenteFinanciamiento) Or IsNull(oRsTmp1.Fields!IdFormaPago) Then
               oRsGrid.AddNew
               oRsGrid.Fields!NroCuenta = Trim(Str(oRsTmp1.Fields!idCuentaAtencion))
               oRsGrid.Fields!Plan = IIf(IsNull(oRsTmp1.Fields!dPlan), "(no tiene valor)", oRsTmp1.Fields!dPlan)
               oRsGrid.Fields!Tarifa = IIf(IsNull(oRsTmp1.Fields!dTarifario), "(no tiene valor)", oRsTmp1.Fields!dTarifario)
               oRsGrid.Fields!fIngreso = oRsTmp1.Fields!FechaIngreso
               oRsGrid.Fields!EstadoCuenta = oRsTmp1.Fields!EstadoCuenta
               oRsGrid.Fields!EstadoAtencion = oRsTmp1.Fields!EstadoAtencion
               oRsGrid.Update
            Else
    '            lcSql = "select * from FuentesFinanciamientoTarifas where idFuenteFinanciamiento=" & oRstmp1.Fields!idFuenteFinanciamiento & _
    '                  " and  idTipoFinanciamiento=" & oRstmp1.Fields!IdFormaPago
    '            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                
                With oCommand
                          .CommandType = adCmdStoredProc
                          Set .ActiveConnection = oConexion
                          .CommandTimeout = 150
                          .CommandText = "FuentesFinanciamientoTarifasSeleccionarPorIdFuenteFinanciamiento"
                          Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oRsTmp1.Fields!idFuenteFinanciamiento): .Parameters.Append oParameter
                          Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oRsTmp1.Fields!IdFormaPago): .Parameters.Append oParameter
                          Set oRsTmp2 = .Execute
                          Set oRsTmp2.ActiveConnection = Nothing
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
          
                If oRsTmp2.RecordCount = 0 Then
                   oRsGrid.AddNew
                   oRsGrid.Fields!NroCuenta = Trim(Str(oRsTmp1.Fields!idCuentaAtencion))
                   oRsGrid.Fields!Plan = oRsTmp1.Fields!dPlan
                   oRsGrid.Fields!Tarifa = oRsTmp1.Fields!dTarifario
                   oRsGrid.Fields!fIngreso = oRsTmp1.Fields!FechaIngreso
                   oRsGrid.Fields!EstadoCuenta = oRsTmp1.Fields!EstadoCuenta
                   oRsGrid.Fields!EstadoAtencion = oRsTmp1.Fields!EstadoAtencion
                   oRsGrid.Update
                End If
                oRsTmp2.Close
            End If
            oRsTmp1.MoveNext
         Loop
    End If
     oRsTmp1.Close
     oConexion.Close
         
     If oRsGrid.RecordCount > 0 Then
        Dim EXL As Excel.Application
        Set EXL = New Excel.Application
        Dim W As Excel.Workbook
        Set W = EXL.Workbooks.Open("c:\excel.xls")
        Dim s As Excel.Worksheet
        Set s = W.Sheets("Hoja1")
        Dim lnFor As Long, lnFila As Integer, lcRango As String
        lnFila = 3
        oRsGrid.MoveFirst
        Do While Not oRsGrid.EOF
            lcRango = "B" + Trim(Str(lnFila))
            s.Range(lcRango).Value = oRsGrid.Fields!NroCuenta
            lcRango = "C" + Trim(Str(lnFila))
            s.Range(lcRango).Value = oRsGrid.Fields!Plan
            lcRango = "D" + Trim(Str(lnFila))
            s.Range(lcRango).Value = oRsGrid.Fields!Tarifa
            lcRango = "E" + Trim(Str(lnFila))
            s.Range(lcRango).Value = oRsGrid.Fields!fIngreso
            lcRango = "F" + Trim(Str(lnFila))
            s.Range(lcRango).Value = oRsGrid.Fields!EstadoCuenta
            lcRango = "G" + Trim(Str(lnFila))
            s.Range(lcRango).Value = oRsGrid.Fields!EstadoAtencion
            lnFila = lnFila + 1
            oRsGrid.MoveNext
        Loop
        Set s = Nothing
        W.Save
        W.Close
        Set W = Nothing
        Set EXL = Nothing
     End If
     MsgBox "Terminó el proceso"
     Me.MousePointer = 1
End Sub

Private Sub cmdActualizaCuboGalenhos2008_Click()
        Dim oRsTmp As New Recordset
        Dim oRsFox As New Recordset
        Dim oConexionFox As New Connection
        Dim lcSql As String, lcCodDx As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
  
        '
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        '
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=GalenHosSql2008"
        '
        Me.MousePointer = 1
        
        With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "FactCatalogoBienesInsumosSeleccionarTodo"
              Set oRsTmp = .Execute
              Set oRsTmp.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        
        ProgressBar1.Max = oRsTmp.RecordCount
        If ProgressBar1.Max > 0 Then
           ProgressBar1.Min = 0
           ProgressBar1.Value = 0
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
              lcSql = "select * from mProducto where medcod='" & Left(oRsTmp.Fields!Codigo, 7) & "'"
              oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
              If oRsFox.RecordCount = 0 Then
                 oRsFox.AddNew
                 oRsFox.Fields!medCod = Left(oRsTmp.Fields!Codigo, 7)
                 oRsFox.Fields!medNom = Left(oRsTmp.Fields!nombre, 100)
                 oRsFox.Fields!medPres = Mid(oRsTmp.Fields!nombre, 101, 70)
                 oRsFox.Fields!medcnc = Mid(oRsTmp.Fields!nombre, 171, 60)
                 oRsFox.Fields!medFF = " "
                 oRsFox.Fields!prdstkprom = 0
                 oRsFox.Fields!mnarcot = " "
                 oRsFox.Fields!mednarcot = " "
                 oRsFox.Fields!fecultprec = Date
                 oRsFox.Fields!prdtipomed = " "
                 oRsFox.Fields!preope_do = 0
                 oRsFox.Fields!predis_do = 0
                 oRsFox.Fields!preadq_do = 0
                 oRsFox.Fields!medRegSan = " "
                 oRsFox.Fields!prdFechUlt = Date
                 oRsFox.Fields!prdAdi = " "
                 oRsFox.Fields!prdSit = "1"
                 oRsFox.Fields!prdptorep = 0
                 oRsFox.Fields!prdstkmax = 0
                 oRsFox.Fields!prdstkmin = 0
                 oRsFox.Fields!prdindact = " "
                 oRsFox.Fields!prdPreOpe = 0
                 oRsFox.Fields!prdPreDist = 0
                 oRsFox.Fields!prdPreAdq = 0
                 oRsFox.Fields!medIndvcto = " "
                 oRsFox.Fields!medAct = " "
                 oRsFox.Fields!medhi = " "
                 oRsFox.Fields!medcs = " "
                 oRsFox.Fields!medps = " "
                 oRsFox.Fields!medcomp = " "
                 oRsFox.Fields!medTraloc = " "
                 oRsFox.Fields!medTraNac = " "
                 oRsFox.Fields!medFactPer = " "
                 oRsFox.Fields!medEstVta = " "
                 oRsFox.Fields!medEst = " "
                 oRsFox.Fields!MedPet = " "
                 oRsFox.Fields!medTip = " "
                 oRsFox.Fields!medNomAbr = " "
                 oRsFox.Fields!mefCod = " "
                 oRsFox.Update
              End If
              oRsFox.Close
              oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
        
        Unload Me
End Sub




Private Sub cmdArreglaENE_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Dim lcEneMal As String, lcEneOK As String
        lcEneMal = ""
        lcEneOK = "Ñ"
        ActualizaLetraEnPacientes lcEneMal, lcEneOK
        lcEneMal = "|"
        lcEneOK = "_"
        ActualizaLetraEnPacientes lcEneMal, lcEneOK
        Unload Me
        
    End If
End Sub

Sub ActualizaLetraEnPacientes(lcEneMal As String, lcEneOK As String)
    Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcNroDocumento As String, lcSegundoNombre As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lcPrimerNombre As String, lcDireccionDomicilio As String, lnRegistros As Long
    Dim oRsTmp1 As New Recordset
    Dim lnFor As Long
  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = wxConexionRed
        .CommandTimeout = 150
        .CommandText = "PacientesBuscarEneMal"
        Set oParameter = .CreateParameter("@EneMal", adVarChar, adParamInput, 20, lcEneMal): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    
    
    lnRegistros = oRsTmp1.RecordCount
    If lnRegistros > 0 Then
       ProgressBar1.Min = 0
       ProgressBar1.Max = lnRegistros
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lcApellidoPaterno = ""
          For lnFor = 1 To Len(oRsTmp1.Fields!ApellidoPaterno)
              If Mid(oRsTmp1.Fields!ApellidoPaterno, lnFor, 1) = lcEneMal Then
                 lcApellidoPaterno = lcApellidoPaterno & lcEneOK
              Else
                 lcApellidoPaterno = lcApellidoPaterno & Mid(oRsTmp1.Fields!ApellidoPaterno, lnFor, 1)
              End If
          Next
          lcApellidoMaterno = ""
          For lnFor = 1 To Len(oRsTmp1.Fields!ApellidoMaterno)
              If Mid(oRsTmp1.Fields!ApellidoMaterno, lnFor, 1) = lcEneMal Then
                 lcApellidoMaterno = lcApellidoMaterno & lcEneOK
              Else
                 lcApellidoMaterno = lcApellidoMaterno & Mid(oRsTmp1.Fields!ApellidoMaterno, lnFor, 1)
              End If
          Next
          lcPrimerNombre = ""
          For lnFor = 1 To Len(oRsTmp1.Fields!PrimerNombre)
              If Mid(oRsTmp1.Fields!PrimerNombre, lnFor, 1) = lcEneMal Then
                 lcPrimerNombre = lcPrimerNombre & lcEneOK
              Else
                 lcPrimerNombre = lcPrimerNombre & Mid(oRsTmp1.Fields!PrimerNombre, lnFor, 1)
              End If
          Next
          lcSegundoNombre = ""
          If Not IsNull(oRsTmp1.Fields!SegundoNombre) Then
                For lnFor = 1 To Len(oRsTmp1.Fields!SegundoNombre)
                    If Mid(oRsTmp1.Fields!SegundoNombre, lnFor, 1) = lcEneMal Then
                       lcSegundoNombre = lcSegundoNombre & lcEneOK
                    Else
                       lcSegundoNombre = lcSegundoNombre & Mid(oRsTmp1.Fields!SegundoNombre, lnFor, 1)
                    End If
                Next
          End If
          lcDireccionDomicilio = ""
          If Not IsNull(oRsTmp1.Fields!DireccionDomicilio) Then
                For lnFor = 1 To Len(oRsTmp1.Fields!DireccionDomicilio)
                    If Mid(oRsTmp1.Fields!DireccionDomicilio, lnFor, 1) = lcEneMal Then
                       lcDireccionDomicilio = lcDireccionDomicilio & lcEneOK
                    Else
                       lcDireccionDomicilio = lcDireccionDomicilio & Mid(oRsTmp1.Fields!DireccionDomicilio, lnFor, 1)
                    End If
                Next
          End If
          lcNroDocumento = ""
          If Not IsNull(oRsTmp1.Fields!NroDocumento) Then
                For lnFor = 1 To Len(oRsTmp1.Fields!NroDocumento)
                    If Mid(oRsTmp1.Fields!NroDocumento, lnFor, 1) = lcEneMal Then
                       lcNroDocumento = lcNroDocumento & lcEneOK
                    Else
                       lcNroDocumento = lcNroDocumento & Mid(oRsTmp1.Fields!NroDocumento, lnFor, 1)
                    End If
                Next
          End If
          
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = wxConexionRed
                .CommandTimeout = 150
                .CommandText = "PacientesModificarDatosEneMal"
                Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oRsTmp1.Fields!idPaciente): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, lcApellidoPaterno): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, lcApellidoMaterno): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, lcPrimerNombre): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, lcDireccionDomicilio): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, lcNroDocumento): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, lcSegundoNombre): .Parameters.Append oParameter
                .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
          
          oRsTmp1.MoveNext
          DoEvents
          If ProgressBar1.Value < lnRegistros Then
             ProgressBar1.Value = ProgressBar1.Value + 1
          End If
          Me.Refresh
       Loop
    End If
    oRsTmp1.Close

End Sub

Private Sub cmdProcesaLabHuaral_Click()
    If UCase(txtClave.Text) = "DEBB" Then
       Dim mo_ReglasCaja  As New SIGHNegocios.ReglasCaja
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oRstmpMDB1 As New Recordset
       Dim oRstmpMDB2 As New Recordset
       Dim oConexion As New Connection
       Dim oConexHBT As New Connection
       Dim oPartidasPresupuestales As New PartidasPresupuestales
       Dim lnFor As Long, lcHoraInicio As String, lcHoraActual As String, lcSql As String
       Dim ldFechaInicio As Date, ldFechaFinal As Date, ldFechaMaxima As Date
       Dim lnCorrelativo As Long, lcMensaje As String
       Dim lcCodigoV As String, lcCodigo As String, lnIdProductoCpt As Long, lnIdProducto As Long
       
       Me.MousePointer = 11
       ProgressBar2.Min = 0
       ProgressBar2.Max = 367
       lcHoraInicio = Time
       '
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       
        oConexHBT.CommandTimeout = 300
        oConexHBT.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablas nuevas galenhos.mdb;"
       
       lcSql = "delete from labitemscpt"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "delete from labitems"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "delete from labitemsgrupos"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       
       
       lcSql = "select * from labitems"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "select * from labitems"
       oRstmpMDB1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       If oRstmpMDB1.RecordCount > 0 Then
          oRstmpMDB1.MoveFirst
          Do While Not oRstmpMDB1.EOF
             oRsTmp1.AddNew
             oRsTmp1.Fields!idItem = oRstmpMDB1.Fields!idItem
             oRsTmp1.Fields!Item = oRstmpMDB1.Fields!Item
             oRsTmp1.Fields!idProductoCpt = oRstmpMDB1.Fields!idProductoCpt
             oRsTmp1.Update
             oRstmpMDB1.MoveNext
          Loop
       End If
       oRstmpMDB1.Close
       oRsTmp1.Close
       
       lcSql = "select * from labitemsgrupos"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "select * from labitemsgrupos"
       oRstmpMDB1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       If oRstmpMDB1.RecordCount > 0 Then
          oRstmpMDB1.MoveFirst
          Do While Not oRstmpMDB1.EOF
             oRsTmp1.AddNew
             oRsTmp1.Fields!idItemGrupo = oRstmpMDB1.Fields!idItemGrupo
             oRsTmp1.Fields!Grupo = oRstmpMDB1.Fields!Grupo
             oRsTmp1.Update
             oRstmpMDB1.MoveNext
          Loop
       End If
       oRstmpMDB1.Close
       oRsTmp1.Close
       
       lcMensaje = ""
       lcSql = "select * from labitemsCpt"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "select * from laboratorio"
       oRstmpMDB1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       If oRstmpMDB1.RecordCount > 0 Then
          oRstmpMDB1.MoveFirst
          Do While Not oRstmpMDB1.EOF
             lcCodigoV = Trim(oRstmpMDB1!campo1)
             lcCodigo = Trim(oRstmpMDB1!campo3)
             lnIdProductoCpt = 0
             lcSql = "select * from factCatalogoServicios where codigo='" & lcCodigoV & "'"
             oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
             If oRsTmp2.RecordCount > 0 Then
                lnIdProductoCpt = oRsTmp2!idProducto
                lcSql = "update factCatalogoServicios set codigo='ER" & lcCodigo & "' where codigo='" & lcCodigo & "'"
                oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                oRsTmp2.Fields!Codigo = lcCodigo
                oRsTmp2.Update
             Else
                oRsTmp2.Close
                lcSql = "select * from factCatalogoServicios where codigo='" & lcCodigo & "'"
                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                If oRsTmp2.RecordCount > 0 Then
                   lnIdProductoCpt = oRsTmp2!idProducto
                End If
             End If
             oRsTmp2.Close
             If lnIdProductoCpt > 0 Then
                lnIdProducto = 0
                lcSql = "select * from factCatalogoServicios where codigo='" & lcCodigo & "'"
                oRstmpMDB2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                If oRstmpMDB2.RecordCount > 0 Then
                   lnIdProducto = oRstmpMDB2!idProducto
                End If
                oRstmpMDB2.Close
                lcSql = "select * from labItemsCpt where idproductoCpt=" & lnIdProducto
                oRstmpMDB2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
                If oRstmpMDB2.RecordCount > 0 Then
                   oRstmpMDB2.MoveFirst
                   Do While Not oRstmpMDB2.EOF
                        oRsTmp1.AddNew
                        oRsTmp1.Fields!idProductoCpt = lnIdProductoCpt
                        oRsTmp1.Fields!ordenXresultado = oRstmpMDB2.Fields!ordenXresultado
                        oRsTmp1.Fields!idGrupo = oRstmpMDB2.Fields!idGrupo
                        oRsTmp1.Fields!idItemGrupo = oRstmpMDB2.Fields!idItemGrupo
                        oRsTmp1.Fields!idItem = oRstmpMDB2.Fields!idItem
                        oRsTmp1.Fields!ValorSiEsCombo = oRstmpMDB2.Fields!ValorSiEsCombo
                        oRsTmp1.Fields!ValorReferencial = oRstmpMDB2.Fields!ValorReferencial
                        oRsTmp1.Fields!Metodo = oRstmpMDB2.Fields!Metodo
                        oRsTmp1.Fields!SoloNumero = oRstmpMDB2.Fields!SoloNumero
                        oRsTmp1.Fields!SoloTexto = oRstmpMDB2.Fields!SoloTexto
                        oRsTmp1.Fields!SoloCombo = oRstmpMDB2.Fields!SoloCombo
                        oRsTmp1.Fields!SoloCheck = oRstmpMDB2.Fields!SoloCheck
                        oRsTmp1.Update
                        oRstmpMDB2.MoveNext
                    Loop
                Else
                    lcMensaje = lcMensaje & Chr(13) & "resultado Huaral no existe: " & oRstmpMDB1!campo3
                End If
                oRstmpMDB2.Close
             Else
                lcMensaje = lcMensaje & Chr(13) & "codigo: " & oRstmpMDB1!campo1
             End If
             oRstmpMDB1.MoveNext
          Loop
       End If
       oRstmpMDB1.Close
       If lcMensaje <> "" Then
          MsgBox lcMensaje
       End If
       oConexion.Close
       Unload Me
    Else
       MsgBox "no es la clave, recuerde que limpia los RESULTADOS, fijese que no tiene HISTORICOS"
    End If
End Sub


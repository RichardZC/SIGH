VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Begin VB.Form CuentaAtencionDetalle 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6285
      Left            =   60
      TabIndex        =   10
      Top             =   1110
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   11086
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos de la atencion"
      TabPicture(0)   =   "CuentaAtencionDetalle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos del paciente"
      TabPicture(1)   =   "CuentaAtencionDetalle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosPaciente"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDatosPaciente 
         Caption         =   "Datos del paciente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3285
         Left            =   -74820
         TabIndex        =   59
         Top             =   390
         Width           =   11715
         Begin VB.TextBox txtNombreMadre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7590
            MaxLength       =   25
            TabIndex        =   69
            Top             =   1335
            Width           =   3885
         End
         Begin VB.TextBox txtNombrePadre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7590
            MaxLength       =   25
            TabIndex        =   68
            Top             =   960
            Width           =   3885
         End
         Begin VB.TextBox txtSegundoNombre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   35
            TabIndex        =   67
            Top             =   1350
            Width           =   4155
         End
         Begin VB.TextBox txtTelefono 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7590
            MaxLength       =   10
            TabIndex        =   66
            Top             =   2070
            Width           =   1515
         End
         Begin VB.TextBox txtApellidoMaterno 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   35
            TabIndex        =   65
            Top             =   615
            Width           =   4155
         End
         Begin VB.TextBox txtApellidoPaterno 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   35
            TabIndex        =   64
            Top             =   240
            Width           =   4155
         End
         Begin VB.TextBox txtPrimerNombre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   35
            TabIndex        =   63
            Top             =   975
            Width           =   4155
         End
         Begin VB.TextBox txtNroDocumento 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            MaxLength       =   9
            TabIndex        =   62
            Top             =   2445
            Width           =   1485
         End
         Begin VB.TextBox txtTercerNombre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   35
            TabIndex        =   61
            Top             =   1710
            Width           =   4155
         End
         Begin VB.TextBox txtObservacion 
            Appearance      =   0  'Flat
            Height          =   705
            Left            =   7590
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   2430
            Width           =   3885
         End
         Begin MSDataListLib.DataCombo cmbIdEstadoCivil 
            Bindings        =   "CuentaAtencionDetalle.frx":0038
            Height          =   315
            Left            =   4440
            TabIndex        =   70
            Top             =   2085
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            ListField       =   "Descripcion1"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox txtFechaNacimiento 
            Height          =   315
            Left            =   7590
            TabIndex        =   71
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdGradoInstruccion 
            Height          =   315
            Left            =   1770
            TabIndex        =   72
            Top             =   2820
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Descripcion1"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbIdDocIdentidad 
            Bindings        =   "CuentaAtencionDetalle.frx":0053
            Height          =   315
            Left            =   1770
            TabIndex        =   73
            Top             =   2445
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            ListField       =   "Descripcion1"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbIdProcedencia 
            Bindings        =   "CuentaAtencionDetalle.frx":006E
            Height          =   315
            Left            =   7590
            TabIndex        =   74
            Top             =   1695
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            ListField       =   "Descripcion1"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbIdTipoOcupacion 
            Bindings        =   "CuentaAtencionDetalle.frx":0087
            Height          =   315
            Left            =   7590
            TabIndex        =   75
            Top             =   600
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            ListField       =   "Descripcion1"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbIdTipoSexo 
            Height          =   315
            Left            =   1770
            TabIndex        =   76
            Top             =   2085
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   93
            Top             =   2520
            Width           =   1545
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupación:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   92
            Top             =   630
            Width           =   1605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tercer Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   91
            Top             =   1755
            Width           =   1605
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Procedencia:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   90
            Top             =   1740
            Width           =   1545
         End
         Begin VB.Label Label36 
            Caption         =   "Nº"
            Height          =   225
            Left            =   3900
            TabIndex        =   89
            Top             =   2520
            Width           =   345
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Madre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   88
            Top             =   1395
            Width           =   1545
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Padre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   87
            Top             =   1005
            Width           =   1545
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Documento:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   86
            Top             =   2490
            Width           =   1605
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Segundo Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   85
            Top             =   1395
            Width           =   1605
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   84
            Top             =   2130
            Width           =   1605
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   83
            Top             =   2130
            Width           =   1545
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Grado de Instruc:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   82
            Top             =   2880
            Width           =   1605
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3450
            TabIndex        =   81
            Top             =   2130
            Width           =   870
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nacimiento:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6060
            TabIndex        =   80
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Materno:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   79
            Top             =   645
            Width           =   1605
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Primer Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   78
            Top             =   1035
            Width           =   1605
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Paterno:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   77
            Top             =   270
            Width           =   1605
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del seguro"
         Height          =   1125
         Left            =   150
         TabIndex        =   48
         Top             =   3840
         Width           =   11745
         Begin VB.TextBox txtNroPoliza 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1830
            TabIndex        =   50
            Top             =   600
            Width           =   1305
         End
         Begin VB.TextBox txtNroPlacaAutomovil 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4320
            TabIndex        =   49
            Top             =   630
            Width           =   1305
         End
         Begin MSDataListLib.DataCombo cmbIdFuenteFinaciamiento 
            Height          =   315
            Left            =   1830
            TabIndex        =   53
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbIdPlan 
            Height          =   315
            Left            =   6240
            TabIndex        =   54
            Top             =   270
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label7 
            Caption         =   "Fuente Financiamiento"
            Height          =   315
            Left            =   150
            TabIndex        =   56
            Top             =   300
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Cobertura"
            Height          =   255
            Left            =   5040
            TabIndex        =   55
            Top             =   300
            Width           =   1665
         End
         Begin VB.Label Label13 
            Caption         =   "Nº Autorización"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label16 
            Caption         =   "Nº Placa"
            Height          =   255
            Left            =   3330
            TabIndex        =   51
            Top             =   690
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de la cuenta"
         Height          =   1125
         Left            =   150
         TabIndex        =   37
         Top             =   2700
         Width           =   11715
         Begin MSMask.MaskEdBox txtFechaApertura 
            Height          =   315
            Left            =   6690
            TabIndex        =   38
            Top             =   300
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoFinanciamiento 
            Height          =   315
            Left            =   1830
            TabIndex        =   39
            Top             =   270
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSMask.MaskEdBox txtFechaCierre 
            Height          =   315
            Left            =   6690
            TabIndex        =   40
            Top             =   660
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdEstadoCuenta 
            Height          =   315
            Left            =   9720
            TabIndex        =   41
            Top             =   270
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Text            =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "Estado Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   8280
            TabIndex        =   47
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Nro Cuenta"
            Height          =   285
            Left            =   150
            TabIndex        =   46
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label lblCuentaAtencion 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1830
            TabIndex        =   45
            Top             =   630
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Financiamiento"
            Height          =   255
            Left            =   90
            TabIndex        =   44
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha apertura"
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
            Left            =   5160
            TabIndex        =   43
            Top             =   330
            Width           =   1545
         End
         Begin VB.Label Label12 
            Caption         =   "Fecha Cierre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5220
            TabIndex        =   42
            Top             =   690
            Width           =   1155
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos de la atención"
         Height          =   2265
         Left            =   150
         TabIndex        =   11
         Top             =   390
         Width           =   11715
         Begin PVCOMBOLibCtl.PVComboBox PVComboBox1 
            Height          =   315
            Left            =   1830
            TabIndex        =   36
            Top             =   300
            Width           =   1875
            _Version        =   524288
            _cx             =   3307
            _cy             =   556
            Appearance      =   0
            Enabled         =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            Locked          =   0   'False
            Style           =   0
            Sorted          =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowPictures    =   0   'False
            ColumnHeaders   =   0   'False
            PrimaryColumn   =   0
            VisibleItems    =   10
            ColumnHeaderHeight=   20
            ListMember      =   ""
            ColumnHeaderForeColor=   0
            ColumnHeaderBackColor=   13160660
            SelectedForeColor=   16777215
            SelectedBackColor=   6956042
            AlternateBackColor=   16777215
            ItemLabelStyle  =   1
            ItemLabelType   =   0
            ItemLabelWidth  =   40
            ItemLabelForeColor=   0
            ItemLabelBackColor=   13160660
            ColumnHeaderStyle=   0
            VerticalGridLines=   -1  'True
            HorizontalGridLines=   -1  'True
            ColumnResize    =   0   'False
            ItemLabelResize =   0   'False
            AllowDBAutoConfig=   -1  'True
            GridLineColor   =   13421772
            List            =   ""
            NullString      =   "[NULL]"
            DropShadow      =   -1  'True
            Text            =   ""
            SortOnColumnHeaderClick=   0   'False
            DropEffect      =   1
            ColumnCount     =   1
            Column0.Heading =   ""
            Column0.Width   =   40
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   ""
            Column0.Format  =   ""
            Column0.Bound   =   0   'False
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            SortKey1.Column =   -1
            SortKey1.Ascending=   -1  'True
            SortKey1.CaseInsensitive=   -1  'True
            SortKey2.Column =   -1
            SortKey2.Ascending=   -1  'True
            SortKey2.CaseInsensitive=   -1  'True
            SortKey3.Column =   -1
            SortKey3.Ascending=   -1  'True
            SortKey3.CaseInsensitive=   -1  'True
            BoundColumn     =   ""
            Border          =   -1  'True
            VertAlign       =   1
            Format          =   ""
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   315
            Left            =   7170
            TabIndex        =   33
            Top             =   1740
            Width           =   315
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   315
            Left            =   2880
            TabIndex        =   31
            Top             =   1380
            Width           =   315
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   315
            Left            =   2880
            TabIndex        =   29
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox txtIdEstablecimientoOrigen 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6090
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1740
            Width           =   1000
         End
         Begin VB.TextBox txtIdServicioIngreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1830
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   1020
            Width           =   1000
         End
         Begin VB.TextBox txtIdMedicoIngreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1830
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   1380
            Width           =   1005
         End
         Begin VB.TextBox txtEdadEnDias 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5160
            TabIndex        =   14
            Top             =   660
            Width           =   735
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   3030
            TabIndex        =   12
            Top             =   660
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaIngreso 
            Height          =   315
            Left            =   1830
            TabIndex        =   13
            Top             =   660
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoServicio 
            Height          =   315
            Left            =   8850
            TabIndex        =   15
            Top             =   1020
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdEspecialidadMedico 
            Height          =   315
            Left            =   8850
            TabIndex        =   20
            Top             =   1380
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoReferenciaOrigen 
            Height          =   315
            Left            =   1830
            TabIndex        =   26
            Top             =   1740
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdViaAdmision 
            Height          =   315
            Left            =   8850
            TabIndex        =   58
            Top             =   660
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label47 
            Caption         =   "Via Admisión"
            Height          =   285
            Left            =   7290
            TabIndex        =   57
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label46 
            Caption         =   "Cita"
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label45 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   7560
            TabIndex        =   34
            Top             =   1740
            Width           =   3975
         End
         Begin VB.Label Label44 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3240
            TabIndex        =   32
            Top             =   1380
            Width           =   3975
         End
         Begin VB.Label Label43 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3240
            TabIndex        =   30
            Top             =   1020
            Width           =   3975
         End
         Begin VB.Label lblIdEstablecimientoOrigen 
            Caption         =   "Establecimiento refer. origen"
            Height          =   315
            Left            =   3870
            TabIndex        =   27
            Top             =   1800
            Width           =   2205
         End
         Begin VB.Label lblIdTipoReferenciaOrigen 
            Caption         =   "Tipo refer. origen"
            Height          =   315
            Left            =   180
            TabIndex        =   25
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblIdServicioIngreso 
            Caption         =   "Servicio ingreso"
            Height          =   315
            Left            =   180
            TabIndex        =   23
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label lblIdMedicoIngreso 
            Caption         =   "Medico ingreso"
            Height          =   315
            Left            =   180
            TabIndex        =   21
            Top             =   1410
            Width           =   1335
         End
         Begin VB.Label lblIdEspecialidadMedico 
            Caption         =   "Especialidad del medico"
            Height          =   315
            Left            =   7260
            TabIndex        =   19
            Top             =   1410
            Width           =   1815
         End
         Begin VB.Label lblIdTipoServicio 
            Caption         =   "Tipo de servicio"
            Height          =   315
            Left            =   7260
            TabIndex        =   18
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label lblFechaIngreso 
            Caption         =   "Fecha ingreso"
            Height          =   315
            Left            =   180
            TabIndex        =   17
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label lblEdadEnDias 
            Caption         =   "Edad "
            Height          =   315
            Left            =   4170
            TabIndex        =   16
            Top             =   690
            Width           =   1005
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2355
         Left            =   -74850
         TabIndex        =   94
         Top             =   3750
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   4154
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos de domicilio"
         TabPicture(0)   =   "CuentaAtencionDetalle.frx":00A0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FraDomicilio"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos de procedencia"
         TabPicture(1)   =   "CuentaAtencionDetalle.frx":00BC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraProcedencia"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Datos de nacimiento"
         TabPicture(2)   =   "CuentaAtencionDetalle.frx":00D8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraNacimiento"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame fraProcedencia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1785
            Left            =   -74850
            TabIndex        =   130
            Top             =   390
            Width           =   11415
            Begin VB.CheckBox chkIgualQueDomicilio 
               Appearance      =   0  'Flat
               Caption         =   "Igual que el domicilio"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1680
               TabIndex        =   131
               Top             =   1020
               Width           =   2775
            End
            Begin MSDataListLib.DataCombo cmbIdCentroPobladoProcedencia 
               Bindings        =   "CuentaAtencionDetalle.frx":00F4
               Height          =   315
               Left            =   1650
               TabIndex        =   132
               Top             =   600
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdDistritoProcedencia 
               Bindings        =   "CuentaAtencionDetalle.frx":010E
               Height          =   315
               Left            =   8490
               TabIndex        =   133
               Top             =   240
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdProvinciaProcedencia 
               Bindings        =   "CuentaAtencionDetalle.frx":0128
               Height          =   315
               Left            =   4860
               TabIndex        =   134
               Top             =   240
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdDepartamentoProcedencia 
               Bindings        =   "CuentaAtencionDetalle.frx":0143
               Height          =   315
               Left            =   1650
               TabIndex        =   135
               Top             =   240
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Departamento:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   139
               Top             =   300
               Width           =   1605
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Provincia:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4020
               TabIndex        =   138
               Top             =   300
               Width           =   705
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7830
               TabIndex        =   137
               Top             =   300
               Width           =   600
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Centro Poblado: "
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   136
               Top             =   660
               Width           =   1605
            End
         End
         Begin VB.Frame fraNacimiento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1785
            Left            =   -74850
            TabIndex        =   120
            Top             =   420
            Width           =   11385
            Begin VB.CheckBox chkIgualUQueDomicilioNac 
               Appearance      =   0  'Flat
               Caption         =   "Igual que el domicilio"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1680
               TabIndex        =   121
               Top             =   1020
               Width           =   2775
            End
            Begin MSDataListLib.DataCombo cmbIdCentroPobladoNacimiento 
               Bindings        =   "CuentaAtencionDetalle.frx":0161
               Height          =   315
               Left            =   1650
               TabIndex        =   122
               Top             =   600
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdDistritoNacimiento 
               Bindings        =   "CuentaAtencionDetalle.frx":017B
               Height          =   315
               Left            =   8490
               TabIndex        =   123
               Top             =   240
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdProvinciaNacimiento 
               Bindings        =   "CuentaAtencionDetalle.frx":0195
               Height          =   315
               Left            =   4860
               TabIndex        =   124
               Top             =   240
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdDepartamentoNacimiento 
               Bindings        =   "CuentaAtencionDetalle.frx":01B0
               Height          =   315
               Left            =   1650
               TabIndex        =   125
               Top             =   240
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Departamento:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   129
               Top             =   300
               Width           =   1605
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Provincia:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4020
               TabIndex        =   128
               Top             =   300
               Width           =   705
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7830
               TabIndex        =   127
               Top             =   300
               Width           =   600
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Centro Poblado: "
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   126
               Top             =   660
               Width           =   1605
            End
         End
         Begin VB.Frame FraDomicilio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1785
            Left            =   150
            TabIndex        =   95
            Top             =   420
            Width           =   11415
            Begin VB.TextBox txtEtapaDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   10200
               MaxLength       =   5
               TabIndex        =   102
               Top             =   1320
               Width           =   1000
            End
            Begin VB.TextBox txtSectorDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   8490
               MaxLength       =   5
               TabIndex        =   101
               Top             =   1320
               Width           =   1000
            End
            Begin VB.TextBox txtPisoDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6465
               MaxLength       =   5
               TabIndex        =   100
               Top             =   1320
               Width           =   1000
            End
            Begin VB.TextBox txtLoteDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   10200
               MaxLength       =   5
               TabIndex        =   99
               Top             =   960
               Width           =   1000
            End
            Begin VB.TextBox txtManzanaDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   8490
               MaxLength       =   5
               TabIndex        =   98
               Top             =   960
               Width           =   1000
            End
            Begin VB.TextBox txtNroDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6465
               MaxLength       =   5
               TabIndex        =   97
               Top             =   960
               Width           =   1000
            End
            Begin VB.TextBox txtDireccionDomicilio 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1650
               MaxLength       =   50
               TabIndex        =   96
               Top             =   960
               Width           =   4005
            End
            Begin MSDataListLib.DataCombo cmbIdCentroPobladoDomicilio 
               Bindings        =   "CuentaAtencionDetalle.frx":01CE
               Height          =   315
               Left            =   1650
               TabIndex        =   103
               Top             =   600
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdDistritoDomicilio 
               Bindings        =   "CuentaAtencionDetalle.frx":01E8
               Height          =   315
               Left            =   8490
               TabIndex        =   104
               Top             =   240
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdProvinciaDomicilio 
               Bindings        =   "CuentaAtencionDetalle.frx":0202
               Height          =   315
               Left            =   4860
               TabIndex        =   105
               Top             =   240
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdDepartamentoDomicilio 
               Bindings        =   "CuentaAtencionDetalle.frx":021D
               Height          =   315
               Left            =   1650
               TabIndex        =   106
               Top             =   240
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbIdPaisDomicilio 
               Height          =   315
               Left            =   8490
               TabIndex        =   107
               Top             =   600
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin VB.Label lblEtapaDomicilio 
               Caption         =   "Etapa"
               Height          =   315
               Left            =   9645
               TabIndex        =   119
               Top             =   1350
               Width           =   525
            End
            Begin VB.Label lblSectorDomicilio 
               Caption         =   "Sector"
               Height          =   315
               Left            =   7605
               TabIndex        =   118
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lblPisoDomicilio 
               Caption         =   "Piso"
               Height          =   315
               Left            =   5940
               TabIndex        =   117
               Top             =   1350
               Width           =   1005
            End
            Begin VB.Label lblLoteDomicilio 
               Appearance      =   0  'Flat
               Caption         =   "Lote"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   9660
               TabIndex        =   116
               Top             =   990
               Width           =   555
            End
            Begin VB.Label lblManzanaDomicilio 
               Appearance      =   0  'Flat
               Caption         =   "Manzana"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   7620
               TabIndex        =   115
               Top             =   990
               Width           =   795
            End
            Begin VB.Label lblNroDomicilio 
               Appearance      =   0  'Flat
               Caption         =   "Nro"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   5940
               TabIndex        =   114
               Top             =   990
               Width           =   465
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "País:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7740
               TabIndex        =   113
               Top             =   630
               Width           =   600
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dirección:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   990
               Width           =   1605
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Departamento:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   111
               Top             =   300
               Width           =   1605
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Provincia:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4020
               TabIndex        =   110
               Top             =   240
               Width           =   705
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7710
               TabIndex        =   109
               Top             =   300
               Width           =   600
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Centro Poblado: "
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   660
               Width           =   1605
            End
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   90
      TabIndex        =   7
      Top             =   7380
      Width           =   11985
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4680
         Picture         =   "CuentaAtencionDetalle.frx":023B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   2670
         Picture         =   "CuentaAtencionDetalle.frx":0727
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del paciente"
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3270
         TabIndex        =   6
         Top             =   270
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.CommandButton btnBuscarHistoriaClinica 
         Caption         =   "..."
         Height          =   315
         Left            =   8190
         TabIndex        =   2
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtNroHistoriaClinica 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   270
         Width           =   1250
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   630
         Width           =   6525
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Historia:"
         Height          =   225
         Left            =   330
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Nombres"
         Height          =   285
         Left            =   330
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CuentaAtencionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POCuentasAtencion
'        Autor: William Castro Grijalva
'        Fecha: 31/08/2004 09:24:18 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

'Dim mo_Teclado As New SIGHComun.Teclado
'Dim mo_Formulario As New SIGHComun.Formulario
'Dim mo_CuentasAtencion As New DOCuentaAtencion
'Dim ml_IdUsuario As Long
'Dim ms_MensajeError As String
'Dim mi_Opcion As sghOpciones
'Dim mb_ExistenDatos As Boolean
'Dim ml_IdCuentaAtencion As Long
'Sub CargarComboBoxes()
'Dim sSQL As String
'Dim sMensaje As String
'
'       cmbIdFuenteFinanciamiento.BoundColumn = "IdFuenteFinanciamiento"
'       cmbIdFuenteFinanciamiento.ListField = "DescripcionLarga"
'       Set cmbIdFuenteFinanciamiento.RowSource = "mo_AdminServiciosHosp.XXXXXObtenerTodos()"
'       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError + Chr(13)
'       cmbIdEstado.BoundColumn = "IdEstado"
'       cmbIdEstado.ListField = "DescripcionLarga"
'       Set cmbIdEstado.RowSource = "mo_AdminServiciosHosp.XXXXXObtenerTodos()"
'       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError + Chr(13)
'       cmbIdTipoFinanciamiento.BoundColumn = "IdTipoFinanciamiento"
'       cmbIdTipoFinanciamiento.ListField = "DescripcionLarga"
'       Set cmbIdTipoFinanciamiento.RowSource = "mo_AdminServiciosHosp.XXXXXObtenerTodos()"
'       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError + Chr(13)
'       cmbIdPlan.BoundColumn = "IdPlan"
'       cmbIdPlan.ListField = "DescripcionLarga"
'       Set cmbIdPlan.RowSource = "mo_AdminServiciosHosp.XXXXXObtenerTodos()"
'       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError + Chr(13)
'
'       cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
'       cmbIdTipoServicio.ListField = "DescricpcionLarga"
'       Set cmbIdTipoServicio.RowSource = "mo_AdminServiciosHosp.XXXXXObtenerTodos()"
'       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError + Chr(13)
'
'       If sMensaje <> "" Then
'           MsgBox mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption
'       End If
'
'End Sub
'Property Let ExistenDatos(bValue As Boolean)
'   mb_ExistenDatos = bValue
'End Property
'Property Get ExistenDatos() As Boolean
'   ExistenDatos = mb_ExistenDatos
'End Property
'Property Let Opcion(iValue As sghOpciones)
'   mi_Opcion = iValue
'End Property
'Property Get Opcion() As sghOpciones
'   Opcion = mi_Opcion
'End Property
'Property Let MensajeError(sValue As String)
'   ms_MensajeError = sValue
'End Property
'Property Get MensajeError() As String
'   MensajeError = ms_MensajeError
'End Property
'Property Let IdUsuario(lValue As Long)
'   ml_IdUsuario = lValue
'End Property
'Property Get IdUsuario() As Long
'   IdUsuario = ml_IdUsuario
'End Property
'Property Let IdCuentaAtencion(lValue As Long)
'   ml_IdCuentaAtencion = lValue
'End Property
'Property Get IdCuentaAtencion() As Long
'   IdCuentaAtencion = ml_IdCuentaAtencion
'End Property
'Property Let IdAtencion(lValue As Long)
'   ml_IdAtencion = lValue
'End Property
'Property Get IdAtencion() As Long
'   IdAtencion = ml_IdAtencion
'End Property
'Dim ml_IdAtencion As Long
'
'Private Sub cmbIdFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, cmbIdFuenteFinanciamiento
'AdministrarKeyPreview KeyCode
'End Sub
'
'Private Sub cmbIdFuenteFinanciamiento_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub cmbIdEstado_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEstado
'AdministrarKeyPreview KeyCode
'End Sub
'
'Private Sub cmbIdEstado_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub cmbIdTipoFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoFinanciamiento
'AdministrarKeyPreview KeyCode
'End Sub
'
'Private Sub cmbIdTipoFinanciamiento_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub txtNombreTitular_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtNombreTitular
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtNombreTitular_LostFocus()
'txtNombreTitular.Text = mo_Teclado.CapitalizarNombres(txtNombreTitular.Text)
'   mo_Formulario.MarcarComoVacio txtNombreTitular
'End Sub
'
'Private Sub txtNombreTitular_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub txtNroPlaca_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtNroPlaca
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtNroPlaca_LostFocus()
'   mo_Formulario.MarcarComoVacio txtNroPlaca
'End Sub
'
'Private Sub txtNroPlaca_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub txtNroAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtNroAutorizacion
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtNroAutorizacion_LostFocus()
'   mo_Formulario.MarcarComoVacio txtNroAutorizacion
'End Sub
'
'Private Sub txtNroAutorizacion_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub txtFechaCierre_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtFechaCierre
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtFechaCierre_LostFocus()
'   mo_Formulario.MarcarComoVacio txtFechaCierre
'End Sub
'
'Private Sub txtFechaCierre_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub txtFechaApertura_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtFechaApertura
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtFechaApertura_LostFocus()
'   mo_Formulario.MarcarComoVacio txtFechaApertura
'End Sub
'
'Private Sub txtFechaApertura_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'
'Private Sub cmbIdPlan_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, cmbIdPlan
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub cmbIdPlan_LostFocus()
'   If cmbIdPlan.Text <> "" Then
'       cmbIdPlan.BoundText = Val(Split(cmbIdPlan.Text, " = ")(0))
'   End If
'   mo_Formulario.MarcarComoVacio cmbIdPlan
'End Sub
'
'Private Sub cmbIdPlan_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
''------------------------------------------------------------------------------------
''   CargarDatosAlFormulario
''   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub CargarDatosAlFormulario()
'
' Select Case mi_Opcion
'     Case sghAgregar
'     Case sghModificar
'         CargarDatosALosControles
'     Case sghConsultar
'         CargarDatosALosControles
'     Case sghEliminar
'         CargarDatosALosControles
' End Select
'End Sub
'
''------------------------------------------------------------------------------------
''   CargarDatosAlFormulario
''   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub Form_Load()
'       Select Case mi_Opcion
'       Case sghAgregar
'           Me.Caption = "Agregar CuentasAtencion"
'       Case sghModificar
'           Me.Caption = "Modificar CuentasAtencion"
'       Case sghConsultar
'           Me.Caption = "Consultar CuentasAtencion"
'       Case sghEliminar
'           Me.Caption = "Eliminar CuentasAtencion"
'       End Select
'
'       CargarComboBoxes
'       CargarDatosAlFormulario
'       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
'End Sub
'
''------------------------------------------------------------------------------------
''   CargarDatosAlFormulario
''   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub Form_Activate()
'   If mi_Opcion <> sghAgregar Then
'       If Not mb_ExistenDatos Then
'           Me.Visible = False
'       End If
'   End If
'End Sub
'Sub AdministrarKeyPreview(KeyCode As Integer)
'   Select Case KeyCode
'       Case vbKeyEscape
'           btnCancelar_Click
'       Case vbKeyF2
'           btnAceptar_Click
'       End Select
'End Sub
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   AdministrarKeyPreview KeyCode
'End Sub
'
'Private Sub btnAceptar_Click()
'   Select Case mi_Opcion
'   Case sghAgregar
'       If ValidarDatosObligatorios() Then
'           If ValidarReglas() Then
'               If AgregarDatos() Then
'                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
'                   LimpiarFormulario
'               Else
'                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'       End If
'   Case sghModificar
'       If ValidarDatosObligatorios() Then
'           If ValidarReglas() Then
'               If ModificarDatos() Then
'                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
'                   Me.Visible = False
'               Else
'                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'       End If
'   Case sghEliminar
'           If ValidarReglas() Then
'               If EliminarDatos() Then
'                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
'                   Me.Visible = False
'               Else
'                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
'               End If
'           End If
'   End Select
'End Sub
'
'Private Sub btnCancelar_Click()
'   Me.Visible = False
'End Sub
'
'Function ValidarDatosObligatorios() As Boolean
'   Dim sMensaje As String
'   ValidarDatosObligatorios = False
'   If IdCuentaAtencion = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de IdCuentaAtencion" + Chr(13)
'   End If
'   If IdAtencion = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de IdAtencion" + Chr(13)
'   End If
'   If Me.cmbIdFuenteFinanciamiento.BoundText = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de IdFuenteFinanciamiento" + Chr(13)
'   End If
'   If Me.cmbIdEstado.BoundText = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de IdEstado" + Chr(13)
'   End If
'   If Me.cmbIdTipoFinanciamiento.BoundText = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de IdTipoFinanciamiento" + Chr(13)
'   End If
'   If Me.txtNombreTitular.Text = "" Then
'       sMensaje = sMensaje + "Ingrese el valor de NombreTitular" + Chr(13)
'   End If
'   If Me.txtNroPlaca.Text = "" Then
'       sMensaje = sMensaje + "Ingrese el valor de NroPlaca" + Chr(13)
'   End If
'   If Me.txtNroAutorizacion.Text = "" Then
'       sMensaje = sMensaje + "Ingrese el valor de NroAutorizacion" + Chr(13)
'   End If
'   If Me.txtFechaCierre.Text = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de FechaCierre" + Chr(13)
'   End If
'   If Me.txtFechaApertura.Text = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de FechaApertura" + Chr(13)
'   End If
'   If Me.cmbIdPlan.BoundText = 0 Then
'       sMensaje = sMensaje + "Ingrese el valor de IdPlan" + Chr(13)
'   End If
'   If sMensaje <> "" Then
'       MsgBox sMensaje, vbInformation, Me.Caption
'       Exit Function
'   End If
'   ValidarDatosObligatorios = True
'End Function
'Function ValidarReglas() As Boolean
'   ValidarReglas = False
'   ValidarReglas = True
'End Function
''------------------------------------------------------------------------------------
''   Cargar datos al objetos de datos
''   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub CargaDatosAlObjetosDeDatos()
'
'   With mo_Atenciones
'           .IdAtencion = Me.IdAtencion
'           .IdTipoCondicionALEstab = Me.cmbIdTipoCondicionALEstab.BoundText
'           .IdTipoCondicionAlServicio = Me.cmbIdTipoCondicionAlServicio.BoundText
'           .IdDestinoAtencion = Me.cmbIdDestinoAtencion.BoundText
'           .IdTipoReferenciaDestino = Me.cmbIdTipoReferenciaDestino.BoundText
'           .IdTipoReferenciaOrigen = Me.cmbIdTipoReferenciaOrigen.BoundText
'           .IdEstablecimientoDestino = Me.txtIdEstablecimientoDestino.Text
'           .IdEstablecimientoOrigen = Me.txtIdEstablecimientoOrigen.Text
'           .HoraEgreso = Me.txtHoraEgreso.Text
'           .FechaEgreso = Me.txtFechaEgreso.Text
'           .HoraIngreso = Me.txtHoraIngreso.Text
'           .FechaIngreso = Me.txtFechaIngreso.Text
'           .IdTipoServicio = Me.cmbIdTipoServicio.BoundText
'           .EdadEnDias = Me.txtEdadEnDias.Text
'           .IdPaciente = Me.txtIdPaciente.Text
'   End With
'
'
'   With mo_CuentasAtencion
'           .IdCuentaAtencion = Me.IdCuentaAtencion
'           .IdAtencion = Me.IdAtencion
'           .IdFuenteFinanciamiento = Me.cmbIdFuenteFinanciamiento.BoundText
'           .IdEstado = Me.cmbIdEstado.BoundText
'           .IdTipoFinanciamiento = Me.cmbIdTipoFinanciamiento.BoundText
'           .NombreTitular = Me.txtNombreTitular.Text
'           .NroPlaca = Me.txtNroPlaca.Text
'           .NroAutorizacion = Me.txtNroAutorizacion.Text
'           .FechaCierre = Me.txtFechaCierre.Text
'           .FechaApertura = Me.txtFechaApertura.Text
'           .IdPlan = Me.cmbIdPlan.BoundText
'   End With
'
'End Sub
'
''------------------------------------------------------------------------------------
''        Agregar Datos
''------------------------------------------------------------------------------------
'
'Function AgregarDatos() As Boolean
'
'   CargaDatosAlObjetosDeDatos
'   AgregarDatos = mo_AdminServiciosComunes.CuentasAtencionAgregar(mo_CuentasAtencion)
'
'End Function
'
''------------------------------------------------------------------------------------
''        Modificar Datos
''------------------------------------------------------------------------------------
'
'Function ModificarDatos() As Boolean
'
'   CargaDatosAlObjetosDeDatos
'   ModificarDatos = mo_AdminServiciosComunes.CuentasAtencionModificar(mo_CuentasAtencion)
'
'End Function
'
''------------------------------------------------------------------------------------
''        Eliminar Datos
''------------------------------------------------------------------------------------
'
'Function EliminarDatos() As Boolean
'
'   CargaDatosAlObjetosDeDatos
'   EliminarDatos = mo_AdminServiciosComunes.CuentasAtencionModificar(mo_CuentasAtencion)
'
'End Function
'
''------------------------------------------------------------------------------------
''   Llenar Datos Al Formulario
''   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub CargarDatosALosControles()
'
'        Set mo_CuentasAtencion = mo_AdminProgramacionMedica.SeleccionarTurnoPorId(Me.IdCuentaAtencion)
'        If mo_AdminServiciosComunes.MensajeError <> "" Then
'             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption"
'             mb_ExistenDatos = False
'             Exit Sub
'        End If
'        If Not mo_CuentasAtencion Is Nothing Then
'           With mo_CuentasAtencion
'           Me.IdCuentaAtencion = .IdCuentaAtencion
'           Me.IdAtencion = .IdAtencion
'           Me.cmbIdFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
'           Me.cmbIdEstado.BoundText = .IdEstado
'           Me.cmbIdTipoFinanciamiento.BoundText = .IdTipoFinanciamiento
'           Me.txtNombreTitular.Text = .NombreTitular
'           Me.txtNroPlaca.Text = .NroPlaca
'           Me.txtNroAutorizacion.Text = .NroAutorizacion
'           Me.txtFechaCierre.Text = .FechaCierre
'           Me.txtFechaApertura.Text = .FechaApertura
'           Me.cmbIdPlan.BoundText = .IdPlan
'               mb_ExistenDatos = True
'           End With
'       Else
'           mb_ExistenDatos = False
'           Exit Sub
'       End If
'
'        Set mo_Atenciones = mo_AdminProgramacionMedica.SeleccionarTurnoPorId(Me.IdAtencion)
'        If mo_AdminServiciosComunes.MensajeError <> "" Then
'             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption"
'             mb_ExistenDatos = False
'             Exit Sub
'        End If
'        If Not mo_Atenciones Is Nothing Then
'            With mo_Atenciones
'            Me.IdAtencion = .IdAtencion
'            Me.cmbIdTipoCondicionALEstab.BoundText = .IdTipoCondicionALEstab
'            Me.cmbIdTipoCondicionAlServicio.BoundText = .IdTipoCondicionAlServicio
'            Me.cmbIdDestinoAtencion.BoundText = .IdDestinoAtencion
'            Me.cmbIdTipoReferenciaDestino.BoundText = .IdTipoReferenciaDestino
'            Me.cmbIdTipoReferenciaOrigen.BoundText = .IdTipoReferenciaOrigen
'            Me.txtIdEstablecimientoDestino.Text = .IdEstablecimientoDestino
'            Me.txtIdEstablecimientoOrigen.Text = .IdEstablecimientoOrigen
'            Me.txtHoraEgreso.Text = .HoraEgreso
'            Me.txtFechaEgreso.Text = .FechaEgreso
'            Me.txtHoraIngreso.Text = .HoraIngreso
'            Me.txtFechaIngreso.Text = .FechaIngreso
'            Me.cmbIdTipoServicio.BoundText = .IdTipoServicio
'            Me.txtEdadEnDias.Text = .EdadEnDias
'            Me.txtIdPaciente.Text = .IdPaciente
'                mb_ExistenDatos = True
'            End With
'        Else
'            mb_ExistenDatos = False
'            Exit Sub
'        End If
'
'End Sub
'
''------------------------------------------------------------------------------------
''   Llenar Datos Al Formulario
''   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
''   Parámetros:     Ninguno
''------------------------------------------------------------------------------------
'
'Sub LimpiarFormulario()
'
'           Me.IdCuentaAtencion = 0
'           Me.IdAtencion = 0
'           Me.cmbIdFuenteFinanciamiento.BoundText = ""
'           Me.cmbIdEstado.BoundText = ""
'           Me.cmbIdTipoFinanciamiento.BoundText = ""
'           Me.txtNombreTitular.Text = ""
'           Me.txtNroPlaca.Text = ""
'           Me.txtNroAutorizacion.Text = ""
'           Me.txtFechaCierre.Text = ""
'           Me.txtFechaApertura.Text = ""
'           Me.cmbIdPlan.BoundText = ""
'
'End Sub
'
'
'Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub cmbIdTipoServicio_LostFocus()
'   If cmbIdTipoServicio.Text <> "" Then
'       cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
'   End If
'   mo_Formulario.MarcarComoVacio cmbIdTipoServicio
'End Sub
'
'Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
'Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtFechaIngreso_LostFocus()
'   mo_Formulario.MarcarComoVacio txtFechaIngreso
'End Sub
'
'Private Sub txtFechaIngreso_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'Private Sub txtHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtHoraIngreso
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtHoraIngreso_LostFocus()
'   mo_Formulario.MarcarComoVacio txtHoraIngreso
'End Sub
'
'Private Sub txtHoraIngreso_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'Private Sub txtEdadEnDias_KeyDown(KeyCode As Integer, Shift As Integer)
'   mo_Teclado.RealizarNavegacion KeyCode, txtEdadEnDias
'AdministrarKeyPreview KeyCode
'End Sub
'
'
'Private Sub txtEdadEnDias_LostFocus()
'   mo_Formulario.MarcarComoVacio txtEdadEnDias
'End Sub
'
'Private Sub txtEdadEnDias_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
'End Sub
'
Private Sub SSTab1_DblClick()

End Sub

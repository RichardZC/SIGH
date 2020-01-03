VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AdmisionAlojDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10890
   Icon            =   "AdmisionAlojDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   885
      Left            =   30
      TabIndex        =   16
      Top             =   4980
      Width           =   10845
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         DisabledPicture =   "AdmisionAlojDetalle.frx":0CCA
         DownPicture     =   "AdmisionAlojDetalle.frx":118E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5640
         Picture         =   "AdmisionAlojDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionAlojDetalle.frx":1B66
         DownPicture     =   "AdmisionAlojDetalle.frx":1FC6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4065
         Picture         =   "AdmisionAlojDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   150
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab tabAdmision 
      Height          =   4935
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "1. Ingreso (F10)"
      TabPicture(0)   =   "AdmisionAlojDetalle.frx":28B0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "UcPacienteDatosAloj1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "2. Egreso (F11)"
      TabPicture(1)   =   "AdmisionAlojDetalle.frx":28CC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin SISGalenPlus.UcPacienteDatosAloj UcPacienteDatosAloj1 
         Height          =   3195
         Left            =   210
         TabIndex        =   0
         Top             =   420
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   5636
      End
      Begin VB.Frame Frame1 
         Caption         =   "Egreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74850
         TabIndex        =   25
         Top             =   570
         Width           =   10485
         Begin VB.CommandButton cmdBuscaCamaEgreso 
            Caption         =   "..."
            Height          =   315
            Left            =   2550
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1290
            Width           =   315
         End
         Begin VB.TextBox txtNroCamaEgreso 
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
            Left            =   1620
            TabIndex        =   43
            Top             =   1290
            Width           =   885
         End
         Begin VB.TextBox txtNombreAcompañante 
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
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   33
            Top             =   2010
            Width           =   3930
         End
         Begin VB.TextBox txtDiasEstancia 
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
            Left            =   4980
            TabIndex        =   41
            Top             =   1650
            Width           =   525
         End
         Begin VB.ComboBox cmbIdDestinoAtencion 
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
            Left            =   1620
            TabIndex        =   29
            Top             =   540
            Width           =   3930
         End
         Begin VB.TextBox lblNombreMedicoEgreso 
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
            Left            =   1620
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   930
            Width           =   3000
         End
         Begin VB.TextBox lblNombreServicioEgreso 
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
            Left            =   2580
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   180
            Width           =   2970
         End
         Begin VB.TextBox txtIdMedicoEgreso 
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
            Left            =   4650
            TabIndex        =   27
            Top             =   930
            Width           =   885
         End
         Begin VB.TextBox txtIdServicioEgreso 
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
            Left            =   1620
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Width           =   885
         End
         Begin MSMask.MaskEdBox txtHoraEgreso 
            Height          =   315
            Left            =   3015
            TabIndex        =   32
            Top             =   1650
            Width           =   735
            _ExtentX        =   1296
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
         Begin MSMask.MaskEdBox txtFechaEgreso 
            Height          =   315
            Left            =   1620
            TabIndex        =   31
            Top             =   1650
            Width           =   1335
            _ExtentX        =   2355
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
         Begin VB.Label lblNroCamaEgreso 
            Caption         =   "Nro Cama egreso"
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
            TabIndex        =   44
            Top             =   1350
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "Estancia"
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
            Left            =   4230
            TabIndex        =   40
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Quien Recibe "
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
            TabIndex        =   39
            Top             =   2070
            Width           =   1455
         End
         Begin VB.Label Label43 
            Caption         =   "Médico egreso"
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
            TabIndex        =   37
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Destino"
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
            TabIndex        =   36
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lblFechaAlta 
            Caption         =   "Fecha alta"
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
            TabIndex        =   35
            Top             =   1710
            Width           =   1230
         End
         Begin VB.Label Label49 
            Caption         =   "Servicio egreso"
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
            TabIndex        =   34
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   4740
         TabIndex        =   20
         Top             =   420
         Width           =   5955
         Begin VB.CommandButton cmdBuscaMadre 
            Caption         =   "..."
            Height          =   315
            Left            =   1560
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Busca a la Madre"
            Top             =   2490
            Width           =   315
         End
         Begin VB.TextBox lblMadre 
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
            Left            =   1890
            TabIndex        =   51
            Top             =   2490
            Width           =   3315
         End
         Begin VB.CommandButton btnQuitarMadre 
            DisabledPicture =   "AdmisionAlojDetalle.frx":28E8
            DownPicture     =   "AdmisionAlojDetalle.frx":2C73
            Height          =   315
            Left            =   5220
            Picture         =   "AdmisionAlojDetalle.frx":3006
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2490
            Width           =   615
         End
         Begin VB.ComboBox cmbIdViasAdmision 
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
            Left            =   1890
            TabIndex        =   1
            Top             =   210
            Width           =   3960
         End
         Begin VB.TextBox txtEdadEnDias 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            TabIndex        =   7
            Top             =   1770
            Width           =   585
         End
         Begin VB.TextBox lblNombreMedico 
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
            Left            =   1890
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   990
            Width           =   3015
         End
         Begin VB.ComboBox cmbServicioIngreso 
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
            Left            =   1890
            TabIndex        =   2
            Top             =   600
            Width           =   3960
         End
         Begin VB.TextBox txtNroCamaIngreso 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            TabIndex        =   9
            Top             =   2130
            Width           =   885
         End
         Begin VB.CommandButton btnVerDisponibilidadDeCamas 
            Caption         =   "..."
            Height          =   315
            Left            =   2820
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2130
            Width           =   315
         End
         Begin VB.ComboBox cmbIdTipoEdad 
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
            ItemData        =   "AdmisionAlojDetalle.frx":3397
            Left            =   2550
            List            =   "AdmisionAlojDetalle.frx":3399
            TabIndex        =   8
            Top             =   1770
            Width           =   1545
         End
         Begin VB.TextBox txtIdMedicoIngreso 
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
            Left            =   4920
            TabIndex        =   4
            Top             =   990
            Width           =   915
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   3300
            TabIndex        =   6
            Top             =   1380
            Width           =   780
            _ExtentX        =   1376
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
         Begin MSMask.MaskEdBox txtFechaIngreso 
            Height          =   315
            Left            =   1890
            TabIndex        =   5
            Top             =   1380
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
         Begin VB.Label Label4 
            Caption         =   "Nombre Madre"
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
            Left            =   60
            TabIndex        =   53
            Top             =   2550
            Width           =   1215
         End
         Begin VB.Label lblViaAdmision 
            Caption         =   "Origen"
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
            Left            =   60
            TabIndex        =   46
            Top             =   270
            Width           =   1605
         End
         Begin VB.Label lblNroCamaIngreso 
            Caption         =   "Nro cama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   42
            Top             =   2160
            Width           =   1275
         End
         Begin VB.Label lblFecha 
            Caption         =   "Fecha ingreso"
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
            Left            =   60
            TabIndex        =   24
            Top             =   1410
            Width           =   1215
         End
         Begin VB.Label lblEdadEnDias 
            Caption         =   "Edad en Alojamiento"
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
            Left            =   60
            TabIndex        =   23
            Top             =   1800
            Width           =   1785
         End
         Begin VB.Label lblIdServicioIngreso 
            Caption         =   "Servicio ingreso"
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
            Left            =   60
            TabIndex        =   22
            Top             =   645
            Width           =   1395
         End
         Begin VB.Label lblIdMedicoIngreso 
            Caption         =   "Medico ingreso"
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
            Left            =   60
            TabIndex        =   21
            Top             =   1035
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   210
         TabIndex        =   17
         Top             =   3690
         Width           =   10515
         Begin VB.TextBox txtNroCuenta 
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
            Height          =   315
            Left            =   6480
            TabIndex        =   12
            Top             =   180
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
            Height          =   330
            Left            =   1680
            TabIndex        =   11
            Top             =   210
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbFormaPago 
            Height          =   330
            Left            =   1680
            TabIndex        =   47
            Top             =   630
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SISGalenPlus.ucMensajeParpadeando ucMensajeParpadeando1 
            Height          =   345
            Left            =   4770
            TabIndex        =   49
            Top             =   660
            Visible         =   0   'False
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   609
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   90
            TabIndex        =   48
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label lblEstadoCta 
            AutoSize        =   -1  'True
            Caption         =   "."
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
            Left            =   8010
            TabIndex        =   38
            Top             =   240
            Width           =   2310
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   4740
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fte.Financiam/IAFA"
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
            Left            =   120
            TabIndex        =   18
            Top             =   270
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "AdmisionAlojDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Alojados
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_Opcion As sghOpciones
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdViasAdmision As New sighentidades.ListaDespleglable
Dim mo_cmbIdDestinoAtencion As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New sighentidades.ListaDespleglable
Dim mo_cmbServicioIngreso As New sighentidades.ListaDespleglable

Dim mo_Atenciones As New DOAtencion
Dim mo_Pacientes  As New doPaciente
Dim mo_CuentasAtencion As New DOCuentaAtencion
Dim oDOOcupacion As New DOEstanciaHospitalaria
Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales

Dim oRsFuentesFinanciamiento As New Recordset
Dim oRsFormaPago As New Recordset

Dim ml_idAtencion As Long
Dim ml_IdUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnIdNacimientoSeleccionado As Long
'
Dim lnFocusCuandoCargeFrm As Long
Dim lbUltimaTeclaPulsoENTER As Boolean
Dim lnEspecialidadServicio As Long
Dim lcCaptionTab2 As String
Dim ms_MensajeError As String
Dim lbPacienteNN As Boolean


Property Let lcNombrePc(lValue As String)
  mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
  mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let IdUsuario(lValue As Long)
  ml_IdUsuario = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
  mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let idAtencion(lValue As Long)
  ml_idAtencion = lValue
End Property

Property Get idAtencion() As Long
  idAtencion = ml_idAtencion
End Property

Private Sub btnAceptar_Click()
  If btnAceptar.Enabled = False Then Exit Sub
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  
  Select Case mi_Opcion
    Case sghAgregar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If AgregarDatos() Then
            Me.txtNroCuenta = mo_Atenciones.idCuentaAtencion
            MsgBox " Los datos se agregaron correctamente, para la Historia Nª: " & mo_Pacientes.NroHistoriaClinica & Chr(13) & Chr(13) & "N° Cuenta " & txtNroCuenta.Text, vbInformation, Me.Caption
            Me.Visible = False
          Else
            MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghModificar
       If ValidarDatosObligatorios() Then
         CargaDatosAlObjetosDeDatos
         If ValidarReglas() Then
           If ModificarDatos() Then
             MsgBox " Los datos se modificaron correctamente, para la Cuenta N° " & txtNroCuenta.Text, vbInformation, Me.Caption
             Me.Visible = False
           Else
             MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
           End If
         End If
       End If
    Case sghEliminar
      'If ValidarReglas() Then
        CargaDatosAlObjetosDeDatos
        If EliminarDatos(oConexion) Then
          MsgBox "Los datos se eliminaron correctamente, para la Cuenta N° " & txtNroCuenta.Text, vbInformation, Me.Caption
          Me.Visible = False
        Else
          MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
        End If
      'End If
  End Select
  oConexion.Close
  Set oConexion = Nothing
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   Dim oConexion As New Connection
   oConexion.Open sighentidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   
   ValidarDatosObligatorios = False
   UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
   
   If mo_Pacientes.ApellidoPaterno = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Paterno " + Chr(13)
   End If
   If mo_Pacientes.ApellidoMaterno = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Materno " + Chr(13)
   End If
   If mo_Pacientes.PrimerNombre = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Primer Nombre" + Chr(13)
   End If
   If mo_Pacientes.idTipoSexo = 0 Then
       sMensaje = sMensaje + "Elija el Sexo" + Chr(13)
   End If
   If cmbIdViasAdmision.Text = "" Then
       sMensaje = sMensaje + "Elija el Origen " + Chr(13)
   End If
   If cmbServicioIngreso.Text = "" Then
       sMensaje = sMensaje + "Elija el Servicio de Ingreso" + Chr(13)
   End If
   If txtIdMedicoIngreso.Text = "" Then
       sMensaje = sMensaje + "Ingrese el Médico que recibe" + Chr(13)
   End If
   If txtFechaIngreso.Text = sighentidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Registre la Fecha de Ingreso " + Chr(13)
   End If
   If txtHoraIngreso.Text = sighentidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Registre la Hora de Ingreso" + Chr(13)
   End If
   If txtEdadEnDias.Text = "" Then
       sMensaje = sMensaje + "Ingrese la Edad" + Chr(13)
   End If
   If cmbIdTipoEdad.Text = "" Then
       sMensaje = sMensaje + "Elija el Tipo de Edad" + Chr(13)
   End If
   If txtNroCamaIngreso.Text = "" Then
       sMensaje = sMensaje + "Por favor asigne la Cama" + Chr(13)
   End If
   If Val(cmbFormaPago.BoundText) = 0 Then
      sMensaje = sMensaje + "Por favor elija el Tipo de Financiamiento" + Chr(13)
   End If
   If mi_Opcion = sghAgregar Then
      If mo_AdminAdmision.BuscaSiEstaHospitalizado(mo_Pacientes.idPaciente, oConexion, sghHospitalizacion) = True Then 'debb-05/12/2015
         Exit Function
      End If
   End If
   If txtFechaEgreso.Text <> sighentidades.FECHA_VACIA_DMY Then
      If cmbIdDestinoAtencion.Text = "" Then
         sMensaje = sMensaje + "Elija el Destino" + Chr(13)
      End If
      If lblNombreMedicoEgreso.Text = "" Then
         sMensaje = sMensaje + "Ingrese el Médico que dió el Egreso " + Chr(13)
      End If
      If txtHoraEgreso.Text = sighentidades.HORA_VACIA_HM Then
         sMensaje = sMensaje + "Ingrese la Hora de Egreso " + Chr(13)
      End If
      If txtNombreAcompañante.Text = "" Then
         sMensaje = sMensaje + "Ingrese el que recibe" + Chr(13)
      End If
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   oConexion.Close
   Set oConexion = Nothing
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
    ValidarReglas = False
    If Not mo_AdminAdmision.ValidaEdadMaximaYSexoSegunServicioHosp(Val(txtEdadEnDias.Text), _
              Val(mo_cmbIdTipoEdad.BoundText), mo_Pacientes.idTipoSexo, mo_Atenciones.IdServicioEgreso, True) Then
        Exit Function
    End If
    If Me.txtFechaEgreso.Text <> sighentidades.FECHA_VACIA_DMY And Me.txtHoraEgreso.Text <> sighentidades.HORA_VACIA_HM Then
        If CDate(Me.txtFechaEgreso & " " & Me.txtHoraEgreso) < CDate(Me.txtFechaIngreso & " " & Me.txtHoraIngreso) Then
            MsgBox "La FECHA DE EGRESO no puede ser menor a la FECHA INGRESO", vbInformation, "Mensaje"
            Exit Function
        End If
    End If
    'Valida que el recien nacido, si ha nacido en el Hospital,  tenga asociada la Cuenta de la MADRE
    'If Year(mo_Atenciones.FechaIngreso) > 2010 Then
       '**ojo***eliminar esta validacion 2010 cuando sea necesaria
       If lnIdNacimientoSeleccionado = 0 Then
             MsgBox "Por favor ingrese el Nombre de la Madre", vbInformation, Me.Caption
             Exit Function
       End If
    'End If
    '
    ValidarReglas = True
End Function

Private Sub btnCancelar_Click()
       Me.Visible = False
End Sub



Private Sub cmbIdDestinoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDestinoAtencion
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbIdDestinoAtencion_LostFocus()
  On Error Resume Next
  Me.lblNombreMedicoEgreso.SetFocus
End Sub

Private Sub cmbIdTipoEdad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoEdad
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbIdTipoEdad_LostFocus()
   If txtFechaIngreso.Text <> sighentidades.FECHA_VACIA_DMY And txtHoraIngreso.Text <> sighentidades.HORA_VACIA_HM And cmbIdTipoEdad.Text <> "" Then
      Dim ldFechaNacimiento As Date
      ldFechaNacimiento = sighentidades.DevuelveFechaNacimiento(txtFechaIngreso.Text, txtHoraIngreso.Text, Val(txtEdadEnDias.Text), Val(mo_cmbIdTipoEdad.BoundText))
      UcPacienteDatosAloj1.ActualizaFechaHoraNacimiento ldFechaNacimiento
   End If
   If txtNroCamaIngreso.Text = "" Then
      btnVerDisponibilidadDeCamas_Click
   Else
      btnAceptar.SetFocus
   End If
End Sub

Private Sub cmbIdViasAdmision_GotFocus()
    cmbIdViasAdmision.SetFocus
End Sub

Private Sub cmbIdViasAdmision_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdViasAdmision
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbIdViasAdmision_LostFocus()
    If lbUltimaTeclaPulsoENTER = True Then
        lbUltimaTeclaPulsoENTER = False
        'cmbIdViasAdmision.SetFocus
        cmbIdViasAdmision_GotFocus
    End If
End Sub

Private Sub cmbServicioIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbServicioIngreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbServicioIngreso_LostFocus()
   If mo_cmbServicioIngreso.BoundText <> "" Then
      txtIdServicioEgreso.Tag = mo_cmbServicioIngreso.BoundText
      txtIdServicioEgreso.Text = mo_cmbServicioIngreso.BoundText
      lblNombreServicioEgreso.Text = cmbServicioIngreso.Text
      lnEspecialidadServicio = mo_AdminServiciosComunes.DevuelveEspecialidadDelServicio(Val(mo_cmbServicioIngreso.BoundText))
   End If
   On Error Resume Next
   lblNombreMedico.SetFocus
End Sub

Private Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
        Select Case lnFocusCuandoCargeFrm
        Case 0  'Se ingresa por primera vez
             tabAdmision.Tab = 1
             On Error Resume Next
             cmbIdDestinoAtencion.SetFocus
        End Select
   Else
        Select Case lnFocusCuandoCargeFrm
        Case 0  'Se ingresa por primera vez
             tabAdmision.Tab = 0
             On Error Resume Next
             UcPacienteDatosAloj1.SetFocusOnApellidoPaterno
        End Select
   End If
   lnFocusCuandoCargeFrm = 100
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdViasAdmision.MiComboBox = cmbIdViasAdmision
    Set mo_cmbIdDestinoAtencion.MiComboBox = cmbIdDestinoAtencion
    Set mo_cmbIdTipoEdad.MiComboBox = cmbIdTipoEdad
    Set mo_cmbServicioIngreso.MiComboBox = cmbServicioIngreso
End Sub

Sub Form_Load()
       mo_Formulario.HabilitarDeshabilitar txtIdMedicoIngreso, False
       mo_Formulario.HabilitarDeshabilitar lblNombreServicioEgreso, False
       mo_Formulario.HabilitarDeshabilitar Me.txtIdServicioEgreso, False
       mo_Formulario.HabilitarDeshabilitar txtIdMedicoEgreso, False
       mo_Formulario.HabilitarDeshabilitar txtDiasEstancia, False
       mo_Formulario.HabilitarDeshabilitar Me.txtNroCuenta, False
       mo_Formulario.HabilitarDeshabilitar txtNroCamaIngreso, False
       mo_Formulario.HabilitarDeshabilitar txtNroCamaEgreso, False
       mo_Formulario.HabilitarDeshabilitar lblMadre, False
       '
       CargaCombox
       '
       UcPacienteDatosAloj1.IdTipoGenHistoriaClinica = sghHistoriaTemporalAlojamiento
       UcPacienteDatosAloj1.Opcion = mi_Opcion
       UcPacienteDatosAloj1.inicializar
       '
       lnFocusCuandoCargeFrm = 0
       lnEspecialidadServicio = 0
       Me.txtFechaIngreso = Date
       Me.txtHoraIngreso = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
       '
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Alojados"
       Case sghModificar
           Me.Caption = "Modificar Alojados"
       Case sghConsultar
           Me.Caption = "Consultar Alojados"
       Case sghEliminar
           Me.Caption = "Eliminar Alojados"
       End Select
       CargarDatosAlFormulario
End Sub

Sub CargarDatosAlFormulario()
     Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
     End Select
End Sub

Sub CargaCombox()
    mo_cmbIdViasAdmision.BoundColumn = "IdOrigenAtencion"
    mo_cmbIdViasAdmision.ListField = "DescripcionLarga"
    Set mo_cmbIdViasAdmision.RowSource = mo_AdminAdmision.TiposOrigenAtencionSeleccionarViasDeHospitalizacion(sghSoloPacAlojados)
    '
    mo_cmbIdDestinoAtencion.BoundColumn = "IdDestinoAtencion"
    mo_cmbIdDestinoAtencion.ListField = "DescripcionLarga"
    Set mo_cmbIdDestinoAtencion.RowSource = mo_AdminAdmision.TiposDestinoAtencionSeleccionarDestinosDeHospitalizacion(sghSoloPacAlojados)
    '
    mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
    mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoEdad.RowSource = mo_AdminServiciosComunes.TiposEdadSeleccionarTodos
    mo_cmbIdTipoEdad.BoundText = "4"    'Default HORAS
    '
    Set oRsFuentesFinanciamiento = mo_AdminServiciosComunes.FuentesFinanciamientoSegunFiltro("EsUsadoEnCaja=1")
    Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
    cmbFuenteFinanciamiento.ListField = "Descripcion"
    cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
    cmbFuenteFinanciamiento.BoundText = "5"   'Credito Hospitalario
    '
    Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
    Dim lcEspecialidadesDelUsuario As String
    lcEspecialidadesDelUsuario = mo_AdminAdmision.DevuelveEspecialidadesServicioSegunUsuarioSistema(sghEspecialidadesHosp, ml_IdUsuario)
    mo_cmbServicioIngreso.BoundColumn = "IdServicio"
    mo_cmbServicioIngreso.ListField = "DservicioHosp"
    Set mo_cmbServicioIngreso.RowSource = oBuscaServicios.DevuelveServiciosDelHospital("(3)", lcEspecialidadesDelUsuario, sghFiltraSoloActivos, sghPorDescTipoServicio)
    Set oBuscaServicios = Nothing
    '
    Set oRsFormaPago = mo_AdminServiciosComunes.TiposFinanciamientoSegunFiltro("esFuenteFinanciamiento=1")
    Set cmbFormaPago.RowSource = oRsFormaPago
    cmbFormaPago.ListField = "Descripcion"
    cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
    mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
End Sub

Private Sub lblNombreMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreMedico
   AdministrarKeyPreview KeyCode

End Sub

Private Sub lblNombreMedico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If

End Sub

Private Sub lblNombreMedico_LostFocus()
        If lblNombreMedico.Locked = False And lbUltimaTeclaPulsoENTER = True Then
           lbUltimaTeclaPulsoENTER = False
           CompletarDatosDeMedico txtIdMedicoIngreso, lblNombreMedico, lnEspecialidadServicio, lblNombreMedico.Text
           On Error Resume Next
           txtFechaIngreso.SetFocus
        End If
End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.IdEspecialidad = lIdEspecialidad
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreMedico = lcFiltraMedico
    End If
    'oBusqueda.Show 1
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
       End If
    End If
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub lblNombreMedicoEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreMedicoEgreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub lblNombreMedicoEgreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If

End Sub

Private Sub lblNombreMedicoEgreso_LostFocus()
        If lblNombreMedicoEgreso.Locked = False And lbUltimaTeclaPulsoENTER = True Then
           lbUltimaTeclaPulsoENTER = False
           CompletarDatosDeMedico txtIdMedicoEgreso, lblNombreMedicoEgreso, lnEspecialidadServicio, lblNombreMedicoEgreso.Text
           On Error Resume Next
           txtFechaEgreso.SetFocus
        End If
End Sub

Private Sub tabAdmision_Click(PreviousTab As Integer)
   On Error Resume Next
   Select Case tabAdmision.Tab
   Case 0
       UcPacienteDatosAloj1.SetFocusOnApellidoPaterno
   Case 1
       cmbIdDestinoAtencion.SetFocus
   End Select
End Sub

Private Sub txtEdadEnDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadEnDias
   AdministrarKeyPreview KeyCode

End Sub



Private Sub txtFechaEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaEgreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtHoraEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraEgreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtHoraEgreso_LostFocus()
    If txtFechaIngreso.Text <> sighentidades.FECHA_VACIA_DMY And txtHoraIngreso.Text <> sighentidades.HORA_VACIA_HM And txtFechaEgreso.Text <> sighentidades.FECHA_VACIA_DMY And txtHoraEgreso.Text <> sighentidades.HORA_VACIA_HM Then
       txtDiasEstancia.Text = sighentidades.DiasDeEstanciaEnHospitalizacionEmergencia(txtFechaIngreso.Text, txtHoraIngreso.Text, txtFechaEgreso.Text, txtHoraEgreso.Text)
    End If
End Sub

Private Sub txtHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraIngreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNombreAcompañante_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraEgreso
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNombreAcompañante_LostFocus()
    btnAceptar.SetFocus
End Sub

Private Sub UcPacienteDatosAloj1_SePresionoTeclaEspecial(KeyCode As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyReturn
         Dim oConexion As New Connection
         oConexion.Open sighentidades.CadenaConexion
         oConexion.CursorLocation = adUseClient
         DeudasPendientesDeAnterioresAtenciones oConexion
         oConexion.Close
         Set oConexion = Nothing
         lbUltimaTeclaPulsoENTER = True: cmbIdViasAdmision.SetFocus
    Case Else
         AdministrarKeyPreview KeyCode
    End Select
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
     Case vbKeyF10
         Me.tabAdmision.Tab = 0
         On Error Resume Next
         UcPacienteDatosAloj1.SetFocusOnApellidoPaterno
     Case vbKeyF11
         Me.tabAdmision.Tab = 1
         On Error Resume Next
         cmbIdDestinoAtencion.SetFocus
     Case vbKeyF12
         Me.tabAdmision.Tab = 1
    End Select
       
End Sub

Private Sub cmdBuscaCamaEgreso_Click()
Dim oBusqueda As New CamasBusqueda
Dim oDOCama As New DOCama
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    oBusqueda.idTipoServicio = sghHospitalizacion
    oBusqueda.IdServicio = Val(mo_cmbServicioIngreso.BoundText)
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
       'CargaCamaSeleccionada (oBusqueda.IdRegistroSeleccionado)
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOCama Is Nothing Then
            If oDOCama.idPaciente = mo_Atenciones.idPaciente Or oDOCama.idPaciente = 0 Then
                Me.txtNroCamaEgreso.Text = oDOCama.Codigo
                Me.txtNroCamaEgreso.Tag = oDOCama.idCama
            Else
                MsgBox "La cama seleccionada no puede usarla", vbInformation, Me.Caption
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDOCama = Nothing
End Sub


Private Sub btnVerDisponibilidadDeCamas_Click()
Dim oBusqueda As New CamasBusqueda
Dim oDOCama As New DOCama
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.idTipoServicio = sghHospitalizacion
    oBusqueda.IdServicio = Val(mo_cmbServicioIngreso.BoundText)
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
       CargaCamaSeleccionada (oBusqueda.idRegistroSeleccionado)
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOCama Is Nothing Then
            If oDOCama.idPaciente = mo_Atenciones.idPaciente Or oDOCama.idPaciente = 0 Then
                Me.txtNroCamaIngreso.Text = oDOCama.Codigo
                Me.txtNroCamaIngreso.Tag = oDOCama.idCama
                   Me.txtNroCamaEgreso.Text = oDOCama.Codigo
                   Me.txtNroCamaEgreso.Tag = oDOCama.idCama
                btnAceptar.SetFocus
            Else
                MsgBox "La cama seleccionada no puede usarla", vbInformation, Me.Caption
                Me.txtNroCamaIngreso.Text = ""
                Me.txtNroCamaIngreso.Tag = ""
            End If
        End If
    End If
    Set oBusqueda = Nothing
    Set oDOCama = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CargaCamaSeleccionada(idCama As Long)
        Dim oDOCama As New DOCama
        Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(idCama, oConexion)
        If Not oDOCama Is Nothing Then
            If Val(mo_cmbServicioIngreso.BoundText) = oDOCama.IdServicioUbicacionActual Then
                Me.txtNroCamaIngreso.Text = oDOCama.Codigo
                Me.txtNroCamaIngreso.Tag = oDOCama.idCama
                If Me.txtNroCamaEgreso.Text = "" Then
                   Me.txtNroCamaEgreso.Text = oDOCama.Codigo
                   Me.txtNroCamaEgreso.Tag = oDOCama.idCama
                End If
            Else
                MsgBox "La cama seleccionada no pertenece al mismo servicio de ingreso", vbInformation, Me.Caption
                Me.txtNroCamaIngreso.Text = ""
                Me.txtNroCamaIngreso.Tag = ""
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oDOCama = Nothing

End Sub

Sub CargaDatosAlObjetosDeDatos()
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
    '********mo_Pacientes****** YA SE CARGO EN VALIDADATOSOBLIGATORIOS()
    '
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
    '---------------------------------------------------------------------------------
    Select Case mi_Opcion
    Case sghAgregar
        With mo_CuentasAtencion
                .idPaciente = mo_Pacientes.idPaciente
                .TotalAsegurado = 0
                .TotalExonerado = 0
                .TotalPagado = 0
                .TotalPorPagar = 0
                .IdEstado = sghEstadoCuenta.sghAbierto
                .FechaApertura = Me.txtFechaIngreso.Text
                .HoraApertura = Me.txtHoraIngreso.Text
                .fechaCierre = 0
                .HoraCierre = ""
                .IdUsuarioAuditoria = ml_IdUsuario
        End With
    Case Else
        mo_CuentasAtencion.IdUsuarioAuditoria = ml_IdUsuario
        If Me.txtFechaEgreso.Text <> sighentidades.FECHA_VACIA_DMY Then
           mo_CuentasAtencion.IdEstado = sghEstadoCuenta.sghConAltaMedica
        End If
    End Select
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
           .idAtencion = Me.idAtencion
           .IdEspecialidadMedico = 0
           .IdMedicoIngreso = Val(Me.txtIdMedicoIngreso.Tag)
           .IdMedicoEgreso = Val(Me.txtIdMedicoEgreso.Tag)
           '.IdMedicoRespNacimiento = 0
           .IdServicioIngreso = Val(mo_cmbServicioIngreso.BoundText)
           .IdOrigenAtencion = Val(mo_cmbIdViasAdmision.BoundText)
'           .IdTipoReferenciaOrigen = 0
'           .idEstablecimientoOrigen = 0
'           .IdEstablecimientoNoMinsaOrigen = 0
           .IdDestinoAtencion = Val(mo_cmbIdDestinoAtencion.BoundText)
'           .IdTipoReferenciaDestino = 0
'           .idEstablecimientoDestino = 0
'           .IdEstablecimientoNoMinsaDestino = 0
           .HoraIngreso = IIf(Me.txtHoraIngreso.Text = sighentidades.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
           .FechaIngreso = IIf(Me.txtFechaIngreso.Text = sighentidades.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
           .FechaEgresoAdministrativo = IIf(Me.txtFechaEgreso = sighentidades.FECHA_VACIA_DMY, 0, Me.txtFechaEgreso)
           .HoraEgresoAdministrativo = IIf(Me.txtHoraEgreso = sighentidades.HORA_VACIA_HM, "", Me.txtHoraEgreso)
           .idTipoServicio = sghHospitalizacion
           .Edad = Me.txtEdadEnDias.Text
           .IdTipoEdad = Val(mo_cmbIdTipoEdad.BoundText)
           .idPaciente = mo_Pacientes.idPaciente
           .IdUsuarioAuditoria = ml_IdUsuario
           '.RecienNacido = 0
           '.Observacion = Me.txtObservacion
           'Estos datos llenaran  en el modulo de registro de atenciones
            .IdTipoCondicionALEstab = 1
            .IdTipoCondicionAlServicio = 1
            .fechaEgreso = 0
            .HoraEgreso = sighentidades.HORA_VACIA_HM
            .IdCamaIngreso = Val(Me.txtNroCamaIngreso.Tag)
            .IdCamaEgreso = Val(Me.txtNroCamaEgreso.Tag)
            .IdCondicionAlta = 0
            .IdServicioEgreso = Val(Me.txtIdServicioEgreso.Tag)
            .IdServicioEgreso = IIf(mi_Opcion = sghAgregar, Val(mo_cmbServicioIngreso.BoundText), Val(Me.txtIdServicioEgreso.Tag))
            .IdTipoAlta = 0
'            .TieneNecropsia = False
'            .HuboInfeccionIntraHospitalaria = False
            .IdTipoGravedad = 0
            .IdFormaPago = Val(cmbFormaPago.BoundText)
            .IdFuenteFinanciamiento = Val(cmbFuenteFinanciamiento.BoundText)
            .IdEstadoAtencion = 1
   End With
   With mo_DoAtencionDatosAdicionales
           .IdMedicoRespNacimiento = 0
           .IdTipoReferenciaOrigen = 0
           .IdEstablecimientoOrigen = 0
           .IdEstablecimientoNoMinsaOrigen = 0
           .IdTipoReferenciaDestino = 0
           .idEstablecimientoDestino = 0
           .IdEstablecimientoNoMinsaDestino = 0
           .RecienNacido = 0
           .TieneNecropsia = False
           .HuboInfeccionIntraHospitalaria = False

           '.DireccionDomicilio
           .NombreAcompaniante = txtNombreAcompañante.Text
           '.Observacion
   End With
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE ESTANCIA
    '---------------------------------------------------------------------------------
    oDOOcupacion.IdServicio = Val(mo_cmbServicioIngreso.BoundText)
    oDOOcupacion.IdMedicoOrdena = Val(Me.txtIdMedicoIngreso.Tag)
    oDOOcupacion.FechaOcupacion = Me.txtFechaIngreso
    oDOOcupacion.HoraOcupacion = Me.txtHoraIngreso
    oDOOcupacion.idCama = Val(Me.txtNroCamaIngreso.Tag)
    oDOOcupacion.IdUsuarioAuditoria = ml_IdUsuario
    oDOOcupacion.LlegoAlServicio = 1
    oDOOcupacion.Secuencia = 1
    oDOOcupacion.IdUsuarioAuditoria = ml_IdUsuario
    If mi_Opcion = sghModificar And mo_Atenciones.fechaEgreso <> 0 Then
       oDOOcupacion.DiasEstancia = sighentidades.DiasDeEstanciaEnHospitalizacionEmergencia(mo_Atenciones.FechaIngreso, mo_Atenciones.HoraIngreso, mo_Atenciones.fechaEgreso, mo_Atenciones.HoraEgreso)
       oDOOcupacion.FechaDesocupacion = mo_Atenciones.fechaEgreso
       oDOOcupacion.HoraDesocupacion = mo_Atenciones.HoraEgreso
    End If
    '
    mo_DoAtencionDatosAdicionales.RecienNacido = sighentidades.CalculaSiEsRecienNacido(mo_Pacientes.FechaNacimiento, CDate(mo_Atenciones.FechaIngreso & " " & mo_Atenciones.HoraIngreso))
    '
    lcCaptionTab2 = tabAdmision.Caption
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminAdmision.AdmisionAlojadosAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, oDOOcupacion, lbPacienteNN, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(tabAdmision.Caption) & "/" & Trim(lcCaptionTab2), lnIdNacimientoSeleccionado, mo_DoAtencionDatosAdicionales)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminAdmision.AdmisionAlojadosModificar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, oDOOcupacion, lbPacienteNN, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(tabAdmision.Caption) & "/" & Trim(lcCaptionTab2), lnIdNacimientoSeleccionado, mo_DoAtencionDatosAdicionales)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos(oConexion As Connection) As Boolean
    ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
    If ms_MensajeError = "" Then
        mo_CuentasAtencion.IdEstado = 9 'anulado
        mo_Atenciones.IdEstadoAtencion = 0  'anulado
        EliminarDatos = mo_AdminAdmision.AdmisionAlojadosAnular(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, oDOOcupacion, lbPacienteNN, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(tabAdmision.Caption) & "/" & Trim(lcCaptionTab2), lnIdNacimientoSeleccionado)
        ms_MensajeError = mo_AdminAdmision.MensajeError
    Else
        MsgBox ms_MensajeError & Chr(13) & "La Anulación tendrá que realizarlo FACTURACION ", vbInformation, "Consulta externa"
    End If
End Function

Sub CargarDatosAlosControles()
Dim lcEstadoAtencion As String
Dim oRsTmp As New Recordset
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        '1do:   CARGAR DATOS DE LA ATENCION
        Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
        If mo_Atenciones.idAtencion = 0 Then
            'El registro ha sido eliminado, pero no se hizo el refresh
             Exit Sub
        End If
        With mo_Atenciones
                mo_cmbIdDestinoAtencion.BoundText = .IdDestinoAtencion
                mo_cmbServicioIngreso.BoundText = .IdServicioIngreso
                Me.txtIdMedicoIngreso.Tag = .IdMedicoIngreso
                Me.txtIdMedicoEgreso.Tag = .IdMedicoEgreso
                mo_cmbIdViasAdmision.BoundText = .IdOrigenAtencion
                Me.txtHoraIngreso.Text = IIf(.HoraIngreso = "", sighentidades.HORA_VACIA_HM, .HoraIngreso)
                Me.txtFechaIngreso.Text = IIf(.FechaIngreso = 0, sighentidades.FECHA_VACIA_DMY, .FechaIngreso)
                Me.txtHoraEgreso.Text = IIf(.HoraEgresoAdministrativo = "", sighentidades.HORA_VACIA_HM, .HoraEgresoAdministrativo)
                Me.txtFechaEgreso.Text = IIf(.FechaEgresoAdministrativo = 0, sighentidades.FECHA_VACIA_DMY, Format(.FechaEgresoAdministrativo, sighentidades.FormatoFechaCorta))
                'Se guarda en estas variables para validar si el paciente ya esta de alta o no
                Me.txtHoraEgreso.Tag = Me.txtHoraEgreso.Text
                Me.txtFechaEgreso.Tag = Me.txtFechaEgreso.Text
                Me.txtEdadEnDias.Text = .Edad
                Me.txtEdadEnDias.Tag = .Edad
                mo_cmbIdTipoEdad.BoundText = .IdTipoEdad
                cmbIdTipoEdad.Tag = .IdTipoEdad
'
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoIngreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoIngreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedico = ""
                End If
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoEgreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoEgreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedicoEgreso = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedicoEgreso = ""
                End If
                'Cama de ingreso
                Me.txtNroCamaIngreso.Tag = .IdCamaIngreso
                Dim oDOCama As New DOCama
                Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(.IdCamaIngreso, oConexion)
                Me.txtNroCamaIngreso.Text = oDOCama.Codigo
                Me.txtNroCamaIngreso.Tag = .IdCamaIngreso
                Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(.IdCamaEgreso, oConexion)
                Me.txtNroCamaEgreso.Text = oDOCama.Codigo
                Me.txtNroCamaEgreso.Tag = .IdCamaEgreso
                Set oDOCama = Nothing
                
                cmbFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
                cmbFormaPago.BoundText = .IdFormaPago
                Select Case .IdEstadoAtencion
                Case 0
                    lcEstadoAtencion = "Anulado"
                    btnAceptar.Enabled = False
                Case 1
                    lcEstadoAtencion = "Registrado"
                Case 2
                    lcEstadoAtencion = "Cerrado"
                    btnAceptar.Enabled = False
                End Select
                '
                txtIdServicioEgreso.Tag = mo_cmbServicioIngreso.BoundText
                txtIdServicioEgreso.Text = mo_cmbServicioIngreso.BoundText
                lblNombreServicioEgreso.Text = cmbServicioIngreso.Text
                Set oRsTmp = mo_ReglasArchivoClinico.ServiciosSeleccionarXidentificador(Val(mo_cmbServicioIngreso.BoundText))
                If oRsTmp.RecordCount > 0 Then
                   lnEspecialidadServicio = oRsTmp.Fields!IdEspecialidad
                End If
                '
                Set oRsTmp = mo_AdminAdmision.AtencionesEstanciaHospitalariaSeleccionarPorIdAtencion(.idAtencion)
                If oRsTmp.RecordCount > 0 Then
                   oDOOcupacion.IdEstanciaHospitalaria = oRsTmp.Fields!IdEstanciaHospitalaria
                End If
                '
                
        End With
        '
        Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(Me.idAtencion, oConexion)
        With mo_DoAtencionDatosAdicionales
            Me.txtNombreAcompañante = .NombreAcompaniante
        End With
        '
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_Atenciones.idCuentaAtencion, oConexion)
        lblEstadoCta = mo_ReglasFarmacia.DevuelveEstadoActualDeEstadoCuenta("idEstado=" & mo_CuentasAtencion.IdEstado, oConexion)
        If mo_CuentasAtencion.IdEstado <> 1 And mo_CuentasAtencion.IdEstado <> 12 Then
            btnAceptar.Enabled = False
        End If
        txtNroCuenta.Text = mo_CuentasAtencion.idCuentaAtencion
        '3to:   CARGAR DATOS DEL PACIENTE
        UcPacienteDatosAloj1.idPaciente = mo_Atenciones.idPaciente
        UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
        '
        DeudasPendientesDeAnterioresAtenciones oConexion
        '
        UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
        Me.Caption = Trim(Me.Caption) & "                HC: " & Trim(mo_Pacientes.NroHistoriaClinica) & " " & Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & "     (Estado: " & lcEstadoAtencion & ")"
        '
        oRsTmp.Close
        Set oRsTmp = Nothing
        Set oDoMedico = Nothing
        Set oDOEmpleado = Nothing
        Set oDOEspecialidades = Nothing
        '
        UcPacienteDatosAloj1.NroHistoriaClinica = mo_Pacientes.NroHistoriaClinica
        '
        CargaDatosDeLaMadre oConexion
        '
        oConexion.Close
        Set oConexion = Nothing
End Sub

Sub CargaDatosDeLaMadre(oConexion As Connection)
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_AdminAdmision.AtencionesNacimientosSeleccionarXidPaciente(mo_Atenciones.idPaciente, oConexion)
    If oRsTmp.RecordCount > 0 Then
       lnIdNacimientoSeleccionado = oRsTmp.Fields!idNacimiento
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    '
    If lnIdNacimientoSeleccionado > 0 Then
       lblMadre.Text = mo_AdminAdmision.DevuelveDatosDeLaMadreDelPacienteActual(lnIdNacimientoSeleccionado, Me.UcPacienteDatosAloj1.idTipoSexo, oConexion)
    End If
End Sub

Sub DeudasPendientesDeAnterioresAtenciones(oConexion As Connection)
        'Deudas
        ms_MensajeError = mo_AdminFacturacion.DevuelveDeudaPacienteDeAntencionesAnteriores(mo_Atenciones.idPaciente, oConexion, mo_CuentasAtencion.idCuentaAtencion)
        If ms_MensajeError <> "" Then
           MsgBox "Tiene Deudas Pendientes por Pagar" & Chr(13) & Chr(13) & ms_MensajeError, vbInformation, Me.Caption
           '
           ucMensajeParpadeando1.Visible = True
           ucMensajeParpadeando1.MensajeDeTexto = "Deudas:  " & ms_MensajeError
        Else
           '
           ucMensajeParpadeando1.Visible = False
           ucMensajeParpadeando1.MensajeDeTexto = ""
        End If
        ms_MensajeError = ""

End Sub


Private Sub cmbFuenteFinanciamiento_Click(Area As Integer)
        Set oRsFormaPago = mo_AdminFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(Val(cmbFuenteFinanciamiento.BoundText))
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, True
        If oRsFormaPago.RecordCount = 1 Then
           cmbFormaPago.BoundText = oRsFormaPago.Fields!idTipoFinanciamiento
        ElseIf Val(cmbFuenteFinanciamiento.BoundText) = 5 Then
           cmbFormaPago.BoundText = "1"
        End If
End Sub


Private Sub cmdBuscaMadre_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaMadre
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       lnIdNacimientoSeleccionado = oBusqueda.IdNacimientoSeleccionado
       lblMadre.Text = mo_AdminAdmision.DevuelveDatosDeLaMadreDelPacienteActual(lnIdNacimientoSeleccionado, Me.UcPacienteDatosAloj1.idTipoSexo, oConexion)
    End If
    Set oBusqueda = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub btnQuitarMadre_Click()
    lblMadre.Text = ""
    lnIdNacimientoSeleccionado = 0
End Sub



VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AtencionAmbulatoriaDetalle 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   90
      TabIndex        =   40
      Top             =   5490
      Width           =   8985
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4560
         Picture         =   "AtencionDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3000
         Picture         =   "AtencionDetalle.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   60
      TabIndex        =   8
      Top             =   1500
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   7064
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos de la atención"
      TabPicture(0)   =   "AtencionDetalle.frx":0961
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Diagnosticos"
      TabPicture(1)   =   "AtencionDetalle.frx":097D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   3435
         Left            =   180
         TabIndex        =   44
         Top             =   390
         Width           =   8685
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   315
            Left            =   2190
            TabIndex        =   54
            Top             =   630
            Width           =   315
         End
         Begin VB.CommandButton btnQuitar 
            Caption         =   "Quitar"
            Height          =   345
            Left            =   7080
            TabIndex        =   47
            Top             =   1620
            Width           =   1425
         End
         Begin VB.CommandButton btnAgregar 
            Caption         =   "Agregar"
            Height          =   345
            Left            =   7080
            TabIndex        =   46
            Top             =   1230
            Width           =   1425
         End
         Begin VB.TextBox txtPrecioPlan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            TabIndex        =   45
            Top             =   630
            Width           =   1065
         End
         Begin UltraGrid.SSUltraGrid grdPlanProducto 
            Height          =   2085
            Left            =   150
            TabIndex        =   48
            Top             =   1200
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   3678
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108864
            Caption         =   "Diagnosticos"
         End
         Begin MSDataListLib.DataCombo cmbIdPlan 
            Height          =   315
            Left            =   1080
            TabIndex        =   49
            Top             =   255
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   315
            Left            =   4365
            TabIndex        =   52
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2550
            TabIndex        =   55
            Top             =   630
            Width           =   5955
         End
         Begin VB.Label Label9 
            Caption         =   "Subtipo"
            Height          =   315
            Left            =   3630
            TabIndex        =   53
            Top             =   285
            Width           =   675
         End
         Begin VB.Label lblIdTipoFinanciamiento 
            Caption         =   "Tipo"
            Height          =   315
            Left            =   210
            TabIndex        =   51
            Top             =   300
            Width           =   705
         End
         Begin VB.Label lblIdFuenteFinanciamiento 
            Caption         =   "Código"
            Height          =   315
            Left            =   210
            TabIndex        =   50
            Top             =   660
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos de la atención"
         Height          =   3435
         Left            =   -74820
         TabIndex        =   9
         Top             =   390
         Width           =   8715
         Begin VB.TextBox txtEdadEnDias 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6465
            TabIndex        =   15
            Top             =   180
            Width           =   1000
         End
         Begin VB.TextBox txtIdEstablecimientoOrigen 
            Height          =   315
            Left            =   2280
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1650
            Width           =   1000
         End
         Begin VB.TextBox txtIdEstablecimientoDestino 
            Height          =   315
            Left            =   2265
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2370
            Width           =   1000
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6465
            TabIndex        =   12
            Top             =   540
            Width           =   1000
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   315
            Left            =   3330
            TabIndex        =   11
            Top             =   1650
            Width           =   315
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   315
            Left            =   3330
            TabIndex        =   10
            Top             =   2370
            Width           =   315
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   3480
            TabIndex        =   16
            Top             =   570
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
            Left            =   2280
            TabIndex        =   17
            Top             =   570
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   315
            Left            =   2265
            TabIndex        =   18
            Top             =   210
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   315
            Left            =   3480
            TabIndex        =   19
            Top             =   930
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   315
            Left            =   2280
            TabIndex        =   20
            Top             =   930
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoReferenciaOrigen 
            Height          =   315
            Left            =   2280
            TabIndex        =   21
            Top             =   1290
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoReferenciaDestino 
            Height          =   315
            Left            =   2295
            TabIndex        =   22
            Top             =   2010
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoCondicionALEstab 
            Height          =   315
            Left            =   2265
            TabIndex        =   23
            Top             =   2760
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdDestinoAtencion 
            Height          =   315
            Left            =   6450
            TabIndex        =   24
            Top             =   900
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoCondicionAlServicio 
            Height          =   315
            Left            =   6810
            TabIndex        =   25
            Top             =   2760
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin VB.Label lblEdadEnDias 
            Caption         =   "EdadEnDias"
            Height          =   315
            Left            =   5340
            TabIndex        =   39
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblFechaIngreso 
            Caption         =   "FechaIngreso"
            Height          =   315
            Left            =   240
            TabIndex        =   38
            Top             =   660
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "IdTipoServicio"
            Height          =   315
            Left            =   270
            TabIndex        =   37
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "FechaIngreso"
            Height          =   315
            Left            =   240
            TabIndex        =   36
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label lblIdTipoReferenciaOrigen 
            Caption         =   "IdTipoReferenciaOrigen"
            Height          =   315
            Left            =   210
            TabIndex        =   35
            Top             =   1290
            Width           =   1905
         End
         Begin VB.Label lblIdEstablecimientoOrigen 
            Caption         =   "IdEstablecimientoOrigen"
            Height          =   315
            Left            =   210
            TabIndex        =   34
            Top             =   1650
            Width           =   2025
         End
         Begin VB.Label lblIdTipoReferenciaDestino 
            Caption         =   "IdTipoReferenciaDestino"
            Height          =   315
            Left            =   210
            TabIndex        =   33
            Top             =   1980
            Width           =   1905
         End
         Begin VB.Label lblIdTipoCondicionALEstab 
            Caption         =   "IdTipoCondicionALEstab"
            Height          =   315
            Left            =   210
            TabIndex        =   32
            Top             =   2760
            Width           =   2085
         End
         Begin VB.Label lblIdEstablecimientoDestino 
            Caption         =   "IdEstablecimientoDestino"
            Height          =   315
            Left            =   240
            TabIndex        =   31
            Top             =   2370
            Width           =   1935
         End
         Begin VB.Label lblIdDestinoAtencion 
            Caption         =   "IdDestinoAtencion"
            Height          =   315
            Left            =   4980
            TabIndex        =   30
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label lblIdTipoCondicionAlServicio 
            Caption         =   "IdTipoCondicionAlServicio"
            Height          =   315
            Left            =   4500
            TabIndex        =   29
            Top             =   2790
            Width           =   1845
         End
         Begin VB.Label Label4 
            Caption         =   "IdCita"
            Height          =   315
            Left            =   5850
            TabIndex        =   28
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3690
            TabIndex        =   27
            Top             =   1650
            Width           =   4845
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3690
            TabIndex        =   26
            Top             =   2370
            Width           =   4845
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del paciente"
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9015
      Begin VB.TextBox txtNroHistoriaClinica 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         Top             =   600
         Width           =   1250
      End
      Begin VB.CommandButton btnBuscarHistoriaClinica 
         Caption         =   "..."
         Height          =   315
         Left            =   8460
         TabIndex        =   2
         Top             =   600
         Width           =   315
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3540
         TabIndex        =   1
         Top             =   600
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label8 
         Caption         =   "Nº de cuenta"
         Height          =   255
         Left            =   330
         TabIndex        =   7
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2250
         TabIndex        =   43
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label15 
         Caption         =   "Nombres"
         Height          =   285
         Left            =   330
         TabIndex        =   6
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Historia:"
         Height          =   225
         Left            =   330
         TabIndex        =   5
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2250
         TabIndex        =   4
         Top             =   960
         Width           =   6525
      End
   End
End
Attribute VB_Name = "AtencionAmbulatoriaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

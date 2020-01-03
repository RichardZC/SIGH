VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AtencionHospDetalle 
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   Icon            =   "AtencionHospDetalle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoriaClinica 
      Caption         =   "Datos de la historia clínica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   12015
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1860
         MaxLength       =   35
         TabIndex        =   9
         Top             =   600
         Width           =   7065
      End
      Begin VB.TextBox txtIdNroHistoria 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   240
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo cmbIdTipoGenHistoriaClinica 
         Height          =   315
         Left            =   2970
         TabIndex        =   6
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   315
         Left            =   7800
         TabIndex        =   11
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   10260
         TabIndex        =   13
         Top             =   600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
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
         Left            =   9390
         TabIndex        =   14
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha Nacimiento:"
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
         Left            =   6270
         TabIndex        =   12
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos y nombres"
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
         Left            =   210
         TabIndex        =   10
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Nro &Historia:"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   7320
      Width           =   12015
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AtencionHospDetalle.frx":08CA
         DownPicture     =   "AtencionHospDetalle.frx":0D2A
         Height          =   700
         Left            =   4410
         Picture         =   "AtencionHospDetalle.frx":119F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   6420
         Picture         =   "AtencionHospDetalle.frx":1614
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   705
         Left            =   150
         Picture         =   "AtencionHospDetalle.frx":1B00
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1245
      End
   End
   Begin TabDlg.SSTab tabAdmision 
      Height          =   6165
      Left            =   60
      TabIndex        =   8
      Top             =   1140
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   10874
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos del ingreso (F4)"
      TabPicture(0)   =   "AtencionHospDetalle.frx":1FD9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdDiagnosticos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Datos de egreso (F6)"
      TabPicture(1)   =   "AtencionHospDetalle.frx":1FF5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Informacion de mortalidad (F8)"
      TabPicture(2)   =   "AtencionHospDetalle.frx":2011
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Nac. y muerte fetal perinatal (F11)"
      TabPicture(3)   =   "AtencionHospDetalle.frx":202D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Datos de egreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   -74850
         TabIndex        =   50
         Top             =   390
         Width           =   11745
         Begin VB.CommandButton btnBuscarMedicosEgreso 
            Caption         =   "..."
            Height          =   315
            Left            =   2580
            TabIndex        =   55
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox txtIdMedicoEgreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            TabIndex        =   54
            Top             =   600
            Width           =   885
         End
         Begin VB.TextBox txtIdServicioEgreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            TabIndex        =   53
            Top             =   240
            Width           =   885
         End
         Begin VB.CommandButton btnBuscarServicioEgreso 
            Caption         =   "..."
            Height          =   315
            Left            =   2580
            TabIndex        =   52
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtNroCamaEgreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7350
            TabIndex        =   51
            Top             =   960
            Width           =   1125
         End
         Begin MSMask.MaskEdBox txtHoraEgreso 
            Height          =   315
            Left            =   2790
            TabIndex        =   56
            Top             =   960
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaEgreso 
            Height          =   315
            Left            =   1620
            TabIndex        =   57
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtHoraEgresoAdm 
            Height          =   315
            Left            =   2790
            TabIndex        =   58
            Top             =   1320
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaEgresoAdm 
            Height          =   315
            Left            =   1620
            TabIndex        =   59
            Top             =   1320
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbCondicionAlta 
            Height          =   315
            Left            =   7350
            TabIndex        =   60
            Top             =   240
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbTipoAlta 
            Height          =   315
            Left            =   7350
            TabIndex        =   61
            Top             =   600
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label lblNombreMedicoEgreso 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2970
            TabIndex        =   70
            Top             =   600
            Width           =   2595
         End
         Begin VB.Label Label43 
            Caption         =   "Medico egreso"
            Height          =   315
            Left            =   120
            TabIndex        =   69
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha egreso"
            Height          =   315
            Left            =   120
            TabIndex        =   68
            Top             =   990
            Width           =   1005
         End
         Begin VB.Label Label44 
            Caption         =   "Fecha egreso adm"
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Top             =   1350
            Width           =   1605
         End
         Begin VB.Label Label46 
            Caption         =   "Condición alta"
            Height          =   315
            Left            =   5850
            TabIndex        =   66
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label48 
            Caption         =   "Tipo alta"
            Height          =   285
            Left            =   5850
            TabIndex        =   65
            Top             =   630
            Width           =   1425
         End
         Begin VB.Label Label49 
            Caption         =   "Servicio ingreso"
            Height          =   315
            Left            =   120
            TabIndex        =   64
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label lblNombreServicioEgreso 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2970
            TabIndex        =   63
            Top             =   240
            Width           =   2595
         End
         Begin VB.Label Label54 
            Caption         =   "Nro Cama egreso"
            Height          =   225
            Left            =   5880
            TabIndex        =   62
            Top             =   990
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos del ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   150
         TabIndex        =   30
         Top             =   390
         Width           =   11745
         Begin VB.TextBox txtNroCamaIngreso 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7320
            TabIndex        =   36
            Top             =   960
            Width           =   885
         End
         Begin VB.TextBox txtIdMedicoIngreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7320
            TabIndex        =   35
            Top             =   600
            Width           =   885
         End
         Begin VB.TextBox txtIdServicioIngreso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7320
            TabIndex        =   34
            Top             =   240
            Width           =   885
         End
         Begin VB.CommandButton btnBuscarServicios 
            Caption         =   "..."
            Height          =   315
            Left            =   8280
            TabIndex        =   33
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton btnBuscarMedicos 
            Caption         =   "..."
            Height          =   315
            Left            =   8280
            TabIndex        =   32
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox txtEdadEnDias 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4830
            TabIndex        =   31
            Top             =   960
            Width           =   735
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   2850
            TabIndex        =   37
            Top             =   960
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdTipoServicio 
            Height          =   315
            Left            =   1650
            TabIndex        =   38
            Top             =   240
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSMask.MaskEdBox txtFechaIngreso 
            Height          =   315
            Left            =   1650
            TabIndex        =   39
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbIdViasAdmision 
            Height          =   315
            Left            =   1650
            TabIndex        =   40
            Top             =   600
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label lblNroCamaIngreso 
            Caption         =   "Nro cama"
            Height          =   345
            Left            =   5820
            TabIndex        =   49
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label lblViaAdmision 
            Caption         =   "Origen"
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   630
            Width           =   1155
         End
         Begin VB.Label lblIdMedicoIngreso 
            Caption         =   "Medico ingreso"
            Height          =   315
            Left            =   5790
            TabIndex        =   47
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lblIdServicioIngreso 
            Caption         =   "Servicio ingreso"
            Height          =   315
            Left            =   5760
            TabIndex        =   46
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label lblNombreServicio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8670
            TabIndex        =   45
            Top             =   240
            Width           =   2595
         End
         Begin VB.Label lblNombreMedico 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8670
            TabIndex        =   44
            Top             =   600
            Width           =   2595
         End
         Begin VB.Label lblIdTipoServicio 
            Caption         =   "Tipo de servicio"
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label lblEdadEnDias 
            Caption         =   "Edad "
            Height          =   315
            Left            =   4230
            TabIndex        =   42
            Top             =   990
            Width           =   1005
         End
         Begin VB.Label lblFecha 
            Caption         =   "Fecha ingreso"
            Height          =   315
            Left            =   150
            TabIndex        =   41
            Top             =   990
            Width           =   1005
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Condición del paciente"
         Height          =   675
         Left            =   150
         TabIndex        =   24
         Top             =   5340
         Width           =   11745
         Begin MSDataListLib.DataCombo cmbIdCondicionEnElServicio 
            Height          =   315
            Left            =   1680
            TabIndex        =   25
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbIdCondicionEnElEstablecimiento 
            Height          =   315
            Left            =   6870
            TabIndex        =   26
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label55 
            Caption         =   "En el servicio"
            Height          =   285
            Left            =   300
            TabIndex        =   28
            Top             =   300
            Width           =   1785
         End
         Begin VB.Label Label56 
            Caption         =   "En el establecimiento"
            Height          =   285
            Left            =   5010
            TabIndex        =   27
            Top             =   300
            Width           =   2265
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   150
         TabIndex        =   15
         Top             =   1800
         Width           =   11745
         Begin VB.CommandButton btnQuitarDx 
            Caption         =   "Quitar"
            Height          =   315
            Left            =   5730
            TabIndex        =   19
            Top             =   630
            Width           =   1335
         End
         Begin VB.CommandButton btnAgregarDx 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   4200
            TabIndex        =   18
            Top             =   630
            Width           =   1305
         End
         Begin VB.CommandButton btnBusquedaDiagnostico 
            Caption         =   ".."
            Height          =   315
            Left            =   2580
            TabIndex        =   17
            Top             =   240
            Width           =   345
         End
         Begin VB.TextBox txtIdDiagnostico 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            TabIndex        =   16
            Top             =   240
            Width           =   1005
         End
         Begin MSDataListLib.DataCombo cmbIdTipoDiagnostico 
            Height          =   315
            Left            =   1500
            TabIndex        =   20
            Top             =   600
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label45 
            Caption         =   "Tipo diagnóstico"
            Height          =   285
            Left            =   150
            TabIndex        =   23
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label lblDescripcionDx 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2970
            TabIndex        =   22
            Top             =   240
            Width           =   7965
         End
         Begin VB.Label Label16 
            Caption         =   "Diagnostico"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   1065
         End
      End
      Begin UltraGrid.SSUltraGrid grdDiagnosticos 
         Height          =   2325
         Left            =   150
         TabIndex        =   29
         Top             =   2970
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   4101
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Lista de diagnósticos"
      End
   End
End
Attribute VB_Name = "AtencionHospDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim ml_TipoServicio As Long
'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA ATENCION
'------------------------------------------------------------------------------------
Dim mo_Atenciones As New DOAtencion
Dim ml_IdCuentaAtencion  As Long
Dim ml_IdAtencion As Long

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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property


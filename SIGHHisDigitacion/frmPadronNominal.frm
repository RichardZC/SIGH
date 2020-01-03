VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPadronNominal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Padrón Nominal"
   ClientHeight    =   4365
   ClientLeft      =   3435
   ClientTop       =   3645
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   10545
   Begin VB.Frame Frame8 
      Caption         =   " "
      Height          =   1095
      Left            =   120
      TabIndex        =   36
      Top             =   3240
      Width           =   10335
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmPadronNominal.frx":0000
         DownPicture     =   "frmPadronNominal.frx":04C4
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
         Left            =   5130
         Picture         =   "frmPadronNominal.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmPadronNominal.frx":0E9C
         DownPicture     =   "frmPadronNominal.frx":12FC
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
         Left            =   3600
         Picture         =   "frmPadronNominal.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   255
         Width           =   1365
      End
   End
   Begin VB.Frame fraDetalleAtencion 
      Caption         =   "---"
      Height          =   2295
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   10335
      Begin VB.ComboBox cmbParasitologico 
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
         ItemData        =   "frmPadronNominal.frx":1BE6
         Left            =   8040
         List            =   "frmPadronNominal.frx":1BF0
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtHemoglobina 
         Alignment       =   1  'Right Justify
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
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   14
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox txtApeMaterno 
         Alignment       =   2  'Center
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
         Left            =   7440
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Presione ENTER para Buscar en la Base de Datos Local"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtApePaterno 
         Alignment       =   2  'Center
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
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Presione ENTER para Buscar en la Base de Datos Local"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
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
         Left            =   1630
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1800
         Width           =   1050
      End
      Begin VB.TextBox txtTalla 
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
         Left            =   2685
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1800
         Width           =   800
      End
      Begin VB.ComboBox CmbDiagNutricional 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtNroHC_FF_COD 
         Alignment       =   2  'Center
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
         MaxLength       =   6
         TabIndex        =   1
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox txtNombresCompletos 
         Alignment       =   2  'Center
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
         MaxLength       =   40
         TabIndex        =   5
         ToolTipText     =   "Presione ENTER para Buscar en la Base de Datos Local"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.ComboBox cmbFinanciador 
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cmbTipoDocumento 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cmbSexo 
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
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtNroDocumento 
         Alignment       =   2  'Center
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
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   2
         ToolTipText     =   "Presione ENTER para Buscar en la Base de Datos Local"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNroAfiliacion 
         Alignment       =   2  'Center
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
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFecNacimiento 
         Height          =   315
         Left            =   5040
         TabIndex        =   7
         Top             =   1200
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
      Begin MSMask.MaskEdBox txtFecEvaluacion 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
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
      Begin VB.Label Label 
         Caption         =   "Parasitológico en heces"
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
         Index           =   6
         Left            =   8040
         TabIndex        =   44
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Hemoglobina"
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
         Index           =   1
         Left            =   6870
         TabIndex        =   43
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Fec de Evaluac."
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
         Index           =   5
         Left            =   150
         TabIndex        =   40
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Diagnostico Nutricional"
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
         Index           =   1
         Left            =   3510
         TabIndex        =   39
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Peso Kg"
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
         Index           =   12
         Left            =   1680
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Talla Cm"
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
         Index           =   13
         Left            =   2670
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Historia Clínica"
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
         Index           =   2
         Left            =   150
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Fec. de Nac."
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
         Index           =   4
         Left            =   5070
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Nº de Afiliación"
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
         Index           =   3
         Left            =   8760
         TabIndex        =   32
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Tipo de Seguro"
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
         Index           =   2
         Left            =   6360
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label label3 
         Caption         =   "Sexo"
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
         Left            =   3630
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nombres"
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
         Index           =   3
         Left            =   150
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido Materno"
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
         Index           =   2
         Left            =   7470
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido Paterno"
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
         Index           =   1
         Left            =   4830
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo Doc."
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
         Index           =   10
         Left            =   1580
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Edición de Atención"
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
         TabIndex        =   25
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Nro Doc."
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
         Index           =   0
         Left            =   3390
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
   End
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
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.TextBox txtCodUbigeo 
         Alignment       =   2  'Center
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
         TabIndex        =   42
         Top             =   435
         Width           =   1455
      End
      Begin VB.TextBox txtCodEstablec 
         Alignment       =   2  'Center
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
         TabIndex        =   19
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox txtUbigeoEstablecimiento 
         Alignment       =   2  'Center
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
         Left            =   1320
         TabIndex        =   18
         Top             =   435
         Width           =   3615
      End
      Begin VB.TextBox txtCodigoEstadistico 
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
         Left            =   6360
         TabIndex        =   17
         Top             =   435
         Width           =   3705
      End
      Begin VB.Label Label 
         Caption         =   "Cod Ubigeo"
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
         Index           =   0
         Left            =   4950
         TabIndex        =   34
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Establecimiento"
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
         Left            =   1350
         TabIndex        =   21
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Responsable Digitación"
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
         Index           =   0
         Left            =   6390
         TabIndex        =   20
         Top             =   180
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmPadronNominal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica en donde se ingresaran las atenciones del Pdron Nominal.
'        Programado por: Palomino F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbTipoDocumento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbFinanciador As New SIGHEntidades.ListaDespleglable
Dim mo_cmbSexo As New SIGHEntidades.ListaDespleglable
Dim mo_CmbDiagNutricional As New SIGHEntidades.ListaDespleglable
Dim mo_cmbParasitologico As New SIGHEntidades.ListaDespleglable
Dim mo_PadronNominalDetalle As New ReglasHISGalenos
Dim mb_PrimerIngresoCabeceraAtencion As Boolean
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim lcBuscaParametro As New SIGHDatos.Parametros

'---------------------------- variables de manejo de negocio -------------------------------
Dim oCabeceraAtencion As New DOHIS_Cabecera             'Contiene los datos de la cabecera de atencion
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_DatosParametro As New SIGHDatos.Parametros       'Representa la fecha y hora del servidor

Dim oRcs_DetalleAtencion As New Recordset               'Representa el detalle de las Atencion
Dim oRcs_DetalleAtencionTemp As New Recordset
Dim oRcs_Diagnosticos As New Recordset                  'Representa el detalle de Diagnosticos de la Atencion
Dim oRcs_DiagnosticosTemp As New Recordset              'Representa el detalle de Diagnosticos para una Atencion, solo existe por Atencion.

Dim oPadronNominal_Detalle As New DoPadronNominal_Detalle
Dim oMensaje As New SIGHNegocios.clMensaje

Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim ml_IdPadNominal As Long
Dim ml_IdLote As Long
'Dim ml_IdEstablecimiento As Long
Dim ml_IdUsuario As Long
Dim ml_IdUsuarioRegistro As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim mi_Opcion As sghOpciones
Dim mi_BotonPresionado As sghBotonDetallePresionado

'Datos de Inicio de Formulario
Dim ml_IdDepartamentoActual As Long: Dim ms_NombreDepActual As String
Dim ml_IdProvinciaActual As Long: Dim ms_NombreProvActual As String
Dim ml_IdDistritoActual As Long: Dim ms_NombreDistrActual As String
Dim ml_IdEstablecimiento As Long: Dim ms_CodigoEstablecimiento As String: Dim ms_NombreEstablecimientoActual As String
Dim mb_SeleccionoLote As Boolean: Dim mb_SeleccionoMedico As Boolean
Dim mb_PesoTallaHabilitados As Boolean

Dim mo_LoteActual As New DOHIS_Lotes
Dim ml_CodigoResponsableDigitacion As Long: Dim ms_NombreRespDigitacion As String

Dim mb_FaltaGrabarAtencion As Boolean
Dim ml_IdPacienteGalenHos As Long

Dim IdTipoActividad As Integer
Dim IdDetalleDiagnostico As Integer
Dim ml_IdMedicoResponsable As Long
Dim ml_IdDistritoAtencion As Long
Dim ml_IdNacionalidadAtencion As Long
Const mi_CantidadMaxDiagnosticos As Integer = 6

'Datos de Cuadros de Dialogo
Dim IdCodigoActividad As Long
Dim IdCodigoNacionalidad As Long
Dim ldZ_PE As Double
Dim ldZ_PT As Double
Dim ldZ_TE As Double
Dim ldiddxnutricionalPE As Double
Dim ldiddxnutricionalPT As Double
Dim ldiddxnutricionalTE As Double

'Colores de Indicacion
Const MI_IDNACIONALIDAD As Integer = 166
Const MS_NOMBRE_NAC As String = "PER"
Const ml_ColorCorrecto As Long = &HFFFFFF
Const ml_ColorError As Long = &HFF6347
Const ml_ColorMensaje As Long = &HFCD33E

'========================================== PROPIEDADES ====================================
'Propiedades de Modulo
Property Let CabeceraAtencion(oValue As DOHIS_Cabecera)
   Set oCabeceraAtencion = oValue
End Property

Property Get CabeceraAtencion() As DOHIS_Cabecera
   CabeceraAtencion = oCabeceraAtencion
End Property

Property Let IdPadNominal(lValue As Long)
   ml_IdPadNominal = lValue
End Property

Property Get IdPadNominal() As Long
   IdPadNominal = ml_IdPadNominal
End Property

Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub btnCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub CmbDiagNutricional_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.txtHemoglobina
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub Command1_Click()
    Unload Me
    Me.Visible = False
    'LimpiarVariablesDeMemoria
End Sub

Private Sub cmbParasitologico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, btnAceptar
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Sub DeshabilitarControles()
    mo_Formulario.HabilitarDeshabilitar txtNroHC_FF_COD, False
    mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtNroDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtApePaterno, False
    mo_Formulario.HabilitarDeshabilitar txtApeMaterno, False
    
    mo_Formulario.HabilitarDeshabilitar txtNombresCompletos, False
    mo_Formulario.HabilitarDeshabilitar cmbSexo, False
    mo_Formulario.HabilitarDeshabilitar txtFecNacimiento, False
    mo_Formulario.HabilitarDeshabilitar cmbFinanciador, False
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion, False
    
    mo_Formulario.HabilitarDeshabilitar txtFecEvaluacion, False
    mo_Formulario.HabilitarDeshabilitar txtPeso, False
    mo_Formulario.HabilitarDeshabilitar txtTalla, False
    mo_Formulario.HabilitarDeshabilitar CmbDiagNutricional, False
    mo_Formulario.HabilitarDeshabilitar txtHemoglobina, False
    mo_Formulario.HabilitarDeshabilitar cmbParasitologico, False
    mo_Formulario.HabilitarDeshabilitar btnAceptar, False
'    Me.btnAceptar.Visible = False
End Sub

Sub CargarDatosDelEstablecimiento()
            ml_IdDepartamentoActual = Val(Left("0" & Right(lcBuscaParametro.SeleccionaFilaParametro(242), 6), 2))
            'ms_NombreDepActual = oRcs_Temp!NombreDepartamento
            ml_IdProvinciaActual = Val(Left(Right("0" & lcBuscaParametro.SeleccionaFilaParametro(242), 6), 4))
            'ms_NombreProvActual = oRcs_Temp!NombreProvincia
            ml_IdDistritoActual = Val(lcBuscaParametro.SeleccionaFilaParametro(242))
            'ms_NombreDistrActual = oRcs_Temp!NombreDistrito
            ms_CodigoEstablecimiento = lcBuscaParametro.SeleccionaFilaParametro(280)
            ms_NombreEstablecimientoActual = lcBuscaParametro.SeleccionaFilaParametro(205)
End Sub

Private Sub Form_Load()
    CargarDatosDelEstablecimiento
    CargarComboBoxes
    CargarDatosAlFormulario
    mo_Formulario.HabilitarDeshabilitar txtCodEstablec, False
    mo_Formulario.HabilitarDeshabilitar txtUbigeoEstablecimiento, False
    mo_Formulario.HabilitarDeshabilitar txtCodUbigeo, False
    mo_Formulario.HabilitarDeshabilitar txtCodigoEstadistico, False
    mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, False
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar registro en el padrón nominal"
        mb_PesoTallaHabilitados = False
        mb_PrimerIngresoCabeceraAtencion = True
    Case sghModificar, sghConsultar, sghEliminar
        If mi_Opcion = sghModificar Then
            Me.btnAceptar.Visible = True
            Me.Caption = "Modificar registro en el padrón nominal"
        ElseIf mi_Opcion = sghConsultar Then
            Me.btnAceptar.Visible = False
            Me.Caption = "Consultar registro en el padrón nominal"
        ElseIf mi_Opcion = sghEliminar Then
            Me.Caption = "Eliminar registro en el padrón nominal"
            Me.btnAceptar.Visible = True
        End If
    End Select
    
    'ASIGNAR LOS VALORES POR DEFECTO DEL REGISTRO DE ATENCION
    If mi_Opcion = sghAgregar Then
        ControlesAtencionPorDefecto
    End If
    Me.Refresh
End Sub

'========================================== EVENTOS ========================================
Private Sub Form_Initialize()
    Set mo_cmbTipoDocumento.MiComboBox = Me.cmbTipoDocumento
    Set mo_cmbFinanciador.MiComboBox = Me.cmbFinanciador
    Set mo_CmbDiagNutricional.MiComboBox = Me.CmbDiagNutricional
    Set mo_cmbSexo.MiComboBox = Me.cmbSexo
   ' Set mo_cmbParasitologico.MiComboBox = Me.cmbParasitologico
End Sub

Sub CargarComboBoxes()
    'Tipo de Documento
    Dim oRcs_Lista As New Recordset
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposDocumento
    oRcs_Lista.MoveFirst
    mo_cmbTipoDocumento.BoundColumn = "IdDocIdentidad"
    mo_cmbTipoDocumento.ListField = "DescripcionLarga"
    Set mo_cmbTipoDocumento.RowSource = oRcs_Lista
    
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Genero
    
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposSexo
    oRcs_Lista.MoveFirst
    mo_cmbSexo.BoundColumn = "IdTipoSexo"
    mo_cmbSexo.ListField = "Descripcionlarga"
    Set mo_cmbSexo.RowSource = oRcs_Lista
    Set oRcs_Lista = Nothing
    
    'Codigo del Tipo de Financiamiento
    Set oRcs_Lista = mo_ReglasHIS.ListaFuentesFinanciamiento
    oRcs_Lista.MoveFirst
    mo_cmbFinanciador.BoundColumn = "IdCodigoFinancHis"
    mo_cmbFinanciador.ListField = "DescripcionLarga"
    Set mo_cmbFinanciador.RowSource = oRcs_Lista
    
    Set oRcs_Lista = Nothing
    
    'Codigo Liata de diagnostico nutricional
    Set oRcs_Lista = mo_ReglasHIS.ListaDiagNutricional
    oRcs_Lista.MoveFirst
    mo_CmbDiagNutricional.BoundColumn = "IdCodigoDx"
    mo_CmbDiagNutricional.ListField = "DescripcionLarga"
    Set mo_CmbDiagNutricional.RowSource = oRcs_Lista
    
    
'    Dim orsTemp As New ADODB.Recordset
'    With orsTemp
'          .Fields.Append "Codigo", adInteger
'          .Fields.Append "Valor", adVarChar, 100, adFldIsNullable
'          .CursorType = adOpenDynamic
'          .LockType = adLockOptimistic
'          .Open
'    End With
'    orsTemp.AddNew
'    orsTemp.Fields!Codigo = 1
'    orsTemp.Fields!Valor = "SI"
'    orsTemp.Update
'    orsTemp.AddNew
'    orsTemp.Fields!Codigo = 2
'    orsTemp.Fields!Valor = "NO"
'    orsTemp.Update
'
'    oRcs_Lista.MoveFirst
'    mo_cmbParasitologico.BoundColumn = "Codigo"
'    mo_cmbParasitologico.ListField = "Valor"
'    Set mo_cmbParasitologico.RowSource = orsTemp
'    mo_cmbParasitologico.BoundText = 2
    
    Set oRcs_Lista = Nothing
'    Set orsTemp = Nothing
End Sub
Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub cmbTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumento
   AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub cmbTipoDocumento_LostFocus()
   If cmbTipoDocumento.Text <> "" Then
     mo_Formulario.MarcarComoVacio cmbTipoDocumento
   End If
End Sub

Private Sub cmbTipoDocumento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFecNacimiento
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub cmbSexo_LostFocus()
   If cmbSexo.Text <> "" Then
      mo_Formulario.MarcarComoVacio cmbSexo
   End If
End Sub

'cmbFinanciador
Private Sub cmbFinanciador_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFecEvaluacion
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub cmbFinanciador_LostFocus()
   If cmbFinanciador.Text <> "" Then
        On Error Resume Next
       mo_cmbFinanciador.BoundText = Val(Split(cmbFinanciador.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbFinanciador
End Sub

Private Sub cmbFinanciador_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Public Sub MostrarFormulario()
Me.Show 1
End Sub

Private Sub txtApeMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombresCompletos
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtApeMaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
    Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(" "), Asc("Ñ"), Asc("ñ")
    Case Else
    KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtApeMaterno_LostFocus()
    txtApeMaterno.Text = UCase(txtApeMaterno.Text)
End Sub

Private Sub txtApePaterno_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtApeMaterno
       AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtApePaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
    Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(" "), Asc("Ñ"), Asc("ñ")
    Case Else
    KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtApePaterno_LostFocus()
    txtApePaterno.Text = UCase(txtApePaterno.Text)
End Sub


Private Sub txtFecEvaluacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPeso
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtFecEvaluacion_LostFocus()
'    mo_CmbDiagNutricional.BoundText = CalculaIdDiagnostico
End Sub

Private Sub txtFecNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbFinanciador
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtFecNacimiento_LostFocus()
'    mo_CmbDiagNutricional.BoundText = CalculaIdDiagnostico
End Sub

Private Sub txtHemoglobina_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbParasitologico
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtHemoglobina_KeyPress(KeyAscii As Integer)
    If Not (mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNombresCompletos_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbSexo
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtNombresCompletos_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
    Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(" "), Asc("Ñ"), Asc("ñ")
    Case Else
    KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtNombresCompletos_LostFocus()
    txtNombresCompletos.Text = UCase(txtNombresCompletos.Text)
End Sub

Private Sub txtNroAfiliacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFecEvaluacion
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtNroAfiliacion_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtNroHC_FF_COD_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumento
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtNroHC_FF_COD_LostFocus()
    Dim oRcs_Temp As Recordset
    'Buscar datos
    If Trim(txtNroHC_FF_COD.Text) <> "" Then
        Set oRcs_Temp = mo_ReglasHIS.PadNominalBuscarDatosxNroHC(Trim(txtNroHC_FF_COD.Text))
        If oRcs_Temp.RecordCount <> 0 Then
            oRcs_Temp.MoveFirst
            mo_cmbTipoDocumento.BoundText = oRcs_Temp.Fields!IdTipoDoc
            Me.txtNroDocumento.Text = IIf(IsNull(oRcs_Temp.Fields!NumDocumento), "", Right("000" & Trim(Str(oRcs_Temp.Fields!NumDocumento)), 8))
            Me.txtApePaterno.Text = oRcs_Temp.Fields!ApellidoPaterno
            Me.txtApeMaterno.Text = oRcs_Temp.Fields!ApellidoMaterno
            Me.txtNombresCompletos.Text = oRcs_Temp.Fields!Nombres
            mo_cmbSexo.BoundText = oRcs_Temp.Fields!idSexo
            Me.txtFecNacimiento.Text = oRcs_Temp.Fields!FecNacimiento
            If mi_Opcion = sghAgregar Then
                mo_cmbFinanciador.BoundText = oRcs_Temp.Fields!IdTipoSeguro
                Me.txtNroAfiliacion.Text = IIf(IsNull(oRcs_Temp.Fields!NumAfiliacion), "", oRcs_Temp.Fields!NumAfiliacion)
            End If
            'Me.txtFecEvaluacion.SetFocus
        End If
        oRcs_Temp.Close
        Set oRcs_Temp = Nothing
    End If
End Sub

Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtTalla
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
       If Not (mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Or KeyAscii = 8) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtPeso_LostFocus()
'    mo_CmbDiagNutricional.BoundText = CalculaIdDiagnostico
End Sub

Private Sub txtTalla_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, btnAceptar
    AdministrarKeyPreview CInt(KeyCode)
End Sub

'LIMPIAMOS LOS CONTROLES CON SUS VALORES POR DEFECTO
Private Sub ControlesAtencionPorDefecto()
    On Error GoTo ControlesAtencionPorDefecto_Error
    mo_cmbTipoDocumento.BoundText = "1"
    txtNroDocumento.Text = ""
    mo_cmbFinanciador.BoundText = "2"
    mo_cmbSexo.BoundText = "1"
    On Error GoTo 0
    Exit Sub
ControlesAtencionPorDefecto_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ControlesAtencionPorDefecto of Formulario frmMantenimientoHIS"
End Sub

Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    'solo para enter y tab
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And mi_Opcion = sghAgregar Then
        'Solo para DNI
        If Val(mo_cmbTipoDocumento.BoundText) = 1 Then
            Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
            Dim mo_HisGalenhos As New SIGHNegocios.ReglasHISGalenos
            Dim mo_DatosFechas As New SIGHEntidades.FechaHora
            Dim o_RcsDatosPaciente As New Recordset
            Dim o_RcsDatosPacienteSISAfiliacion As New Recordset
            
            Set o_RcsDatosPaciente = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
            Set o_RcsDatosPacienteSISAfiliacion = mo_HisGalenhos.SisafilicianesObtenerNombreApellidosPorNroDocumentoYtipo(Me.txtNroDocumento.Text, Val(mo_cmbTipoDocumento.BoundText))
            
            
            'Si Encuentra apellidos y nombres del paciente en SISAfiliacion (SIGHExterna) EYPE
            If o_RcsDatosPacienteSISAfiliacion.RecordCount <> 0 Then
                o_RcsDatosPacienteSISAfiliacion.MoveFirst
                Me.txtApePaterno.Text = o_RcsDatosPacienteSISAfiliacion!Paterno
                Me.txtApeMaterno.Text = o_RcsDatosPacienteSISAfiliacion!Materno
                Me.txtNombresCompletos = o_RcsDatosPacienteSISAfiliacion!Pnombre & " " & o_RcsDatosPacienteSISAfiliacion!ONombres
                Me.txtNroAfiliacion = o_RcsDatosPacienteSISAfiliacion!AfiliacionNroFormato
                Me.txtFecNacimiento = o_RcsDatosPacienteSISAfiliacion!Fnacimiento
                If o_RcsDatosPacienteSISAfiliacion!Genero = 1 Then
                    mo_cmbSexo.BoundText = o_RcsDatosPacienteSISAfiliacion!Genero
                Else
                    mo_cmbSexo.BoundText = 2
                End If
                
            End If
                If KeyCode = vbKeyReturn Then
                    mo_Teclado.RealizarNavegacion KeyCode, txtNroHC_FF_COD
                End If
                'Si Encuentra algun Dato del Paciente
                If o_RcsDatosPaciente.RecordCount <> 0 Then
                    o_RcsDatosPaciente.MoveFirst
        
                    'IDPACIENTE
                    ml_IdPacienteGalenHos = CLng(o_RcsDatosPaciente!IdPaciente)
                    mo_cmbSexo.BoundText = CStr(o_RcsDatosPaciente!IdTipoSexo)
                End If
        End If
    ElseIf mi_Opcion = sghModificar Then
        If KeyCode = vbKeyReturn Then
            mo_Teclado.RealizarNavegacion KeyCode, txtNroHC_FF_COD
        End If
    End If
    AdministrarKeyPreview CInt(KeyCode)
End Sub

'CARGA DE DATOS INICIAL EN EL FORMUALRIO PRINCIPAL
Sub CargarDatosAlFormulario()
    mb_FaltaGrabarAtencion = False
    Dim oRcs_Temp As New ADODB.Recordset
    
    ml_IdUsuarioRegistro = 0
    If mi_Opcion = sghOpciones.sghConsultar Or mi_Opcion = sghOpciones.sghModificar Or mi_Opcion = sghOpciones.sghEliminar Then
        Dim oPadronNominal_Detalle As New DoPadronNominal_Detalle
        oPadronNominal_Detalle.IdPaNomDetalle = ml_IdPadNominal
        If mo_ReglasHIS.PadronNominal_DetalleSeleccionarPorId(oPadronNominal_Detalle) Then
            mo_cmbTipoDocumento.BoundText = oPadronNominal_Detalle.IdTipoDoc
            Me.txtNroDocumento.Text = IIf(oPadronNominal_Detalle.NumDocumento = 0, "", Right("000" & Trim(Str(oPadronNominal_Detalle.NumDocumento)), 8))
            Me.txtNroHC_FF_COD.Text = oPadronNominal_Detalle.HistClinica
            Me.txtApePaterno.Text = oPadronNominal_Detalle.ApellidoPaterno
            Me.txtApeMaterno.Text = oPadronNominal_Detalle.ApellidoMaterno
            Me.txtNombresCompletos.Text = oPadronNominal_Detalle.Nombres
            mo_cmbSexo.BoundText = oPadronNominal_Detalle.idSexo
            Me.txtFecNacimiento.Text = oPadronNominal_Detalle.FecNacimiento
            mo_cmbFinanciador.BoundText = oPadronNominal_Detalle.IdTipoSeguro
            Me.txtNroAfiliacion.Text = oPadronNominal_Detalle.NumAfiliacion
            Me.txtFecEvaluacion.Text = oPadronNominal_Detalle.FecEvaluacion
            Me.txtPeso.Text = oPadronNominal_Detalle.Peso
            Me.txtTalla.Text = oPadronNominal_Detalle.Talla
            mo_CmbDiagNutricional.BoundText = oPadronNominal_Detalle.IdDiagNutricional
            Me.txtHemoglobina.Text = oPadronNominal_Detalle.hemoglobina
            cmbParasitologico.ListIndex = IIf(oPadronNominal_Detalle.Heces = "SI", 0, 1)
'            ml_IdEstablecimiento = oPadronNominal_Detalle.IdEstablecimiento
'            If ml_IdUsuario <> oPadronNominal_Detalle.IdUsuario Then
'                DeshabilitarControles
'            End If
'            ml_IdUsuarioRegistro = oPadronNominal_Detalle.IdUsuario
        End If
    End If
    
'    'CARGA DATOS DEL ESTABLECIMIENTO ACTUAL
'    Set oRcs_Temp = mo_ReglasHIS.HIS_DatosEstablecimientoXidEstablecimiento(ml_IdEstablecimiento)
'    If oRcs_Temp.RecordCount <> 0 Then
'        oRcs_Temp.MoveFirst
'        Do While Not oRcs_Temp.EOF
'            ml_IdDepartamentoActual = oRcs_Temp!IdDepartamento
'            ms_NombreDepActual = oRcs_Temp!NombreDepartamento
'            ml_IdProvinciaActual = oRcs_Temp!IdProvincia
'            ms_NombreProvActual = oRcs_Temp!NombreProvincia
'            ml_IdDistritoActual = oRcs_Temp!IdDistrito
'            ms_NombreDistrActual = oRcs_Temp!NombreDistrito
''            ml_IdEstablecimientoActual = oRcs_Temp!IdEstablecimiento
'            ms_CodigoEstablecimiento = oRcs_Temp!Codigo
'            ms_NombreEstablecimientoActual = oRcs_Temp!NombreEstablecimiento
'            oRcs_Temp.MoveNext
'        Loop
'    End If
    
    'CARGAR DATOS DEL DIGITADOR ACTUAL
    If ml_IdUsuarioRegistro = 0 Then
        Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosDigitador(ml_IdUsuario)
    Else
        Set oRcs_Temp = mo_ReglasHIS.ObtenerDatosDigitador(ml_IdUsuarioRegistro)
    End If
    
    'verifica si tiene grabado el codigo de responsable de digitacion
    If IsNull(oRcs_Temp.Fields(2).Value) Then
        MsgBox "No tiene configurado el Codigo de Responsable de Digitacion", vbInformation, "HIS SIGH"
        ml_CodigoResponsableDigitacion = 0
        ms_NombreRespDigitacion = oRcs_Temp.Fields(1)
    Else
        ml_CodigoResponsableDigitacion = oRcs_Temp.Fields(2)
        ms_NombreRespDigitacion = oRcs_Temp.Fields(1)
    End If
        
    'INGRESO DE VALORES CONTROLES VISUALES
    txtCodEstablec.Text = ms_CodigoEstablecimiento
    txtUbigeoEstablecimiento.Text = ms_NombreEstablecimientoActual
    txtCodigoEstadistico.Text = ml_CodigoResponsableDigitacion & " - " & ms_NombreRespDigitacion
    txtCodUbigeo.Text = ml_IdDistritoActual
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnAceptar_Click
' Author    : User
' Date      : 22/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : btnAceptar_Click
' Author    : User
' Date      : 22/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : btnAceptar_Click
' Author    : User
' Date      : 22/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
        If ValidarDatosObligatorios() Then
             If AgregarDatos() Then
                   MsgBox " Los datos se agregarón exitosamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_PadronNominalDetalle.MensajeError, vbExclamation, Me.Caption
               End If
        End If
        
   Case sghModificar
        If ValidarDatosObligatorios() Then
             If ModificarDatos() Then
                   MsgBox " Los datos se modificarón exitosamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_PadronNominalDetalle.MensajeError, vbExclamation, Me.Caption
               End If
        End If
   Case sghEliminar
        If MsgBox("¿Desea eliminar el registro del padrón nominal?", vbOKCancel, Me.Caption) = vbOK Then
             If EliminarDatos() Then
                 MsgBox " Los datos se eliminarón correctamente", vbInformation, Me.Caption
                 Me.Visible = False
             Else
                 MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_PadronNominalDetalle.MensajeError, vbExclamation, Me.Caption
             End If
         End If
   End Select

   On Error GoTo 0
   Exit Sub
btnAceptar_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAceptar_Click of Formulario frmPadronNominal"

End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   Dim sMensajeNutricional As String
   ValidarDatosObligatorios = False
    
   If Len(Trim(Me.txtNroHC_FF_COD.Text)) = 0 Then
       sMensaje = sMensaje + "- Ingrese el número de historia clínica" + Chr(13)
   End If
   If Len(Trim(Me.txtNroDocumento.Text)) <> 0 Then
        If Len(Trim(Me.txtNroDocumento.Text)) < 8 Then
            sMensaje = sMensaje + "- El número de documento de identidad no debe tener menos de 8 digitos" + Chr(13)
        End If
   End If
   If Len(Trim(Me.txtApePaterno.Text)) = 0 Then
       sMensaje = sMensaje + "- Ingrese el apellido paterno" + Chr(13)
   End If
   If Len(Trim(Me.txtApeMaterno.Text)) = 0 Then
       sMensaje = sMensaje + "- Ingrese el apellido materno" + Chr(13)
   End If
   If Len(Trim(Me.txtNombresCompletos.Text)) = 0 Then
       sMensaje = sMensaje + "- Ingrese nombres completos" + Chr(13)
   End If
    If Me.txtFecNacimiento.Text <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecNacimiento.Text, "DD/MM/AAAA") Then
            sMensaje = sMensaje + "- La fecha de nacimiento no tiene formato correcto" + Chr(13)
        End If
    Else
         sMensaje = sMensaje + "- Ingrese la fecha de nacimiento" + Chr(13)
    End If
    If Me.txtFecEvaluacion.Text <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecEvaluacion.Text, "DD/MM/AAAA") Then
            sMensaje = sMensaje + "- La fecha de evaluación no tiene formato correcto" + Chr(13)
        Else
            If CDate(Me.txtFecEvaluacion.Text) > Date Then
                sMensaje = sMensaje + "- La fecha de evaluación no puede ser mayor que la fecha actual" + Chr(13)
            End If
        End If
    Else
         sMensaje = sMensaje + "- Ingrese la fecha de evaluación" + Chr(13)
    End If
    If SIGHEntidades.EsFecha(txtFecNacimiento.Text, "DD/MM/AAAA") And SIGHEntidades.EsFecha(txtFecEvaluacion.Text, "DD/MM/AAAA") Then
        If CDate(Me.txtFecNacimiento.Text) > CDate(Me.txtFecEvaluacion.Text) Then
            sMensaje = sMensaje + "- La fecha de evaluación no debe ser menor que la fecha de nacimiento" + Chr(13)
        Else
            If SIGHEntidades.EdadActualEnDias(CDate(Me.txtFecNacimiento.Text), CDate(Me.txtFecEvaluacion.Text)) > 1825 Then
                sMensaje = sMensaje + "- La diferencia de la fecha de evaluación y fecha de nacimiento no debe superar los 5 años" + Chr(13)
            End If
        End If
   End If
   If Len(Trim(txtPeso.Text)) = 0 Then
        sMensaje = sMensaje + "- Ingrese el peso" + Chr(13)
   End If
   If Len(Trim(txtTalla.Text)) = 0 Then
        sMensaje = sMensaje + "- Ingrese la talla" + Chr(13)
   End If
   If Trim(Me.cmbSexo.Text) = "" Then
        sMensaje = sMensaje + "- No ha seleccionado el tipo de documento" + Chr(13)
   End If
   If Trim(Me.cmbFinanciador.Text) = "" Then
        sMensaje = sMensaje + "- No ha seleccionado el tipo de seguro" + Chr(13)
   End If
   If Trim(Me.CmbDiagNutricional.Text) = "" Then
        sMensaje = sMensaje + "- No ha seleccionado el diagnóstico " + Chr(13)
   End If
   
   If txtHemoglobina.Text = "" Then
        If lcBuscaParametro.SeleccionaFilaParametro(335) = "S" Then
            sMensaje = sMensaje + "- Ingrese el dato de la Hemoglobina" + Chr(13)
        End If
   Else
        If InStr(txtHemoglobina.Text, ".") = Len(txtHemoglobina.Text) Then
            sMensaje = sMensaje + "- El dato de la Hemoglobina no tiene el formato correcto" + Chr(13)
        Else
            If Not (txtHemoglobina.Text >= 5 And txtHemoglobina.Text <= 20) Then
                sMensaje = sMensaje + "- El dato de la Hemoglobina debe estar entre 5 y 20" + Chr(13)
            End If
        End If
   End If
   
    ldZ_PE = 0
    ldZ_PT = 0
    ldZ_TE = 0
    ldiddxnutricionalPE = 0
    ldiddxnutricionalPT = 0
    ldiddxnutricionalTE = 0

   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   Else
        CalculaIdDiagnostico
        sMensajeNutricional = ""
        'Advierte valores ingresados
        If (ldZ_PE < -3 Or ldZ_PE > 3) Then
            sMensajeNutricional = "P/E"
        End If
        If (ldZ_PT < -3 Or ldZ_PT > 3) Then
            If sMensajeNutricional = "" Then
                sMensajeNutricional = "P/T"
            Else
                sMensajeNutricional = sMensajeNutricional & ",P/T"
            End If
        End If
        If (ldZ_TE < -3 Or ldZ_TE > 3) Then
            If sMensajeNutricional = "" Then
                sMensajeNutricional = "T/E"
            Else
                sMensajeNutricional = sMensajeNutricional & " y T/E"
            End If
        End If
        
        If sMensajeNutricional <> "" Then
            oMensaje.MostrarFormulario Chr(13) & ("¿Está seguro que los datos ingresados son los correctos?. Existen incongruencias en " & sMensajeNutricional & " ¿desea grabar?"), Me.Caption, 14, True, sghRojo, True
            If oMensaje.BotonPresionado = sghCancelar Then
                Exit Function
            End If
        End If
   End If
   
   ValidarDatosObligatorios = True
End Function
              
'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_PadronNominalDetalle.PadronNominalDetalleAgregar(oPadronNominal_Detalle)
End Function

Function ModificarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_PadronNominalDetalle.PadronNominalDetalleModificar(oPadronNominal_Detalle)
End Function

Function EliminarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_PadronNominalDetalle.PadronNominalDetalleEliminar(oPadronNominal_Detalle)
End Function

Sub LimpiarFormulario()
    Me.txtNroDocumento = ""
    Me.txtNroHC_FF_COD.Text = ""
    Me.txtApePaterno.Text = ""
    Me.txtApeMaterno.Text = ""
    Me.txtNombresCompletos.Text = ""
    Me.txtFecNacimiento.Text = "__/__/____"
    Me.txtNroAfiliacion.Text = ""
    Me.txtFecEvaluacion = "__/__/____"
    Me.txtPeso.Text = ""
    Me.txtHemoglobina.Text = ""
    Me.txtTalla.Text = ""
    mo_cmbParasitologico.BoundText = "N"
End Sub

Sub CargaDatosAlObjetosDeDatos()
   With oPadronNominal_Detalle
        .IdPaNomDetalle = ml_IdPadNominal
        .IdTipoDoc = Val(mo_cmbTipoDocumento.BoundText)
        .NumDocumento = IIf(Me.txtNroDocumento.Text = "", 0, Val(Me.txtNroDocumento.Text))
        .HistClinica = Me.txtNroHC_FF_COD.Text
        .ApellidoPaterno = Me.txtApePaterno.Text
        .ApellidoMaterno = Me.txtApeMaterno.Text
        .Nombres = Me.txtNombresCompletos.Text
        .idSexo = Val(mo_cmbSexo.BoundText)
        .FecNacimiento = Me.txtFecNacimiento.Text
        .IdTipoSeguro = Val(mo_cmbFinanciador.BoundText)
        .NumAfiliacion = Me.txtNroAfiliacion.Text
        .FecEvaluacion = Me.txtFecEvaluacion.Text
        .Peso = Me.txtPeso.Text
        .Talla = Me.txtTalla.Text
        .IdDiagNutricional = Val(mo_CmbDiagNutricional.BoundText)
        .CodRenaes = CLng(ms_CodigoEstablecimiento)
        .IdDiagPE = ldiddxnutricionalPE
        .IdDiagPT = ldiddxnutricionalPT
        .IdDiagTE = ldiddxnutricionalTE
        .hemoglobina = Val(Me.txtHemoglobina.Text)
        .Heces = cmbParasitologico.Text
        '.IdUsuario = ml_IdUsuario
'        .ParasitHeces = mo_cmbParasitologico.BoundText
'        .IdEstablecimiento = ml_IdEstablecimiento
   End With
End Sub
Private Sub txtTalla_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub
Private Sub txtTalla_LostFocus()
'    mo_CmbDiagNutricional.BoundText = CalculaIdDiagnostico
End Sub
'Actualiza valores y Devuelve percentil IMC de la ATENCION ACTUAL DEL PACIENTE
Sub CalculaIdDiagnostico()
    Dim orsTemp As Recordset
    Dim ldPesoKg As Double
    Dim lcTallaCM As String
    Dim lnEdadDias As Long
    Dim lnColValor1 As Long
    Dim lnColValor2 As Long
    
    If Me.txtFecNacimiento.Text <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecNacimiento.Text, "DD/MM/AAAA") Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    If Me.txtFecEvaluacion.Text <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecEvaluacion.Text, "DD/MM/AAAA") Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    If CDate(Me.txtFecNacimiento.Text) > CDate(Me.txtFecEvaluacion.Text) Then
       Exit Sub
    End If

    If InStr(txtPeso.Text, "_") >= 1 Then
        Exit Sub
    End If
    
    If Len(Trim(txtTalla.Text)) = 0 Then
        Exit Sub
    End If

    lnEdadDias = DateDiff("d", CDate(Me.txtFecNacimiento.Text), CDate(Me.txtFecEvaluacion.Text))
    ldPesoKg = Val(txtPeso.Text)
    lcTallaCM = txtTalla.Text
    
    If mo_cmbSexo.BoundText = "1" Then ' Hombre
        lnColValor1 = 4
        lnColValor2 = 5
    Else ' Mujer
        lnColValor1 = 20
        lnColValor2 = 21
    End If
    
    On Error Resume Next
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\CRED WHO.xls")

    Dim s As Excel.Worksheet
    s.Cells(1861, lnColValor1).Value = lnEdadDias
    s.Cells(1862, lnColValor2).Value = ldPesoKg
    ldZ_PE = s.Cells(1869, lnColValor1).Value
    
    Set s = W.Sheets("T-E")
    s.Cells(1861, lnColValor1).Value = lnEdadDias
    s.Cells(1862, lnColValor2).Value = lcTallaCM
    ldZ_TE = s.Cells(1869, lnColValor1).Value
    
    Set s = W.Sheets("P-T")
    s.Cells(655, lnColValor1).Value = lcTallaCM
    s.Cells(656, lnColValor2).Value = ldPesoKg
    ldZ_PT = s.Cells(663, lnColValor1).Value

    Set orsTemp = mo_ReglasHIS.PadNominalSeleccionarDxNutricionalPorRangoZ("PE", ldZ_PE)
    If orsTemp.RecordCount > 0 Then
        orsTemp.MoveFirst
        ldiddxnutricionalPE = orsTemp.Fields!IdDiagnostico
    End If

    Set orsTemp = mo_ReglasHIS.PadNominalSeleccionarDxNutricionalPorRangoZ("PT", ldZ_PT)
    If orsTemp.RecordCount > 0 Then
        orsTemp.MoveFirst
        ldiddxnutricionalPT = orsTemp.Fields!IdDiagnostico
    End If

    
    Set orsTemp = mo_ReglasHIS.PadNominalSeleccionarDxNutricionalPorRangoZ("TE", ldZ_TE)
    If orsTemp.RecordCount > 0 Then
        orsTemp.MoveFirst
        ldiddxnutricionalTE = orsTemp.Fields!IdDiagnostico
    End If
    
LimpiarVariables:
    orsTemp.Close
    W.Close False
    Set s = Nothing
    Set W = Nothing
    Set EXL = Nothing
    Set orsTemp = Nothing
    
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


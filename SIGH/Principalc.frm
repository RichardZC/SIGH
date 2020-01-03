VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PrincipalC2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Información para Clínica"
   ClientHeight    =   9090
   ClientLeft      =   1260
   ClientTop       =   840
   ClientWidth     =   14505
   Icon            =   "Principalc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SISGalenPlus.ucContanciasDeAtencion ucContanciasAtencion 
      Height          =   615
      Left            =   5160
      TabIndex        =   54
      Top             =   7605
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
   End
   Begin SISGalenPlus.ucHCelectronicaLista ucHCelectronicaLista1 
      Height          =   510
      Left            =   5085
      TabIndex        =   53
      Top             =   6960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   900
   End
   Begin SISGalenPlus.UcSISfuaLista UcSISfuaLista1 
      Height          =   540
      Left            =   5025
      TabIndex        =   52
      Top             =   6315
      Visible         =   0   'False
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   953
   End
   Begin SISGalenPlus.ucHISListaLotes ucHISListaLotes 
      Height          =   735
      Left            =   4980
      TabIndex        =   51
      Top             =   5475
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   1296
   End
   Begin SISGalenPlus.ucSIcitasLista ucSIlistasCitas1 
      Height          =   645
      Left            =   4995
      TabIndex        =   50
      Top             =   4740
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   1138
   End
   Begin SISGalenPlus.UcImagenesLista UcImagenesLista1 
      Height          =   525
      Left            =   4995
      TabIndex        =   49
      Top             =   4170
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   926
   End
   Begin SISGalenPlus.ucFacturacionLaboratorio ucFactOrdenesLaboratorio 
      Height          =   555
      Left            =   855
      TabIndex        =   46
      Top             =   7440
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   979
   End
   Begin SISGalenPlus.ucEstablecimientosNoMinsaL ucEstablecimientosNoMinsaLista1 
      Height          =   375
      Left            =   9465
      TabIndex        =   45
      Top             =   5805
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
   End
   Begin SISGalenPlus.ucEspecialidadesLista ucEspecialidadesLista1 
      Height          =   540
      Left            =   9300
      TabIndex        =   44
      Top             =   5205
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   953
   End
   Begin SISGalenPlus.ucDiagnosticosLista ucDiagnosticosLista1 
      Height          =   465
      Left            =   9255
      TabIndex        =   43
      Top             =   4710
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   820
   End
   Begin SISGalenPlus.ucServiciosLista ucServiciosLista1 
      Height          =   420
      Left            =   9300
      TabIndex        =   42
      Top             =   4230
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   741
   End
   Begin SISGalenPlus.ucConfiguraResLab ucConfiguraResLab2 
      Height          =   570
      Left            =   10230
      TabIndex        =   41
      Top             =   8715
      Visible         =   0   'False
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   1005
   End
   Begin SISGalenPlus.ucFactPaquetesLista ucFactPaquetesLista1 
      Height          =   465
      Left            =   10200
      TabIndex        =   40
      Top             =   8250
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
   End
   Begin SISGalenPlus.ucFuentesFinanLista ucFuentesFinanciamientoLista1 
      Height          =   495
      Left            =   11835
      TabIndex        =   39
      Top             =   7485
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
   End
   Begin SISGalenPlus.ucPartidasLista ucPartidasLista1 
      Height          =   720
      Left            =   11685
      TabIndex        =   38
      Top             =   7950
      Visible         =   0   'False
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   767
   End
   Begin SISGalenPlus.ucTiposFinanciamientoLista ucTiposFinanciamientoLista1 
      Height          =   495
      Left            =   11985
      TabIndex        =   37
      Top             =   7020
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
   End
   Begin SISGalenPlus.ucCatalogoServiciosLista ucCatalogoServiciosLista1 
      Height          =   585
      Left            =   11955
      TabIndex        =   36
      Top             =   6375
      Visible         =   0   'False
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1032
   End
   Begin SISGalenPlus.ucTiposTarifaLista ucTiposTarifaLista1 
      Height          =   615
      Left            =   11940
      TabIndex        =   35
      Top             =   5745
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   1508
   End
   Begin SISGalenPlus.ucPacienteExternos ucPacienteExternos1 
      Height          =   630
      Left            =   9165
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   1111
   End
   Begin SISGalenPlus.ucReembolsosLista ucReembolsosLista1 
      Height          =   510
      Left            =   8880
      TabIndex        =   33
      Top             =   2775
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   900
   End
   Begin SISGalenPlus.ucFacturacionOrdenesLista ucFacturacionGeneralLista 
      Height          =   420
      Left            =   8775
      TabIndex        =   32
      Top             =   2250
      Visible         =   0   'False
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   741
   End
   Begin SISGalenPlus.ucEstadoCuenta ucEstadoCuenta1 
      Height          =   495
      Left            =   8700
      TabIndex        =   31
      Top             =   1575
      Visible         =   0   'False
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   873
   End
   Begin SISGalenPlus.ucFarmHpreciosLista ucFarmHpreciosLista1 
      Height          =   330
      Left            =   4845
      TabIndex        =   30
      Top             =   3675
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   582
   End
   Begin SISGalenPlus.ucFarmDespachoDonaciones ucFarmDespachoDonaciones1 
      Height          =   315
      Left            =   4740
      TabIndex        =   29
      Top             =   3225
      Visible         =   0   'False
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   556
   End
   Begin SISGalenPlus.ucFarmDependExtLista ucFarmDependExtLista1 
      Height          =   285
      Left            =   4620
      TabIndex        =   28
      Top             =   2925
      Visible         =   0   'False
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   503
   End
   Begin SISGalenPlus.ucFarmIntervencionLista ucFarmIntervencionLista1 
      Height          =   315
      Left            =   4710
      TabIndex        =   27
      Top             =   2595
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
   End
   Begin SISGalenPlus.ucSolicitudHistoriasLista ucSolicitudHistoriasLista1 
      Height          =   495
      Left            =   11865
      TabIndex        =   26
      Top             =   4260
      Visible         =   0   'False
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   873
   End
   Begin SISGalenPlus.ucMovimientoHistoriasLista ucMovimientoHistoriasLista1 
      Height          =   480
      Left            =   11880
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   847
   End
   Begin SISGalenPlus.ucHistoriaClinicaLista ucHistoriaClinicaLista1 
      Height          =   435
      Left            =   11850
      TabIndex        =   24
      Top             =   3300
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   767
   End
   Begin SISGalenPlus.ucMedicosLista ucMedicosLista1 
      Height          =   645
      Left            =   11835
      TabIndex        =   23
      Top             =   2580
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1138
   End
   Begin SISGalenPlus.ucTurnosLista ucTurnosLista1 
      Height          =   570
      Left            =   11775
      TabIndex        =   22
      Top             =   2115
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1005
   End
   Begin SISGalenPlus.ucProgramacionLista ucProgramacionLista1 
      Height          =   780
      Left            =   11745
      TabIndex        =   21
      Top             =   1515
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1376
   End
   Begin SISGalenPlus.ucCamasLista ucCamasLista1 
      Height          =   795
      Left            =   870
      TabIndex        =   18
      Top             =   5490
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   1402
   End
   Begin SISGalenPlus.ucArchivadoresLista ucArchivadoresLista1 
      Height          =   615
      Left            =   840
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   1085
   End
   Begin SISGalenPlus.ucRecetasLista ucRecetasLista1 
      Height          =   615
      Left            =   825
      TabIndex        =   16
      Top             =   3585
      Visible         =   0   'False
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   1085
   End
   Begin SISGalenPlus.ucAdmisionLista ucAdmisionCE 
      Height          =   780
      Left            =   840
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1376
   End
   Begin SISGalenPlus.ucAtencionesTriaje ucAtencionesTriaje1 
      Height          =   765
      Left            =   795
      TabIndex        =   14
      Top             =   2025
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1349
   End
   Begin SISGalenPlus.ucCitasLista ucCitasLista1 
      Height          =   690
      Left            =   765
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   1217
   End
   Begin SISGalenPlus.ucCajaLista ucCajaLista1 
      Height          =   480
      Left            =   8265
      TabIndex        =   12
      Top             =   945
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   847
   End
   Begin SISGalenPlus.ucCajaNotaCredito ucCajaNotaCredito1 
      Height          =   465
      Left            =   8265
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   820
   End
   Begin SISGalenPlus.ucGestionCaja ucGestionCaja1 
      Height          =   540
      Left            =   8295
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   953
   End
   Begin SISGalenPlus.ucFarmNiLista ucFarmNiLista1 
      Height          =   435
      Left            =   4695
      TabIndex        =   9
      Top             =   2085
      Visible         =   0   'False
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   767
   End
   Begin SISGalenPlus.ucFarmNsLista ucFarmNsLista1 
      Height          =   525
      Left            =   4530
      TabIndex        =   8
      Top             =   1620
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   926
   End
   Begin SISGalenPlus.ucFarmInventarioLista ucFarmInventarioLista1 
      Height          =   450
      Left            =   4575
      TabIndex        =   7
      Top             =   1185
      Visible         =   0   'False
      Width           =   3720
      _ExtentX        =   4895
      _ExtentY        =   794
   End
   Begin SISGalenPlus.ucFarmAlmacenes ucFarmAlmacenes1 
      Height          =   480
      Left            =   4785
      TabIndex        =   6
      Top             =   645
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   847
   End
   Begin SISGalenPlus.ucCatalogoBienesInsumosL ucCatalogoBienesInsumosLista1 
      Height          =   885
      Left            =   11925
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6271
      _ExtentY        =   1561
   End
   Begin SISGalenPlus.ucRolesLista ucRolesLista1 
      Height          =   375
      Left            =   10980
      TabIndex        =   4
      Top             =   -30
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5953
      _ExtentY        =   661
   End
   Begin SISGalenPlus.ucEmpleadosLista ucEmpleadosLista1 
      Height          =   735
      Left            =   10980
      TabIndex        =   3
      Top             =   375
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   1296
   End
   Begin SISGalenPlus.ucFarmVentasLista ucFarmVentasLista1 
      Height          =   1185
      Left            =   4755
      TabIndex        =   2
      Top             =   -45
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   2090
   End
   Begin SISGalenPlus.ucPacientesLista ucPacientesLista1 
      Height          =   930
      Left            =   795
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   4895
      _ExtentY        =   1640
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -60
      Top             =   5325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":0152
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":2904
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":2A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":3000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":3452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":35AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":39FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principalc.frx":3B58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   3  'Align Left
      Height          =   9090
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   16034
      ButtonWidth     =   3757
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            ImageIndex      =   2
            Object.Width           =   1500
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Key             =   "Modificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            Key             =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "Consultar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reembolso x Cuenta"
            Key             =   "reembolsoXcuenta"
            Description     =   "Reembolso x Cuenta"
            Object.ToolTipText     =   "Registro de Reembolso de una CUENTA"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alta Médica"
            Key             =   "AltaMedica"
            Description     =   "Alta Médica"
            Object.ToolTipText     =   "Alta Médica del Paciente actual"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Visita de Enfermera"
            Key             =   "VisitaEnfermera"
            Description     =   "Visita de Enfermera"
            Object.ToolTipText     =   "Visita de la Enfermera al Paciente actual"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SALIR"
            Key             =   "SALIR"
            Description     =   "salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin SISGalenPlus.ucAdmisionLista ucAdmisionConsEmerg 
      Height          =   780
      Left            =   795
      TabIndex        =   19
      Top             =   4815
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1376
   End
   Begin SISGalenPlus.ucAdmisionLista ucAdmisionHospitalizacion 
      Height          =   780
      Left            =   870
      TabIndex        =   20
      Top             =   6495
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1376
   End
   Begin SISGalenPlus.ucFacturacionLaboratorio ucFacturacionBS 
      Height          =   555
      Left            =   855
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   979
   End
   Begin SISGalenPlus.ucFacturacionLaboratorio ucFacturacionOrdenesPatologia 
      Height          =   555
      Left            =   885
      TabIndex        =   48
      Top             =   8610
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   979
   End
End
Attribute VB_Name = "PrincipalC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programa Principal del SIstema, muestra MENU
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim LbEsConsultorioAsignado As Boolean
Dim ms_ModuloSeleccionado As String
Dim mo_LastControl As Control
'Dim mo_LoginForm As Login
Dim mb_MantenerValoresCitas As Boolean
Dim mo_AdmisionHospDetalle As New AdmisionHospDetalle
Dim mo_AdmisionHospEgreso As New AdmisionHospEgreso
''Visitas
Dim mo_VisitasEnfermeras As New VisitasEnfermeras
''Referencias a reglas de negocios
Dim mo_FuenteFinanciamientoDetalle As New SIGHCatalogos.clFuenteFinancDetalle
Dim mo_PartidasDetalle As New SIGHCatalogos.clPartidaDetalle
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
'Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
'Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
'Dim mo_AdminServHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdmisionCEDetalle As New AdmisionCEatenciones
'Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
''Referencia a forms
Dim ml_IdUsuarioAuditoria As Long
'Dim mb_LoadingForm As Boolean
'Dim mrs_ListItems As New Recordset
'Dim ml_ToolbarHeightAdd As Long
Dim mb_abrioCaja As Boolean
Dim lc_NombrePc As String
Dim ms_ListBarItemClave As String

'Dim lbVisualizaListaMedicamentosVencidos As Boolean
'
''mgaray201503
Dim lbCajeroEmiteSoloServicios As Boolean
'Dim mb_UsuarioActualEsCajero As Boolean
Dim moDOCajaGestion As DOCajaGestion
'
Property Let ListBarItemClave(lValue As String)
    ms_ListBarItemClave = lValue
End Property

'
Property Let MuestraOpcionElegida(lValue As String)
    ml_IdUsuarioAuditoria = sighentidades.USUARIO
    ms_ModuloSeleccionado = lValue
    Toolbar.Buttons.Item(6).Visible = False
    Toolbar.Buttons.Item(7).Visible = False
    Toolbar.Buttons.Item(8).Visible = False
    Select Case ms_ModuloSeleccionado
    Case "PacienteCE"
         ucPacientesLista1.TipoFiltro = sghFiltrarTodos
         ConfigurarControl ucPacientesLista1
    Case "AdmisionCE"
        Me.ucCitasLista1.lbCargaTablasUnaVez = True
        ucCitasLista1.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucCitasLista1
        mb_MantenerValoresCitas = True
    Case "AtencionesTriaje"
        ConfigurarControl ucAtencionesTriaje1
    Case "AtencionesCE"
        mo_AdmisionCEDetalle.lbCargaTablasUnaVez = True
        ucAdmisionCE.TipoFiltro = sghFiltrarConsultaExterna
        ucAdmisionCE.Titulo = "Atenciones de consulta externa"
        ucAdmisionCE.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucAdmisionCE
    Case "RecetasCE"
        ucRecetasLista1.idTipoServicio = sghConsultaExterna
        ConfigurarControl ucRecetasLista1
    Case "idConsultorioAsignado"
        LbEsConsultorioAsignado = True
        Me.ucArchivadoresLista1.TipoBusqueda = sghHistoriaEnPrestamo
        ConfigurarControl ucArchivadoresLista1
         
    Case "AdmisionConsultorioEmerg"
        mo_AdmisionHospDetalle.lbCargaTablasUnaVez = True
        mo_AdmisionHospEgreso.lbCargaTablasUnaVez = True
        Toolbar.Buttons.Item(7).Visible = True
        Toolbar.Buttons.Item(8).Visible = True
        ucAdmisionConsEmerg.Titulo = "Admisión de emergencia"
        ucAdmisionConsEmerg.TipoFiltro = sghFiltrarEmergencia
        ucAdmisionConsEmerg.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucAdmisionConsEmerg
    Case "CamasEmergencia"
        ConfigurarControl ucCamasLista1
        ucCamasLista1.idTipoServicio = sghEmergenciaConsultorios
        ucCamasLista1.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
        ucCamasLista1.EsListaParaMantenimiento = True
        ucCamasLista1.RealizarBusqueda
    Case "RecetasE"
        ucRecetasLista1.idTipoServicio = sghEmergenciaConsultorios
        ConfigurarControl ucRecetasLista1
        
        
    Case "AdmisionHospitalizacion"
        mo_AdmisionHospDetalle.lbCargaTablasUnaVez = True
        mo_AdmisionHospEgreso.lbCargaTablasUnaVez = True
        Toolbar.Buttons.Item(7).Visible = True
        Toolbar.Buttons.Item(8).Visible = True
        ucAdmisionHospitalizacion.Titulo = "Admisión de hospitalización"
        ucAdmisionHospitalizacion.TipoFiltro = sghFiltrarHospitalizacion
        ucAdmisionHospitalizacion.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucAdmisionHospitalizacion
    Case "CamasHospitalizacion"
        ConfigurarControl ucCamasLista1
        ucCamasLista1.idTipoServicio = sghHospitalizacion
        ucCamasLista1.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
        ucCamasLista1.EsListaParaMantenimiento = True
        ucCamasLista1.RealizarBusqueda
        
    Case "Programacion"
        ucProgramacionLista1.idUsuario = ml_IdUsuarioAuditoria
        ucProgramacionLista1.lnIdTablaLISTBARITEMS = 401
        ucProgramacionLista1.lcNombrePc = lc_NombrePc
        ConfigurarControl ucProgramacionLista1
    Case "Turno"
        ConfigurarControl ucTurnosLista1
    Case "Medico"
        ConfigurarControl ucMedicosLista1
        
    Case "HistoriaClinica"
        ConfigurarControl ucHistoriaClinicaLista1
    Case "MovimientoHistoria"
        Me.ucMovimientoHistoriasLista1.TipoBusqueda = sghTodasHistorias
        ConfigurarControl ucMovimientoHistoriasLista1
        ucMovimientoHistoriasLista1.Inicializar
    Case "SolicitudHistorias"
        Me.ucSolicitudHistoriasLista1.TipoBusqueda = sghHistoriaEnPrestamo
        If ucSolicitudHistoriasLista1.IdArchivero = 0 Then
            ucSolicitudHistoriasLista1.IdArchivero = ml_IdUsuarioAuditoria
        End If
        ConfigurarControl ucSolicitudHistoriasLista1
    Case "Archivero"
        LbEsConsultorioAsignado = False
        Me.ucArchivadoresLista1.TipoBusqueda = sghHistoriaEnPrestamo
        ConfigurarControl ucArchivadoresLista1
        
       
        
    Case "EstadoCuenta"
        Me.ucEstadoCuenta1.idUsuario = ml_IdUsuarioAuditoria
        Me.ucEstadoCuenta1.lcNombrePc = lc_NombrePc
        Me.ucEstadoCuenta1.lnIdTablaLISTBARITEMS = 613
        Me.ucEstadoCuenta1.lnHWnd = Me.hwnd
        ConfigurarControl Me.ucEstadoCuenta1
    Case "FacturacionGeneral"
        ucFacturacionGeneralLista.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFacturacionGeneralLista
        ucFacturacionGeneralLista.PuntoCarga = 1 'General
        ucFacturacionGeneralLista.idTipoFinanciamiento = 1
        ucFacturacionGeneralLista.Titulo = "Consumo en el Servicio"
    Case "FactReembolsos"
        Toolbar.Buttons.Item(6).Visible = True
        ucReembolsosLista1.idUsuario = ml_IdUsuarioAuditoria
        ucReembolsosLista1.lnHWnd = Me.hwnd
        ConfigurarControl ucReembolsosLista1
    Case "PacExtConSeguro"
        ucPacienteExternos1.idUsuario = ml_IdUsuarioAuditoria
        ucPacienteExternos1.EsPacienteSinSeguro = False
        ConfigurarControl ucPacienteExternos1
        
         
         
    Case "CatalogoBienes"
        ucCatalogoBienesInsumosLista1.idUsuario = ml_IdUsuarioAuditoria
        ucCatalogoBienesInsumosLista1.lnHWnd = Me.hwnd
        ConfigurarControl Me.ucCatalogoBienesInsumosLista1
    Case "TipoTarifa"
        ConfigurarControl ucTiposTarifaLista1
    Case "FacturacionCatalogoServicios"
        ConfigurarControl Me.ucCatalogoServiciosLista1
    Case "TiposFinanciamiento"
        ConfigurarControl ucTiposFinanciamientoLista1
    Case "FacturacionPartidas"
        ConfigurarControl ucPartidasLista1
    Case "PqteServicio"
        ConfigurarControl Me.ucFactPaquetesLista1
    Case "ConfiguracionResLab"
        ConfigurarControl ucConfiguraResLab2
    Case "FuentesFinanciamiento"
        ConfigurarControl ucFuentesFinanciamientoLista1
        
    Case "Servicios"
        ConfigurarControl ucServiciosLista1
    Case "Diagnosticos"
        ConfigurarControl ucDiagnosticosLista1
    Case "Especialidades"
        ConfigurarControl ucEspecialidadesLista1
    Case "EstablecimientosNoMinsa"
        ConfigurarControl ucEstablecimientosNoMinsaLista1
        ucEstablecimientosNoMinsaLista1.ConfigurarEstablecimientos
        
        
    Case "OrdenesLaboratorio"
        ucFactOrdenesLaboratorio.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFactOrdenesLaboratorio
        ucFactOrdenesLaboratorio.HabilitarPuntoCarga = False
        ucFactOrdenesLaboratorio.PuntoCarga = 2 'Patología Clínica
        ucFactOrdenesLaboratorio.idTipoFinanciamiento = 1
        ucFactOrdenesLaboratorio.Titulo = "Órdenes para Análisis de Laboratorio (Patología Clínica)"
        ucFactOrdenesLaboratorio.AreaTrabajo = 69
        ucFactOrdenesLaboratorio.lcNombrePc = lc_NombrePc
        ucFactOrdenesLaboratorio.lnIdTablaLISTBARITEMS = 1312
    Case "OrdenesPatologia"
        ucFactOrdenesLaboratorio.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFacturacionOrdenesPatologia
        ucFacturacionOrdenesPatologia.HabilitarPuntoCarga = False
        ucFacturacionOrdenesPatologia.PuntoCarga = 3 'Anatomía Patológica
        ucFacturacionOrdenesPatologia.idTipoFinanciamiento = 1
        ucFacturacionOrdenesPatologia.Titulo = "Órdenes para Análisis de Laboratorio (Anatomía Patológica)"
        ucFacturacionOrdenesPatologia.AreaTrabajo = 70
        ucFacturacionOrdenesPatologia.lcNombrePc = lc_NombrePc
        ucFacturacionOrdenesPatologia.lnIdTablaLISTBARITEMS = 1321
     Case "BS"
        ucFacturacionBS.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFacturacionBS
        ucFacturacionBS.HabilitarPuntoCarga = False
        ucFacturacionBS.PuntoCarga = 11 'Banco de Sangre
        ucFacturacionBS.idTipoFinanciamiento = 1
        ucFacturacionBS.Titulo = "Órdenes del Banco de Sangre"
        ucFacturacionBS.AreaTrabajo = 69
        ucFacturacionOrdenesPatologia.lcNombrePc = lc_NombrePc
        ucFacturacionOrdenesPatologia.lnIdTablaLISTBARITEMS = 1322
        
        
    Case "ImagRayosX"
        ConfigurarControl UcImagenesLista1
        UcImagenesLista1.PuntoCarga = 21 'PuntoCarga.Rayos X
        UcImagenesLista1.Titulo = "Rayos X"
    Case "ImagEcografiaG"
        ConfigurarControl UcImagenesLista1
        UcImagenesLista1.PuntoCarga = 20 'PuntoCarga.EcografiaGeneral
        UcImagenesLista1.Titulo = "Ecografía General"
    Case "ImagTomografia"
        ConfigurarControl UcImagenesLista1
        UcImagenesLista1.PuntoCarga = 22 'PuntoCarga.tomografia
        UcImagenesLista1.Titulo = "Tomografía"
    Case "ImagEcografiaO"
        ConfigurarControl UcImagenesLista1
        UcImagenesLista1.PuntoCarga = 23 'PuntoCarga.EcografiaObstetrica
        UcImagenesLista1.Titulo = "Ecografía Obstétrica"
    Case "prgImagen"
        ConfigurarControl Me.ucHISListaLotes
    Case "FacturacionImagenologia"
        ucSIlistasCitas1.Area = sghImageneología
        ConfigurarControl ucSIlistasCitas1
        
        
         
    Case "Cajas"
        ConfigurarControl ucCajaLista1
    Case "NotaCredito"
        ucCajaNotaCredito1.lnHWnd = Me.hwnd
        ucCajaNotaCredito1.Inicializar
        ConfigurarControl ucCajaNotaCredito1
    Case "GestionCaja"
        ucGestionCaja1.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucGestionCaja1
        AperturaCaja
        If mb_abrioCaja Then
            ucGestionCaja1.NombreCajero = mo_AdminCaja.SeleccionaDatosCajero(sighentidades.USUARIO, sghApellidosYnombres)
            ucGestionCaja1.lnIdTablaLISTBARITEMS = 702
            ucGestionCaja1.lcNombrePc = lc_NombrePc
        End If
    Case "FarmAlmacen"
        ConfigurarControl ucFarmAlmacenes1
    Case "Inventario"
        ConfigurarControl ucFarmInventarioLista1
    Case "Ventas"
         ConfigurarControl ucFarmVentasLista1
    Case "NI", "NIF", "FARMADOP"                                                                         'debb2014
        ucFarmNiLista1.NIsoloParaFarmacia = IIf(ms_ModuloSeleccionado = "NI", False, True)    'debb2014
        ucFarmNiLista1.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFarmNiLista1
        If ms_ModuloSeleccionado = "FARMADOP" Then
           ucFarmNiLista1.Titulo = "ARMADO DE PAQUETES"
        End If
    Case "NS", "NSF"                                                                         'debb2014
        ucFarmNsLista1.NSsoloParaFarmacia = IIf(ms_ModuloSeleccionado = "NS", False, True)   'debb2014
        ucFarmNsLista1.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFarmNsLista1
    Case "IntervencionS"
        ucFarmIntervencionLista1.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFarmIntervencionLista1
    Case "DependenciaExt"
        ConfigurarControl ucFarmDependExtLista1
    Case "DespachoDonaciones"
        ucFarmDespachoDonaciones1.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucFarmDespachoDonaciones1
     Case "FarmPrecios"                              'debb2014b
        ConfigurarControl ucFarmHpreciosLista1
        
         
    Case "Roles"
        ConfigurarControl ucRolesLista1
        Set ucRolesLista1.DataSource = mo_AdminSeguridad.RolesSeleccionarTodos()
    Case "Empleado"
         ConfigurarControl ucEmpleadosLista1
         
    Case "Constancias"
        ucContanciasAtencion.idUsuario = ml_IdUsuarioAuditoria
        ConfigurarControl ucContanciasAtencion
        ucContanciasAtencion.Titulo = "Constancias de Atención y Hospitalización"
    
    Case "Fua"
        ConfigurarControl Me.UcSISfuaLista1
         
    Case "HcElectronica"
         ConfigurarControl ucHCelectronicaLista1
         
    End Select

End Property

Sub AperturaCaja()
Dim oApertura As New AperturaDecaja
Dim oDOEmpleado As dOEmpleado
Dim sNombreCajero As String
Dim oRsPermisos As New Recordset
Dim lbUsuarioRealizaApertura As Boolean
        '
        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
        If oRsPermisos.RecordCount > 0 Then
           Do While Not oRsPermisos.EOF
              Select Case oRsPermisos.Fields!IdPermiso
              Case 201    'Caja - Realizar Apertura
                   lbUsuarioRealizaApertura = True
              End Select
              oRsPermisos.MoveNext
           Loop
        End If
        Set oRsPermisos = Nothing
        '
        If lbUsuarioRealizaApertura = True Then
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(ml_IdUsuarioAuditoria)
            sNombreCajero = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            oApertura.NombreCajero = sNombreCajero
            oApertura.idUsuario = ml_IdUsuarioAuditoria
            oApertura.lcNombrePc = lc_NombrePc
            oApertura.Show 1
            If oApertura.AperturoCajaOK = True Then
                
                'debb-15/03/2016 (inicio)
                If oApertura.IdTurno = 0 Then
                   MsgBox "Tiene problemas con el TURNO", vbInformation, ""
                   Exit Sub
                End If
                'debb-15/03/2016 (fin)
                
                mb_abrioCaja = Me.ucGestionCaja1.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, oApertura.IdCaja, oApertura.IdTurno, oApertura.EmiteSoloServicio)
                
                '/****************************INO***************************************/
               ' mb_abrioCaja = Me.ucGestionDevolucion2.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, oApertura.IdCaja, oApertura.IdTurno, oApertura.EmiteSoloServicio)
                '/****************************INO***************************************/
                
                'mgaray201503
                Set moDOCajaGestion = New DOCajaGestion
                moDOCajaGestion.IdCaja = oApertura.IdCaja
                moDOCajaGestion.IdCajero = oApertura.IdTurno
                lbCajeroEmiteSoloServicios = oApertura.EmiteSoloServicio
            End If
            Unload oApertura
            
        Else
            MsgBox "El Usuario no tiene permiso para realizar APERTURA DE CAJA", vbInformation, Me.Caption
        End If

End Sub
'
'Property Get bAbrioCaja() As Boolean
'   bAbrioCaja = mb_abrioCaja
'End Property
'
'Property Get oDOCajaGestion() As DOCajaGestion
'   Set oDOCajaGestion = moDOCajaGestion
'End Property
'
'Property Get bCajeroEmiteSoloServicios() As Boolean
'   bCajeroEmiteSoloServicios = lbCajeroEmiteSoloServicios
'End Property
'
'Property Get UsuarioActualEsCajero() As Boolean
'   UsuarioActualEsCajero = mb_UsuarioActualEsCajero
'End Property
'
''Franco Temporal
'Property Get Turno() As Integer
'    Dim Hora As Integer
'    Hora = Val(Format(Now, "HH"))
'    If Hora >= 6 And Hora <= 13 Then
'        Turno = 1
'    ElseIf Hora >= 14 And Hora <= 21 Then
'        Turno = 2
'    ElseIf Hora >= 22 Or (Hora >= 0 And Hora <= 5) Then
'        Turno = 3
'    End If
'End Property
'
'Property Set LoginForm(oValue As Login)
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Set mo_LoginForm = oValue
'    ml_IdUsuarioAuditoria = oValue.IdUsuarioAutenticado
'    status.Panels(2).Text = "Usuario: " & oValue.NombreUsuarioAutenticado
'    status.Panels(3).Text = "Servidor: " & lcBuscaParametro.RetornaNombreDeServidor
'    status.Panels(4).Text = "PC: " & lc_NombrePc
'    status.Panels(5).Text = lcBuscaParametro.SeleccionaFilaParametro(205)
'    status.Panels(6).Text = WxLcVersionSisGalenPlus
'    status.Panels(7).Text = lcBuscaParametro.SeleccionaFilaParametro(314) & " " & lcBuscaParametro.RetornaVersionServidorSQLserver
'    wxParametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
'    Set lcBuscaParametro = Nothing
'End Property
'
'Private Sub CentrarImagen()
'  Dim lcBuscaParametro As New SIGHDatos.Parametros
'  If lcBuscaParametro.SeleccionaFilaParametro(282) = "S" Then
'     pctLogo.Picture = LoadPicture(App.Path & "\Imagenes\principalcs.jpg")
'  Else
'     pctLogo.Picture = LoadPicture(App.Path & "\Imagenes\principal.jpg")
'  End If
'  'Centrar imagen
'  Dim to_x As Single
'  Dim to_y As Single
'  If pctLogo.Picture = 0 Then Exit Sub
'  Cls
'  to_x = (ScaleWidth - pctLogo.ScaleWidth) / 2
'  'to_y = (ScaleHeight - pctLogo.ScaleHeight) / 2
'  to_y = 0
'
'  Me.PaintPicture pctLogo.Picture, to_x, to_y ', , , , , , &H373842
'  Set lcBuscaParametro = Nothing
'End Sub
'
'Sub MuestraReportes(lnIdUsuarioSistema As Long)
'    '********Chequea Si tiene acceso a las Opciones del Menu Reporte - daniel barrantes (inicio)
'    Dim oRsTmp As New Recordset
'    Dim lcSql As String
'    Set oRsTmp = mo_AdminSeguridad.RetornaOpcionesReporteQueNoTieneAcceso(lnIdUsuarioSistema)
'    If oRsTmp.RecordCount > 0 Then
'       oRsTmp.MoveFirst
'       Do While Not oRsTmp.EOF
'          lcSql = oRsTmp.Fields!id_menuReporte
'          'Me.Toolbar.Tools.Item(lcSql).Visible = False
'          oRsTmp.MoveNext
'       Loop
'    End If
'    oRsTmp.Close
'    Set oRsTmp = Nothing
'
'End Sub
'
'Private Sub Form_Activate()
'    If sighentidades.Parametro282valorInt = "1" Then
'        MuestraReportes 0
'        CargaSetup_X_PC
'        Exit Sub
'    End If
'
'    If mb_LoadingForm Then
'       ' If mo_LoginForm.Autenticado Then
'            Dim Grupo As SSGroup
'            Dim ListItem As SSListItem
'            Dim rsGrupos As Recordset
'            Dim rsItems As Recordset
'            Dim lbSigue As Boolean
'            'mgaray201504
'            Dim bConAccesoGestionCaja As Boolean
'            'debb-25/08/2015 (inicio)
'            Dim lbElUsuarioTieneRolAdministrador As Boolean, lbMuestraReportePacientesSISconMas180diasEstancia As Boolean
'            lbElUsuarioTieneRolAdministrador = mo_AdminSeguridad.TieneRolAdministrador(ml_IdUsuarioAuditoria)
'            'debb-25/08/2015 (fin)
'            bConAccesoGestionCaja = False
'
'            lbVisualizaListaMedicamentosVencidos = mo_AdminSeguridad.EmpleadoVisualizaListaMedicamentosVencidos(ml_IdUsuarioAuditoria)
'            'MsgBox "paso LOGIN"
'            If sighentidades.Parametro282valorInt <> "1" Then
'                Set rsItems = mo_AdminSeguridad.RolesItemsSeleccionarItemsPorUsuarioYGrupoSql2000(ml_IdUsuarioAuditoria, 0)
'                Set rsGrupos = mo_AdminSeguridad.RolesItemsSeleccionarGruposPorUsuarioSql2000(ml_IdUsuarioAuditoria)
'                Do While Not rsGrupos.EOF
'                        Set Grupo = SecurityListbar.Groups.Add()
'                        Grupo.Key = rsGrupos!Clave
'                        'Grupo.Index = rsGrupos!Indice
'                        Grupo.Caption = rsGrupos!Texto
'                        '
'                        Set rsItems = mo_AdminSeguridad.RolesItemsSeleccionarItemsPorUsuarioYGrupoSql2000(ml_IdUsuarioAuditoria, rsGrupos!IdListGrupo)
'                        rsItems.Filter = "IdListGrupo=" & rsGrupos!IdListGrupo
'                        '
'                        Do While Not rsItems.EOF
'                            Set ListItem = Grupo.ListItems.Add()
'                            'ListItem.Index = rsItems!Indice
'                            ListItem.Key = rsItems!Clave
'                            ListItem.Text = rsItems!Texto
'                            ListItem.TagVariant = rsItems!IdListItem
'                            ListItem.IconLarge = Trim(rsItems!KeyIcon)
'                            'If rsItems!IdListItem = 1307 Then
'                            '   lbVisualizaListaMedicamentosVencidos = True
'                            'End If
'                            mrs_ListItems.AddNew
'                            mrs_ListItems!IdListItem = rsItems!IdListItem
'                            mrs_ListItems!Clave = rsItems!Clave
'                            'Admision Emergencia y Admision Hospitalizacion
'                            If rsItems!IdListItem = sghOpcionGalenHos.sghAdmisionEmergencia Or rsItems!IdListItem = sghOpcionGalenHos.sghAdmisionHospitalizacion Then
'                               mo_AdmisionHospDetalle.lbCargaTablasUnaVez = True
'                               mo_AdmisionHospEgreso.lbCargaTablasUnaVez = True
'                               'debb-25/08/2015 (inicio)
'                               If lbElUsuarioTieneRolAdministrador = False And rsItems!IdListItem = sghOpcionGalenHos.sghAdmisionHospitalizacion Then
'                                  lbMuestraReportePacientesSISconMas180diasEstancia = True
'                               End If
'                               'debb-25/08/2015 (fin)
'                            ElseIf rsItems!IdListItem = sghOpcionGalenHos.sghRegistroCitaCE Then
'                               'Admision - Consulta Externa
'                               'Me.ucCitasLista1.lbCargaTablasUnaVez = True
'                            ElseIf rsItems!IdListItem = sghOpcionGalenHos.sghRegistroAtencionCE Then
'                               'Registro de Atencion - Consulta Externa
'                               mo_AdmisionCEDetalle.lbCargaTablasUnaVez = True
'                            End If
'                            'mgaray201504
'                            If rsItems!IdListItem = sghOpcionGalenHos.sghGestionGaja Then
'                               bConAccesoGestionCaja = True
'                            End If
'
'                            '
'                            rsItems.MoveNext
'                        Loop
'                        rsItems.Close
'                        'Reportes
'                        rsGrupos.MoveNext
'                Loop
'            End If
'            mb_UsuarioActualEsCajero = UsuarioEsCajero(bConAccesoGestionCaja)
'            '***************daniel barrantes**************
'            '********Chequea Si tiene acceso a las Opciones del Menu Reporte - daniel barrantes (inicio)
''            MuestraReportes ml_IdUsuarioAuditoria
''            Dim oRsTmp As New Recordset
''            Dim lcSql As String
''            Set oRsTmp = mo_AdminSeguridad.RetornaOpcionesReporteQueNoTieneAcceso(ml_IdUsuarioAuditoria)
''            If oRsTmp.RecordCount > 0 Then
''               oRsTmp.MoveFirst
''               Do While Not oRsTmp.EOF
''                  lcSql = oRsTmp.Fields!id_MenuReporte
''                  Me.toolbar.Tools.Item(lcSql).Visible = False
''                  oRsTmp.MoveNext
''               Loop
''            End If
''            oRsTmp.Close
''            Set oRsTmp = Nothing
'            '********Chequea Si tiene acceso a las Opciones del Menu Reporte - daniel barrantes (fin)
'
'            '***************Franklin Cachay**************
'            '******** Reportes solo usados en algunos Hospitales - Franklin Cachay (inicio)
'            'Hospital Ayacucho
'            Dim lcBuscaParametro As New SIGHDatos.Parametros
'            If Not (Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3543 Or lcBuscaParametro.SeleccionaFilaParametro(8) = "0") Then
''                Me.toolbar.Tools.Item("ID_ResumenCentroCosto").Visible = False
''                Me.toolbar.Tools.Item("ID_DetalleporcadaCentroCosto").Visible = False
'            End If
'            '******** Reportes solo usados en algunos Hospitales - Franklin Cachay (fin)
'
'            rsGrupos.Close
''            SecurityListbar.Groups.Remove SecurityListbar.Groups.Item("Dummy")
'            'Eliminar las citas que quedaron bloqueadas por este usuario
'            mo_AdminAdmision.CitasBloqueadasEliminarPorUsuario ml_IdUsuarioAuditoria
'        End If
'        mb_LoadingForm = False
'    'End If
'    If lbVisualizaListaMedicamentosVencidos = True Then
'        lbVisualizaListaMedicamentosVencidos = False
'        Dim oRptProdXvencer As New SighFarmacia.RepProductoPorVencer
'        oRptProdXvencer.mostrarReporte = True
'        oRptProdXvencer.EjecutaFormulario
'        Set oRptProdXvencer = Nothing
'    End If
'    'debb-25/08/2015 (inicio)
'    If lbMuestraReportePacientesSISconMas180diasEstancia = True Then
'        lbMuestraReportePacientesSISconMas180diasEstancia = False
'        Dim oRptIngHosp As New SIGHProxies.clReporteIngrHosp
'        oRptIngHosp.IdTipoReporte = sighentidades.sghReporteIngresosHospitalario
'        oRptIngHosp.mostrarReporte = True
'        oRptIngHosp.EjecutaFormulario
'        Set oRptIngHosp = Nothing
'    End If
'    'debb-25/08/2015 (fin)
'    CargaSetup_X_PC
'
'End Sub
'
'Private Sub Form_Initialize()
'    If sighentidades.Parametro282valorInt = "1" Then
'       mb_LoadingForm = True
'       Exit Sub
'    End If
'
'
'    On Error Resume Next
'    Me.Picture = LoadPicture(App.Path + "\Imagenes\principal.jpg")
'
'    mb_LoadingForm = True
'
'    'GenerarRecordsetDeListItems
'
'End Sub
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    AdministrarKeyPreview KeyCode
'End Sub
'Sub AdministrarKeyPreview(KeyCode As Integer)
'On Error Resume Next
'
'    Select Case KeyCode
'    Case vbKeyEscape
'
'        'WCG 04/06/2006
'        Select Case ms_ModuloSeleccionado
'
'        'EFGL 14/06/2006
'        Case "GestionCaja", "FacturacionProcedimientos", "FacturacionPatologiaClinica", "FacturacionAnatomiaPatologica", "FacturacionImaginologia", "EstadoCuenta"
'        'fin EFGL 14/06/2006
'        Case Else
'            mo_LastControl.Visible = False
'        End Select
'
'    Case vbKeyF2
'    Case vbKeyF6
'
'        RealizarBusquedas
'    Case vbKeyF7
'        LimpiarFiltro
'    Case vbKeyF8
'    Case vbKeyF9
'    Case vbKeyF10
'    Case vbKeyF11
'    Case vbKeyF12
'
'    End Select
'
'End Sub
'Sub RealizarBusquedas()
'    Select Case ms_ModuloSeleccionado
''    'MODULO AMBULATORIO
''    Case "AdmisionCE"
''        'ucCitasLista1
'    Case "PacienteCE"
'        ucPacientesLista1.RealizarBusqueda
''    Case "AtencionesCE"
''        ucAdmisionCE.RealizarBusqueda False
''    Case "InterconsultasCE"
''    'MODULO DE CONSULTORIOS DE EMERGENCIA
''    Case "PacienteEmerg", "PacienteObservacionEmerg"
''        ucPacientesLista1.RealizarBusqueda
''    Case "AdmisionConsultorioEmerg"
''
''        ucAdmisionConsEmerg.RealizarBusqueda False
''    Case "AtencionesConsultorioEmerg"
''        ucAdmisionConsEmerg.RealizarBusqueda False
''    Case "InterconsultasConsEmerg"
''
''    'MODULO OBSERVACION EMERGENCIA
''    Case "AdmisionObservacionEmerg"
''        ucAdmisionObservacion.RealizarBusqueda False
''    Case "InterconsultasObsEmerg"
''
''    Case "CamasEmergencia"
''        ucCamasLista1.RealizarBusqueda
''
''    'MODULO DE HOSPITALIZACION
''    Case "PacienteHosp"
''        ucPacientesLista1.RealizarBusqueda
''    Case "AdmisionHospitalizacion"
''        ucAdmisionHospitalizacion.RealizarBusqueda False
''    Case "AtencionesHospitalizacion"
''        ucAdmisionHospitalizacion.RealizarBusqueda False
''    Case "CamasHospitalizacion"
''        ucCamasLista1.RealizarBusqueda
''    Case "InterconsultasHosp"
''
''
''    'MODULO PROGRAMACION MEDICA
''    Case "Programacion"
''
''    Case "Turno"
''        ucTurnosLista1.RealizarBusqueda
''    Case "Medico"
''        ucMedicosLista1.RealizarBusqueda
''    'MODULO ARCHIVO CLINICO
''    Case "HistoriaClinica"
''        ucHistoriaClinicaLista1.RealizarBusqueda
''    Case "MovimientoHistoria"
''        ucMovimientoHistoriasLista1.RealizarBusqueda
''    Case "SolicitudHistorias"
''        ucSolicitudHistoriasLista1.RealizarBusqueda
''    Case "Archivero"
''        ucArchivadoresLista1.RealizarBusqueda
''
''    'MODULO GENERAL
''    Case "Empleado"
''        ucEmpleadosLista1.RealizarBusqueda
''    Case "Servicios"
''    Case "Diagnosticos"
''        ucDiagnosticosLista1.RealizarBusqueda
''    Case "Procedimientos"
''        ucProcedimientosLista1.RealizarBusqueda
''    Case "TiposFinanciamiento"
''        ucTiposFinanciamientoLista1.RealizarBusqueda
''    Case "FuentesFinanciamiento"
''        ucFuentesFinanciamientoLista1.RealizarBusqueda
''    Case "EstablecimientosNoMinsa"
''        ucEstablecimientosNoMinsaLista1.RealizarBusqueda
''    Case "Especialidades"
''        ucEspecialidadesLista1.RealizarBusqueda
''
''    'MZD Ini 01/06/2005
''    'MODULO CAJA
''    Case "MovimientosCaja"
''
''    'MZD Fin 01/06/2005
''    'FIN GENERAL
''    'SEGURIDAD
''    Case "Roles"
''    'mgaray20141009
''    Case "AtencionesTriaje":
''        Me.ucAtencionesTriaje1.RealizarBusqueda
''    'mgaray201411f
''    'IMAGENOLOGIA
''    Case "ImagTipoModalidadSala":
''        Me.ucImagTipoModalidadSala1.RealizarBusqueda
''    Case "ImagSala":
''        Me.ucImagSala1.RealizarBusqueda
''    Case "ImagCatalgoServicioDuracion":
''        Me.ucImagCatalgoServicioDuracion1.RealizarBusqueda
''    Case "IntegracionSistema"
''        Me.ucInteoIntegracionSistema1.RealizarBusqueda
'    End Select
'
'End Sub
'Sub LimpiarFiltro()
'
''    Select Case ms_ModuloSeleccionado
''    'MODULO AMBULATORIO
''    Case "AdmisionCE"
''        'ucCitasLista1
''    Case "PacienteCE"
''        ucPacientesLista1.LimpiarFiltro
''    Case "AtencionesCE"
''        ucAdmisionCE.LimpiarFiltro False 'Actualizado 14102014
''    Case "InterconsultasCE"
''
''
''    'MODULO DE CONSULTORIOS DE EMERGENCIA
''    Case "PacienteEmerg", "PacienteObservacionEmerg"
''        ucPacientesLista1.LimpiarFiltro
''    Case "AdmisionConsultorioEmerg"
''        ucAdmisionConsEmerg.LimpiarFiltro False 'Actualizado 14102014
''    Case "AtencionesConsultorioEmerg"
''        ucAdmisionConsEmerg.LimpiarFiltro False 'Actualizado 14102014
''    Case "InterconsultasConsEmerg"
''
''
''    'MODULO OBSERVACION EMERGENCIA
''    Case "AdmisionObservacionEmerg"
''        ucAdmisionObservacion.LimpiarFiltro False 'Actualizado 14102014
''    Case "InterconsultasObsEmerg"
''
''    Case "CamasEmergencia"
''        'ucCamasLista1.LimpiarFiltro
''
''    'MODULO DE HOSPITALIZACION
''    Case "PacienteHosp"
''        ucPacientesLista1.LimpiarFiltro
''    Case "AdmisionHospitalizacion"
''        ucAdmisionHospitalizacion.LimpiarFiltro False 'Actualizado 14102014
''    Case "AtencionesHospitalizacion"
''        ucAdmisionHospitalizacion.LimpiarFiltro False 'Actualizado 14102014
''    Case "CamasHospitalizacion"
''        'ucCamasLista1.LimpiarFiltro
''    Case "InterconsultasHosp"
''
''
''    'MODULO PROGRAMACION MEDICA
''    Case "Programacion"
''
''    Case "Turno"
''        ucTurnosLista1.LimpiarFiltro
''    Case "Medico"
''        ucMedicosLista1.LimpiarFiltro
''    'MODULO ARCHIVO CLINICO
''    Case "HistoriaClinica"
''        ucHistoriaClinicaLista1.LimpiarFiltro
''    Case "MovimientoHistoria"
''        ucMovimientoHistoriasLista1.LimpiarFiltro
''    Case "SolicitudHistorias"
''        ucSolicitudHistoriasLista1.LimpiarFiltro
''    Case "Archivero"
''        ucArchivadoresLista1.LimpiarFiltro
''
''    'MODULO GENERAL
''    Case "Empleado"
''        ucEmpleadosLista1.LimpiarFiltro
''    Case "Servicios"
''    Case "Diagnosticos"
''        ucDiagnosticosLista1.LimpiarFiltro
''    Case "Procedimientos"
''        ucProcedimientosLista1.LimpiarFiltro
''    Case "TiposFinanciamiento"
''        'ucTiposFinanciamientoLista1.LimpiarFiltro
''    Case "FuentesFinanciamiento"
''        'ucFuentesFinanciamientoLista1.LimpiarFiltro
''    Case "EstablecimientosNoMinsa"
''        ucEstablecimientosNoMinsaLista1.LimpiarFiltro
''    Case "Especialidades"
''        ucEspecialidadesLista1.LimpiarFiltro
''
''    'MZD Ini 01/06/2005
''    'MODULO CAJA
''    Case "MovimientosCaja"
''
''    'MZD Fin 01/06/2005
''    'FIN GENERAL
''    'SEGURIDAD
''    Case "Roles"
''    'mgaray20141009
''    Case "AtencionesTriaje":
''        Me.ucAtencionesTriaje1.LimpiarFiltro
''    'mgaray201411f
''    'IMAGENOLOGIA
''    Case "ImagTipoModalidadSala":
''        Me.ucImagTipoModalidadSala1.LimpiarFiltro
''    Case "ImagSala":
''        Me.ucImagSala1.LimpiarFiltro
''    Case "ImagCatalgoServicioDuracion":
''        Me.ucImagCatalgoServicioDuracion1.LimpiarFiltro
''    Case "IntegracionSistema"
''        Me.ucInteoIntegracionSistema1.LimpiarFiltro
''    End Select
'
'End Sub
Private Sub Form_Load()
'    SkinConfigura
'    If sighentidades.Parametro282valorInt = "1" Then
'        SkinConfigura
''        ml_ToolbarHeightAdd = 0
        mb_MantenerValoresCitas = False
        lc_NombrePc = sighentidades.RetornaNombrePC
        ml_IdUsuarioAuditoria = sighentidades.USUARIO
''        EliminaArchivosOpenOffice
''
''        Exit Sub
'    End If

   ' ml_ToolbarHeightAdd = 0
   ' mb_MantenerValoresCitas = False
    'lc_NombrePc = sighentidades.RetornaNombrePC
    OcultaBotonXdelFormulario Me.hwnd
   ' EliminaArchivosOpenOffice

End Sub
'Sub SkinConfigura()
'  On Error GoTo ErrSkin
'  Skin1.LoadSkin App.Path & "\" & WxSkin
'  Skin1.ApplySkin Me.hwnd
'ErrSkin:
'End Sub
'Private Sub EliminaArchivosOpenOffice()
'   Dim Archivo As String, viejo As String
'   Dim flag As Boolean
'   Dim c As Integer
'   On Error GoTo ElimArOP
'    flag = True
'    viejo = "xxx"
'    While (flag = True)
'        Archivo = Dir(App.Path & "\plantillas\*.ods")
'        If Archivo = "" Or Archivo = viejo Then
'            flag = False
'        Else
'            If InStr("1234567890", Left(Archivo, 1)) > 0 Then
'                Kill App.Path & "\plantillas\" & Archivo
'            Else
'                viejo = Archivo
'            End If
'        End If
'    Wend
'ElimArOP:
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If sighentidades.Parametro282valorInt = "1" Then
'       GalenhosKillExcelApplication
'       Me.Visible = False
'    End If
'
'
'    GalenhosKillExcelApplication
'    End
'End Sub
'
Private Sub Form_Resize()
    On Error Resume Next
    mo_LastControl.Top = Me.Top
    mo_LastControl.Left = Me.Left + Toolbar.Width + 50
    mo_LastControl.Width = Me.Width - 800
    mo_LastControl.Height = Me.Height
End Sub
'Sub ConfigurarPermisosDelItemSeleccionado(lIdUsuario As Long, lIdListItem As Long, sKey As String)
'
''    Me.Toolbar.Tools.Item("ID_Agregar").Enabled = True
''    Me.Toolbar.Tools.Item("ID_Modificar").Enabled = True
''    Me.Toolbar.Tools.Item("ID_Consultar").Enabled = True
''    Me.Toolbar.Tools.Item("ID_Eliminar").Enabled = True
''
''    Dim rsPermisos As Recordset
''    Set rsPermisos = mo_AdminSeguridad.RolesItemsSeleccionarPermisosPorEmpleadoYListItem(lIdUsuario, lIdListItem)
''    If Not (rsPermisos.EOF And rsPermisos.BOF) Then
''        If Not IsNull(rsPermisos!Agregar) Then
''           Me.Toolbar.Tools.Item("ID_Agregar").Enabled = (rsPermisos!Agregar > 0)
''        End If
''        If Not IsNull(rsPermisos!Modificar) Then
''           Me.Toolbar.Tools.Item("ID_Modificar").Enabled = (rsPermisos!Modificar > 0)
''        End If
''        If Not IsNull(rsPermisos!Consultar) Then
''           Me.Toolbar.Tools.Item("ID_Consultar").Enabled = (rsPermisos!Consultar > 0)
''        End If
''        If Not IsNull(rsPermisos!Eliminar) Then
''           Me.Toolbar.Tools.Item("ID_Eliminar").Enabled = (rsPermisos!Eliminar > 0)
''        End If
''    End If
''    rsPermisos.Close
''
''    'Manejo de excepciones
''    Select Case sKey
''    Case "AdmisionCE"
''        Me.ucCitasLista1.MenuAgregarEnabled = Me.Toolbar.Tools.Item("ID_Agregar").Enabled
''        Me.ucCitasLista1.MenuEliminarEnabled = Me.Toolbar.Tools.Item("ID_Eliminar").Enabled
''        Me.ucCitasLista1.MenuModificarEnabled = Me.Toolbar.Tools.Item("ID_Modificar").Enabled
''        Me.ucCitasLista1.MenuConsultarEnabled = Me.Toolbar.Tools.Item("ID_Consultar").Enabled
''    Case "AtencionesCE"
''        Me.Toolbar.Tools.Item("ID_Agregar").Enabled = False
''    Case "Programacion"
''        Me.ucProgramacionLista1.MenuAgregarEnabled = Me.Toolbar.Tools.Item("ID_Agregar").Enabled
''        Me.ucProgramacionLista1.MenuEliminarEnabled = Me.Toolbar.Tools.Item("ID_Eliminar").Enabled
''        Me.ucProgramacionLista1.MenuModificarEnabled = Me.Toolbar.Tools.Item("ID_Modificar").Enabled
''        Me.ucProgramacionLista1.MenuConsultarEnabled = Me.Toolbar.Tools.Item("ID_Consultar").Enabled
''    Case "AdmisionEmergencia"
'''        Me.toolbar.Tools.Item("ID_EmergenciaAObservacion").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
'''        Me.toolbar.Tools.Item("ID_EmergenciaAHospitalizacion").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
'''        Me.toolbar.Tools.Item("ID_EmergenciaAltaPaciente").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
'''        Me.toolbar.Tools.Item("ID_EmergenciaTransferencias").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
''        Me.Toolbar.Tools.Item("ID_EmergenciaAObservacion").Visible = False
''        Me.Toolbar.Tools.Item("ID_EmergenciaAHospitalizacion").Visible = False
''        Me.Toolbar.Tools.Item("ID_EmergenciaAltaPaciente").Visible = False
''        Me.Toolbar.Tools.Item("ID_EmergenciaTransferencias").Visible = False
''    Case "AdmisionHospitalizacion"
'''        Me.toolbar.Tools.Item("ID_HospitalizacionAlojamientoConjunto").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
'''        Me.toolbar.Tools.Item("ID_HospitalizacionAltaPaciente").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
'''        Me.toolbar.Tools.Item("ID_HospitalizacionTransferencias").Enabled = Me.toolbar.Tools.Item("ID_Modificar").Enabled
''        'Me.toolbar.Tools.Item("ID_HospitalizacionAlojamientoConjunto").Visible = False
''        'Me.toolbar.Tools.Item("ID_HospitalizacionAltaPaciente").Visible = False
''        'Me.toolbar.Tools.Item("ID_HospitalizacionTransferencias").Visible = False
''    Case "GestionCaja"
'''        Me.toolbar.Tools("ID_ParteDiario").Visible = False
'''        Me.toolbar.Tools("ID_ConsolidadoServ").Visible = False
'''        Me.toolbar.Tools("ID_ConsolFarmacia").Visible = False
''    End Select
'
'End Sub
'
'Private Sub Form_Terminate()
'    If sighentidades.Parametro282valorInt = "1" Then
'       Me.Visible = False
'    End If
'
'     mo_AdminSeguridad.LogueaUsuario 0, sighentidades.USUARIO, lc_NombrePc
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If sighentidades.Parametro282valorInt = "1" Then
'       Me.Visible = False
'    End If
'
'    mo_AdminSeguridad.LogueaUsuario 0, sighentidades.USUARIO, lc_NombrePc
'End Sub
'
'
'Sub EligeOpcion()
'    Select Case ms_ModuloSeleccionado
'    Case "PacienteCE"
'
'        ucPacientesLista1.Inicializar
'        ucPacientesLista1.TipoFiltro = sghFiltrarTodos
'
'        ConfigurarControl ucPacientesLista1
'
'
'    'MODULO CONSULTA EXTERNA
''    Case "AtencionesCE"
'''        ucAdmisionCE.TipoFiltro = sghFiltrarConsultaExterna
'''        ucAdmisionCE.Titulo = "Atenciones de consulta externa"
'''        ucAdmisionCE.idUsuario = ml_IdUsuarioAuditoria
'''        ConfigurarControl ucAdmisionCE
''    Case "AtencionesTriaje"
''        ConfigurarControl ucAtencionesTriaje1     'debb-jamo
''    Case "RecetasCE"
'''        ucRecetasLista1.IdTipoServicio = sghConsultaExterna
'''        ConfigurarControl ucRecetasLista1
''    Case "idConsultorioAsignado"
'''        LbEsConsultorioAsignado = True
'''        Me.ucArchivadoresLista1.TipoBusqueda = sghHistoriaEnPrestamo
'''        ConfigurarControl ucArchivadoresLista1
''
''    'HIS GALENOS  - JVG
''    Case "HisLoteCE"
''         'ucHISListaLotes.idUsuario = ml_IdUsuarioAuditoria
''         'ucHISListaLotes.Inicializar
''         ConfigurarControl ucHISListaLotes
''    Case "HisCE"
''         ucHISListaAtencion.idUsuario = ml_IdUsuarioAuditoria
''         'ucHISListaAtencion.Inicializar
''         ConfigurarControl ucHISListaAtencion
''    Case "HisPMMR"
''         ucHISListaProgramacion.idUsuario = ml_IdUsuarioAuditoria
''         'ucHISListaProgramacion.Inicializar
''         ConfigurarControl ucHISListaProgramacion
''    Case "HisREMR"
''        'ucHISEstablecimientos.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucHISEstablecimientos
''    Case "HisPN"
''        UcHISPadronNominal.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcHISPadronNominal
''    Case "HisCalidad"
''        UcHISCalidad.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcHISCalidad
''    'Seguimiento
''    Case "HcElectronica"
''         ConfigurarControl ucHCelectronicaLista1
''    Case "Sprogramas"
''    Case "Sadscripcion"
''
''    'MODULO DE EMERGENCIA
''    Case "PacienteEmerg", "PacienteObservacionEmerg"
''        ucPacientesLista1.inicializar
''        ucPacientesLista1.TipoFiltro = sghFiltrarTodos
''        ConfigurarControl ucPacientesLista1
''    Case "AdmisionConsultorioEmerg"
''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 1
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(1).Name = "Alta Médica"
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(2).Name = "."
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(3).Name = "."
''        toolbar.Tools("ID_HospitalizacionTransferencias").Visible = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Visible = False
'''        toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = False
'''        toolbar.Tools("ID_EmergenciaAObservacion").Enabled = False
'''        toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = False
'''        toolbar.Tools("ID_EmergenciaTransferencias").Enabled = False
''        ucAdmisionConsEmerg.Titulo = "Admisión de emergencia"
''        ucAdmisionConsEmerg.TipoFiltro = sghFiltrarEmergencia
''        ucAdmisionConsEmerg.idUsuario = ml_IdUsuarioAuditoria
'''        toolbar.Toolbars("Admisión Emergencia").Visible = True
'''        toolbar.Toolbars("Admisión Emergencia").DockedRow = 3
'''        toolbar.Toolbars("Admisión Emergencia").DockedColumn = 3
''        ConfigurarControl ucAdmisionConsEmerg
''    Case "CamasEmergencia"                                 '09/08/2011
''        ConfigurarControl ucCamasLista1
''        ucCamasLista1.IdTipoServicio = sghEmergenciaConsultorios
''        ucCamasLista1.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
''        ucCamasLista1.EsListaParaMantenimiento = True
''        ucCamasLista1.RealizarBusqueda
''    Case "RecetasE"
''        ucRecetasLista1.IdTipoServicio = sghEmergenciaConsultorios
''        ConfigurarControl ucRecetasLista1
''
''    'MODULO DE HOSPITALIZACION
''    Case "PacienteHosp"
''        ucPacientesLista1.inicializar
''        ucPacientesLista1.TipoFiltro = sghFiltrarTodos
''        ConfigurarControl ucPacientesLista1
''    Case "AdmisionHospitalizacion"
''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 1
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(1).Name = "Alta Médica"
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(2).Name = "."
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(3).Name = "."
''        toolbar.Tools("ID_HospitalizacionTransferencias").Visible = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Visible = False
''        'toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = False
''        'toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = False
''        'toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = False
''        ucAdmisionHospitalizacion.Titulo = "Admisión de hospitalización"
''        ucAdmisionHospitalizacion.TipoFiltro = sghFiltrarHospitalizacion
''        ucAdmisionHospitalizacion.idUsuario = ml_IdUsuarioAuditoria
'''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
'''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
'''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 3
''        ConfigurarControl ucAdmisionHospitalizacion
''    Case "AlojadosHospitalizacion"
''        toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = False
''        toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = False
''        ucAdmisionHospitalizacion.Titulo = "Admisión de Alojados"
''        ucAdmisionHospitalizacion.TipoFiltro = sghFiltrarHospitalizacion
''        ucAdmisionHospitalizacion.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucAdmisionHospitalizacion
''    Case "CamasHospitalizacion"
''        ConfigurarControl ucCamasLista1
''        ucCamasLista1.IdTipoServicio = sghHospitalizacion
''        ucCamasLista1.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
''        ucCamasLista1.EsListaParaMantenimiento = True
''        ucCamasLista1.RealizarBusqueda
''    Case "RecetasH"
''        ucRecetasLista1.IdTipoServicio = sghHospitalizacion
''        ConfigurarControl ucRecetasLista1
''
''    'MODULO PROGRAMACION MEDICA
''    Case "Programacion"
''        ucProgramacionLista1.idUsuario = ml_IdUsuarioAuditoria
''        ucProgramacionLista1.lnIdTablaLISTBARITEMS = 401
''        ucProgramacionLista1.lcNombrePc = lc_NombrePc
''        ConfigurarControl ucProgramacionLista1
''
''    Case "Turno"
''        ConfigurarControl ucTurnosLista1
''
''    Case "Medico"
''        ConfigurarControl ucMedicosLista1
''
''    'MODULO ARCHIVO CLINICO
''    Case "HistoriaClinica"
''        ConfigurarControl ucHistoriaClinicaLista1
''    Case "MovimientoHistoria"
''        Me.ucMovimientoHistoriasLista1.TipoBusqueda = sghTodasHistorias
''        ConfigurarControl ucMovimientoHistoriasLista1
''        ucMovimientoHistoriasLista1.inicializar
''    Case "SolicitudHistorias"
''        Me.ucSolicitudHistoriasLista1.TipoBusqueda = sghHistoriaEnPrestamo
''        If ucSolicitudHistoriasLista1.IdArchivero = 0 Then
''            ucSolicitudHistoriasLista1.IdArchivero = ml_IdUsuarioAuditoria
''        End If
''        ConfigurarControl ucSolicitudHistoriasLista1
''    Case "Archivero"
''        LbEsConsultorioAsignado = False
''        Me.ucArchivadoresLista1.TipoBusqueda = sghHistoriaEnPrestamo
''        ConfigurarControl ucArchivadoresLista1
''    Case "MovFormatosHC"
''        ConfigurarControl ucMovimientoFormatoHcLista1
''        ucMovimientoFormatoHcLista1.inicializar
''
''    'FACTURACION SERVICIOS
''    Case "FacturacionGeneral"
''        ucFacturacionGeneralLista.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFacturacionGeneralLista
''        ucFacturacionGeneralLista.PuntoCarga = 1 'General
''        ucFacturacionGeneralLista.idTipoFinanciamiento = 1
''        ucFacturacionGeneralLista.Titulo = "Consumo en el Servicio"
''    Case "FacturacionPatologiaClinica"
''        ucFactPatologiaClinica.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactPatologiaClinica
''        ucFactPatologiaClinica.HabilitarPuntoCarga = False
''        ucFactPatologiaClinica.PuntoCarga = 2 'Patología Clínica
''        ucFactPatologiaClinica.idTipoFinanciamiento = 1
''        ucFactPatologiaClinica.Titulo = "Facturacion Laboratorio"
''    Case "FacturacionAnatomiaPatologica"
''        ucFactAnatomiaPatologica.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactAnatomiaPatologica
''        ucFactAnatomiaPatologica.HabilitarPuntoCarga = False
''        ucFactAnatomiaPatologica.PuntoCarga = 3 'Anatomía Patológica
''        ucFactAnatomiaPatologica.idTipoFinanciamiento = 1
''        ucFactAnatomiaPatologica.Titulo = "Facturacion Anatomía Patológica"
''    Case "FacturacionImagenologia"
''        ucFactImagenologia.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactImagenologia
''        ucFactImagenologia.HabilitarPuntoCarga = False
''        ucFactImagenologia.PuntoCarga = 4 'Imagenología
''        ucFactImagenologia.idTipoFinanciamiento = 1
''        ucFactImagenologia.Titulo = "Facturacion Imagenología"
''    Case "FacturacionSalaOperaciones"
''        ucFactSalaOperaciones.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactSalaOperaciones
''        ucFactSalaOperaciones.HabilitarPuntoCarga = False
''        ucFactSalaOperaciones.PuntoCarga = 7 'Sala de Operaciones
''        ucFactSalaOperaciones.idTipoFinanciamiento = 1
''        ucFactSalaOperaciones.Titulo = "Facturacion Sala Operaciones"
''    Case "FactReembolsos"
''        '18/7/11
''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 1
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(1).Name = "Agregar Reembolso x Cuenta"
''        toolbar.Tools("ID_HospitalizacionTransferencias").Visible = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Visible = False
''       ' toolbar.Toolbars("Admisión Hospitalización").Tools.Item(2).Name = "."
''       ' toolbar.Toolbars("Admisión Hospitalización").Tools.Item(3).Name = "."
''        '18/7/11
''        ucReembolsosLista1.idUsuario = ml_IdUsuarioAuditoria
''        ucReembolsosLista1.lnHWnd = Me.hwnd
''        ConfigurarControl ucReembolsosLista1
''    Case "PacExtConSeguro"
''        ucPacienteExternos1.idUsuario = ml_IdUsuarioAuditoria
''        ucPacienteExternos1.EsPacienteSinSeguro = False
''        ConfigurarControl ucPacienteExternos1
''    Case "PacExtParticular"
''        ucPacienteExternos1.idUsuario = ml_IdUsuarioAuditoria
''        ucPacienteExternos1.EsPacienteSinSeguro = True
''        ConfigurarControl ucPacienteExternos1
''
''    'FACTURACION CONFIGURACION
''    Case "FacturacionCatalogoServicios"
''
''        ConfigurarControl Me.ucCatalogoServiciosLista1
''
''    Case "FacturacionCentroCostos"
''
''        ConfigurarControl Me.ucCentrosCostoLista1
''    Case "PqteServicio"
''        ConfigurarControl Me.ucFactPaquetesLista1
''    Case "EstadoCuenta"
''        Me.ucEstadoCuenta1.idUsuario = ml_IdUsuarioAuditoria
''        Me.ucEstadoCuenta1.lcNombrePc = lc_NombrePc
''        Me.ucEstadoCuenta1.lnIdTablaLISTBARITEMS = 613
''        Me.ucEstadoCuenta1.lnHWnd = Me.hwnd
''        'Me.ucEstadoCuenta1.Inicializar
''        ConfigurarControl Me.ucEstadoCuenta1
''
''    Case "Farmacia"
''
''        ucFactFarmacia.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactFarmacia
''
''        ucFactFarmacia.PuntoCarga = 5 'Farmacia
''        ucFactFarmacia.HabilitarPuntoCarga = False
''        ucFactFarmacia.idTipoFinanciamiento = 1
''        ucFactFarmacia.Titulo = "Facturacion Farmacia"
''
''    'modificación samuel 02/06
''    Case "ConfiguracionResLab"
''        ConfigurarControl ucConfiguraResLab2
''
''    'MODULO GENERAL
''    Case "Empleado"
''        ConfigurarControl ucEmpleadosLista1
''    Case "Servicios"
''        ConfigurarControl ucServiciosLista1
''    Case "Diagnosticos"
''        ConfigurarControl ucDiagnosticosLista1
''    Case "Procedimientos"
''        ConfigurarControl ucProcedimientosLista1
''    Case "TiposFinanciamiento"
''        ConfigurarControl ucTiposFinanciamientoLista1
''    Case "FuentesFinanciamiento"
''        ConfigurarControl ucFuentesFinanciamientoLista1
''    Case "FacturacionPartidas"
''        ConfigurarControl ucPartidasLista1
''    Case "EstablecimientosNoMinsa"
''        ConfigurarControl ucEstablecimientosNoMinsaLista1
''        ucEstablecimientosNoMinsaLista1.ConfigurarEstablecimientos
''    Case "DiagnosticosPDF"
''        Dim oShell As New sighentidades.Shell
''        If sighentidades.RutaAdobeReader <> "" Then
''            oShell.ejecutarComando sighentidades.RutaAdobeReader + " " + App.Path + "\archivos\" + "cie10.pdf"
''        Else
''            MsgBox "No tiene instalado el adobe reader", vbInformation, Me.Caption
''        End If
''    Case "Especialidades"
''        ConfigurarControl ucEspecialidadesLista1
''    Case "TipoTarifa"
''        ConfigurarControl ucTiposTarifaLista1
''    'MODULO DE CAJA
''    Case "Cajas"
''        ConfigurarControl ucCajaLista1
''    Case "Cajeros"
''        ConfigurarControl ucCajeroLista1
''    Case "AsignacionTerminales"
''    Case "GestionCaja"
''        toolbar.Toolbars("Edición").Visible = False
''        'toolbar.Toolbars("Gestión de Caja").Visible = True
''        If (mb_abrioCaja) Then
''            If mo_LastControl Is ucGestionCaja1 Then
''                mo_LastControl.Visible = True
''                toolbar.Toolbars("Gestión de Caja").Visible = True
''                Exit Sub
''            End If
''            mo_LastControl.Visible = False
''            ucGestionCaja1.NombreCajero = status.Panels(2).Text
''            ucGestionCaja1.Visible = True
''            Set mo_LastControl = ucGestionCaja1
''            toolbar.Toolbars("Gestión de Caja").Visible = True
''            Exit Sub
''        End If
''
''        ucGestionCaja1.idUsuario = ml_IdUsuarioAuditoria
''        ucGestionCaja1.NombreCajero = status.Panels(2).Text
''        ucGestionCaja1.lnIdTablaLISTBARITEMS = 702
''        ucGestionCaja1.lcNombrePc = lc_NombrePc
''        toolbar.Toolbars("Gestión de Caja").Visible = True
''
''        ConfigurarControl ucGestionCaja1
''
''    '/********************INO****************************/
''    Case "Devoluciones"
''        MsgBox "Se retiro el modulo del SISGalenPlus, posteriormente se estará agregando " & Chr(13) & " el modulo 'Nota de Crédito para las Devoluciones", vbInformation, Me.Caption
''
''    '/********************FCV MAYO2015****************************/
''    Case "NotaCredito"
''        ucCajaNotaCredito1.lnHWnd = Me.hwnd
''        ucCajaNotaCredito1.inicializar
''        ConfigurarControl ucCajaNotaCredito1
''    '/********************INO****************************/
''    'MODULO FARMACIA
''    Case "Inventario"
''        ConfigurarControl ucFarmInventarioLista1
''    Case "NI", "NIF", "FARMADOP"                                                                         'debb2014
''        ucFarmNiLista1.NIsoloParaFarmacia = IIf(ms_ModuloSeleccionado = "NI", False, True)    'debb2014
''        ucFarmNiLista1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmNiLista1
''        If ms_ModuloSeleccionado = "FARMADOP" Then
''           ucFarmNiLista1.Titulo = "ARMADO DE PAQUETES"
''        End If
''    Case "NS", "NSF"                                                                         'debb2014
''        ucFarmNsLista1.NSsoloParaFarmacia = IIf(ms_ModuloSeleccionado = "NS", False, True)   'debb2014
''        ucFarmNsLista1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmNsLista1
''    Case "IntervencionS"
''        ucFarmIntervencionLista1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmIntervencionLista1
''    Case "Ventas"
''        ucFarmVentasLista1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmVentasLista1
''    Case "DependenciaExt"
''        ConfigurarControl ucFarmDependExtLista1
''    Case "DespachoDonaciones"
''        ucFarmDespachoDonaciones1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmDespachoDonaciones1
''    Case "FarmAlmacen"
''        ConfigurarControl ucFarmAlmacenes1
''     Case "FarmPrecios"                              'debb2014b
''        'ConfigurarControl Me.ucFarmHistoricoPrecios1     'debb2014b
''        ConfigurarControl ucFarmHpreciosLista1
''
''    'CATALOGOS
''    Case "CatalogoBienes"
''        ucCatalogoBienesInsumosLista1.idUsuario = ml_IdUsuarioAuditoria
''        ucCatalogoBienesInsumosLista1.lnHWnd = Me.hwnd
''        ConfigurarControl Me.ucCatalogoBienesInsumosLista1
''    'SEGURIDAD
''    Case "Roles"
''        ConfigurarControl ucRolesLista1
''        Set ucRolesLista1.DataSource = mo_AdminSeguridad.RolesSeleccionarTodos()
''    'MODULO IMAGENEOLOGIA
''    Case "ImagRayosX"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 21 'PuntoCarga.Rayos X
''        UcImagenesLista1.Titulo = "Rayos X"
''    Case "ImagEcografiaG"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 20 'PuntoCarga.EcografiaGeneral
''        UcImagenesLista1.Titulo = "Ecografía General"
''    Case "ImagTomografia"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 22 'PuntoCarga.tomografia
''        UcImagenesLista1.Titulo = "Tomografía"
''    Case "ImagEcografiaO"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 23 'PuntoCarga.EcografiaObstetrica
''        UcImagenesLista1.Titulo = "Ecografía Obstétrica"
''    Case "ImagIngresos"
''        UcImagIngresos1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcImagIngresos1
''    Case "ImagSalidas"
''        UcImagSalidas1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcImagSalidas1
''    'mgaray201411f
''    Case "ImagTipoModalidadSala"
'''        ucImagTipoModalidadSala1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucImagTipoModalidadSala1
''    Case "ImagSala"
''        ConfigurarControl ucImagSala1
''    Case "ImagCatalgoServicioDuracion":
''        ConfigurarControl ucImagCatalgoServicioDuracion1
''    Case "IntegracionSistema"
''        ConfigurarControl ucInteoIntegracionSistema1
''    'Módulo LABORATORIO
''    Case "OrdenesLaboratorio"
''        ucFactOrdenesLaboratorio.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactOrdenesLaboratorio
''        ucFactOrdenesLaboratorio.HabilitarPuntoCarga = False
''        ucFactOrdenesLaboratorio.PuntoCarga = 2 'Patología Clínica
''        ucFactOrdenesLaboratorio.idTipoFinanciamiento = 1
''        ucFactOrdenesLaboratorio.Titulo = "Órdenes para Análisis de Laboratorio (Patología Clínica)"
''        ucFactOrdenesLaboratorio.AreaTrabajo = 69
''        ucFactOrdenesLaboratorio.lcNombrePc = lc_NombrePc
''        ucFactOrdenesLaboratorio.lnIdTablaLISTBARITEMS = 1312
''    Case "OrdenesPatologia"
''        ucFactOrdenesLaboratorio.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFacturacionOrdenesPatologia
''        ucFacturacionOrdenesPatologia.HabilitarPuntoCarga = False
''        ucFacturacionOrdenesPatologia.PuntoCarga = 3 'Anatomía Patológica
''        ucFacturacionOrdenesPatologia.idTipoFinanciamiento = 1
''        ucFacturacionOrdenesPatologia.Titulo = "Órdenes para Análisis de Laboratorio (Anatomía Patológica)"
''        ucFacturacionOrdenesPatologia.AreaTrabajo = 70
''        ucFacturacionOrdenesPatologia.lcNombrePc = lc_NombrePc
''        ucFacturacionOrdenesPatologia.lnIdTablaLISTBARITEMS = 1321
''     Case "BS"
''        ucFacturacionBS.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFacturacionBS
''        ucFacturacionBS.HabilitarPuntoCarga = False
''        ucFacturacionBS.PuntoCarga = 11 'Banco de Sangre
''        ucFacturacionBS.idTipoFinanciamiento = 1
''        ucFacturacionBS.Titulo = "Órdenes del Banco de Sangre"
''        ucFacturacionBS.AreaTrabajo = 69
''        ucFacturacionOrdenesPatologia.lcNombrePc = lc_NombrePc
''        ucFacturacionOrdenesPatologia.lnIdTablaLISTBARITEMS = 1322
''     Case "LabIngresos"
''        UcLabIngresos1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcLabIngresos1
''        UcLabIngresos1.idTipoFinanciamiento = 1
''        UcLabIngresos1.Titulo = "Ingreso de Insumos"
''    Case "LabEgresos"
''        UcLabSalidas1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcLabSalidas1
''        UcLabSalidas1.idTipoFinanciamiento = 1
''        UcLabSalidas1.Titulo = "Salida de Insumos"
''
''    'Estadística
''    Case "Constancias"
''        ucContanciasAtencion.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucContanciasAtencion
''        ucContanciasAtencion.Titulo = "Constancias de Atención y Hospitalización"
''    'SIS
''    Case "Fua"
''        ConfigurarControl Me.UcSISfuaLista1
'    End Select
'
'End Sub
'
'
'Private Sub SecurityListbar_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
''Dim oControl As Control
''
''    'Por defecto la barra de gestión de caja esta invisible
''    'y la barra de edición esta visible
''    Toolbar.Toolbars("Edición").Visible = True
''    Toolbar.Toolbars("Gestión de Caja").Visible = False
''    Toolbar.Toolbars("Admisión Emergencia").Visible = False
''    Toolbar.Toolbars("Admisión Hospitalización").Visible = False
''
''    mrs_ListItems.MoveFirst
''    mrs_ListItems.Find "Clave = '" & ItemClicked.Key & "'"
''    If Not (mrs_ListItems.EOF And mrs_ListItems.BOF) Then
''        ConfigurarPermisosDelItemSeleccionado ml_IdUsuarioAuditoria, mrs_ListItems!IdListItem, ItemClicked.Key
''    End If
''
''    'GUARDA LA CLAVE DEL MODULO SELECCIONADO
''    ms_ModuloSeleccionado = ItemClicked.Key
''    ml_ToolbarHeightAdd = 0
''
''    EligeOpcion
'
''    Select Case ms_ModuloSeleccionado
''    'MODULO CONSULTA EXTERNA
''    Case "AdmisionCE"
''        ucCitasLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucCitasLista1
''        mb_MantenerValoresCitas = True
''    Case "PacienteCE"
''        ucPacientesLista1.inicializar
''        ucPacientesLista1.TipoFiltro = sghFiltrarTodos
''
''        ConfigurarControl ucPacientesLista1
''    Case "AtencionesCE"
''        ucAdmisionCE.TipoFiltro = sghFiltrarConsultaExterna
''        ucAdmisionCE.Titulo = "Atenciones de consulta externa"
''        ucAdmisionCE.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucAdmisionCE
''    Case "AtencionesTriaje"
''        ConfigurarControl ucAtencionesTriaje1     'debb-jamo
''    Case "RecetasCE"
''        ucRecetasLista1.idTipoServicio = sghConsultaExterna
''        ConfigurarControl ucRecetasLista1
''    Case "idConsultorioAsignado"
''        LbEsConsultorioAsignado = True
''        Me.ucArchivadoresLista1.TipoBusqueda = sghHistoriaEnPrestamo
''        ConfigurarControl ucArchivadoresLista1
''
''    'HIS GALENOS  - JVG
''    Case "HisLoteCE"
''         ucHISListaLotes.IdUsuario = ml_IdUsuarioAuditoria
''         'ucHISListaLotes.Inicializar
''         ConfigurarControl ucHISListaLotes
''    Case "HisCE"
''         ucHISListaAtencion.IdUsuario = ml_IdUsuarioAuditoria
''         'ucHISListaAtencion.Inicializar
''         ConfigurarControl ucHISListaAtencion
''    Case "HisPMMR"
''         ucHISListaProgramacion.IdUsuario = ml_IdUsuarioAuditoria
''         'ucHISListaProgramacion.Inicializar
''         ConfigurarControl ucHISListaProgramacion
''    Case "HisREMR"
''        ucHISEstablecimientos.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucHISEstablecimientos
''    Case "HisPN"
''        UcHISPadronNominal.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcHISPadronNominal
''    Case "HisCalidad"
''        UcHISCalidad.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcHISCalidad
''    'Seguimiento
''    Case "HcElectronica"
''         ConfigurarControl ucHCelectronicaLista1
''    Case "Sprogramas"
''    Case "Sadscripcion"
''
''    'MODULO DE EMERGENCIA
''    Case "PacienteEmerg", "PacienteObservacionEmerg"
''        ucPacientesLista1.inicializar
''        ucPacientesLista1.TipoFiltro = sghFiltrarTodos
''        ConfigurarControl ucPacientesLista1
''    Case "AdmisionConsultorioEmerg"
''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 1
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(1).Name = "Alta Médica"
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(2).Name = "."
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(3).Name = "."
''        toolbar.Tools("ID_HospitalizacionTransferencias").Visible = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Visible = False
'''        toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = False
'''        toolbar.Tools("ID_EmergenciaAObservacion").Enabled = False
'''        toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = False
'''        toolbar.Tools("ID_EmergenciaTransferencias").Enabled = False
''        ucAdmisionConsEmerg.Titulo = "Admisión de emergencia"
''        ucAdmisionConsEmerg.TipoFiltro = sghFiltrarEmergencia
''        ucAdmisionConsEmerg.IdUsuario = ml_IdUsuarioAuditoria
'''        toolbar.Toolbars("Admisión Emergencia").Visible = True
'''        toolbar.Toolbars("Admisión Emergencia").DockedRow = 3
'''        toolbar.Toolbars("Admisión Emergencia").DockedColumn = 3
''        ConfigurarControl ucAdmisionConsEmerg
''    Case "CamasEmergencia"                                 '09/08/2011
''        ConfigurarControl ucCamasLista1
''        ucCamasLista1.idTipoServicio = sghEmergenciaConsultorios
''        ucCamasLista1.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
''        ucCamasLista1.EsListaParaMantenimiento = True
''        ucCamasLista1.RealizarBusqueda
''    Case "RecetasE"
''        ucRecetasLista1.idTipoServicio = sghEmergenciaConsultorios
''        ConfigurarControl ucRecetasLista1
''
''    'MODULO DE HOSPITALIZACION
''    Case "PacienteHosp"
''        ucPacientesLista1.inicializar
''        ucPacientesLista1.TipoFiltro = sghFiltrarTodos
''        ConfigurarControl ucPacientesLista1
''    Case "AdmisionHospitalizacion"
''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 1
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(1).Name = "Alta Médica"
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(2).Name = "."
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(3).Name = "."
''        toolbar.Tools("ID_HospitalizacionTransferencias").Visible = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Visible = False
''        'toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = False
''        'toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = False
''        'toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = False
''        ucAdmisionHospitalizacion.Titulo = "Admisión de hospitalización"
''        ucAdmisionHospitalizacion.TipoFiltro = sghFiltrarHospitalizacion
''        ucAdmisionHospitalizacion.IdUsuario = ml_IdUsuarioAuditoria
'''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
'''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
'''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 3
''        ConfigurarControl ucAdmisionHospitalizacion
''    Case "AlojadosHospitalizacion"
''        toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = False
''        toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = False
''        ucAdmisionHospitalizacion.Titulo = "Admisión de Alojados"
''        ucAdmisionHospitalizacion.TipoFiltro = sghFiltrarHospitalizacion
''        ucAdmisionHospitalizacion.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucAdmisionHospitalizacion
''    Case "CamasHospitalizacion"
''        ConfigurarControl ucCamasLista1
''        ucCamasLista1.idTipoServicio = sghHospitalizacion
''        ucCamasLista1.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
''        ucCamasLista1.EsListaParaMantenimiento = True
''        ucCamasLista1.RealizarBusqueda
''    Case "RecetasH"
''        ucRecetasLista1.idTipoServicio = sghHospitalizacion
''        ConfigurarControl ucRecetasLista1
''
''    'MODULO PROGRAMACION MEDICA
''    Case "Programacion"
''        ucProgramacionLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ucProgramacionLista1.lnIdTablaLISTBARITEMS = 401
''        ucProgramacionLista1.lcNombrePc = lc_NombrePc
''        ConfigurarControl ucProgramacionLista1
''
''    Case "Turno"
''        ConfigurarControl ucTurnosLista1
''
''    Case "Medico"
''        ConfigurarControl ucMedicosLista1
''
''    'MODULO ARCHIVO CLINICO
''    Case "HistoriaClinica"
''        ConfigurarControl ucHistoriaClinicaLista1
''    Case "MovimientoHistoria"
''        Me.ucMovimientoHistoriasLista1.TipoBusqueda = sghTodasHistorias
''        ConfigurarControl ucMovimientoHistoriasLista1
''        ucMovimientoHistoriasLista1.inicializar
''    Case "SolicitudHistorias"
''        Me.ucSolicitudHistoriasLista1.TipoBusqueda = sghHistoriaEnPrestamo
''        If ucSolicitudHistoriasLista1.IdArchivero = 0 Then
''            ucSolicitudHistoriasLista1.IdArchivero = ml_IdUsuarioAuditoria
''        End If
''        ConfigurarControl ucSolicitudHistoriasLista1
''    Case "Archivero"
''        LbEsConsultorioAsignado = False
''        Me.ucArchivadoresLista1.TipoBusqueda = sghHistoriaEnPrestamo
''        ConfigurarControl ucArchivadoresLista1
''    Case "MovFormatosHC"
''        ConfigurarControl ucMovimientoFormatoHcLista1
''        ucMovimientoFormatoHcLista1.inicializar
''
''    'FACTURACION SERVICIOS
''    Case "FacturacionGeneral"
''        ucFacturacionGeneralLista.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFacturacionGeneralLista
''        ucFacturacionGeneralLista.PuntoCarga = 1 'General
''        ucFacturacionGeneralLista.idTipoFinanciamiento = 1
''        ucFacturacionGeneralLista.Titulo = "Consumo en el Servicio"
''    Case "FacturacionPatologiaClinica"
''        ucFactPatologiaClinica.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactPatologiaClinica
''        ucFactPatologiaClinica.HabilitarPuntoCarga = False
''        ucFactPatologiaClinica.PuntoCarga = 2 'Patología Clínica
''        ucFactPatologiaClinica.idTipoFinanciamiento = 1
''        ucFactPatologiaClinica.Titulo = "Facturacion Laboratorio"
''    Case "FacturacionAnatomiaPatologica"
''        ucFactAnatomiaPatologica.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactAnatomiaPatologica
''        ucFactAnatomiaPatologica.HabilitarPuntoCarga = False
''        ucFactAnatomiaPatologica.PuntoCarga = 3 'Anatomía Patológica
''        ucFactAnatomiaPatologica.idTipoFinanciamiento = 1
''        ucFactAnatomiaPatologica.Titulo = "Facturacion Anatomía Patológica"
''    Case "FacturacionImagenologia"
''        ucFactImagenologia.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactImagenologia
''        ucFactImagenologia.HabilitarPuntoCarga = False
''        ucFactImagenologia.PuntoCarga = 4 'Imagenología
''        ucFactImagenologia.idTipoFinanciamiento = 1
''        ucFactImagenologia.Titulo = "Facturacion Imagenología"
''    Case "FacturacionSalaOperaciones"
''        ucFactSalaOperaciones.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactSalaOperaciones
''        ucFactSalaOperaciones.HabilitarPuntoCarga = False
''        ucFactSalaOperaciones.PuntoCarga = 7 'Sala de Operaciones
''        ucFactSalaOperaciones.idTipoFinanciamiento = 1
''        ucFactSalaOperaciones.Titulo = "Facturacion Sala Operaciones"
''    Case "FactReembolsos"
''        '18/7/11
''        toolbar.Toolbars("Admisión Hospitalización").Visible = True
''        toolbar.Toolbars("Admisión Hospitalización").DockedRow = 3
''        toolbar.Toolbars("Admisión Hospitalización").DockedColumn = 1
''        toolbar.Toolbars("Admisión Hospitalización").Tools.Item(1).Name = "Agregar Reembolso x Cuenta"
''        toolbar.Tools("ID_HospitalizacionTransferencias").Visible = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Visible = False
''       ' toolbar.Toolbars("Admisión Hospitalización").Tools.Item(2).Name = "."
''       ' toolbar.Toolbars("Admisión Hospitalización").Tools.Item(3).Name = "."
''        '18/7/11
''        ucReembolsosLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ucReembolsosLista1.lnHWnd = Me.hwnd
''        ConfigurarControl ucReembolsosLista1
''    Case "PacExtConSeguro"
''        ucPacienteExternos1.IdUsuario = ml_IdUsuarioAuditoria
''        ucPacienteExternos1.EsPacienteSinSeguro = False
''        ConfigurarControl ucPacienteExternos1
''    Case "PacExtParticular"
''        ucPacienteExternos1.IdUsuario = ml_IdUsuarioAuditoria
''        ucPacienteExternos1.EsPacienteSinSeguro = True
''        ConfigurarControl ucPacienteExternos1
''
''    'FACTURACION CONFIGURACION
''    Case "FacturacionCatalogoServicios"
''
''        ConfigurarControl Me.ucCatalogoServiciosLista1
''
''    Case "FacturacionCentroCostos"
''
''        ConfigurarControl Me.ucCentrosCostoLista1
''    Case "PqteServicio"
''        ConfigurarControl Me.ucFactPaquetesLista1
''    Case "EstadoCuenta"
''        Me.ucEstadoCuenta1.IdUsuario = ml_IdUsuarioAuditoria
''        Me.ucEstadoCuenta1.lcNombrePc = lc_NombrePc
''        Me.ucEstadoCuenta1.lnIdTablaLISTBARITEMS = 613
''        Me.ucEstadoCuenta1.lnHWnd = Me.hwnd
''        'Me.ucEstadoCuenta1.Inicializar
''        ConfigurarControl Me.ucEstadoCuenta1
''
''    Case "Farmacia"
''
''        ucFactFarmacia.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactFarmacia
''
''        ucFactFarmacia.PuntoCarga = 5 'Farmacia
''        ucFactFarmacia.HabilitarPuntoCarga = False
''        ucFactFarmacia.idTipoFinanciamiento = 1
''        ucFactFarmacia.Titulo = "Facturacion Farmacia"
''
''    'modificación samuel 02/06
''    Case "ConfiguracionResLab"
''        ConfigurarControl ucConfiguraResLab2
''
''    'MODULO GENERAL
''    Case "Empleado"
''        ConfigurarControl ucEmpleadosLista1
''    Case "Servicios"
''        ConfigurarControl ucServiciosLista1
''    Case "Diagnosticos"
''        ConfigurarControl ucDiagnosticosLista1
''    Case "Procedimientos"
''        ConfigurarControl ucProcedimientosLista1
''    Case "TiposFinanciamiento"
''        ConfigurarControl ucTiposFinanciamientoLista1
''    Case "FuentesFinanciamiento"
''        ConfigurarControl ucFuentesFinanciamientoLista1
''    Case "FacturacionPartidas"
''        ConfigurarControl ucPartidasLista1
''    Case "EstablecimientosNoMinsa"
''        ConfigurarControl ucEstablecimientosNoMinsaLista1
''        ucEstablecimientosNoMinsaLista1.ConfigurarEstablecimientos
''    Case "DiagnosticosPDF"
''        Dim oShell As New sighentidades.Shell
''        If sighentidades.RutaAdobeReader <> "" Then
''            oShell.ejecutarComando sighentidades.RutaAdobeReader + " " + App.Path + "\archivos\" + "cie10.pdf"
''        Else
''            MsgBox "No tiene instalado el adobe reader", vbInformation, Me.Caption
''        End If
''    Case "Especialidades"
''        ConfigurarControl ucEspecialidadesLista1
''    Case "TipoTarifa"
''        ConfigurarControl ucTiposTarifaLista1
''    'MODULO DE CAJA
''    Case "Cajas"
''        ConfigurarControl ucCajaLista1
''    Case "Cajeros"
''        ConfigurarControl ucCajeroLista1
''    Case "AsignacionTerminales"
''    Case "GestionCaja"
''        toolbar.Toolbars("Edición").Visible = False
''        'toolbar.Toolbars("Gestión de Caja").Visible = True
''        If (mb_abrioCaja) Then
''            If mo_LastControl Is ucGestionCaja1 Then
''                mo_LastControl.Visible = True
''                toolbar.Toolbars("Gestión de Caja").Visible = True
''                Exit Sub
''            End If
''            mo_LastControl.Visible = False
''            ucGestionCaja1.NombreCajero = status.Panels(2).Text
''            ucGestionCaja1.Visible = True
''            Set mo_LastControl = ucGestionCaja1
''            toolbar.Toolbars("Gestión de Caja").Visible = True
''            Exit Sub
''        End If
''
''        ucGestionCaja1.IdUsuario = ml_IdUsuarioAuditoria
''        ucGestionCaja1.NombreCajero = status.Panels(2).Text
''        ucGestionCaja1.lnIdTablaLISTBARITEMS = 702
''        ucGestionCaja1.lcNombrePc = lc_NombrePc
''        toolbar.Toolbars("Gestión de Caja").Visible = True
''
''        ConfigurarControl ucGestionCaja1
''
''    '/********************INO****************************/
''    Case "Devoluciones"
''        MsgBox "Se retiro el modulo del SISGalenPlus, posteriormente se estará agregando " & Chr(13) & " el modulo 'Nota de Crédito para las Devoluciones", vbInformation, Me.Caption
'''        toolbar.Toolbars("Edición").Visible = False
'''        'toolbar.Toolbars("Gestión de Caja").Visible = True
'''
'''        If (mb_abrioCaja) Then
'''            If mo_LastControl Is ucGestionDevolucion2 Then
'''                mo_LastControl.Visible = True
'''                'ConfigurarControl ucGestionDevolucion2 '
'''                toolbar.Toolbars("Gestión de Caja").Visible = True
'''                Exit Sub
'''            End If
'''            mo_LastControl.Visible = False
'''
'''            ucGestionDevolucion2.idUsuario = ml_IdUsuarioAuditoria
'''            ucGestionDevolucion2.NombreCajero = status.Panels(2).Text
'''            ucGestionDevolucion2.Visible = True
'''            Set mo_LastControl = ucGestionDevolucion2
'''            toolbar.Toolbars("Gestión de Caja").Visible = True
'''            Exit Sub
'''        End If
'''
'''        ucGestionDevolucion2.idUsuario = ml_IdUsuarioAuditoria
'''        ucGestionDevolucion2.NombreCajero = status.Panels(2).Text
'''        ucGestionDevolucion2.lnIdTablaLISTBARITEMS = 702
'''        ucGestionDevolucion2.lcNombrePc = lc_NombrePc
'''        toolbar.Toolbars("Gestión de Caja").Visible = True
'''
'''        ConfigurarControl ucGestionDevolucion2
'''      '/********************INO****************************/
''
''    '/********************FCV MAYO2015****************************/
''    Case "NotaCredito"
''        ucCajaNotaCredito1.lnHWnd = Me.hwnd
''        ucCajaNotaCredito1.inicializar
''        ConfigurarControl ucCajaNotaCredito1
''    '/********************INO****************************/
''    'MODULO FARMACIA
''    Case "Inventario"
''        ConfigurarControl ucFarmInventarioLista1
''    Case "NI", "NIF", "FARMADOP"                                                                         'debb2014
''        ucFarmNiLista1.NIsoloParaFarmacia = IIf(ms_ModuloSeleccionado = "NI", False, True)    'debb2014
''        ucFarmNiLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmNiLista1
''        If ms_ModuloSeleccionado = "FARMADOP" Then
''           ucFarmNiLista1.Titulo = "ARMADO DE PAQUETES"
''        End If
''    Case "NS", "NSF"                                                                         'debb2014
''        ucFarmNsLista1.NSsoloParaFarmacia = IIf(ms_ModuloSeleccionado = "NS", False, True)   'debb2014
''        ucFarmNsLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmNsLista1
''    Case "IntervencionS"
''        ucFarmIntervencionLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmIntervencionLista1
''    Case "Ventas"
''        ucFarmVentasLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmVentasLista1
''    Case "DependenciaExt"
''        ConfigurarControl ucFarmDependExtLista1
''    Case "DespachoDonaciones"
''        ucFarmDespachoDonaciones1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFarmDespachoDonaciones1
''    Case "FarmAlmacen"
''        ConfigurarControl ucFarmAlmacenes1
''     Case "FarmPrecios"                              'debb2014b
''        'ConfigurarControl Me.ucFarmHistoricoPrecios1     'debb2014b
''        ConfigurarControl ucFarmHpreciosLista1
''
''    'CATALOGOS
''    Case "CatalogoBienes"
''        ucCatalogoBienesInsumosLista1.IdUsuario = ml_IdUsuarioAuditoria
''        ucCatalogoBienesInsumosLista1.lnHWnd = Me.hwnd
''        ConfigurarControl Me.ucCatalogoBienesInsumosLista1
''    'SEGURIDAD
''    Case "Roles"
''        ConfigurarControl ucRolesLista1
''        Set ucRolesLista1.DataSource = mo_AdminSeguridad.RolesSeleccionarTodos()
''    'MODULO IMAGENEOLOGIA
''    Case "ImagRayosX"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 21 'PuntoCarga.Rayos X
''        UcImagenesLista1.Titulo = "Rayos X"
''    Case "ImagEcografiaG"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 20 'PuntoCarga.EcografiaGeneral
''        UcImagenesLista1.Titulo = "Ecografía General"
''    Case "ImagTomografia"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 22 'PuntoCarga.tomografia
''        UcImagenesLista1.Titulo = "Tomografía"
''    Case "ImagEcografiaO"
''        ConfigurarControl UcImagenesLista1
''        UcImagenesLista1.PuntoCarga = 23 'PuntoCarga.EcografiaObstetrica
''        UcImagenesLista1.Titulo = "Ecografía Obstétrica"
''    Case "ImagIngresos"
''        UcImagIngresos1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcImagIngresos1
''    Case "ImagSalidas"
''        UcImagSalidas1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcImagSalidas1
''    'mgaray201411f
''    Case "ImagTipoModalidadSala"
'''        ucImagTipoModalidadSala1.idUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucImagTipoModalidadSala1
''    Case "ImagSala"
''        ConfigurarControl ucImagSala1
''    Case "ImagCatalgoServicioDuracion":
''        ConfigurarControl ucImagCatalgoServicioDuracion1
''    Case "IntegracionSistema"
''        ConfigurarControl ucInteoIntegracionSistema1
''    'Módulo LABORATORIO
''    Case "OrdenesLaboratorio"
''        ucFactOrdenesLaboratorio.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFactOrdenesLaboratorio
''        ucFactOrdenesLaboratorio.HabilitarPuntoCarga = False
''        ucFactOrdenesLaboratorio.PuntoCarga = 2 'Patología Clínica
''        ucFactOrdenesLaboratorio.idTipoFinanciamiento = 1
''        ucFactOrdenesLaboratorio.Titulo = "Órdenes para Análisis de Laboratorio (Patología Clínica)"
''        ucFactOrdenesLaboratorio.AreaTrabajo = 69
''        ucFactOrdenesLaboratorio.lcNombrePc = lc_NombrePc
''        ucFactOrdenesLaboratorio.lnIdTablaLISTBARITEMS = 1312
''    Case "OrdenesPatologia"
''        ucFactOrdenesLaboratorio.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFacturacionOrdenesPatologia
''        ucFacturacionOrdenesPatologia.HabilitarPuntoCarga = False
''        ucFacturacionOrdenesPatologia.PuntoCarga = 3 'Anatomía Patológica
''        ucFacturacionOrdenesPatologia.idTipoFinanciamiento = 1
''        ucFacturacionOrdenesPatologia.Titulo = "Órdenes para Análisis de Laboratorio (Anatomía Patológica)"
''        ucFacturacionOrdenesPatologia.AreaTrabajo = 70
''        ucFacturacionOrdenesPatologia.lcNombrePc = lc_NombrePc
''        ucFacturacionOrdenesPatologia.lnIdTablaLISTBARITEMS = 1321
''     Case "BS"
''        ucFacturacionBS.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucFacturacionBS
''        ucFacturacionBS.HabilitarPuntoCarga = False
''        ucFacturacionBS.PuntoCarga = 11 'Banco de Sangre
''        ucFacturacionBS.idTipoFinanciamiento = 1
''        ucFacturacionBS.Titulo = "Órdenes del Banco de Sangre"
''        ucFacturacionBS.AreaTrabajo = 69
''        ucFacturacionOrdenesPatologia.lcNombrePc = lc_NombrePc
''        ucFacturacionOrdenesPatologia.lnIdTablaLISTBARITEMS = 1322
''     Case "LabIngresos"
''        UcLabIngresos1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcLabIngresos1
''        UcLabIngresos1.idTipoFinanciamiento = 1
''        UcLabIngresos1.Titulo = "Ingreso de Insumos"
''    Case "LabEgresos"
''        UcLabSalidas1.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl UcLabSalidas1
''        UcLabSalidas1.idTipoFinanciamiento = 1
''        UcLabSalidas1.Titulo = "Salida de Insumos"
''
''    'Estadística
''    Case "Constancias"
''        ucContanciasAtencion.IdUsuario = ml_IdUsuarioAuditoria
''        ConfigurarControl ucContanciasAtencion
''        ucContanciasAtencion.Titulo = "Constancias de Atención y Hospitalización"
''    'SIS
''    Case "Fua"
''        ConfigurarControl Me.UcSISfuaLista1
''    End Select
'End Sub
Sub ConfigurarControl(oControl As Control)

        On Error Resume Next

        If oControl Is ucCitasLista1 Then
            If Not mb_MantenerValoresCitas Then
                oControl.Inicializar
            End If
        ElseIf oControl Is Me.ucArchivadoresLista1 Then
            Me.ucArchivadoresLista1.EsConsultorioAsignado = LbEsConsultorioAsignado
        Else
            oControl.Inicializar
        End If

      '  mo_LastControl.Visible = False
        oControl.Visible = True


        Set mo_LastControl = oControl
        Form_Resize


End Sub
'
'Private Sub tmrHora_Timer()
'  status.Panels(1).Text = ""
'End Sub
'
'Private Sub toolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
'
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    '**********************************************************************
'    '   MANEJO DEL MENU PRINCIPAL
'    '   DE ACUERDO AL MODULO SELECCIONADO
'    '**********************************************************************
'
'    Select Case Tool.Id
'    Case "ID_Archivo", "ID_Reportes", "ID_ProgramacionMedica", "ID_ArchivoClinico", "ID_Herramientas", "ID_Ayuda", "ID_ReportesDeFarmacia", "ID_HerrFarmacia"
'        Exit Sub
'    Case "ID_RptHospitalizacion", "ID_Emergencia", "ID_Economia", "ID_Seguros", "ID_Convenios", "ID_HerrConsultaExterna", "ID_Imagenologia", "ID_LaboratorioMod", "ID_ModuloHIS"
'        Exit Sub
'    Case "ID_Salir"
'        mo_AdminSeguridad.LogueaUsuario 0, sighentidades.USUARIO, lc_NombrePc
'        If sighentidades.Parametro282valorInt = "1" Then
'           Me.Visible = False
'        Else
'           End
'        End If
'
'    '*****************************   REPORTES   ******************************
'    Case "ID_ImportaAFILIADOSSIS"
'        AgregaAfiliadosSIS lcBuscaParametro.SeleccionaFilaParametro(313), lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'    Case "id_PadronNominal"                                                 'debb-2/3/2015
'        Dim oRptCEpadronNominal As New RptCEpadronNominal                   'debb-2/3/2015
'        oRptCEpadronNominal.EjecutaFormulario                               'debb-2/3/2015
'        Set RptCEpadronNominal = Nothing                                    'debb-2/3/2015
'        Exit Sub                                                            'debb-2/3/2015
'    Case "ID_RptProgMedica"
'        Dim oProgMedicaRpt As New SIGHReportes.clProgramMedica
'        oProgMedicaRpt.EjecutaFormulario
'        Set oProgMedicaRpt = Nothing
'        Exit Sub
'    Case "ID_RptHistoriaSolicitadas"
'        Dim oSolicitud As New SIGHReportes.clSolicitudHistorias
'        oSolicitud.TipoReporte = "RPT_HISTORIAS_SERVICIO"
'        oSolicitud.idUsuario = ml_IdUsuarioAuditoria
'        oSolicitud.EjecutaFormulario
'        Set oSolicitud = Nothing
'        Exit Sub
'    Case "ID_RptHistoriaSolicitadasMedico"
'        Dim oSolicitudMedico As New SIGHReportes.clSolicitudHistorias
'        oSolicitudMedico.TipoReporte = "RPT_HISTORIAS_MEDICO"
'        oSolicitudMedico.idUsuario = ml_IdUsuarioAuditoria
'        oSolicitudMedico.EjecutaFormulario
'        Set oSolicitudMedico = Nothing
'        Exit Sub
'    Case "ID_EgresosHospitalarios"
'        Dim oRptHosp As New SIGHProxies.clReportesEgreHosp
'        oRptHosp.IdTipoReporte = sighentidades.sghReporteEgresosHospitalario
'        oRptHosp.idTipoServicio = 0
'        oRptHosp.EjecutaFormulario
'        Set oRptHosp = Nothing
'        Exit Sub
'    Case "ID_IngresosHospitalarios"
'        Dim oRptIngHosp As New SIGHProxies.clReporteIngrHosp
'        oRptIngHosp.IdTipoReporte = sighentidades.sghReporteIngresosHospitalario
'        oRptIngHosp.EjecutaFormulario
'        Set oRptIngHosp = Nothing
'        Exit Sub
'    Case "ID_CensoHospitalario"
'        Dim oRptCensoHospitalario As New SIGHReportes.clAtencionesCenso
''        oRptoRptCensoHospitalarioIngHosp.IdTipoReporte = sighEntidades.sghReporteIngresosHospitalario
'        oRptCensoHospitalario.EjecutaFormulario
'        Set oRptCensoHospitalario = Nothing
'        Exit Sub
'    Case "ID_CuposAsignados"
'        Dim oRptCuposAsignados As New SIGHReportes.clCuposAsignadosRep
'        oRptCuposAsignados.EjecutaFormulario
'        Set oRptCuposAsignados = Nothing
'        Exit Sub
'    Case "ID_CambiodeClave"
'        Dim oCambClave As New LoginActualizaClave
'        oCambClave.idUsuario = ml_IdUsuarioAuditoria
'        oCambClave.Show 1
'        Set oCambClave = Nothing
'        Exit Sub
'    Case "ID_AcercaDe"
'        Splash.Show 1
'        Unload Splash
'        Exit Sub
'    End Select
'
'    '***************REPORTES**************
'    Select Case Tool.Id
'    'Consulta externa
'    Case "ID_MorbilidadCE"
'        Dim oRptMorbilidadCE As New SIGHReportes.RptHMorbCE
'        oRptMorbilidadCE.EjecutaFormulario
'        Set oRptMorbilidadCE = Nothing
'        Exit Sub
'    Case "Id_RepMaterno"
'        Dim oRptRepMaterno As New SIGHReportes.clCeMaterno
'        oRptRepMaterno.EjecutaFormulario
'        Set oRptRepMaterno = Nothing
'        Exit Sub
'    Case "Id_RepPerinatal"
'        Dim oRptRepPerinatal As New SIGHReportes.clCePerinatal
'        oRptRepPerinatal.EjecutaFormulario
'        Set oRptRepPerinatal = Nothing
'        Exit Sub
'    'mgaray201411h
'    Case "Id_RepPerinatalIndicadores"
'        Dim oRptRepPerinatalIndicadores As New SIGHReportes.clCePerinatalIndicadores
'        oRptRepPerinatalIndicadores.EjecutaFormulario
'        Set oRptRepPerinatalIndicadores = Nothing
'        Exit Sub
'
'    'MODULO Reportes
'    Case "ID_PacientesmenoresaNanios"
'        Dim oRptMovimientoHistorias As New SIGHReportes.RptAHCpacienteHastaNanio
'        oRptMovimientoHistorias.EjecutaFormulario
'        Set oRptMovimientoHistorias = Nothing
'        Exit Sub
'    Case "ID_MovimientosdeHistorias"
'        Dim oRptAHCMovimEntSal As New SIGHReportes.RptAHCMovimEntSal
'        oRptAHCMovimEntSal.EjecutaFormulario
'        Set oRptAHCMovimEntSal = Nothing
'        Exit Sub
'    Case "ID_MovimientodeFormatosdeHistorias"
'        Dim oRptAHCMovimFormatos As New SIGHReportes.RptAHCMovimFormatos
'        oRptAHCMovimFormatos.EjecutaFormulario
'        Set oRptAHCMovimFormatos = Nothing
'        Exit Sub
'    Case "ID_MovimientoFormatosHCporMes"
'        Dim oRptAHCMovimFormatMes As New SIGHReportes.RptAHCMovimFormatMes
'        oRptAHCMovimFormatMes.EjecutaFormulario
'        Set oRptAHCMovimFormatMes = Nothing
'        Exit Sub
'    Case "ID_HCsolicPorServ"
'        Dim oRpt219 As New SIGHReportes.RptAHSolicPorServ
'        oRpt219.EjecutaFormulario
'        Set oRpt219 = Nothing
'        Exit Sub
'    Case "ID_HCsolicPorMedico"
'        Dim oRpt220 As New SIGHReportes.RptAHSolicPorMedico
'        oRpt220.EjecutaFormulario
'        Set oRpt220 = Nothing
'        Exit Sub
'    Case "ID_HCespeciales"
'        MsgBox "Use el reporte: Relación de historias clinicas de pacientes judiciales"
''        Dim oRpt221 As New SIGHReportes.RptAHSolicPorTipo
''        oRpt221.EjecutaFormulario
''        Set oRpt221 = Nothing
'        Exit Sub
'    Case "ID_HCpaciVIH"
'        Dim oRpt222 As New SIGHReportes.RptAHCconVIH
'        oRpt222.EjecutaFormulario
'        Set oRpt222 = Nothing
'        Exit Sub
'    Case "ID_HCpaciJudiciales"
'        Dim oRpt223 As New SIGHReportes.RptAHSolicPorTipo
'        oRpt223.EjecutaFormulario
'        Set oRpt223 = Nothing
'        Exit Sub
'    Case "ID_HCnoLlegan24hr"
'        Dim oRpt224 As New SIGHReportes.RptAHCEgresoMedico24
'        oRpt224.EjecutaFormulario
'        Set oRpt224 = Nothing
'        Exit Sub
'    Case "ID_HIndicAnual"
'        MsgBox "Use el reporte: Indicadores Hospitalarios por Dpto/Servicio/Especialidad"
''        Dim oRpt11 As New SIGHReportes.RptHIndicadorAnual
''        oRpt11.EjecutaFormulario
''        Set oRpt11 = Nothing
'        Exit Sub
'    Case "ID_HIndicMeses"
'        Dim oRpt22 As New SIGHReportes.RptHIndicadorMeses
'        oRpt22.EjecutaFormulario
'        Set oRpt22 = Nothing
'        Exit Sub
'    Case "ID_HIndicAnual1"
'        Dim oRpt13 As New SIGHReportes.RptHIndicadorAnual
'        oRpt13.EjecutaFormulario
'        Set oRpt13 = Nothing
'        Exit Sub
'    Case "ID_HEgresosHosp"
'        Dim oRpt24 As New SIGHReportes.RptHEgresosHosp
'        oRpt24.EjecutaFormulario
'        Set oRpt24 = Nothing
'        Exit Sub
'    Case "ID_HIngresosHosp"
'        Dim oRpt25 As New SIGHReportes.RptHIngresosHosp
'        oRpt25.EjecutaFormulario
'        Set oRpt25 = Nothing
'        Exit Sub
'    Case "ID_HTransf"
'        Dim oRpt26 As New SIGHReportes.RptHTransferencia
'        oRpt26.EjecutaFormulario
'        Set oRpt26 = Nothing
'        Exit Sub
'    Case "ID_HMortalidadC"
'        MsgBox "Use el reporte: Mortalidad Hospitalaria por causa básica, según ciclos de vida por Dpto/Especialidad"
''        Dim oRpt27 As New SIGHReportes.RptHMortalidad
''        oRpt27.EjecutaFormulario
''        Set oRpt27 = Nothing
'        Exit Sub
'    Case "ID_HMortalidadD"
'        MsgBox "Use el reporte: Mortalidad Hospitalaria por causa básica, según ciclos de vida por Dpto/Especialidad"
''        Dim oRpt28 As New SIGHReportes.RptHMortalidad
''        oRpt28.EjecutaFormulario
''        Set oRpt28 = Nothing
'        Exit Sub
'    Case "ID_HMortalidadE"
'        Dim oRpt29 As New SIGHReportes.RptHMortalidad
'        oRpt29.EjecutaFormulario
'        Set oRpt29 = Nothing
'        Exit Sub
'    Case "ID_HMorbilidadC"
'        MsgBox "Use el reporte: Primeras causas de morbilidad Hospitalaria por Diagnósticos, según ciclos de vida por Dpto/Especialidad"
''        Dim oRpt210 As New SIGHReportes.RptHMorbilidad
''        oRpt210.EjecutaFormulario
''        Set oRpt210 = Nothing
'        Exit Sub
'    Case "ID_HMorbilidadD"
'        MsgBox "Use el reporte: Primeras causas de morbilidad Hospitalaria por Diagnósticos, según ciclos de vida por Dpto/Especialidad"
''        Dim oRpt211 As New SIGHReportes.RptHMorbilidad
''        oRpt211.EjecutaFormulario
''        Set oRpt211 = Nothing
'        Exit Sub
'    Case "ID_HMorbilidadE"
'        Dim oRpt212 As New SIGHReportes.RptHMorbilidad
'        oRpt212.EjecutaFormulario
'        Set oRpt212 = Nothing
'        Exit Sub
'    Case "ID_HProcedimientos"
'        Dim oRpt213 As New SIGHReportes.RptHProcedimientos
'        oRpt213.EjecutaFormulario
'        Set oRpt213 = Nothing
'        Exit Sub
'    Case "ID_HDiasEstancia"
'        Dim oRpt214 As New SIGHReportes.RptHEstanciaH
'        oRpt214.EjecutaFormulario
'        Set oRpt214 = Nothing
'        Exit Sub
'    Case "ID_HIndicPrPermanencia"
'        Dim oRpt215 As New SIGHReportes.RptHPrPermanencia
'        oRpt215.EjecutaFormulario
'        Set oRpt215 = Nothing
'        Exit Sub
'    Case "ID_HCamasH"
'        Dim oRpt216 As New SIGHReportes.RptHCamas
'        oRpt216.EjecutaFormulario
'        Set oRpt216 = Nothing
'        Exit Sub
'    Case "ID_HDiasCamaH"
'        Dim oRpt217 As New SIGHReportes.RptHCamaDias
'        oRpt217.EjecutaFormulario
'        Set oRpt217 = Nothing
'        Exit Sub
'    Case "ID_HDiasPacienteH"
'        Dim oRpt218 As New SIGHReportes.RptHDiasPaciente
'        oRpt218.EjecutaFormulario
'        Set oRpt218 = Nothing
'        Exit Sub
'    Case "ID_EMorbilidad"
'        Dim oRpt225 As New SIGHReportes.RptHMorbEm
'        oRpt225.EjecutaFormulario
'        Set oRpt225 = Nothing
'        Exit Sub
'    Case "ID_MinsaEssalud"
'        Dim oRpt231 As New SIGHReportes.RptEAtencConv
'        oRpt231.EjecutaFormulario
'        Set oRpt231 = Nothing
'        Exit Sub
'    Case "ID_MinsaFospolis"
'      '  Dim oRpt232 As New SIGHReportes.RptEFospolis
'      '  oRpt232.EjecutaFormulario
'      '  Set oRpt232 = Nothing
'      '  Exit Sub
'    Case "ID_ReportedeRegistrodeInformaciónporUsuariodelSistema"
'        Dim oRpt233 As New SIGHReportes.RptHerrUsuarioSistema
'        oRpt233.EjecutaFormulario
'        Set oRpt233 = Nothing
'        Exit Sub
'    Case "ID_ImprimeFormatoHIS"
'        Dim oRpt234 As New SIGHProxies.RptCEhis
'        oRpt234.EjecutaFormulario
'        Set oRpt234 = Nothing
'        Exit Sub
'    Case "ID_GastosdePacientes"
'      '  Dim oRpt235 As New SIGHReportes.RptCEgastosDePacientes
'       ' oRpt235.EjecutaFormulario
'       ' Set oRpt235 = Nothing
'       ' Exit Sub
'    Case "ID_FrecuenciadeDxdePacientesatendidos"
'        Dim oRpt236 As New SIGHReportes.RptCEdx
'        oRpt236.EjecutaFormulario
'        Set oRpt236 = Nothing
'        Exit Sub
'    Case "ID_ConsumoServiciosdePacientesAtendidos"
'        Dim oRpt237 As New SIGHReportes.RptCEservi
'        oRpt237.EjecutaFormulario
'        Set oRpt237 = Nothing
'        Exit Sub
'    Case "ID_CierredeCuentasdeAtención"
'        CierreCtaAtencion
'        Exit Sub
'    Case "ID_EgresosConsultaExterna(Epicrisis)"
'        Dim oRptHosp2 As New SIGHProxies.clReportesEgreHosp
'        oRptHosp2.IdTipoReporte = sighentidades.sghReporteEgresosHospitalario
'        oRptHosp2.idTipoServicio = 2
'        oRptHosp2.EjecutaFormulario
'        Set oRptHosp2 = Nothing
'        Exit Sub
'    Case "ID_EgresosEmergencia(Epicrisis)"
'        Dim oRptHosp1 As New SIGHProxies.clReportesEgreHosp
'        oRptHosp1.IdTipoReporte = sighentidades.sghReporteEgresosHospitalario
'        oRptHosp1.idTipoServicio = 1
'        oRptHosp1.EjecutaFormulario
'        Set oRptHosp2 = Nothing
'        Exit Sub
'    Case "ID_IngresosEmergencia"
'        Dim oRptIngHosp1 As New SIGHProxies.clReporteIngrHosp
'        oRptIngHosp1.IdTipoReporte = sighentidades.sghReporteIngresosHospitalario
'        oRptIngHosp1.idTipoServicio = 1
'        oRptIngHosp1.EjecutaFormulario
'        Set oRptIngHosp1 = Nothing
'        Exit Sub
'    Case "ID_IndicadordeAtencionesvsAtendidos"
'        Dim oRpt238 As New SIGHReportes.RptCEatenciones
'        oRpt238.EjecutaFormulario
'        Set oRpt238 = Nothing
'        Exit Sub
'    Case "Id_hcNOusadas"
'        Dim oRptHCnoUsadas As New SIGHReportes.RptAHhcNOusadas
'        oRptHCnoUsadas.EjecutaFormulario
'        Set oRptHCnoUsadas = Nothing
'        Exit Sub
'    Case "Id_NoLlegaAC"
'        Dim oRptHCnoLlegaAC As New SIGHReportes.RptAHhcNoLlegaAC
'        oRptHCnoLlegaAC.EjecutaFormulario
'        Set oRptHCnoLlegaAC = Nothing
'        Exit Sub
'    Case "ID_AtencionSISHECE"
'        Dim oRptclAtencionesTotales As New SIGHProxies.clAtencionesTotales
'        oRptclAtencionesTotales.EjecutaFormulario
'        Set oRptclAtencionesTotales = Nothing
'        Exit Sub
'
'   'MODULO Farmacia - Reportes
''    Case "ID_ActualizaFVencimiento"   'Adams
''        Dim oActualizaSaldo As New SighFarmacia.ActualizaSaldo
''        oActualizaSaldo.idUsuario = ml_IdUsuarioAuditoria
''        oActualizaSaldo.lcNombrePc = lc_NombrePc
''        oActualizaSaldo.MostrarFormulario
''        Set oActualizaSaldo = Nothing
''        Exit Sub
'    Case "ID_FarmVtaItems"
'        Dim oRptFKardex As New SighFarmacia.RepMovimientoES
'        oRptFKardex.idUsuario = sighentidades.USUARIO
'        oRptFKardex.EjecutaFrm
'        Set oRptFKardex = Nothing
'        Exit Sub
'    Case "id_kardex"
'        Dim oRptVtas As New SighFarmacia.RepKardex
'        oRptVtas.idUsuario = ml_IdUsuarioAuditoria
'        oRptVtas.EjecutaFormulario
'        Set oRptVtas = Nothing
'        Exit Sub
'    Case "id_saldos"
'        Dim oRptFSaldos As New SighFarmacia.RepSaldosPorAlmacen
'        oRptFSaldos.idUsuario = ml_IdUsuarioAuditoria
'        oRptFSaldos.EjecutaFormulario
'        Set oRptFSaldos = Nothing
'        Exit Sub
'    Case "ID_RegenerarSaldos"
'        Dim oRegeneraSaldos As New SIGHProxies.RegeneraSaldos
'        oRegeneraSaldos.idUsuario = ml_IdUsuarioAuditoria
'        oRegeneraSaldos.lcNombrePc = lc_NombrePc
'        oRegeneraSaldos.MostrarFormulario
'        Set oRegeneraSaldos = Nothing
'        Exit Sub
'    Case "ID_FormatoICI"
'        Dim oRptICI As New SIGHProxies.RepICI
'        oRptICI.idUsuario = ml_IdUsuarioAuditoria
'        oRptICI.EjecutaFormulario
'        Set oRptICI = Nothing
'        Exit Sub
'    Case "ID_FormatoIDI"
'        Dim oRptIDI As New SighFarmacia.RepIDI
'        oRptIDI.idUsuario = ml_IdUsuarioAuditoria
'        oRptIDI.EjecutaFormulario
'        Set oRptIDI = Nothing
'        Exit Sub
'    Case "ID_ProductosporVencer"
'        Dim oRptProdXvencer As New SighFarmacia.RepProductoPorVencer
'        oRptProdXvencer.EjecutaFormulario
'        Set oRptProdXvencer = Nothing
'        Exit Sub
'    Case "ID_MovimientodeDocumentosdeEntradaySalida"
'        'Dim oRptMovES As New SIGHProxies.RepMovimientoES
'        Dim oRptMovES As New SighFarmacia.RepMovimientoES
'        oRptMovES.idUsuario = ml_IdUsuarioAuditoria
'        oRptMovES.EjecutaFormulario
'        Set oRptMovES = Nothing
'        Exit Sub
'    Case "ID_AperturaAnual"
'        Dim oAperturaAnual As New SighFarmacia.AperturaAnual
'        oAperturaAnual.lcNombrePc = lc_NombrePc
'        oAperturaAnual.idUsuario = ml_IdUsuarioAuditoria
'        oAperturaAnual.MostrarFormulario
'        Set oAperturaAnual = Nothing
'        Exit Sub
'    Case "ID_MontossegúnPlan"
'        Dim oMontosP As New SighFarmacia.RepMontosXplan
'        oMontosP.idUsuario = ml_IdUsuarioAuditoria
'        oMontosP.EjecutaFormulario
'        Set oMontosP = Nothing
'        Exit Sub
'    Case "ID_RecetasporServicio"
'        Dim oRecetas As New SighFarmacia.RepRecetasXservicio
'        oRecetas.idUsuario = ml_IdUsuarioAuditoria
'        oRecetas.EjecutaFormulario
'        Set oRecetas = Nothing
'        Exit Sub
'    Case "ID_ConsumoporNCuenta"
'        Dim oConsCta As New SighFarmacia.RepConsumoPorCuenta
'        oConsCta.EjecutaFormulario
'        Set oConsCta = Nothing
'        Exit Sub
'    Case "ID_ConsumopromedioAnual"
'        Dim oConsAnual As New SighFarmacia.RepConsumoPromAnual
'        oConsAnual.idUsuario = ml_IdUsuarioAuditoria
'        oConsAnual.EjecutaFormulario
'        Set oConsAnual = Nothing
'        Exit Sub
'    Case "ID_consumoSegunCodigoReceta"
'        Dim oRepXusuario As New SighFarmacia.RepRecetasXusuario
'        oRepXusuario.idUsuario = ml_IdUsuarioAuditoria
'        oRepXusuario.EjecutaFormulario
'        Set oRepXusuario = Nothing
'        Exit Sub
'    Case "ID_AuditoriaFarm"
'        Dim oRepAuditoriaFarm As New SighFarmacia.RepAuditoriaFarmacia
'        oRepAuditoriaFarm.idUsuario = ml_IdUsuarioAuditoria
'        oRepAuditoriaFarm.EjecutaFormulario
'        Set oRepAuditoriaFarm = Nothing
'        Exit Sub
'    Case "ID_ConsumoXservicio"             'debb-04/04/2011
'        Dim oRepConsumoXservicio As New RepConsumoXservicio
'        oRepConsumoXservicio.EjecutaFormulario
'        Set oRepConsumoXservicio = Nothing
'        Exit Sub
'    'MODULO ECONOMIA - Reportes
'    Case "ID_ReembolsosAnuales"
'        Dim oRptERembolsoAnual As New RptERembolsoAnual
'        oRptERembolsoAnual.idUsuario = ml_IdUsuarioAuditoria
'        oRptERembolsoAnual.EjecutaFormulario
'        Set oRptERembolsoAnual = Nothing
'        Exit Sub
'    Case "ID_ConsolidadoRecaudacion"
'       ' Dim RepConsRecaudacion As New RpParteDiario
'       ' RepConsRecaudacion.IdTipoReporte = 4
'       ' RepConsRecaudacion.idUsuario = ml_IdUsuarioAuditoria
'       ' RepConsRecaudacion.Show 1
'       ' Set RepConsRecaudacion = Nothing
'       ' Exit Sub
'        MsgBox "...Reporte en desarrollo..."
'        Exit Sub
'    Case "ID_InformedeRecaudaciondeAltas"
'        Dim oRpt228 As New SIGHReportes.RptERecaudAltas
'        oRpt228.EjecutaFormulario
'        Set oRpt228 = Nothing
'        Exit Sub
'    Case "ID_ExoneracionesGeneral"
'        Dim oRpt229 As New SIGHReportes.RptEExoneraciones
'        oRpt229.EjecutaFormulario
'        Set oRpt229 = Nothing
'        Exit Sub
'    Case "ID_Liquidación"
'        Dim oRptLiq As New SIGHReportes.RptESisSoatExoConv
'        oRptLiq.idUsuario = ml_IdUsuarioAuditoria
'        oRptLiq.EjecutaFormulario
'        Set oRptLiq = Nothing
'        Exit Sub
'    Case "ID_ConsumoporPuntosdeCarga"
'        Dim oRptConsPtoCarga As New SIGHReportes.RptEConsumoXptoCarga
'        oRptConsPtoCarga.idUsuario = ml_IdUsuarioAuditoria
'        oRptConsPtoCarga.EjecutaFormulario
'        Set oRptConsPtoCarga = Nothing
'        Exit Sub
'    Case "ID_ExoneracionesenGeneral"
'        Dim oRpt239 As New SIGHReportes.RptEExoGeneral
'        oRpt239.EjecutaFormulario
'        Set oRpt239 = Nothing
'        Exit Sub
'    Case "Id_ResumenPartida"
'        Dim oRptResumenPartida As New RptEPartidaResumen
'        oRptResumenPartida.EjecutaFormulario
'        Set oRptResumenPartida = Nothing
'        Exit Sub
'    Case "Id_DetallePartida"
'        Dim oRptPartidaDetalle As New RptEpartidaDetalle
'        oRptPartidaDetalle.EjecutaFormulario
'        Set oRptPartidaDetalle = Nothing
'        Exit Sub
'    Case "ID_RecalculoSOATaParticular"    'debb-04/04/2011
'        Dim oRptEconRecalculoSOAT As New SIGHReportes.RptEconRecalculoSOAT
'        oRptEconRecalculoSOAT.EjecutaFormulario
'        Set oRptEconRecalculoSOAT = Nothing
'        Exit Sub
'    Case "ID_TipoTarifa"
'        Dim oRptEtipoTarifa As New SIGHReportes.RptEtipoTarifa
'        oRptEtipoTarifa.EjecutaFormulario
'        Set oRptEtipoTarifa = Nothing
'        Exit Sub
'    'sunat facturador
'    Case "ID_SunatFacturador"
'        Dim orpCajaExportaSunat As New rpCajaExportaSunat
'        orpCajaExportaSunat.idUsuario = ml_IdUsuarioAuditoria
'        orpCajaExportaSunat.lcNombrePc = lc_NombrePc
'        orpCajaExportaSunat.Show 1
'        Set orpCajaExportaSunat = Nothing
'        Exit Sub
'    'sunat facturador
'    '/******************************************************************/
'    '/***************************INO************************************/
'    '/******************************************************************/
'     Case "ID_CajaDevoluciones"
''        Dim oRptCajaDevoluciones As New SIGHReportes.clRptCajaDevoluciones
''        oRptCajaDevoluciones.EjecutaFormulario
''        Set oRptCajaDevoluciones = Nothing
''        Exit Sub
''    '/******************************************************************/
''    '/***************************INO************************************/
''    '/******************************************************************/
'        MsgBox "El reporte esta en reestructuración", vbInformation, "Mensaje"
'
'    'MODULO IMAGENOLOGIA - Reportes
'    Case "ID_ImgMovimientodiario"
'        Dim oRepImgMovDiario As New SIGHImagen.RepMovimientoDiario
'        oRepImgMovDiario.idUsuario = ml_IdUsuarioAuditoria
'        oRepImgMovDiario.EjecutaFormulario
'        Set oRepImgMovDiario = Nothing
'        Exit Sub
'    Case "ID_ImgKardex"
'        Dim oRepImgKardex As New SIGHImagen.RepKardex
'        oRepImgKardex.idUsuario = ml_IdUsuarioAuditoria
'        oRepImgKardex.EjecutaFormulario
'        Set oRepImgKardex = Nothing
'        Exit Sub
'    Case "ID_ImgEGporFechas"
'        Dim oRepEcogGen As New SIGHImagen.RepEcogGen
'        oRepEcogGen.idUsuario = ml_IdUsuarioAuditoria
'        oRepEcogGen.EjecutaFormulario
'        Set oRepEcogGen = Nothing
'        Exit Sub
'    Case "ID_ImgEOporFechas"
'        Dim oRepEcogObs As New SIGHImagen.RepEcogObs
'        oRepEcogObs.idUsuario = ml_IdUsuarioAuditoria
'        oRepEcogObs.EjecutaFormulario
'        Set oRepEcogObs = Nothing
'        Exit Sub
'    Case "ID_ImgTomoPorFechas"
'        Dim oRepTomografia As New SIGHImagen.RepTomografia
'        oRepTomografia.idUsuario = ml_IdUsuarioAuditoria
'        oRepTomografia.EjecutaFormulario
'        Set oRepTomografia = Nothing
'        Exit Sub
'    Case "ID_ImgRayosXporFechas"
'        Dim oRepRayosX As New SIGHImagen.RepRayosX
'        oRepRayosX.idUsuario = ml_IdUsuarioAuditoria
'        oRepRayosX.EjecutaFormulario
'        Set oRepRayosX = Nothing
'        Exit Sub
'    Case "ID_ImgProductividad"
'        Dim oRepProduccion As New SIGHImagen.RepProduccion
'        oRepProduccion.idUsuario = ml_IdUsuarioAuditoria
'        oRepProduccion.EjecutaFormulario
'        Set oRepProduccion = Nothing
'        Exit Sub
''    Case "ID_ImgProductividad1"   'adams
''        Dim oRepProduccion1 As New SIGHImagen.RepProduccion
''        oRepProduccion1.idUsuario = ml_IdUsuarioAuditoria
''        oRepProduccion1.EjecutaFormulario
''        Set oRepProduccion1 = Nothing
''        Exit Sub
'    Case "ID_ImgAuditoria"
'        Dim oRepAuditoriaImg As New SIGHImagen.RepAuditoriaImg
'        oRepAuditoriaImg.idUsuario = ml_IdUsuarioAuditoria
'        oRepAuditoriaImg.EjecutaFormulario
'        Set oRepAuditoriaImg = Nothing
'        Exit Sub
'    Case "ID_ImgConsumodeInsumos"
'        Dim oRepConsumodeInsumos As New SIGHImagen.RepInsumoPorTipoServ
'        oRepConsumodeInsumos.idUsuario = ml_IdUsuarioAuditoria
'        oRepConsumodeInsumos.EjecutaFormulario
'        Set oRepConsumodeInsumos = Nothing
'        Exit Sub
'    Case "ID_ImgProducciónPagosyDeuda"
'        Dim oRepProducciónPagosyDeuda As New SIGHImagen.RepProducPagoDeuda
'        oRepProducciónPagosyDeuda.idUsuario = ml_IdUsuarioAuditoria
'        oRepProducciónPagosyDeuda.EjecutaFormulario
'        Set oRepProducciónPagosyDeuda = Nothing
'        Exit Sub
'    Case "ID_ImgConsumodeInsumosporServicios"
'        Dim oRepConsumodeInsumosporServicios As New SIGHImagen.RepInsumoPorServicio
'        oRepConsumodeInsumosporServicios.idUsuario = ml_IdUsuarioAuditoria
'        oRepConsumodeInsumosporServicios.EjecutaFormulario
'        Set oRepConsumodeInsumosporServicios = Nothing
'        Exit Sub
'    Case "ID_ReprogramacionMedica"
'        Dim oHerrModifPac As New SIGHProxies.clHerrReprogramMedica
'        oHerrModifPac.idUsuario = ml_IdUsuarioAuditoria
'        oHerrModifPac.MostrarFormulario
'        Set oHerrModifPac = Nothing
'        Exit Sub
'    Case "ID_PasaAtenciondeNN"
'        Dim oHerrModificaNN As New HerrModificaPacienteAtencionHE
'        oHerrModificaNN.idUsuario = ml_IdUsuarioAuditoria
'        oHerrModificaNN.Show 1
'        Set oHerrModificaNN = Nothing
'        Exit Sub
'   Case "ID_ExportaSUNASA"
'       Dim oSUNASA As New SIGHProxies.clSunasa
'       oSUNASA.idUsuario = ml_IdUsuarioAuditoria
'       oSUNASA.lcNombrePc = lc_NombrePc
'       oSUNASA.MostrarFormulario
'       Set oSUNASA = Nothing
'   Case "Id_ActualizaParametros"
'       Dim oActParametros As New HerrActualizacionParametros
'       oActParametros.Show 1
'       Set oActParametros = Nothing
'       Exit Sub
'
'    Case "ID_ReporteSIS"
'        Dim oRepSIS As New SIGHProxies.RptEconRepSIS
'        oRepSIS.idUsuario = ml_IdUsuarioAuditoria
'        oRepSIS.EjecutaFormulario
'        Set oRepSIS = Nothing
'        Exit Sub
'    Case "ID_RepConvenios"
'        Dim oRepConvenios As New rptEconRepConvenios
'        oRepConvenios.idUsuario = ml_IdUsuarioAuditoria
'        oRepConvenios.EjecutaFormulario
'        Set oRepConvenios = Nothing
'        Exit Sub
'    Case "ID_AuditoriaCE"
'        Dim oRptCEauditoria As New RptCEauditoria
'        oRptCEauditoria.idUsuario = ml_IdUsuarioAuditoria
'        oRptCEauditoria.EjecutaFormulario
'        Set oRptCEauditoria = Nothing
'        Exit Sub
'    Case "ID_AuditoriaArchivoClínicos"
'        Dim oRptAHCauditoria As New RptAHCauditoria
'        oRptAHCauditoria.idUsuario = ml_IdUsuarioAuditoria
'        oRptAHCauditoria.EjecutaFormulario
'        Set oRptAHCauditoria = Nothing
'        Exit Sub
'    Case "ID_AuditoriaHosp"
'        Dim oRptHauditoria As New SIGHReportes.RptHauditoria
'        oRptHauditoria.idUsuario = ml_IdUsuarioAuditoria
'        oRptHauditoria.EjecutaFormulario
'        Set oRptHauditoria = Nothing
'        Exit Sub
'    Case "ID_AuditoriaEmerg"
'        Dim oRptEmergAuditoria As New RptEmergAuditoria
'        oRptEmergAuditoria.idUsuario = ml_IdUsuarioAuditoria
'        oRptEmergAuditoria.EjecutaFormulario
'        Set oRptEmergAuditoria = Nothing
'        Exit Sub
'    Case "ID_AuditoriaEcon"
'        Dim oRptEauditoria As New RptEauditoria
'        oRptEauditoria.idUsuario = ml_IdUsuarioAuditoria
'        oRptEauditoria.EjecutaFormulario
'        Set oRptEauditoria = Nothing
'        Exit Sub
'
'    'Herramientas
'    Case "ID_ExportadatosalSistemaSEM"
'       Dim oHerrSem As New SIGHProxies.clExportaSem
'       oHerrSem.idUsuario = ml_IdUsuarioAuditoria
'       oHerrSem.lcNombrePc = lc_NombrePc
'       oHerrSem.MostrarFormulario
'       Set oHerrSem = Nothing
'       Exit Sub
'    Case "ID_ExportaHIS"
'       Dim oHerrHIS As New HerrExportaHIS
'       oHerrHIS.idUsuario = ml_IdUsuarioAuditoria
'       oHerrHIS.lcNombrePc = lc_NombrePc
'       oHerrHIS.Show 1
'       Set oHerrHIS = Nothing
'       Exit Sub
'    Case "ID_ExportaSip2000"
'       Dim oHerrSip2000 As New HerrExportaSIP2000
'       oHerrSip2000.idUsuario = ml_IdUsuarioAuditoria
'       oHerrSip2000.lcNombrePc = lc_NombrePc
'       oHerrSip2000.Show 1
'       Set oHerrSip2000 = Nothing
'       Exit Sub
'    Case "ID_ExportaSIS"
'       Dim oHerrSis As New HerrExportaSIS
'       oHerrSis.idUsuario = ml_IdUsuarioAuditoria
'       oHerrSis.lcNombrePc = lc_NombrePc
'       oHerrSis.Show 1
'       Set oHerrSis = Nothing
'       Exit Sub
'
'   Case "ID_CitasWeb"
'       Dim oHerrExportaCitasWeb As New HerrExportaCitasWeb
'       oHerrExportaCitasWeb.idUsuario = ml_IdUsuarioAuditoria
'       oHerrExportaCitasWeb.Show 1
'       Set oHerrExportaCitasWeb = Nothing
'       Exit Sub
'
'    Case "ID_AlojadosporFechas"
'        Dim oRptAlojados As New RptHAlojados
'        oRptAlojados.idUsuario = ml_IdUsuarioAuditoria
'        oRptAlojados.EjecutaFormulario
'        Set oRptAlojados = Nothing
'        Exit Sub
'
'    Case "ID_ExportaURENIS"
'       Dim oHerrUrenis As New HerrExportaUrenis
'       oHerrUrenis.idUsuario = ml_IdUsuarioAuditoria
'       oHerrUrenis.lcNombrePc = lc_NombrePc
'       oHerrUrenis.Show 1
'       Set oHerrUrenis = Nothing
'       Exit Sub
'
'    'Reportes Laboratorio
'    Case "ID_LabAuditoria"
'      Dim orlabAuditoria As New rlabAuditoria
'      orlabAuditoria.idUsuario = ml_IdUsuarioAuditoria
'      orlabAuditoria.EjecutaFormulario
'      Set orlabAuditoria = Nothing
'      Exit Sub
'    Case "ID_LabPruebas"
'      Dim OrLabPruebas As New rLabPruebas
'      OrLabPruebas.idUsuario = ml_IdUsuarioAuditoria
'      OrLabPruebas.EjecutaFormulario
'      Set OrLabPruebas = Nothing
'      Exit Sub
'    Case "ID_LabProductividad"
'      Dim OrlRepProduccion As New SIGHProxies.rlRepProduccion
'      OrlRepProduccion.idUsuario = ml_IdUsuarioAuditoria
'      OrlRepProduccion.EjecutaFormulario
'      Set OrlRepProduccion = Nothing
'      Exit Sub
'    Case "ID_LabProductividadConsolidado" 'Adams
'      Dim OrlRepProducPagoDeuda1 As New rlRepProducPagoDeuda1
'      OrlRepProducPagoDeuda1.idUsuario = ml_IdUsuarioAuditoria
'      OrlRepProducPagoDeuda1.EjecutaFormulario
'      Set OrlRepProducPagoDeuda1 = Nothing
'      Exit Sub
'    Case "ID_LabProduccion"
'      Dim OrlRepProducPagoDeuda As New SIGHProxies.rlRepProducPagoDeuda
'      OrlRepProducPagoDeuda.idUsuario = ml_IdUsuarioAuditoria
'      OrlRepProducPagoDeuda.EjecutaFormulario
'      Set OrlRepProducPagoDeuda = Nothing
'      Exit Sub
'    Case "ID_LabTipoAnalisis"
'      Dim ORrlRepTipoAnalisis As New SIGHProxies.rlRepTipoAnalisis
'
'      ORrlRepTipoAnalisis.idUsuario = ml_IdUsuarioAuditoria
'      ORrlRepTipoAnalisis.EjecutaFormulario
'      Set ORrlRepTipoAnalisis = Nothing
'      Exit Sub
'    Case "ID_LabTipoAnalisisResultados"
'      Dim ORrlRepTipoAnalisisConRes As New rlRepTipoAnalisisConRes
'      ORrlRepTipoAnalisisConRes.idUsuario = ml_IdUsuarioAuditoria
'      ORrlRepTipoAnalisisConRes.EjecutaFormulario
'      Set ORrlRepTipoAnalisisConRes = Nothing
'      Exit Sub
'
'    '---Adams
'    Case "id_mn_CantidadesMortalidad"
'      Dim oRptMN_Cantidades As New SIGHReportes.RptMN_Cantidades
'      'oRptMN_Cantidades.idUsuario = ml_IdUsuarioAuditoria
'      oRptMN_Cantidades.EjecutaFormulario
'      Set oRptMN_Cantidades = Nothing
'      Exit Sub
'
'    '---Adams
'    End Select
'
'
'    '**********************************************************************
'    '   MANEJO DEL TOOLBAR DE GESTIÓN DE CAJA (se supone que este se activa cuando se selecciona la opción de gestión de caja
'    '**********************************************************************
'    Select Case Tool.Id
'    'MODULO DE CAJA
'    Case "ID_CajaApertura"
'        AperturaCaja
'        Exit Sub
'    Case "ID_CajaCierre"
'        CerrarCaja
'        Exit Sub
'    Case "ID_ParteDiario"
'       ' Dim RepPartDiario As New RpParteDiario
'       ' RepPartDiario.IdTipoReporte = 1
'       ' RepPartDiario.idUsuario = ml_IdUsuarioAuditoria
'       ' RepPartDiario.Show 1
'       ' Set RepPartDiario = Nothing
'        MsgBox "...Reporte en desarrollo..."
'        Exit Sub
'    Case "ID_ConsolidadoServ"
'        Dim RepServicio As New RpParteDiario
'        RepServicio.IdTipoReporte = 2
'        RepServicio.idUsuario = ml_IdUsuarioAuditoria
'        RepServicio.Show 1
'        Set RepServicio = Nothing
'        Exit Sub
'    Case "ID_ConsolidadoVentas" 'Adams
'        Dim RepConsolidadoVentas As New RpRegistroVentas
'        RepConsolidadoVentas.IdTipoReporte = 2
'        RepConsolidadoVentas.idUsuario = ml_IdUsuarioAuditoria
'        RepConsolidadoVentas.Show 1
'        Set RepConsolidadoVentas = Nothing
'        Exit Sub
'    Case "ID_ConsolFarmacia"
'        Dim RepFarmacia As New RpParteDiario
'        RepFarmacia.IdTipoReporte = 3
'        RepFarmacia.idUsuario = ml_IdUsuarioAuditoria
'        RepFarmacia.Show 1
'        Set RepFarmacia = Nothing
'        Exit Sub
'    Case "ID_ResumenCentroCosto"
'        If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3543 Or lcBuscaParametro.SeleccionaFilaParametro(8) = "0" Then
'            Dim RepResumCC As New RpParteDiario
'            RepResumCC.IdTipoReporte = 5
'            RepResumCC.idUsuario = ml_IdUsuarioAuditoria
'            RepResumCC.Show 1
'            Set RepResumCC = Nothing
'        Else
'            MsgBox "Este reporte solo lo puede usar el Hospital Regional Ayacucho" & Chr(13) & _
'                   "       use el reporte ECONOMIA -> TIPO TARIFA (CAJA)         ", vbInformation, "Mensaje"
'        End If
'        Exit Sub
'    Case "ID_DetalleporcadaCentroCosto"
'        If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 3543 Or lcBuscaParametro.SeleccionaFilaParametro(8) = "0" Then
'            Dim RepDetalleCC As New RpCajaDetalleCentroCosto
'            RepDetalleCC.idUsuario = ml_IdUsuarioAuditoria
'            RepDetalleCC.Show 1
'            Set RepDetalleCC = Nothing
'        Else
'            MsgBox "Este reporte solo lo puede usar el Hospital Regional Ayacucho" & Chr(13) & _
'                   "       use el reporte ECONOMIA -> TIPO TARIFA (CAJA)         ", vbInformation, "Mensaje"
'        End If
'        Exit Sub
'    End Select
'
'    '**********************************************************************
'    '   MANEJO DEL TOOLBAR DE PUNTO DIGITACIÓN HIS
'    '**********************************************************************
'    Select Case Tool.Id
'    Case "ID_DxOmitidos" 'HIS Digitacion - Frank08082014
'        Dim oRptHisDxOmitidos2 As New SIGHReportes.clRptHisDxOmitidos
'        oRptHisDxOmitidos2.EjecutaFormulario
'        Set oRptHisDxOmitidos2 = Nothing
'        Exit Sub
'    End Select
'
'
'    '**********************************************************************
'    'MANEJO DEL TOOLBAR DE EDICION (AGREGAR, MODIFICAR, CONSULTAR, ELIMINAR)
'    'DE ACUERDO AL MODULO SELECCIONADO
'    '**********************************************************************
'    Select Case ms_ModuloSeleccionado
'    'MODULO AMBULATORIO CE
'    Case "AdmisionCE"
'        EdicionCitas Tool.Id
'    Case "PacienteCE"
'        EdicionPaciente Tool.Id, sghConsultaExterna, 101
'    Case "AtencionesCE"
'        EdicionAdmisionCE Tool.Id, sghConsultaExterna, 103
'    Case "AtencionesTriaje"
'        EdicionTriaje Tool.Id       'debb-jamo
'    Case "RecetasCE"
'        EdicionReceta Tool.Id, 1366, sghConsultaExterna
'    Case "idConsultorioAsignado"
'        EdicionArchiveroServicio Tool.Id
'
'    'MODULO HIS-GALENOS JVG
'    Case "HisCE"
'        EdicionHisCE Tool.Id, 1346, ml_IdUsuarioAuditoria, lc_NombrePc
'    Case "HisPMMR"
'        EdicionProgramacionHIS Tool.Id, 1347, ml_IdUsuarioAuditoria, lc_NombrePc
'    Case "HisLoteCE"
'        EdicionHisLotesCE Tool.Id, 1348, ml_IdUsuarioAuditoria, lc_NombrePc
'    Case "HisREMR"
'        EdicionHisEstablecimientos Tool.Id, 1349, ml_IdUsuarioAuditoria, lc_NombrePc
'    Case "HisPN"
'        EdicionPadronNominal Tool.Id, 1353, ml_IdUsuarioAuditoria, lc_NombrePc
'    Case "HisCalidad"
'        EdicionHisDobleDigitacion Tool.Id, 1354, ml_IdUsuarioAuditoria, lc_NombrePc
''        Calidad Tool.ID, 1354, ml_IdUsuarioAuditoria, lc_NombrePc
'
'
'    'MODULO CONSULTORIOS EMERGENCIA
'    Case "PacienteEmerg", "PacienteObservacionEmerg"
'        EdicionPaciente Tool.Id, sghEmergenciaObservacion, 201
'    Case "AdmisionConsultorioEmerg"
'        EdicionAdmisionEmergencia Tool.Id
'    Case "CamasEmergencia"
'        EdicionCamas Tool.Id, True
'    Case "RecetasE"
'        EdicionReceta Tool.Id, 1343, sghEmergenciaConsultorios
'
'    'MODULO HOSPITALIZACION
'    Case "PacienteHosp"
'        EdicionPaciente Tool.Id, sghHospitalizacion, 301
'    Case "AdmisionHospitalizacion"
'        EdicionAdmisionHospitalizacion Tool.Id
'    Case "CamasHospitalizacion"
'        EdicionCamas Tool.Id, False
'    Case "AlojadosHospitalizacion"
'        EdicionAlojados Tool.Id
'    Case "RecetasH"
'        EdicionReceta Tool.Id, 1344, sghHospitalizacion
'
'    'MODULO DE PROGRAMACION
'    Case "Programacion"
'        EdicionProgMedica Tool.Id
'
'    Case "Turno"
'        EdicionTurno Tool.Id
'
'    Case "Medico"
'        EdicionMedico Tool.Id
'
'    'MODULO ARCHIVO CLINICO
'    Case "HistoriaClinica"
'        EdicionHistoriaClinica Tool.Id
'
'    Case "MovimientoHistoria"
'        EdicionMovimientoHistorias Tool.Id
'
'    Case "SolicitudHistorias"
'       'EdicionSolicitudHistorias Tool.ID
'
'    Case "Archivero"
'        EdicionArchiveroServicio Tool.Id
'
'    Case "MovFormatosHC"
'        EdicionMovimientoFormatoHC Tool.Id
'    'MODULO FACTURACION
'    Case "FacturacionGeneral"
'        EdicionOrdenesServicio Tool.Id
'
'    Case "FacturacionPatologiaClinica"
'       ' EdicionOrdenesServicioPatologiaClinica Tool.ID
'
'    Case "FacturacionAnatomiaPatologica"
'       ' EdicionOrdenesServicioAnatomiaPatologia Tool.ID
'
'    Case "FacturacionImagenologia"
'       ' EdicionOrdenesServicioImagenologia Tool.ID
'
'    Case "FacturacionSalaOperaciones"
'       ' EdicionOrdenesServicioSalaOperaciones Tool.ID
'
'    Case "Farmacia"
'      '  EdicionOrdenesServicioFarmacia Tool.ID
'
'    Case "FacturacionCatalogoServicios"
'
'       ' Select Case ucCatalogoServiciosLista1.IdTipoCatalogo
'        'Case 0
'            EdicionCatalogoBaseServicios Tool.Id
'        'Case Else
'        '    EdicionCatalogoServicios Tool.ID
'        'End Select
'
'    Case "FacturacionCentroCostos"
'        EdicionCentrosCosto Tool.Id
'    Case "PqteServicio"
'        EdicionPaqueteServicio Tool.Id
'    Case "FactReembolsos"
'        EdicionReembolsos Tool.Id
'    Case "PacExtConSeguro"
'        EdicionPacExtConSeguro Tool.Id
'    Case "PacExtParticular"
'        'EdicionPacExtParticular Tool.ID
'    'MODULO GENERAL
'    Case "Empleado"
'        EdicionEmpleado Tool.Id
'
'    Case "Servicios"
'        EdicionServicio Tool.Id
'
'    Case "Diagnosticos"
'        EdicionDiagnosticos Tool.Id
'
'    Case "Procedimientos"
'        'EdicionProcedimientos Tool.ID
'
'    Case "TiposFinanciamiento"
'        EdicionTiposFinanciamiento Tool.Id
'
'    Case "FuentesFinanciamiento"
'
'        EdicionFuentesFinanciamiento Tool.Id
'    Case "FacturacionPartidas"
'        EdicionPartidaPresupuestal Tool.Id
'
'    Case "EstablecimientosNoMinsa"
'        EdicionEstablecimientosNoMinsa Tool.Id
'
'    Case "Especialidades"
'        EdicionEspecialidades Tool.Id
'
'    Case "TipoTarifa"
'        EdicionTipoTarifa Tool.Id
'
'    'MODULO CAJA
'    Case "Cajas"
'        EdicionCaja Tool.Id
'
'    'FRANK MAYO
'    Case "NotaCredito"
'        EdicionCajaNotaCredito Tool.Id
'
'    Case "Cajeros"
'        'EdicionCajero Tool.ID
'
'    Case "CatalogoBienes"
'       ' Select Case ucCatalogoBienesInsumosLista1.IdTipoCatalogo
'       ' Case 0
'            EdicionCatalogoBaseBienesInsumos Tool.Id
'       ' Case Else
'       '     EdicionCatalogoBienesInsumos Tool.ID
'       ' End Select
'
'    'MODULO SEGURIDAD
'    Case "Roles"
'        EdicionRoles Tool.Id
'
'    'MODULO FARMACIA
'    Case "Inventario"
'        EdicionInventario Tool.Id
'    Case "NS", "NSF"                                                       '**debb2014
'        EdicionNS Tool.Id, IIf(ms_ModuloSeleccionado = "NS", False, True)  '**debb2014
'    Case "NI", "NIF", "FARMADOP"                                                       '**debb2014"
'        EdicionNI Tool.Id, IIf(ms_ModuloSeleccionado = "NI", False, True)  '**debb2014
'    Case "IntervencionS"
'        EdicionIntervencionS Tool.Id
'    Case "Ventas"
'        EdicionVentas Tool.Id
'    Case "DependenciaExt"
'        EdicionDependenciaExt Tool.Id
'    Case "DespachoDonaciones"
'        EdicionDespachoDonaciones Tool.Id
'    Case "FarmAlmacen"
'        EdicionMantenedorFarmacia Tool.Id
'    Case "FarmPrecios"                                     'debb2014b
'        EdicionMantenedorHistoricoPrecios Tool.Id          'debb2014b
'
'    'MODULO IMAGENEOLOGIA
'    Case "ImagRayosX"
'        EdicionRayosX Tool.Id
'    Case "ImagEcografiaG"
'        EdicionImagEcografiaGen Tool.Id
'    Case "ImagTomografia"
'        EdicionImagTomografia Tool.Id
'    Case "ImagEcografiaO"
'        EdicionImagEcografiaObs Tool.Id
'    Case "ImagIngresos"
'        EdicionImagIngresos Tool.Id
'    Case "ImagSalidas"
'        EdicionImagSalidas Tool.Id
'    'mgaray201411f
'    Case "ImagTipoModalidadSala"
'        EdicionTipoModalidadSala Tool.Id
'    Case "ImagSala"
'        EdicionSala Tool.Id
'    Case "ImagCatalgoServicioDuracion"
'        EdicionImagFactCatalogoServiciosDuracion Tool.Id
'    Case "IntegracionSistema"
'        EdicionIntegracionSistema Tool.Id
'
'    'MODULO LABORATORIO
'    Case "OrdenesLaboratorio"
'        EdicionLaboratorio Tool.Id
'    Case "OrdenesPatologia"
'        EdicionOrdenesServicioAnatomiaPatologia_ Tool.Id
'    Case "BS"
'        EdicionOrdenesBS_ Tool.Id
'    Case "ResultadosLaboratorio"
'        EdicionResultados Tool.Id
'    Case "MuestrasExamenes"
'        EdicionMuestras Tool.Id
'    Case "LabIngresos"
'        EdicionLabIngresos Tool.Id
'    Case "LabEgresos"
'        EdicionLabSalidas Tool.Id
'
'    'Constancias de Atención
'    Case "Constancias"
'      EdicionConstancias Tool.Id
'    'Sis
'    Case "Fua"
'      EdicionFua Tool.Id
'
'    Case "ConfiguracionResLab" ' modificacion samuel
'        EdicionConfiguracionResLab Tool.Id
'
'    End Select
'
'    Set lcBuscaParametro = Nothing
'End Sub
'
'Sub CierreCtaAtencion()
''        Dim oCierreCtas As New CierreCtaAtencion
''        oCierreCtas.IdUsuario = ml_IdUsuarioAuditoria
''        oCierreCtas.Show 1
''        Unload oCierreCtas
'
'End Sub
'
Sub EdicionConfiguracionResLab(sToolId As String) 'nuevo Samuel
    Dim oConfiguracionReslab As New SIGHLaboratorio.clConfiguarcionResLab

        Select Case sToolId
        Case "ID_Agregar":
           oConfiguracionReslab.Opcion = sghAgregar
        Case "ID_Modificar":
           oConfiguracionReslab.Opcion = sghModificar
           oConfiguracionReslab.idProducto = ucConfiguraResLab2.idRegistroSeleccionado
        Case "ID_Consultar":
           oConfiguracionReslab.Opcion = sghConsultar
           oConfiguracionReslab.idProducto = ucConfiguraResLab2.idRegistroSeleccionado
        Case "ID_Eliminar":
           oConfiguracionReslab.Opcion = sghEliminar
           oConfiguracionReslab.idProducto = ucConfiguraResLab2.idRegistroSeleccionado
        End Select
       oConfiguracionReslab.idUsuario = ml_IdUsuarioAuditoria
       oConfiguracionReslab.lcNombrePc = lc_NombrePc
       oConfiguracionReslab.lnIdTablaLISTBARITEMS = 1303
       oConfiguracionReslab.MostrarFormulario
       Set oConfiguracionReslab = Nothing

End Sub
'
'
''debb-jamo
Sub EdicionTriaje(sToolId As String)
Dim oTriaje As New SIGHCatalogos.clTriaje
    Dim oRs As New ADODB.Recordset

        Select Case sToolId
        Case "ID_Agregar":
           oTriaje.Opcion = sghAgregar
        Case "ID_Modificar":
           oTriaje.Opcion = sghModificar
           oTriaje.idAtencion = ucAtencionesTriaje1.idRegistroSeleccionado

           Set oRs = Me.ucAtencionesTriaje1.DataSource
            If oRs Is Nothing Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If oRs.State = 0 Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If oRs.RecordCount = 0 Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
        Case "ID_Consultar":
           oTriaje.Opcion = sghConsultar
           oTriaje.idAtencion = ucAtencionesTriaje1.idRegistroSeleccionado
           Set oRs = Me.ucAtencionesTriaje1.DataSource
            If oRs Is Nothing Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If oRs.State = 0 Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If oRs.RecordCount = 0 Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
        Case "ID_Eliminar":
           oTriaje.Opcion = sghEliminar
           oTriaje.idAtencion = ucAtencionesTriaje1.idRegistroSeleccionado
           Set oRs = Me.ucAtencionesTriaje1.DataSource
            If oRs Is Nothing Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If oRs.State = 0 Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If oRs.RecordCount = 0 Then
                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
       oTriaje.idUsuario = ml_IdUsuarioAuditoria
       oTriaje.lcNombrePc = lc_NombrePc
       oTriaje.lnIdTablaLISTBARITEMS = 1303
       oTriaje.MostrarFormulario
       If oTriaje.GuardoTriaje Then ucAtencionesTriaje1.RealizarBusqueda
       Set oTriaje = Nothing
End Sub

'''*******************************INO*************************************
''Sub EdicionTriajeOftalmologico(sToolId As String)
''Dim oTriajeOftalmologico As New SIGHCatalogos.clTriajeOftalomologico
''
''    Dim oRs As New ADODB.Recordset
''
''        Select Case sToolId
''        Case "ID_Agregar":
''           oTriajeOftalmologico.Opcion = sghAgregar
''        Case "ID_Modificar":
''           oTriajeOftalmologico.Opcion = sghModificar
''           oTriajeOftalmologico.idAtencion = ucAtencionesTriajeOftalmologico1.idRegistroSeleccionado
''
''           Set oRs = Me.ucAtencionesTriajeOftalmologico1.DataSource
''            If oRs Is Nothing Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            If oRs.State = 0 Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            If oRs.RecordCount = 0 Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''        Case "ID_Consultar":
''           oTriajeOftalmologico.Opcion = sghConsultar
''           oTriajeOftalmologico.idAtencion = ucAtencionesTriajeOftalmologico1.idRegistroSeleccionado
''           Set oRs = Me.ucAtencionesTriajeOftalmologico1.DataSource
''            If oRs Is Nothing Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            If oRs.State = 0 Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            If oRs.RecordCount = 0 Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''        Case "ID_Eliminar":
''           oTriajeOftalmologico.Opcion = sghEliminar
''           oTriajeOftalmologico.idAtencion = ucAtencionesTriajeOftalmologico1.idRegistroSeleccionado
''           Set oRs = Me.ucAtencionesTriajeOftalmologico1.DataSource
''            If oRs Is Nothing Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            If oRs.State = 0 Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            If oRs.RecordCount = 0 Then
''                MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''        End Select
''       oTriajeOftalmologico.idUsuario = ml_IdUsuarioAuditoria
''       oTriajeOftalmologico.lcNombrePc = lc_NombrePc
''       oTriajeOftalmologico.lnIdTablaLISTBARITEMS = 1303
''       oTriajeOftalmologico.MostrarFormulario
''       Set oTriajeOftalmologico = Nothing
''       ucAtencionesTriajeOftalmologico1.RealizarBusqueda
''End Sub
'''*******************************INO*************************************
'
'
Function SeleccionarOpcion(sToolId As String) As sghOpciones

        Select Case sToolId
        Case "ID_Agregar":
            SeleccionarOpcion = sghAgregar
        Case "ID_Modificar":
            SeleccionarOpcion = sghModificar
        Case "ID_Consultar":
            SeleccionarOpcion = sghConsultar
        Case "ID_Eliminar":
            SeleccionarOpcion = sghEliminar
        End Select

End Function
'
Sub EdicionTurno(sToolId As String)
Dim mo_TurnoDetalle As New SIGHCatalogos.clTurnoDetalle

        mo_TurnoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_TurnoDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_TurnoDetalle.lnIdTablaLISTBARITEMS = 402
        mo_TurnoDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_TurnoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_TurnoDetalle.IdTurno = Me.ucTurnosLista1.idRegistroSeleccionado
            If mo_TurnoDetalle.IdTurno = -1 Or mo_TurnoDetalle.IdTurno = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_TurnoDetalle.MostrarFormulario
        Set mo_TurnoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

        Set ucTurnosLista1.DataSource = mo_AdminProgramacionMedica.TurnosSeleccionarTodos()


End Sub
''
Sub EdicionEmpleado(sToolId As String)
Dim mo_EmpleadoDetalle As New SIGHCatalogos.clEmpleadoDetalle

        mo_EmpleadoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_EmpleadoDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_EmpleadoDetalle.lnIdTablaLISTBARITEMS = 1301
        mo_EmpleadoDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_EmpleadoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_EmpleadoDetalle.IdEmpleado = Me.ucEmpleadosLista1.idRegistroSeleccionado
            If mo_EmpleadoDetalle.IdEmpleado = -1 Or mo_EmpleadoDetalle.IdEmpleado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_EmpleadoDetalle.MostrarFormulario
        Set mo_EmpleadoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
''
Sub EdicionServicio(sToolId As String)
Dim mo_ServicioDetalle As New SIGHProxies.clServicioDetalle


        mo_ServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_ServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_ServicioDetalle.lnIdTablaLISTBARITEMS = 1201
        mo_ServicioDetalle.lcNombrePc = lc_NombrePc
        If ucServiciosLista1.idTipoServicio = 0 Then
            MsgBox "Por favor seleccione el tipo de servicio", vbInformation, Me.Caption
            Exit Sub
        End If

        mo_ServicioDetalle.idTipoServicio = ucServiciosLista1.idTipoServicio

        Select Case mo_ServicioDetalle.Opcion
        Case sghAgregar

        Case sghModificar, sghConsultar, sghEliminar
            mo_ServicioDetalle.IdServicio = Me.ucServiciosLista1.idRegistroSeleccionado
            If mo_ServicioDetalle.IdServicio = -1 Or mo_ServicioDetalle.IdServicio = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_ServicioDetalle.MostrarFormulario
        Set mo_ServicioDetalle = Nothing

        Me.ucServiciosLista1.ActualizarGrilla
        Me.ucServiciosLista1.ActualizarJerarquia

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
Sub EdicionEspecialidades(sToolId As String)
Dim mo_EspecialidadDetalle As New SIGHCatalogos.clEspecialidadDetalle

        mo_EspecialidadDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_EspecialidadDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_EspecialidadDetalle.lnIdTablaLISTBARITEMS = 1206
        mo_EspecialidadDetalle.lcNombrePc = lc_NombrePc

        Select Case mo_EspecialidadDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_EspecialidadDetalle.IdEspecialidad = Me.ucEspecialidadesLista1.idRegistroSeleccionado
            If mo_EspecialidadDetalle.IdEspecialidad = -1 Or mo_EspecialidadDetalle.IdEspecialidad = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_EspecialidadDetalle.MostrarFormulario
        Set mo_EspecialidadDetalle = Nothing

        Me.ucEspecialidadesLista1.ActualizarGrilla

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select


End Sub
''
Sub EdicionMedico(sToolId As String)
Dim mo_MedicoDetalle As New SIGHCatalogos.clMedicoDetalle

        mo_MedicoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_MedicoDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_MedicoDetalle.lnIdTablaLISTBARITEMS = 403
        mo_MedicoDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_MedicoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_MedicoDetalle.idMedico = Me.ucMedicosLista1.idRegistroSeleccionado
            If mo_MedicoDetalle.idMedico = -1 Or mo_MedicoDetalle.idMedico = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_MedicoDetalle.MostrarFormulario
        Set mo_MedicoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
            'Set ucMedicosLista1.DataSource = mo_AdminProgramacionMedica.MedicosSeleccionarTodos()
        Case "ID_Modificar":
            'Set ucMedicosLista1.DataSource = mo_AdminProgramacionMedica.MedicosSeleccionarTodos()
        Case "ID_Consultar":
        Case "ID_Eliminar":
            'Set ucMedicosLista1.DataSource = mo_AdminProgramacionMedica.MedicosSeleccionarTodos()
        End Select

End Sub
Sub EdicionAdmisionCE(sToolId As String, lTipoServicio As sghTipoServicio, lnIdTablaLISTBARITEMS As Long)

        mo_AdmisionCEDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_AdmisionCEDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_AdmisionCEDetalle.TipoVistaForm = sghVistaAtencion

        Select Case sToolId
        Case "ID_Agregar", "ID_Modificar", "ID_Consultar", "ID_Eliminar"
            Select Case mo_AdmisionCEDetalle.Opcion
            Case sghAgregar

            Case sghModificar, sghConsultar, sghEliminar
                Dim oRs As ADODB.Recordset

                Set oRs = Me.ucAdmisionCE.DataSource

                If oRs Is Nothing Then
                    MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                    Exit Sub
                End If
                If oRs.RecordCount = 0 Then
                    MsgBox "Seleccione un Registro", vbInformation, Me.Caption
                    Exit Sub
                End If
                mo_AdmisionCEDetalle.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
                mo_AdmisionCEDetalle.lcNombrePc = lc_NombrePc
                mo_AdmisionCEDetalle.IdCita = Me.ucAdmisionCE.idRegistroSeleccionado
                If Me.ucAdmisionCE.NoPagoConsultaEnCaja = True Then
                    MsgBox "El Paciente es Pagante, no pagó por Consulta en CAJA", vbInformation, Me.Caption
                    Exit Sub
                ElseIf Me.ucAdmisionCE.NoPasoPorTriaje = True Then
                    MsgBox "El Paciente no pasó por Triaje" & Chr(13) & Chr(13) & "Este Consultorio se configuró para que los Pacientes pasen por Triaje antes de su Atención", vbInformation, Me.Caption
                    Exit Sub
                ElseIf mo_AdmisionCEDetalle.IdCita = -1 Or mo_AdmisionCEDetalle.IdCita = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
            End Select
        Case "ID_Exonerar"

            Exit Sub
        Case "ID_PendientePago"

            Exit Sub
        Case "ID_EstadoCuenta"

            Exit Sub
        End Select

        If Me.ucAdmisionCE.HoraEgreso = "" Then 'Actualizado 21102014
            mo_AdmisionCEDetalle.OcultarBotonesImpresionReceta False
        Else
            mo_AdmisionCEDetalle.OcultarBotonesImpresionReceta True
        End If

        mo_AdmisionCEDetalle.Icon = Me.Icon
        mo_AdmisionCEDetalle.lbNuevoMovimiento = True
        mo_AdmisionCEDetalle.Show 1

        Select Case sToolId
        Case "ID_Agregar"

        Case "ID_Modificar"
              If lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then  'CE: Registro DX
                 Me.ucAdmisionCE.RealizarBusqueda False
                 Me.ucAdmisionCE.FocusEnGrilla
              End If
        Case "ID_Consultar"
        Case "ID_Eliminar"

        End Select

End Sub
''
''
''
Sub EdicionHistoriaClinica(sToolId As String)
Dim mo_HistoriaClinicaDetalle As New HistoriaClinicaDetalle

        mo_HistoriaClinicaDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_HistoriaClinicaDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_HistoriaClinicaDetalle.lnIdTablaLISTBARITEMS = 501
        mo_HistoriaClinicaDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_HistoriaClinicaDetalle.Opcion
        Case sghAgregar

        Case sghModificar, sghConsultar, sghEliminar
            mo_HistoriaClinicaDetalle.IdHistoriaClinica = Me.ucHistoriaClinicaLista1.idRegistroSeleccionado
            If mo_HistoriaClinicaDetalle.IdHistoriaClinica = -1 Or mo_HistoriaClinicaDetalle.IdHistoriaClinica = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_HistoriaClinicaDetalle.Icon = Me.Icon
        mo_HistoriaClinicaDetalle.Show 1
        Unload mo_HistoriaClinicaDetalle

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
''
Sub EdicionMovimientoHistorias(sToolId As String)
Dim mo_MovimientoHistoriaDetalle As New MovimientoHistoriaDetalle

        mo_MovimientoHistoriaDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_MovimientoHistoriaDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_MovimientoHistoriaDetalle.lnIdTablaLISTBARITEMS = 502
        mo_MovimientoHistoriaDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_MovimientoHistoriaDetalle.Opcion
        Case sghAgregar

        Case sghModificar, sghConsultar, sghEliminar
            mo_MovimientoHistoriaDetalle.IdMovimiento = Me.ucMovimientoHistoriasLista1.idRegistroSeleccionado
            mo_MovimientoHistoriaDetalle.idPacienteSeleccionado = Me.ucMovimientoHistoriasLista1.idPacienteSeleccionado
            If mo_MovimientoHistoriaDetalle.IdMovimiento = -1 Or mo_MovimientoHistoriaDetalle.IdMovimiento = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_MovimientoHistoriaDetalle.Icon = Me.Icon
        mo_MovimientoHistoriaDetalle.Show 1
        Unload mo_MovimientoHistoriaDetalle
        ucMovimientoHistoriasLista1.RealizarBusqueda
        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
''
Sub EdicionSolicitudHistorias(sToolId As String)
Dim mo_SolicitudHistoriaDetalle As New SolicitudHistoriaDetalle

        mo_SolicitudHistoriaDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_SolicitudHistoriaDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_SolicitudHistoriaDetalle.lnIdTablaLISTBARITEMS = 503
        mo_SolicitudHistoriaDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_SolicitudHistoriaDetalle.Opcion
        Case sghAgregar

        Case sghModificar, sghConsultar, sghEliminar
            mo_SolicitudHistoriaDetalle.IdHistoriaSolicitada = Me.ucSolicitudHistoriasLista1.idRegistroSeleccionado
            If mo_SolicitudHistoriaDetalle.IdHistoriaSolicitada = -1 Or mo_SolicitudHistoriaDetalle.IdHistoriaSolicitada = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_SolicitudHistoriaDetalle.Icon = Me.Icon
        mo_SolicitudHistoriaDetalle.Show 1

        Unload mo_SolicitudHistoriaDetalle

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
Sub EdicionArchiveroServicio(sToolId As String)
Dim mo_ArchiveroServicio As New SIGHCatalogos.clArchiveroServicioD
        mo_ArchiveroServicio.EsConsultorioAsignado = LbEsConsultorioAsignado
        mo_ArchiveroServicio.Opcion = SeleccionarOpcion(sToolId)
        mo_ArchiveroServicio.idUsuario = ml_IdUsuarioAuditoria
        mo_ArchiveroServicio.lcNombrePc = lc_NombrePc
        mo_ArchiveroServicio.lnIdTablaLISTBARITEMS = 504
        Select Case mo_ArchiveroServicio.Opcion
        Case sghAgregar

        Case sghModificar, sghConsultar, sghEliminar
            mo_ArchiveroServicio.IdEmpleado = Me.ucArchivadoresLista1.idRegistroSeleccionado
            If mo_ArchiveroServicio.IdEmpleado = -1 Or mo_ArchiveroServicio.IdEmpleado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_ArchiveroServicio.MostrarFormulario
        Set mo_ArchiveroServicio = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
''
Sub EdicionAdmisionEmergencia(sToolId As String)
Dim rsEmergencia As New Recordset

        '-----------AGREGADO FRANKLIN CACHAY 0403
        If sToolId = "ID_HospitalizacionVisitaEnfermera" Then
           If ucAdmisionConsEmerg.DataSource Is Nothing Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           Set rsEmergencia = ucAdmisionConsEmerg.DataSource

           If rsEmergencia.RecordCount = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If

           If (rsEmergencia!idAtencion = -1 Or rsEmergencia!idAtencion = 0) Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           mo_VisitasEnfermeras.Opcion = sghModificar
           mo_VisitasEnfermeras.idCuentaAtencion = rsEmergencia!idCuentaAtencion
           mo_VisitasEnfermeras.TipoServicio = rsEmergencia!idTipoServicio
           mo_VisitasEnfermeras.lcNombrePc = lc_NombrePc
           mo_VisitasEnfermeras.lnIdTablaLISTBARITEMS = 302
           mo_VisitasEnfermeras.lbNuevoMovimiento = True
           mo_VisitasEnfermeras.CargaUnaSolaVez = True
           mo_VisitasEnfermeras.idUsuario = ml_IdUsuarioAuditoria
           mo_VisitasEnfermeras.Show 1
           Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If sToolId = "ID_HospitalizacionAlojamientoConjunto" Then
            If ucAdmisionConsEmerg.idRegistroSeleccionado = 0 Then
                Exit Sub
            End If

            mo_AdmisionHospEgreso.idUsuario = ml_IdUsuarioAuditoria
            mo_AdmisionHospEgreso.Opcion = sghModificar
            mo_AdmisionHospEgreso.TipoAccionDeAdmision = sghAdmisionNormal
            Set rsEmergencia = ucAdmisionConsEmerg.DataSource
            If rsEmergencia.State = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If rsEmergencia.RecordCount = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
            mo_AdmisionHospEgreso.idAtencion = rsEmergencia!idAtencion
            mo_AdmisionHospEgreso.idCuentaAtencion = rsEmergencia!idCuentaAtencion
            mo_AdmisionHospEgreso.TipoServicio = rsEmergencia!idTipoServicio
            If mo_AdmisionHospEgreso.idAtencion = -1 Or mo_AdmisionHospEgreso.idAtencion = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If

            mo_AdmisionHospEgreso.lcNombrePc = lc_NombrePc
            mo_AdmisionHospEgreso.lnIdTablaLISTBARITEMS = 202
            mo_AdmisionHospEgreso.lbNuevoMovimiento = True
            mo_AdmisionHospEgreso.Show 1
        Else
            mo_AdmisionHospDetalle.idUsuario = ml_IdUsuarioAuditoria
            Select Case sToolId
            Case "ID_Agregar", "ID_Modificar", "ID_Consultar", "ID_Eliminar"
                mo_AdmisionHospDetalle.Opcion = SeleccionarOpcion(sToolId)
                mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghAdmisionNormal
                Select Case mo_AdmisionHospDetalle.Opcion
                Case sghAgregar
                    mo_AdmisionHospDetalle.TipoServicio = sghEmergenciaConsultorios
                Case sghModificar, sghConsultar, sghEliminar
                    Set rsEmergencia = ucAdmisionConsEmerg.DataSource

                    If ucAdmisionConsEmerg.idRegistroSeleccionado = 0 Then
                        Exit Sub
                    End If
                    If rsEmergencia.State = 0 Then
                        MsgBox "Seleccione un registro", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    If rsEmergencia.RecordCount = 0 Then
                        MsgBox "Seleccione un registro", vbInformation, Me.Caption
                        Exit Sub
                    End If

                    mo_AdmisionHospDetalle.idAtencion = rsEmergencia!idAtencion
                    mo_AdmisionHospDetalle.idCuentaAtencion = rsEmergencia!idCuentaAtencion

                    mo_AdmisionHospDetalle.TipoServicio = rsEmergencia!idTipoServicio
                    If mo_AdmisionHospDetalle.idAtencion = -1 Or mo_AdmisionHospDetalle.idAtencion = 0 Then
                        MsgBox "Seleccione un registro", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End Select
            Case "ID_EmergenciaAObservacion"

                mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghEnviarAObservacion
                Set rsEmergencia = ucAdmisionConsEmerg.DataSource
                mo_AdmisionHospDetalle.idPaciente = rsEmergencia!idPaciente
                mo_AdmisionHospDetalle.idCuentaAtencion = rsEmergencia!idCuentaAtencion
                mo_AdmisionHospDetalle.TipoServicio = sghEmergenciaObservacion
                mo_AdmisionHospDetalle.IdAtencionPadre = rsEmergencia!idAtencion
                mo_AdmisionHospDetalle.Opcion = sghAgregar

            Case "ID_EmergenciaAHospitalizacion"

                mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghTrasladarAHospitalizacion
                Set rsEmergencia = ucAdmisionConsEmerg.DataSource
                mo_AdmisionHospDetalle.idPaciente = rsEmergencia!idPaciente
                mo_AdmisionHospDetalle.idCuentaAtencion = rsEmergencia!idCuentaAtencion
                mo_AdmisionHospDetalle.TipoServicio = sghHospitalizacion
                mo_AdmisionHospDetalle.IdAtencionPadre = rsEmergencia!idAtencion
                mo_AdmisionHospDetalle.Opcion = sghAgregar

            Case "ID_EmergenciaAltaPaciente"

                    mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghDarDeAlta
                    Set rsEmergencia = ucAdmisionConsEmerg.DataSource
                    mo_AdmisionHospDetalle.idPaciente = rsEmergencia!idPaciente
                    mo_AdmisionHospDetalle.idCuentaAtencion = rsEmergencia!idCuentaAtencion
                    mo_AdmisionHospDetalle.idAtencion = rsEmergencia!idAtencion
                    mo_AdmisionHospDetalle.TipoServicio = rsEmergencia!idTipoServicio
                    mo_AdmisionHospDetalle.Opcion = sghModificar
                    If mo_AdmisionHospDetalle.idAtencion = -1 Or mo_AdmisionHospDetalle.idAtencion = 0 Then
                        MsgBox "Seleccione un registro", vbInformation, Me.Caption
                        Exit Sub
                    End If
            Case "ID_EmergenciaTransferencias"

                    mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghTransferencias
                    Set rsEmergencia = ucAdmisionConsEmerg.DataSource
                    mo_AdmisionHospDetalle.idPaciente = rsEmergencia!idPaciente
                    mo_AdmisionHospDetalle.idCuentaAtencion = rsEmergencia!idCuentaAtencion
                    mo_AdmisionHospDetalle.idAtencion = rsEmergencia!idAtencion
                    mo_AdmisionHospDetalle.TipoServicio = rsEmergencia!idTipoServicio
                    mo_AdmisionHospDetalle.Opcion = sghModificar
                    If mo_AdmisionHospDetalle.idAtencion = -1 Or mo_AdmisionHospDetalle.idAtencion = 0 Then
                        MsgBox "Seleccione un registro", vbInformation, Me.Caption
                        Exit Sub
                    End If
            Case "ID_Exonerar"

                Exit Sub
            Case "ID_PendientePago"

                Exit Sub
            Case "ID_EstadoCuenta"

                Exit Sub

            End Select

            mo_AdmisionHospDetalle.lcNombrePc = lc_NombrePc
            mo_AdmisionHospDetalle.lnIdTablaLISTBARITEMS = 202
            mo_AdmisionHospDetalle.lbNuevoMovimiento = True
            mo_AdmisionHospDetalle.Show 1
      End If

End Sub
''
Sub EdicionAdmisionHospitalizacion(sToolId As String)
Dim rsHospitalizacion As New Recordset
        If sToolId = "ID_HospitalizacionAlojamientoConjunto" Then

           If ucAdmisionHospitalizacion.DataSource Is Nothing Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
           If rsHospitalizacion.State = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           If rsHospitalizacion.RecordCount = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           If (rsHospitalizacion!idAtencion = -1 Or rsHospitalizacion!idAtencion = 0) Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           mo_AdmisionHospEgreso.TipoAccionDeAdmision = sghAdmisionNormal
           mo_AdmisionHospEgreso.Opcion = sghModificar
           mo_AdmisionHospEgreso.idCuentaAtencion = IIf(IsNull(rsHospitalizacion!idCuentaAtencion), 0, rsHospitalizacion!idCuentaAtencion)
           mo_AdmisionHospEgreso.idAtencion = rsHospitalizacion!idAtencion
           mo_AdmisionHospEgreso.TipoServicio = sghHospitalizacion
           mo_AdmisionHospEgreso.lcNombrePc = lc_NombrePc
           mo_AdmisionHospEgreso.lnIdTablaLISTBARITEMS = 302
           mo_AdmisionHospEgreso.lbNuevoMovimiento = True
           mo_AdmisionHospEgreso.idUsuario = ml_IdUsuarioAuditoria
           mo_AdmisionHospEgreso.Show 1
           Exit Sub
        End If

        '-----------AGREGADO FRANKLIN CACHAY 0403
        If sToolId = "ID_HospitalizacionVisitaEnfermera" Then
           If ucAdmisionHospitalizacion.DataSource Is Nothing Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
           If rsHospitalizacion.State = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           If rsHospitalizacion.RecordCount = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           If (rsHospitalizacion!idAtencion = -1 Or rsHospitalizacion!idAtencion = 0) Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
           End If
           mo_VisitasEnfermeras.Opcion = sghModificar
           mo_VisitasEnfermeras.idCuentaAtencion = IIf(IsNull(rsHospitalizacion!idCuentaAtencion), 0, rsHospitalizacion!idCuentaAtencion)
           mo_VisitasEnfermeras.TipoServicio = sghHospitalizacion
           mo_VisitasEnfermeras.lcNombrePc = lc_NombrePc
           mo_VisitasEnfermeras.lnIdTablaLISTBARITEMS = 302
           mo_VisitasEnfermeras.lbNuevoMovimiento = True
           mo_VisitasEnfermeras.CargaUnaSolaVez = True
           mo_VisitasEnfermeras.idUsuario = ml_IdUsuarioAuditoria
           mo_VisitasEnfermeras.Show 1
           Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        mo_AdmisionHospDetalle.idUsuario = ml_IdUsuarioAuditoria

        Select Case sToolId
        Case "ID_Agregar", "ID_Modificar", "ID_Consultar", "ID_Eliminar"

            mo_AdmisionHospDetalle.Opcion = SeleccionarOpcion(sToolId)
            mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghAdmisionNormal
            Select Case mo_AdmisionHospDetalle.Opcion
            Case sghAgregar
                mo_AdmisionHospDetalle.TipoServicio = sghHospitalizacion
                mo_AdmisionHospDetalle.IdServicioConCamaDisponible = ucAdmisionHospitalizacion.IdServicioConCamaDisponible
            Case sghModificar, sghConsultar, sghEliminar

                Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
                If ucAdmisionHospitalizacion.idRegistroSeleccionado = 0 Then
                    Exit Sub
                End If
                If rsHospitalizacion.State = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
                If rsHospitalizacion.RecordCount = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
                If rsHospitalizacion!idAtencion = -1 Or rsHospitalizacion!idAtencion = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
                mo_AdmisionHospDetalle.idCuentaAtencion = IIf(IsNull(rsHospitalizacion!idCuentaAtencion), 0, rsHospitalizacion!idCuentaAtencion)
                mo_AdmisionHospDetalle.idAtencion = rsHospitalizacion!idAtencion
                mo_AdmisionHospDetalle.TipoServicio = sghHospitalizacion
            End Select


        Case "ID_HospitalizacionAlojamientoConjunto"    'Crea una hospitalizacion adicional de tipo AC
            mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghIngresarUnAlojamientoConjunto
            Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
            mo_AdmisionHospDetalle.idPaciente = rsHospitalizacion!idPaciente
            mo_AdmisionHospDetalle.idCuentaAtencion = rsHospitalizacion!idCuentaAtencion
            mo_AdmisionHospDetalle.IdAtencionPadre = rsHospitalizacion!idAtencion
            mo_AdmisionHospDetalle.TipoServicio = sghHospitalizacion
            mo_AdmisionHospDetalle.Opcion = sghAgregar

        Case "ID_HospitalizacionAltaPaciente"           'Modifica los datos de hosp, enfocandose en el alta del paciente
                mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghDarDeAlta
                Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
                mo_AdmisionHospDetalle.idPaciente = rsHospitalizacion!idPaciente
                mo_AdmisionHospDetalle.idCuentaAtencion = rsHospitalizacion!idCuentaAtencion
                mo_AdmisionHospDetalle.idAtencion = rsHospitalizacion!idAtencion
                mo_AdmisionHospDetalle.TipoServicio = sghHospitalizacion
                mo_AdmisionHospDetalle.Opcion = sghModificar
                If mo_AdmisionHospDetalle.idAtencion = -1 Or mo_AdmisionHospDetalle.idAtencion = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If

        Case "ID_HospitalizacionTransferencias"         'Modifica los datos del paciente, enfocandose en las transferencias
                mo_AdmisionHospDetalle.TipoAccionDeAdmision = sghTransferencias
                Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
                mo_AdmisionHospDetalle.idPaciente = rsHospitalizacion!idPaciente
                mo_AdmisionHospDetalle.idCuentaAtencion = rsHospitalizacion!idCuentaAtencion
                mo_AdmisionHospDetalle.idAtencion = rsHospitalizacion!idAtencion
                mo_AdmisionHospDetalle.TipoServicio = sghHospitalizacion
                mo_AdmisionHospDetalle.Opcion = sghModificar
                If mo_AdmisionHospDetalle.idAtencion = -1 Or mo_AdmisionHospDetalle.idAtencion = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
        Case "ID_Exonerar"

            Exit Sub
        Case "ID_PendientePago"

            Exit Sub
        Case "ID_EstadoCuenta"

            Exit Sub
        End Select
        mo_AdmisionHospDetalle.lcNombrePc = lc_NombrePc
        mo_AdmisionHospDetalle.lnIdTablaLISTBARITEMS = 302
        'mo_AdmisionHospDetalle.Icon = Me.Icon  'tener en cuenta que esto automaticamente hace un Load del form
        mo_AdmisionHospDetalle.lbNuevoMovimiento = True
        mo_AdmisionHospDetalle.Show 1

End Sub
''
''
''
''Sub EdicionPreLiquidacion(sToolId As String)
''
''
''End Sub
Sub EdicionDiagnosticos(sToolId As String)
Dim mo_DiagnosticoDetalle As New SIGHCatalogos.clDiagnosticoDetalle

        mo_DiagnosticoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_DiagnosticoDetalle.idUsuario = ml_IdUsuarioAuditoria

        Select Case mo_DiagnosticoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_DiagnosticoDetalle.idDiagnostico = Me.ucDiagnosticosLista1.idRegistroSeleccionado
            If mo_DiagnosticoDetalle.idDiagnostico = -1 Or mo_DiagnosticoDetalle.idDiagnostico = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_DiagnosticoDetalle.MostrarFormulario
        Set mo_DiagnosticoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
Sub EdicionTiposFinanciamiento(sToolId As String)
Dim mo_TipoFinanciamientoDetalle As New SIGHCatalogos.clTipoFinanciamDetalle

        mo_TipoFinanciamientoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_TipoFinanciamientoDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_TipoFinanciamientoDetalle.lnIdTablaLISTBARITEMS = 611
        mo_TipoFinanciamientoDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_TipoFinanciamientoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_TipoFinanciamientoDetalle.idTipoFinanciamiento = Me.ucTiposFinanciamientoLista1.idRegistroSeleccionado
            If mo_TipoFinanciamientoDetalle.idTipoFinanciamiento = -1 Or mo_TipoFinanciamientoDetalle.idTipoFinanciamiento = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_TipoFinanciamientoDetalle.MostrarFormulario
        Set mo_TipoFinanciamientoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
''
Sub EdicionFuentesFinanciamiento(sToolId As String)

        mo_FuenteFinanciamientoDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_FuenteFinanciamientoDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_FuenteFinanciamientoDetalle.lcNombrePc = lc_NombrePc
        mo_FuenteFinanciamientoDetalle.lnIdTablaLISTBARITEMS = 1311
        Select Case mo_FuenteFinanciamientoDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_FuenteFinanciamientoDetalle.IdFuenteFinanciamiento = Me.ucFuentesFinanciamientoLista1.idRegistroSeleccionado
            If mo_FuenteFinanciamientoDetalle.IdFuenteFinanciamiento = -1 Or mo_FuenteFinanciamientoDetalle.IdFuenteFinanciamiento = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

       mo_FuenteFinanciamientoDetalle.MostrarFormulario
       Set mo_FuenteFinanciamientoDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select


End Sub
Sub EdicionPartidaPresupuestal(sToolId As String)

        mo_PartidasDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_PartidasDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_PartidasDetalle.lnIdTablaLISTBARITEMS = 612
        mo_PartidasDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_PartidasDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_PartidasDetalle.IdPartidaPresupuestal = Me.ucPartidasLista1.idRegistroSeleccionado
            If mo_PartidasDetalle.IdPartidaPresupuestal = -1 Or mo_PartidasDetalle.IdPartidaPresupuestal = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_PartidasDetalle.MostrarFormulario
        Set mo_PartidasDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select


End Sub
''
''
''
''
''
Sub EdicionEstablecimientosNoMinsa(sToolId As String)
'Dim mo_EstablecimientoNoMinsaDetalle As New EstablecimientoNoMinsaDetalle
Dim mo_EstablecimientoNoMinsaDetalle As New SIGHCatalogos.clEstablecNoMinsaDetalle

        mo_EstablecimientoNoMinsaDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_EstablecimientoNoMinsaDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_EstablecimientoNoMinsaDetalle.lnIdTablaLISTBARITEMS = 1204
        mo_EstablecimientoNoMinsaDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_EstablecimientoNoMinsaDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_EstablecimientoNoMinsaDetalle.IdEstablecimientoNoMINSA = Me.ucEstablecimientosNoMinsaLista1.idRegistroSeleccionado
            If mo_EstablecimientoNoMinsaDetalle.IdEstablecimientoNoMINSA = -1 Or mo_EstablecimientoNoMinsaDetalle.IdEstablecimientoNoMINSA = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

'        mo_EstablecimientoNoMinsaDetalle.Icon = Me.Icon
'        mo_EstablecimientoNoMinsaDetalle.Show 1
'        Unload mo_EstablecimientoNoMinsaDetalle
        mo_EstablecimientoNoMinsaDetalle.MostrarFormulario

        Set mo_EstablecimientoNoMinsaDetalle = Nothing
        ucEstablecimientosNoMinsaLista1.RealizarBusqueda
        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select



End Sub
''
''
''
''Sub EdicionFactExamenes(sToolId As String)
''
''End Sub
''
''Sub EdicionFactRecetas(sToolId As String)
''End Sub
''
Sub EdicionCamas(sToolId As String, lbEsEmergencia As Boolean)
        Dim mo_camas As New SIGHProxies.CamaDetalleProxy
        mo_camas.Opcion = SeleccionarOpcion(sToolId)
        mo_camas.idUsuario = ml_IdUsuarioAuditoria
        If lbEsEmergencia = True Then
           mo_camas.lnIdTablaLISTBARITEMS = 203
        Else
           mo_camas.lnIdTablaLISTBARITEMS = 303
        End If
        mo_camas.lcNombrePc = lc_NombrePc
        Select Case mo_camas.Opcion
        Case sghAgregar
            'mgaray20141014
            If ucCamasLista1.IdServicio = 0 Then
                MsgBox "Seleccione Servicio", vbInformation, "Agregar Camas"
                Exit Sub
            End If
            ucCamasLista1.SetDataServicioBusqueda
            mo_camas.IdServicio = ucCamasLista1.IdServicio
            mo_camas.CodigoServicio = ucCamasLista1.CodigoServicio
            mo_camas.NombreServicio = ucCamasLista1.NombreServicio
        Case sghModificar, sghConsultar, sghEliminar
            If ucCamasLista1.TieneRegistros = False Then Exit Sub
            mo_camas.idCama = ucCamasLista1.idRegistroSeleccionado
            If ucCamasLista1.idRegistroSeleccionado = -1 Or ucCamasLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_camas.idTipoServicio = sghHospitalizacion
        mo_camas.MostrarDialogo IIf(lbEsEmergencia = True, 2, 3)
        ucCamasLista1.RealizarBusqueda
End Sub
''
Sub EdicionCitas(sToolId As String)
        Me.ucCitasLista1.lcNombrePc = lc_NombrePc
        Me.ucCitasLista1.lbNuevoMovimiento = True
        Select Case sToolId
        Case "ID_Agregar":
            Me.ucCitasLista1.mnuDiarioAgregarCita_Click
        Case "ID_Modificar":
            Me.ucCitasLista1.mnuModificarDiarioCita_Click
        Case "ID_Consultar":
            Me.ucCitasLista1.mnuDiarioConsultarCita_Click
        Case "ID_Eliminar":
            Me.ucCitasLista1.mnuDiarioEliminarCita_Click
        End Select

End Sub
''
''
''
Sub EdicionProgMedica(sToolId As String)
        Me.ucProgramacionLista1.lnIdTablaLISTBARITEMS = 401
        Me.ucProgramacionLista1.lcNombrePc = lc_NombrePc
        Select Case sToolId
        Case "ID_Agregar":
            Me.ucProgramacionLista1.mnuDiarioAgregarProgramacion_Click
        Case "ID_Modificar":
            Me.ucProgramacionLista1.mnuDiarioModificarProgramacion_Click
        Case "ID_Consultar":
            Me.ucProgramacionLista1.mnuDiarioConsultarProgramacion_Click
        Case "ID_Eliminar":
            Me.ucProgramacionLista1.mnuDiarioEliminarProgramacion_Click
        End Select

End Sub
''
Sub EdicionRoles(sToolId As String)
Dim mo_RolDetalle As New RolesDetalle

        mo_RolDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_RolDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_RolDetalle.lnIdTablaLISTBARITEMS = 1302
        mo_RolDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_RolDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_RolDetalle.IdRol = Me.ucRolesLista1.idRegistroSeleccionado
            If mo_RolDetalle.IdRol = -1 Or mo_RolDetalle.IdRol = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_RolDetalle.Icon = Me.Icon
        mo_RolDetalle.Show 1
        Unload mo_RolDetalle

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

        Set ucRolesLista1.DataSource = mo_AdminSeguridad.RolesSeleccionarTodos()

End Sub
''
''Sub GenerarRecordsetDeListItems()
''
''    With mrs_ListItems
''          .Fields.Append "IdListItem", adInteger, 4
''          .Fields.Append "Clave", adVarChar, 50
''          .CursorType = adOpenStatic
''          .LockType = adLockOptimistic
''          .Open
''    End With
''
''End Sub
'
'
'
'Private Sub ucAdmisionConsEmerg_OnClick(oRecordset As ADODB.Recordset)
'
'    ml_ToolbarHeightAdd = 0
'    On Error Resume Next
'    If Not IsDate(oRecordset!FechaEgresoAdministrativo) Then
'        ml_ToolbarHeightAdd = 500
'        Select Case oRecordset!idTipoServicio
'        Case 2
''            toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = True
''            toolbar.Tools("ID_EmergenciaAObservacion").Enabled = True
''            toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = True
''            toolbar.Tools("ID_EmergenciaTransferencias").Enabled = True
'        Case 4
''            toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = True
''            toolbar.Tools("ID_EmergenciaAObservacion").Enabled = False
''            toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = True
''            toolbar.Tools("ID_EmergenciaTransferencias").Enabled = True
'        End Select
'    Else
''            toolbar.Tools("ID_EmergenciaAltaPaciente").Enabled = False
''            toolbar.Tools("ID_EmergenciaAObservacion").Enabled = False
''            toolbar.Tools("ID_EmergenciaAHospitalizacion").Enabled = False
''            toolbar.Tools("ID_EmergenciaTransferencias").Enabled = False
'    End If
'
'End Sub
'
'Private Sub ucAdmisionHospitalizacion_OnClick(oRecordset As ADODB.Recordset)
'    ml_ToolbarHeightAdd = 0
'    On Error Resume Next
'
'    If Not IsDate(oRecordset!FechaEgresoAdministrativo) Then
''        ml_ToolbarHeightAdd = 500
''        toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = True
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = True
''        toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = True
'    Else
''        toolbar.Tools("ID_HospitalizacionAlojamientoConjunto").Enabled = False
''        toolbar.Tools("ID_HospitalizacionAltaPaciente").Enabled = False
''        toolbar.Tools("ID_HospitalizacionTransferencias").Enabled = False
'    End If
'End Sub
'
Sub EdicionCatalogoBaseBienesInsumos(sToolId As String)
Dim mo_CatalogoBienesEInsumosDetalle As New SIGHCatalogos.clCatalogoBaseBienesDet

    mo_CatalogoBienesEInsumosDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_CatalogoBienesEInsumosDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_CatalogoBienesEInsumosDetalle.lnIdTablaLISTBARITEMS = 803
    mo_CatalogoBienesEInsumosDetalle.lcNombrePc = lc_NombrePc
    Select Case mo_CatalogoBienesEInsumosDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        mo_CatalogoBienesEInsumosDetalle.idProducto = Me.ucCatalogoBienesInsumosLista1.idRegistroSeleccionado
        If mo_CatalogoBienesEInsumosDetalle.idProducto = -1 Or mo_CatalogoBienesEInsumosDetalle.idProducto = 0 Then
            MsgBox "Seleccione un registro", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select

    mo_CatalogoBienesEInsumosDetalle.MostrarFormulario
    Set mo_CatalogoBienesEInsumosDetalle = Nothing

    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select

End Sub
''Sub EdicionCatalogoBienesInsumos(sToolId As String)
''Dim mo_CatalogoBienesInsumosDetalle As New SIGHCatalogos.clCatalogoBienesInsumoDet
''    mo_CatalogoBienesInsumosDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_CatalogoBienesInsumosDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_CatalogoBienesInsumosDetalle.TipoCatalogo = Me.ucCatalogoBienesInsumosLista1.IdTipoCatalogo
''
''
''    Select Case mo_CatalogoBienesInsumosDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''        Exit Sub
''    End Select
''
''    mo_CatalogoBienesInsumosDetalle.MostrarFormulario
''    Set mo_CatalogoBienesInsumosDetalle = Nothing
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
Sub EdicionCatalogoBaseServicios(sToolId As String)
Dim mo_CatalogoServiciosDetalle As New SIGHCatalogos.clCatalogoBaseServicDet

    mo_CatalogoServiciosDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_CatalogoServiciosDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_CatalogoServiciosDetalle.lnIdTablaLISTBARITEMS = 610
    mo_CatalogoServiciosDetalle.lcNombrePc = lc_NombrePc
    Select Case mo_CatalogoServiciosDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        mo_CatalogoServiciosDetalle.idProducto = Me.ucCatalogoServiciosLista1.idRegistroSeleccionado
        If mo_CatalogoServiciosDetalle.idProducto = -1 Or mo_CatalogoServiciosDetalle.idProducto = 0 Then
            MsgBox "Seleccione un registro", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select

     mo_CatalogoServiciosDetalle.MostrarFormulario
     Set mo_CatalogoServiciosDetalle = Nothing

    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select

End Sub
''Sub EdicionCatalogoServicios(sToolId As String)
''Dim mo_CatalogoServiciosDetalle As New SIGHCatalogos.clCatalogoServicioDetalle
''
''    mo_CatalogoServiciosDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_CatalogoServiciosDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_CatalogoServiciosDetalle.TipoCatalogo = Me.ucCatalogoServiciosLista1.IdTipoCatalogo
''
''    Select Case mo_CatalogoServiciosDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''        Exit Sub
''    End Select
''
''    mo_CatalogoServiciosDetalle.MostrarFormulario
''    Set mo_CatalogoServiciosDetalle = Nothing
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
''
''Sub EdicionCentrosCosto(sToolId As String)
''Dim mo_CentrosCostoDetalle As New SIGHCatalogos.clCentroCostosDetalle
''
''
''    mo_CentrosCostoDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_CentrosCostoDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_CentrosCostoDetalle.lnIdTablaLISTBARITEMS = 609
''    mo_CentrosCostoDetalle.lcNombrePc = lc_NombrePc
''    Select Case mo_CentrosCostoDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''        mo_CentrosCostoDetalle.IdCentroCosto = Me.ucCentrosCostoLista1.idRegistroSeleccionado
''        If mo_CentrosCostoDetalle.IdCentroCosto = -1 Or mo_CentrosCostoDetalle.IdCentroCosto = 0 Then
''            MsgBox "Seleccione un registro", vbInformation, Me.Caption
''            Exit Sub
''        End If
''    End Select
''
''    mo_CentrosCostoDetalle.MostrarFormulario
''    Set mo_CentrosCostoDetalle = Nothing
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
'
'
'Sub AperturaCaja()
'Dim oApertura As New AperturaDecaja
'Dim oDOEmpleado As DOEmpleado
'Dim sNombreCajero As String
'Dim oRsPermisos As New Recordset
'Dim lbUsuarioRealizaApertura As Boolean
'        '
'        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
'        If oRsPermisos.RecordCount > 0 Then
'           Do While Not oRsPermisos.EOF
'              Select Case oRsPermisos.Fields!IdPermiso
'              Case 201    'Caja - Realizar Apertura
'                   lbUsuarioRealizaApertura = True
'              End Select
'              oRsPermisos.MoveNext
'           Loop
'        End If
'        Set oRsPermisos = Nothing
'        '
'        If lbUsuarioRealizaApertura = True Then
'            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(ml_IdUsuarioAuditoria)
'            sNombreCajero = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
'            oApertura.NombreCajero = sNombreCajero
'            oApertura.idUsuario = ml_IdUsuarioAuditoria
'            oApertura.lcNombrePc = lc_NombrePc
'            oApertura.Show 1
'            If oApertura.AperturoCajaOK = True Then
'
'                'debb-15/03/2016 (inicio)
'                If oApertura.IdTurno = 0 Then
'                   MsgBox "Tiene problemas con el TURNO", vbInformation, ""
'                   Exit Sub
'                End If
'                'debb-15/03/2016 (fin)
'
'                mb_abrioCaja = Me.ucGestionCaja1.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, oApertura.IdCaja, oApertura.IdTurno, oApertura.EmiteSoloServicio)
'
'                '/****************************INO***************************************/
'                mb_abrioCaja = Me.ucGestionDevolucion2.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, oApertura.IdCaja, oApertura.IdTurno, oApertura.EmiteSoloServicio)
'                '/****************************INO***************************************/
'
'                Me.Toolbar.Tools("ID_CajaApertura").Enabled = False
'                Me.Toolbar.Tools("ID_CajaCierre").Enabled = True
'                'mgaray201503
'                Set moDOCajaGestion = New DOCajaGestion
'                moDOCajaGestion.IdCaja = oApertura.IdCaja
'                moDOCajaGestion.IdCajero = oApertura.IdTurno
'                lbCajeroEmiteSoloServicios = oApertura.EmiteSoloServicio
'            End If
'            Unload oApertura
'
'        Else
'            MsgBox "El Usuario no tiene permiso para realizar APERTURA DE CAJA", vbInformation, Me.Caption
'        End If
'
'End Sub
'
'Sub CerrarCaja()
'Dim oRsPermisos As New Recordset
'Dim lbUsuarioRealizaCierre As Boolean
'        '
'        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
'        If oRsPermisos.RecordCount > 0 Then
'           Do While Not oRsPermisos.EOF
'              Select Case oRsPermisos.Fields!IdPermiso
'              Case 202    'Caja - Realizar Apertura
'                   lbUsuarioRealizaCierre = True
'              End Select
'              oRsPermisos.MoveNext
'           Loop
'        End If
'        Set oRsPermisos = Nothing
'        '
'        If lbUsuarioRealizaCierre = True Then
'
'            If Not mb_abrioCaja Then
'                Exit Sub
'            End If
'            If MsgBox("¿Esta seguro de realizar el CIERRE DE CAJA ?", vbYesNo, Me.Caption) = vbYes Then
'                If ucGestionCaja1.RealizarCierreDeCaja() Then
'                    Me.Toolbar.Tools("ID_CajaApertura").Enabled = True
'                    mb_abrioCaja = False
'                End If
'
'                '/******************************INO*************************************
'                 If ucGestionDevolucion2.RealizarCierreDeCaja() Then
'                    Me.Toolbar.Tools("ID_CajaApertura").Enabled = True
'                    mb_abrioCaja = False
'                End If
'                '/******************************INO*************************************
'
'            Else
'                ucGestionCaja1.MuestraTabEmisionDocumentos (False)
'                Me.Toolbar.Tools("ID_CajaApertura").Enabled = True
'                mb_abrioCaja = False
'            End If
'            Me.Toolbar.Tools("ID_CajaApertura").Enabled = True
'            Me.Toolbar.Tools("ID_CajaCierre").Enabled = False
'        Else
'            MsgBox "El USUARIO no tiene permiso para realizar el  CIERRE"
'        End If
'End Sub
'
'
'Private Sub ucCajeroServicios1_HizoClickEnEscape()
'
'    mo_LastControl.Visible = False
'    Toolbar.Toolbars("Edición").Visible = True
'    Toolbar.Toolbars("Gestión de Caja").Visible = False
'
'End Sub
'
'Private Sub ucEstadoCuenta1_HizoClickEnCancelar()
'    mo_LastControl.Visible = False
'
'End Sub
'
Sub EdicionOrdenesServicio(sToolId As String)
Dim mo_FacOrdenServicioDetalle As New FacOrdenServicioDetalle

    mo_FacOrdenServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_FacOrdenServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_FacOrdenServicioDetalle.idTipoFinanciamiento = Me.ucFacturacionGeneralLista.idTipoFinanciamiento
    mo_FacOrdenServicioDetalle.PuntoCarga = Me.ucFacturacionGeneralLista.PuntoCarga
    mo_FacOrdenServicioDetalle.lnIdTablaLISTBARITEMS = 601
    mo_FacOrdenServicioDetalle.lcNombrePc = lc_NombrePc
    Select Case mo_FacOrdenServicioDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
    End Select

    mo_FacOrdenServicioDetalle.IdOrden = Me.ucFacturacionGeneralLista.idRegistroSeleccionado
    mo_FacOrdenServicioDetalle.Show 1
    Unload mo_FacOrdenServicioDetalle
    ucFacturacionGeneralLista.RealizarBusqueda
    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select

End Sub
''
''Sub EdicionOrdenesServicioPatologiaClinica(sToolId As String)
''Dim mo_FacOrdenServicioDetalle As New FacOrdenServicioDetalle
''
''    mo_FacOrdenServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_FacOrdenServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_FacOrdenServicioDetalle.idTipoFinanciamiento = Me.ucFactPatologiaClinica.idTipoFinanciamiento
''    mo_FacOrdenServicioDetalle.PuntoCarga = Me.ucFactPatologiaClinica.PuntoCarga
''    Select Case mo_FacOrdenServicioDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''    End Select
''
''    mo_FacOrdenServicioDetalle.IdOrden = Me.ucFactPatologiaClinica.idRegistroSeleccionado
''    mo_FacOrdenServicioDetalle.Show 1
''    Unload mo_FacOrdenServicioDetalle
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
''Sub EdicionOrdenesServicioAnatomiaPatologia(sToolId As String)
''Dim mo_FacOrdenServicioDetalle As New FacOrdenServicioDetalle
''
''    mo_FacOrdenServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_FacOrdenServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_FacOrdenServicioDetalle.idTipoFinanciamiento = Me.ucFactAnatomiaPatologica.idTipoFinanciamiento
''    mo_FacOrdenServicioDetalle.PuntoCarga = Me.ucFactAnatomiaPatologica.PuntoCarga
''
''    Select Case mo_FacOrdenServicioDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''    End Select
''
''    mo_FacOrdenServicioDetalle.IdOrden = Me.ucFactAnatomiaPatologica.idRegistroSeleccionado
''    mo_FacOrdenServicioDetalle.Show 1
''    Unload mo_FacOrdenServicioDetalle
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
''
''Sub EdicionOrdenesServicioImagenologia(sToolId As String)
''Dim mo_FacOrdenServicioDetalle As New FacOrdenServicioDetalle
''
''    mo_FacOrdenServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_FacOrdenServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_FacOrdenServicioDetalle.idTipoFinanciamiento = Me.ucFactImagenologia.idTipoFinanciamiento
''    mo_FacOrdenServicioDetalle.PuntoCarga = Me.ucFactImagenologia.PuntoCarga
''
''    Select Case mo_FacOrdenServicioDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''    End Select
''
''    mo_FacOrdenServicioDetalle.IdOrden = Me.ucFactImagenologia.idRegistroSeleccionado
''    mo_FacOrdenServicioDetalle.Show 1
''    Unload mo_FacOrdenServicioDetalle
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
''Sub EdicionOrdenesServicioSalaOperaciones(sToolId As String)
''Dim mo_FacOrdenServicioDetalle As New FacOrdenServicioDetalle
''
''    mo_FacOrdenServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_FacOrdenServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_FacOrdenServicioDetalle.idTipoFinanciamiento = Me.ucFactSalaOperaciones.idTipoFinanciamiento
''    mo_FacOrdenServicioDetalle.PuntoCarga = Me.ucFactSalaOperaciones.PuntoCarga
''    mo_FacOrdenServicioDetalle.lcNombrePc = lc_NombrePc
''    mo_FacOrdenServicioDetalle.lnIdTablaLISTBARITEMS = 607
''
''    Select Case mo_FacOrdenServicioDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''    End Select
''
''    mo_FacOrdenServicioDetalle.IdOrden = Me.ucFactSalaOperaciones.idRegistroSeleccionado
''    mo_FacOrdenServicioDetalle.Show 1
''    Unload mo_FacOrdenServicioDetalle
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
''Sub EdicionOrdenesServicioFarmacia(sToolId As String)
''
''End Sub
'
'Sub ImprimirParteDiario()
'Dim oRptCaja As New RptCaja
'
'    oRptCaja.IdGestionCaja = Me.ucGestionCaja1.IdGestionCaja
'
'    If oRptCaja.IdGestionCaja <> -1 Then
'        oRptCaja.CrearParteDiario Me.hwnd
'    End If
'
'End Sub
'
'Sub ImprimirConsolidadoServicio()
'Dim oRptCaja As New RptCaja
'
'    oRptCaja.IdGestionCaja = Me.ucGestionCaja1.IdGestionCaja
'
'    If oRptCaja.IdGestionCaja <> -1 Then
'        oRptCaja.CrearReporteConsolidadoServicios False
'    End If
'
'End Sub
'
Sub EdicionCaja(sToolId As String)
Dim mo_cajaDetalle As New SIGHCatalogos.clCajaDetalle

        mo_cajaDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_cajaDetalle.idUsuario = ml_IdUsuarioAuditoria
        mo_cajaDetalle.lnIdTablaLISTBARITEMS = 705
        mo_cajaDetalle.lcNombrePc = lc_NombrePc
        Select Case mo_cajaDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_cajaDetalle.IdCaja = Me.ucCajaLista1.idRegistroSeleccionado
            If mo_cajaDetalle.IdCaja = -1 Or mo_cajaDetalle.IdCaja = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_cajaDetalle.MostrarFormulario
        Set mo_cajaDetalle = Nothing

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
''
Sub EdicionInventario(sToolId As String)
        Dim mo_Inventario As New SIGHProxies.Inventario
        mo_Inventario.Opcion = SeleccionarOpcion(sToolId)
        mo_Inventario.idUsuario = ml_IdUsuarioAuditoria
        mo_Inventario.lnIdTablaLISTBARITEMS = 801
        mo_Inventario.lcNombrePc = lc_NombrePc
        Select Case mo_Inventario.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_Inventario.idInventario = ucFarmInventarioLista1.idRegistroSeleccionado
            If ucFarmInventarioLista1.idRegistroSeleccionado = -1 Or ucFarmInventarioLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_Inventario.MostrarFormularioInventario
        ucFarmInventarioLista1.RealizarBusqueda
End Sub
'''**debb2014
Sub EdicionNS(sToolId As String, lbNSsoloParaFarmacia As Boolean)
        Dim mo_Ns As New SighFarmacia.NotaSalida
        Dim lcMovimiento As String
        lcMovimiento = Right("0" + Trim(Str(ucFarmNsLista1.idRegistroSeleccionado)), 9)
        mo_Ns.Opcion = SeleccionarOpcion(sToolId)
        mo_Ns.idUsuario = ml_IdUsuarioAuditoria
        mo_Ns.lcNombrePc = lc_NombrePc
        If lbNSsoloParaFarmacia = True Then
           mo_Ns.lnIdTablaLISTBARITEMS = 1358
        Else
           mo_Ns.lnIdTablaLISTBARITEMS = 1305
        End If
        Select Case mo_Ns.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_Ns.movNumero = lcMovimiento
            If ucFarmNsLista1.idRegistroSeleccionado = -1 Or ucFarmNsLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_Ns.MostrarFormularioNotaSalida
        ucFarmNsLista1.RealizarBusqueda
End Sub
'''**debb2014
Sub EdicionNI(sToolId As String, lbNIsoloParaFarmacia As Boolean)
        Dim lcMovimiento As String
        If ms_ModuloSeleccionado = "FARMADOP" Then
            Dim niArmado As New SighFarmacia.NotaSalida
            lcMovimiento = Right("0" + Trim(Str(ucFarmNiLista1.idRegistroSeleccionado)), 9)
            niArmado.Opcion = SeleccionarOpcion(sToolId)
            niArmado.idUsuario = ml_IdUsuarioAuditoria
            niArmado.lcNombrePc = lc_NombrePc
            niArmado.lnIdTablaLISTBARITEMS = 1357
            Select Case niArmado.Opcion
            Case sghAgregar
            Case sghModificar, sghConsultar, sghEliminar
                niArmado.movNumero = lcMovimiento
                If ucFarmNiLista1.idRegistroSeleccionado = -1 Or ucFarmNiLista1.idRegistroSeleccionado = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
            End Select
            niArmado.MostrarFormularioPaquetes
            Set niArmado = Nothing
        Else
            Dim mo_Ni As New SighFarmacia.NotaIngreso

            lcMovimiento = Right("0" + Trim(Str(ucFarmNiLista1.idRegistroSeleccionado)), 9)
            mo_Ni.Opcion = SeleccionarOpcion(sToolId)
            mo_Ni.idUsuario = ml_IdUsuarioAuditoria
            mo_Ni.lcNombrePc = lc_NombrePc
            If lbNIsoloParaFarmacia = True Then
               mo_Ni.lnIdTablaLISTBARITEMS = 1357
            Else
               mo_Ni.lnIdTablaLISTBARITEMS = 1304
            End If
            Select Case mo_Ni.Opcion
            Case sghAgregar
            Case sghModificar, sghConsultar, sghEliminar
                mo_Ni.movNumero = lcMovimiento
                If ucFarmNiLista1.idRegistroSeleccionado = -1 Or ucFarmNiLista1.idRegistroSeleccionado = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
            End Select
            mo_Ni.MostrarFormularioNotaIngreso
            Set mo_Ni = Nothing
        End If
        ucFarmNiLista1.RealizarBusqueda
End Sub
Sub EdicionIntervencionS(sToolId As String)
        Dim mo_IntervencionS As New SighFarmacia.IntervencionS
        Dim lcMovimiento As String
        lcMovimiento = Right("0" + Trim(Str(ucFarmIntervencionLista1.idRegistroSeleccionado)), 9)
        mo_IntervencionS.Opcion = SeleccionarOpcion(sToolId)
        mo_IntervencionS.idUsuario = ml_IdUsuarioAuditoria
        mo_IntervencionS.lnIdTablaLISTBARITEMS = 1308
        mo_IntervencionS.lcNombrePc = lc_NombrePc
        Select Case mo_IntervencionS.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_IntervencionS.movNumero = lcMovimiento
            If ucFarmIntervencionLista1.idRegistroSeleccionado = -1 Or ucFarmIntervencionLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_IntervencionS.MostrarFormularioIntervencion
        ucFarmIntervencionLista1.RealizarBusqueda
End Sub
''
Sub EdicionVentas(sToolId As String)
        Dim mo_Ventas As New SighFarmacia.Ventas
        Dim lcMovimiento As String
        If ucFarmVentasLista1.TipoVentaSeleccionada = 0 Then  'Venta Directa - farmMovimientos
           lcMovimiento = Right("0" + Trim(Str(ucFarmVentasLista1.idRegistroSeleccionado)), 9)
        Else    'preventas - farmPreVentas
           lcMovimiento = Trim(Str(ucFarmVentasLista1.idRegistroSeleccionado))
        End If
        mo_Ventas.Opcion = SeleccionarOpcion(sToolId)
        mo_Ventas.idUsuario = ml_IdUsuarioAuditoria
        mo_Ventas.TipoVentaSeleccionada = ucFarmVentasLista1.TipoVentaSeleccionada
        mo_Ventas.lnIdTablaLISTBARITEMS = 1307
        mo_Ventas.lcNombrePc = lc_NombrePc
        Select Case mo_Ventas.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_Ventas.movNumero = lcMovimiento
            If ucFarmVentasLista1.idRegistroSeleccionado = -1 Or ucFarmVentasLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_Ventas.MostrarFormulario

        'ucFarmVentasLista1.inicializar
        ucFarmVentasLista1.RealizarBusqueda
End Sub
''
Sub EdicionDependenciaExt(sToolId As String)
        Dim mo_DependenciaExt As New SighFarmacia.DependenciaExt
        mo_DependenciaExt.Opcion = SeleccionarOpcion(sToolId)
        mo_DependenciaExt.idUsuario = ml_IdUsuarioAuditoria
        mo_DependenciaExt.lnIdTablaLISTBARITEMS = 1310
        mo_DependenciaExt.lcNombrePc = lc_NombrePc
        Select Case mo_DependenciaExt.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_DependenciaExt.IdDependenciaExt = ucFarmDependExtLista1.idRegistroSeleccionado
            If ucFarmDependExtLista1.idRegistroSeleccionado = -1 Or ucFarmDependExtLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_DependenciaExt.MostrarFormulario
        ucFarmDependExtLista1.RealizarBusqueda
End Sub
''
Sub EdicionRayosX(sToolId As String)
        Dim mo_RayosX As New SIGHImagen.RayosX
        mo_RayosX.Opcion = SeleccionarOpcion(sToolId)
        mo_RayosX.idUsuario = ml_IdUsuarioAuditoria
        mo_RayosX.lnIdTablaLISTBARITEMS = 1318
        mo_RayosX.lcNombrePc = lc_NombrePc
        Select Case mo_RayosX.Opcion
        Case sghAgregar
             If UcImagenesLista1.SeEligioGridBoleta = True Then
                mo_RayosX.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
                mo_RayosX.SeEligioGridBoleta = UcImagenesLista1.SeEligioGridBoleta
             End If
        Case sghModificar, sghConsultar, sghEliminar
            mo_RayosX.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
            If UcImagenesLista1.idRegistroSeleccionado = -1 Or UcImagenesLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_RayosX.MostrarFormulario
        UcImagenesLista1.RealizarBusqueda
        UcImagenesLista1.SeEligioGridBoleta = False
End Sub
''
''Sub EdicionImagIngresos(sToolId As String)
''        Dim mo_ImagIngresos As New SIGHImagen.Ingresos
''        mo_ImagIngresos.Opcion = SeleccionarOpcion(sToolId)
''        mo_ImagIngresos.idUsuario = ml_IdUsuarioAuditoria
''        mo_ImagIngresos.lnIdTablaLISTBARITEMS = 1315
''        mo_ImagIngresos.lcNombrePc = lc_NombrePc
''        Select Case mo_ImagIngresos.Opcion
''        Case sghAgregar
''        Case sghModificar, sghConsultar, sghEliminar
''            mo_ImagIngresos.IdMovimiento = UcImagIngresos1.idRegistroSeleccionado
''            If UcImagIngresos1.idRegistroSeleccionado = -1 Or UcImagIngresos1.idRegistroSeleccionado = 0 Then
''                MsgBox "Seleccione un registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''        End Select
''        mo_ImagIngresos.MostrarFormulario
''        UcImagIngresos1.RealizarBusqueda
''End Sub
''
''Sub EdicionImagSalidas(sToolId As String)
''        Dim mo_ImagSalidas As New SIGHImagen.Salidas
''        mo_ImagSalidas.Opcion = SeleccionarOpcion(sToolId)
''        mo_ImagSalidas.idUsuario = ml_IdUsuarioAuditoria
''        mo_ImagSalidas.lnIdTablaLISTBARITEMS = 1316
''        mo_ImagSalidas.lcNombrePc = lc_NombrePc
''        Select Case mo_ImagSalidas.Opcion
''        Case sghAgregar
''        Case sghModificar, sghConsultar, sghEliminar
''            mo_ImagSalidas.IdMovimiento = UcImagSalidas1.idRegistroSeleccionado
''            If UcImagSalidas1.idRegistroSeleccionado = -1 Or UcImagSalidas1.idRegistroSeleccionado = 0 Then
''                MsgBox "Seleccione un registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''        End Select
''        mo_ImagSalidas.MostrarFormulario
''        UcImagSalidas1.RealizarBusqueda
''End Sub
''
Sub EdicionImagEcografiaObs(sToolId As String)
        Dim mo_EcogObs As New SIGHImagen.EcogObs
        mo_EcogObs.Opcion = SeleccionarOpcion(sToolId)
        mo_EcogObs.idUsuario = ml_IdUsuarioAuditoria
        mo_EcogObs.lnIdTablaLISTBARITEMS = 1320
        mo_EcogObs.lcNombrePc = lc_NombrePc
        Select Case mo_EcogObs.Opcion
        Case sghAgregar
             If UcImagenesLista1.SeEligioGridBoleta = True Then
                mo_EcogObs.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
                mo_EcogObs.SeEligioGridBoleta = UcImagenesLista1.SeEligioGridBoleta
             End If
        Case sghModificar, sghConsultar, sghEliminar
            mo_EcogObs.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
            If UcImagenesLista1.idRegistroSeleccionado = -1 Or UcImagenesLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_EcogObs.MostrarFormulario
        UcImagenesLista1.RealizarBusqueda
        UcImagenesLista1.SeEligioGridBoleta = False
End Sub
''
Sub EdicionImagEcografiaGen(sToolId As String)
        Dim mo_EcogGen As New SIGHImagen.EcogGen
        mo_EcogGen.Opcion = SeleccionarOpcion(sToolId)
        mo_EcogGen.idUsuario = ml_IdUsuarioAuditoria
        mo_EcogGen.lcNombrePc = lc_NombrePc
        mo_EcogGen.lnIdTablaLISTBARITEMS = 1317
        Select Case mo_EcogGen.Opcion
        Case sghAgregar
             If UcImagenesLista1.SeEligioGridBoleta = True Then
                mo_EcogGen.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
                mo_EcogGen.SeEligioGridBoleta = UcImagenesLista1.SeEligioGridBoleta
             End If
        Case sghModificar, sghConsultar, sghEliminar
            If UcImagenesLista1.SeEligioGridBoleta = True Then
            Else
               mo_EcogGen.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
            End If
            If UcImagenesLista1.idRegistroSeleccionado = -1 Or UcImagenesLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_EcogGen.MostrarFormulario
        UcImagenesLista1.RealizarBusqueda
        UcImagenesLista1.SeEligioGridBoleta = False
End Sub

Sub EdicionImagTomografia(sToolId As String)
        Dim mo_tomog As New SIGHImagen.Tomog
        mo_tomog.Opcion = SeleccionarOpcion(sToolId)
        mo_tomog.idUsuario = ml_IdUsuarioAuditoria
        mo_tomog.lnIdTablaLISTBARITEMS = 1319
        mo_tomog.lcNombrePc = lc_NombrePc
        Select Case mo_tomog.Opcion
        Case sghAgregar
             If UcImagenesLista1.SeEligioGridBoleta = True Then
                mo_tomog.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
                mo_tomog.SeEligioGridBoleta = UcImagenesLista1.SeEligioGridBoleta
             End If
        Case sghModificar, sghConsultar, sghEliminar
            mo_tomog.IdMovimiento = UcImagenesLista1.idRegistroSeleccionado
            If UcImagenesLista1.idRegistroSeleccionado = -1 Or UcImagenesLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_tomog.MostrarFormulario
        UcImagenesLista1.RealizarBusqueda
        UcImagenesLista1.SeEligioGridBoleta = False
End Sub
''
Sub EdicionLaboratorio(sToolId As String)
  Dim mo_laboratorio As New SIGHLaboratorio.laboratorio
  mo_laboratorio.Opcion = SeleccionarOpcion(sToolId)
  mo_laboratorio.idUsuario = ml_IdUsuarioAuditoria
  mo_laboratorio.PuntoCarga = 2
  mo_laboratorio.lnIdTablaLISTBARITEMS = 1312
  mo_laboratorio.lcNombrePc = lc_NombrePc
  mo_laboratorio.AreaTrabajo = ucFactOrdenesLaboratorio.AreaTrabajo
  Select Case mo_laboratorio.Opcion
  Case sghAgregar
       If ucFactOrdenesLaboratorio.SeEligioGridBoleta = True Then
          mo_laboratorio.IdMovimiento = ucFactOrdenesLaboratorio.idRegistroSeleccionado
          mo_laboratorio.SeEligioGridBoleta = ucFactOrdenesLaboratorio.SeEligioGridBoleta
       End If
  Case sghModificar, sghConsultar, sghEliminar
       If ucFactOrdenesLaboratorio.SeEligioGridBoleta = True Then
       Else
           mo_laboratorio.IdMovimiento = ucFactOrdenesLaboratorio.idRegistroSeleccionado
       End If
       If ucFactOrdenesLaboratorio.idRegistroSeleccionado = -1 Or ucFactOrdenesLaboratorio.idRegistroSeleccionado = 0 Then
          MsgBox "Seleccione un registro", vbInformation, Me.Caption
          Exit Sub
       End If
  End Select
  mo_laboratorio.MostrarFormulario
  ucFactOrdenesLaboratorio.RealizarBusqueda
  ucFactOrdenesLaboratorio.SeEligioGridBoleta = False
End Sub
''
''Sub EdicionOrdenesServicioPatologiaClinica_(sToolId As String)
''  Dim mo_FacOrdenServicioDetalle As New FacOrdenServicioDetalle
''
''    mo_FacOrdenServicioDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_FacOrdenServicioDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_FacOrdenServicioDetalle.idTipoFinanciamiento = Me.ucFactOrdenesLaboratorio.idTipoFinanciamiento
''    mo_FacOrdenServicioDetalle.PuntoCarga = Me.ucFactOrdenesLaboratorio.PuntoCarga
''
''    Select Case mo_FacOrdenServicioDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''    End Select
''
''    mo_FacOrdenServicioDetalle.IdOrden = Me.ucFactOrdenesLaboratorio.idRegistroSeleccionado
''    mo_FacOrdenServicioDetalle.Show 1
''    Unload mo_FacOrdenServicioDetalle
''
''    Select Case sToolId
''    Case "ID_Agregar":
''    Case "ID_Modificar":
''    Case "ID_Consultar":
''    Case "ID_Eliminar":
''    End Select
''
''End Sub
''
'''Frank 29042015
Sub EdicionOrdenesServicioAnatomiaPatologia_(sToolId As String)
  Dim mo_laboratorio As New SIGHLaboratorio.laboratorio
  mo_laboratorio.Opcion = SeleccionarOpcion(sToolId)
  mo_laboratorio.idUsuario = ml_IdUsuarioAuditoria
  mo_laboratorio.PuntoCarga = 3
  mo_laboratorio.lnIdTablaLISTBARITEMS = 1312
  mo_laboratorio.lcNombrePc = lc_NombrePc
  mo_laboratorio.AreaTrabajo = ucFacturacionOrdenesPatologia.AreaTrabajo
  Select Case mo_laboratorio.Opcion
  Case sghAgregar
       If ucFacturacionOrdenesPatologia.SeEligioGridBoleta = True Then
          mo_laboratorio.IdMovimiento = ucFacturacionOrdenesPatologia.idRegistroSeleccionado
          mo_laboratorio.SeEligioGridBoleta = ucFacturacionOrdenesPatologia.SeEligioGridBoleta
       End If
  Case sghModificar, sghConsultar, sghEliminar
       If ucFacturacionOrdenesPatologia.SeEligioGridBoleta = True Then
       Else
           mo_laboratorio.IdMovimiento = ucFacturacionOrdenesPatologia.idRegistroSeleccionado
       End If
       If ucFacturacionOrdenesPatologia.idRegistroSeleccionado = -1 Or ucFacturacionOrdenesPatologia.idRegistroSeleccionado = 0 Then
          MsgBox "Seleccione un registro", vbInformation, Me.Caption
          Exit Sub
       End If
  End Select
  mo_laboratorio.MostrarFormulario
  ucFacturacionOrdenesPatologia.RealizarBusqueda
  ucFacturacionOrdenesPatologia.SeEligioGridBoleta = False
End Sub
''
Sub EdicionOrdenesBS_(sToolId As String)
  Dim mo_laboratorio As New SIGHLaboratorio.laboratorio
  mo_laboratorio.Opcion = SeleccionarOpcion(sToolId)
  mo_laboratorio.idUsuario = ml_IdUsuarioAuditoria
  mo_laboratorio.PuntoCarga = 11
  mo_laboratorio.lnIdTablaLISTBARITEMS = 1312
  mo_laboratorio.lcNombrePc = lc_NombrePc
  mo_laboratorio.AreaTrabajo = ucFacturacionBS.AreaTrabajo
  Select Case mo_laboratorio.Opcion
  Case sghAgregar
       If ucFacturacionBS.SeEligioGridBoleta = True Then
          mo_laboratorio.IdMovimiento = ucFacturacionBS.idRegistroSeleccionado
          mo_laboratorio.SeEligioGridBoleta = ucFacturacionBS.SeEligioGridBoleta
       End If
  Case sghModificar, sghConsultar, sghEliminar
    If ucFacturacionBS.SeEligioGridBoleta = True Then
    Else
       mo_laboratorio.IdMovimiento = ucFacturacionBS.idRegistroSeleccionado
    End If
    If ucFacturacionBS.idRegistroSeleccionado = -1 Or ucFacturacionBS.idRegistroSeleccionado = 0 Then
      MsgBox "Seleccione un registro", vbInformation, Me.Caption
      Exit Sub
    End If
  End Select
  mo_laboratorio.MostrarFormulario
  ucFacturacionBS.RealizarBusqueda
End Sub
''
''Sub EdicionResultados(sToolId As String)
''  Dim mo_LabIngresos As New SIGHLaboratorio.laboratorio
''  mo_LabIngresos.Opcion = SeleccionarOpcion(sToolId)
''  mo_LabIngresos.idUsuario = ml_IdUsuarioAuditoria
''  Select Case mo_LabIngresos.Opcion
''  Case sghAgregar
''  Case sghModificar, sghConsultar, sghEliminar
''    mo_LabIngresos.IdMovimiento = UcLabIngresos1.idRegistroSeleccionado
''    If UcLabIngresos1.idRegistroSeleccionado = -1 Or UcLabIngresos1.idRegistroSeleccionado = 0 Then
''      MsgBox "Seleccione un registro", vbInformation, Me.Caption
''      Exit Sub
''    End If
''  End Select
''  mo_LabIngresos.MostrarFormulario
''  UcLabIngresos1.RealizarBusqueda
''End Sub
''
''Sub EdicionMuestras(sToolId As String)
''  Dim mo_LabSalidas As New SIGHLaboratorio.laboratorio
''  mo_LabSalidas.Opcion = SeleccionarOpcion(sToolId)
''  mo_LabSalidas.idUsuario = ml_IdUsuarioAuditoria
''  Select Case mo_LabSalidas.Opcion
''  Case sghAgregar
''  Case sghModificar, sghConsultar, sghEliminar
''    mo_LabSalidas.IdMovimiento = UcLabSalidas1.idRegistroSeleccionado
''    If UcLabSalidas1.idRegistroSeleccionado = -1 Or UcLabSalidas1.idRegistroSeleccionado = 0 Then
''      MsgBox "Seleccione un registro", vbInformation, Me.Caption
''      Exit Sub
''    End If
''  End Select
''  mo_LabSalidas.MostrarFormulario
''  UcLabSalidas1.RealizarBusqueda
''End Sub
''
''Sub EdicionLabIngresos(sToolId As String)
''  Dim mo_LabIngresos As New SIGHLaboratorio.Ingresos
''  mo_LabIngresos.Opcion = SeleccionarOpcion(sToolId)
''  mo_LabIngresos.idUsuario = ml_IdUsuarioAuditoria
''  mo_LabIngresos.idPuntoCarga = UcLabIngresos1.PuntoCarga
''  mo_LabIngresos.lnIdTablaLISTBARITEMS = 1313
''  mo_LabIngresos.lcNombrePc = lc_NombrePc
''  If UcLabIngresos1.PuntoCarga = -1 Or UcLabIngresos1.PuntoCarga = 0 Then
''    MsgBox "Escoja un punto de Carga para Agregar/Modificar un registro de Ingreso de Insumos.", vbInformation, Me.Caption
''    Exit Sub
''  End If
''  Select Case mo_LabIngresos.Opcion
''  Case sghAgregar
''  Case sghModificar, sghConsultar, sghEliminar
''    mo_LabIngresos.IdMovimiento = UcLabIngresos1.idRegistroSeleccionado
''    If UcLabIngresos1.idRegistroSeleccionado = -1 Or UcLabIngresos1.idRegistroSeleccionado = 0 Then
''      MsgBox "Seleccione un registro para Modificar Ingreso de Insumos.", vbInformation, Me.Caption
''      Exit Sub
''    End If
''  End Select
''  mo_LabIngresos.MostrarFormulario
''  UcLabIngresos1.RealizarBusqueda
''End Sub
''
''Sub EdicionLabSalidas(sToolId As String)
''  Dim mo_LabSalidas As New SIGHLaboratorio.Salidas
''  mo_LabSalidas.Opcion = SeleccionarOpcion(sToolId)
''  mo_LabSalidas.idUsuario = ml_IdUsuarioAuditoria
''  mo_LabSalidas.idPuntoCarga = UcLabSalidas1.PuntoCarga
''  mo_LabSalidas.lnIdTablaLISTBARITEMS = 1314
''  mo_LabSalidas.lcNombrePc = lc_NombrePc
''  If UcLabSalidas1.PuntoCarga = -1 Or UcLabSalidas1.PuntoCarga = 0 Then
''    MsgBox "Escoja un punto de Carga para Agregar/Modificar un registro de Salida de Insumos", vbInformation, Me.Caption
''    Exit Sub
''  End If
''  Select Case mo_LabSalidas.Opcion
''  Case sghAgregar
''  Case sghModificar, sghConsultar, sghEliminar
''    mo_LabSalidas.IdMovimiento = UcLabSalidas1.idRegistroSeleccionado
''    If UcLabSalidas1.idRegistroSeleccionado = -1 Or UcLabSalidas1.idRegistroSeleccionado = 0 Then
''      MsgBox "Seleccione un registro para Modificar Salida de Insumos", vbInformation, Me.Caption
''      Exit Sub
''    End If
''  End Select
''  mo_LabSalidas.MostrarFormulario
''  UcLabSalidas1.RealizarBusqueda
''End Sub
''
''Sub EdicionAlojados(sToolId As String)
''        Dim mo_Alojados As New AdmisionAlojDetalle
''        Dim rsHospitalizacion As New Recordset
''        mo_Alojados.Opcion = SeleccionarOpcion(sToolId)
''        mo_Alojados.idUsuario = ml_IdUsuarioAuditoria
''        mo_Alojados.lnIdTablaLISTBARITEMS = 1323
''        mo_Alojados.lcNombrePc = lc_NombrePc
''        Select Case mo_Alojados.Opcion
''        Case sghAgregar
''        Case sghModificar, sghConsultar, sghEliminar
''            Set rsHospitalizacion = ucAdmisionHospitalizacion.DataSource
''            If ucAdmisionHospitalizacion.idRegistroSeleccionado = 0 Then
''                MsgBox "Seleccione un registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            mo_Alojados.idAtencion = rsHospitalizacion!idAtencion
''        End Select
''        mo_Alojados.Show 1
''        Unload mo_Alojados
''End Sub
''
''
Sub EdicionReembolsos(sToolId As String)
        Dim oReembolsosDetalle As New ReembolsosDetalle
        If sToolId = "ID_HospitalizacionAlojamientoConjunto" Then
            oReembolsosDetalle.Opcion = sghAgregar
            oReembolsosDetalle.SoloSeIngresaUnaCuenta = True
        Else
            oReembolsosDetalle.Opcion = SeleccionarOpcion(sToolId)
            Select Case oReembolsosDetalle.Opcion
            Case sghAgregar
            Case sghModificar, sghConsultar, sghEliminar
                If ucReembolsosLista1.idRegistroSeleccionado = 0 Then
                    MsgBox "Seleccione un registro", vbInformation, Me.Caption
                    Exit Sub
                End If
                oReembolsosDetalle.IdFactReembolso = ucReembolsosLista1.idRegistroSeleccionado
            End Select
        End If
        oReembolsosDetalle.idUsuario = ml_IdUsuarioAuditoria
        oReembolsosDetalle.lnIdTablaLISTBARITEMS = 1331
        oReembolsosDetalle.lcNombrePc = lc_NombrePc
        oReembolsosDetalle.Show 1
        Unload oReembolsosDetalle
        ucReembolsosLista1.RealizarBusqueda
End Sub
''
''Sub EdicionMovimientoFormatoHC(sToolId As String)
''Dim mo_MovimientoFormatoHCDetalle As New MovimientoFormatoHCDetalle
''        mo_MovimientoFormatoHCDetalle.Opcion = SeleccionarOpcion(sToolId)
''        mo_MovimientoFormatoHCDetalle.idUsuario = ml_IdUsuarioAuditoria
''        mo_MovimientoFormatoHCDetalle.lnIdTablaLISTBARITEMS = 1330
''        mo_MovimientoFormatoHCDetalle.lcNombrePc = lc_NombrePc
''        Select Case mo_MovimientoFormatoHCDetalle.Opcion
''        Case sghAgregar
''        Case sghModificar, sghConsultar, sghEliminar
''            mo_MovimientoFormatoHCDetalle.IdMovimiento = Me.ucMovimientoFormatoHcLista1.idRegistroSeleccionado
''            If mo_MovimientoFormatoHCDetalle.IdMovimiento = -1 Or mo_MovimientoFormatoHCDetalle.IdMovimiento = 0 Then
''                MsgBox "Seleccione un registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''        End Select
''        mo_MovimientoFormatoHCDetalle.Icon = Me.Icon
''        mo_MovimientoFormatoHCDetalle.Show 1
''        Unload mo_MovimientoFormatoHCDetalle
''        Select Case sToolId
''        Case "ID_Agregar":
''        Case "ID_Modificar":
''        Case "ID_Consultar":
''        Case "ID_Eliminar":
''        End Select
''End Sub
''
Sub EdicionConstancias(sToolId As String)
  Dim mo_Constancias As New rptConstAtencion
  mo_Constancias.Opcion = SeleccionarOpcion(sToolId)
  mo_Constancias.idUsuario = ml_IdUsuarioAuditoria
  mo_Constancias.lnIdTablaLISTBARITEMS = 1325
  mo_Constancias.lcNombrePc = lc_NombrePc
  Select Case mo_Constancias.Opcion
    Case sghAgregar

    Case sghModificar, sghConsultar, sghEliminar
      mo_Constancias.IdMovimiento = ucContanciasAtencion.idRegistroSeleccionado
      mo_Constancias.Historia = ucContanciasAtencion.Historia
      mo_Constancias.idAtencion = ucContanciasAtencion.idAtencion
      mo_Constancias.idTipoConstancia = ucContanciasAtencion.idTipoConstancia
      mo_Constancias.Recibo = ucContanciasAtencion.Recibo
      mo_Constancias.Observaciones = ucContanciasAtencion.Observaciones
      mo_Constancias.IdServicio = ucContanciasAtencion.IdServicio
      If ucContanciasAtencion.idRegistroSeleccionado = -1 Or ucContanciasAtencion.idRegistroSeleccionado = 0 Then
        MsgBox "Seleccione un registro para consultar datos de la Constancia", vbInformation, Me.Caption
        Exit Sub
       End If
  End Select
  mo_Constancias.EjecutaFormulario
  ucContanciasAtencion.RealizarBusqueda
End Sub
''
Sub EdicionPacExtConSeguro(sToolId As String)
        Dim oFacGeneraCtaPacienteExtSeguro As New FacGeneraCtaPacienteExtSeguro
        oFacGeneraCtaPacienteExtSeguro.Opcion = SeleccionarOpcion(sToolId)
        oFacGeneraCtaPacienteExtSeguro.idUsuario = ml_IdUsuarioAuditoria
        oFacGeneraCtaPacienteExtSeguro.lnIdTablaLISTBARITEMS = 1339
        oFacGeneraCtaPacienteExtSeguro.lcNombrePc = lc_NombrePc
        Select Case oFacGeneraCtaPacienteExtSeguro.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            If ucPacienteExternos1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
            oFacGeneraCtaPacienteExtSeguro.idAtencion = ucPacienteExternos1.idRegistroSeleccionado
        End Select
        oFacGeneraCtaPacienteExtSeguro.Show 1
        Unload oFacGeneraCtaPacienteExtSeguro
        ucPacienteExternos1.RealizarBusqueda
End Sub
'
''Sub EdicionPacExtParticular(sToolId As String)
''        Dim oFacGeneraCtaPacienteExterno As New FacGeneraCtaPacienteExterno
''        oFacGeneraCtaPacienteExterno.Opcion = SeleccionarOpcion(sToolId)
''        oFacGeneraCtaPacienteExterno.idUsuario = ml_IdUsuarioAuditoria
''        oFacGeneraCtaPacienteExterno.lnIdTablaLISTBARITEMS = 1340
''        oFacGeneraCtaPacienteExterno.lcNombrePc = lc_NombrePc
''        oFacGeneraCtaPacienteExterno.idPuntoCarga = 6  'Consulta externa -admision
''        Select Case oFacGeneraCtaPacienteExterno.Opcion
''        Case sghAgregar
''        Case sghModificar, sghConsultar, sghEliminar
''            If ucPacienteExternos1.IdRegistroSeleccionado = 0 Then
''                MsgBox "Seleccione un registro", vbInformation, Me.Caption
''                Exit Sub
''            End If
''            oFacGeneraCtaPacienteExterno.idAtencion = ucPacienteExternos1.IdRegistroSeleccionado
''        End Select
''        oFacGeneraCtaPacienteExterno.Show 1
''        Unload oFacGeneraCtaPacienteExterno
''
''End Sub
''
Sub EdicionPaqueteServicio(sToolId As String)
Dim mo_FacCatalogoPaqueteDetalle As New SIGHProxies.clFactCatalogoPqteDetalle
    mo_FacCatalogoPaqueteDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_FacCatalogoPaqueteDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_FacCatalogoPaqueteDetalle.lnIdTablaLISTBARITEMS = 1341
    mo_FacCatalogoPaqueteDetalle.lcNombrePc = lc_NombrePc
    Select Case mo_FacCatalogoPaqueteDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        mo_FacCatalogoPaqueteDetalle.idFactPaquete = Me.ucFactPaquetesLista1.idRegistroSeleccionado
        If mo_FacCatalogoPaqueteDetalle.idFactPaquete = -1 Or mo_FacCatalogoPaqueteDetalle.idFactPaquete = 0 Then
            MsgBox "Seleccione un registro", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select
     mo_FacCatalogoPaqueteDetalle.MostrarFormulario
     Set mo_FacCatalogoPaqueteDetalle = Nothing

    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select
End Sub
''
Sub EdicionDespachoDonaciones(sToolId As String)
        Dim mo_DespachoDonaciones As New SighFarmacia.DespachoDonaciones
        Dim lcMovimiento As String
        lcMovimiento = Right("0" + Trim(Str(ucFarmDespachoDonaciones1.idRegistroSeleccionado)), 9)
        mo_DespachoDonaciones.Opcion = SeleccionarOpcion(sToolId)
        mo_DespachoDonaciones.idUsuario = ml_IdUsuarioAuditoria
        mo_DespachoDonaciones.lnIdTablaLISTBARITEMS = 1342
        mo_DespachoDonaciones.lcNombrePc = lc_NombrePc
        Select Case mo_DespachoDonaciones.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_DespachoDonaciones.movNumero = lcMovimiento
            If ucFarmDespachoDonaciones1.idRegistroSeleccionado = -1 Or ucFarmDespachoDonaciones1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_DespachoDonaciones.MostrarFormulario
        ucFarmDespachoDonaciones1.RealizarBusqueda
End Sub
'
'
''debb-hra
'Private Sub cmdFechaHoraServidor_Click()
'  'CentrarImagen
'
'  Dim lcBuscaParametro As New SIGHDatos.Parametros
'  status.Panels(1).Text = "      " & lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQL1
'  'status.Panels(7).Text = lcBuscaParametro.SeleccionaFilaParametro(314) & " " & lcBuscaParametro.RetornaVersionServidorSQLserver
'  'status.Panels(7).Width = 3400
'  Set lcBuscaParametro = Nothing
'End Sub
'
'
''Sub EdicionHisCE(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
''    If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ModuloHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub  'Frank0808
''    Dim oRcsTemp1 As New ADODB.Recordset
''    Set oRcsTemp1 = mo_ReglasHIS.ObtenerListaEstablecimientosMR
''    If oRcsTemp1.RecordCount = 0 Then
''        MsgBox "No ha registrado los establecimientos de la MicroRed", vbExclamation, Me.Caption
''        Exit Sub
''    End If
''    Dim mo_HISDetalle As New SIGHhisDigitacion.MantenimientoHIS
''    mo_HISDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_HISDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_HISDetalle.lcNombrePc = lc_NombrePc
''    mo_HISDetalle.IdEstablecimiento = ucHISListaAtencion.DevuelveIdEstablecimiento
''    mo_HISDetalle.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
''    Select Case mo_HISDetalle.Opcion
''    Case sghAgregar
''    Case sghModificar, sghConsultar, sghEliminar
''        mo_HISDetalle.IdRegistroHIS = Me.ucHISListaAtencion.idRegistroSeleccionado
''        If mo_HISDetalle.IdRegistroHIS = -1 Or mo_HISDetalle.IdRegistroHIS = 0 Then
''            MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''            Exit Sub
''        End If
''    End Select
''    mo_HISDetalle.MostrarFormulario
''    'Frank HIS
''    ucHISListaAtencion.RealizarBusqueda
''    Select Case sToolId
''        Case "ID_Agregar":
''        Case "ID_Modificar":
''        Case "ID_Consultar":
''        Case "ID_Eliminar":
''    End Select
''End Sub
'
''Sub EdicionHisDobleDigitacion(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
''    If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub
''    Dim oRcsTemp1 As New ADODB.Recordset
''    Set oRcsTemp1 = mo_ReglasHIS.ObtenerListaEstablecimientosMR
''    If oRcsTemp1.RecordCount = 0 Then
''        MsgBox "No ha registrado los establecimientos de la MicroRed", vbExclamation, Me.Caption
''        Exit Sub
''    End If
''    Dim mo_HISDetalle As New SIGHhisDigitacion.MantRegHisCalidad
''    mo_HISDetalle.Opcion = SeleccionarOpcion(sToolId)
''    mo_HISDetalle.idUsuario = ml_IdUsuarioAuditoria
''    mo_HISDetalle.lcNombrePc = lc_NombrePc
''    mo_HISDetalle.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
''    mo_HISDetalle.IdHisDetalle = UcHISCalidad.idRegistroSeleccionado
''
''    If mo_HISDetalle.IdHisDetalle = -1 Or mo_HISDetalle.IdHisDetalle = 0 Then
''        MsgBox "Seleccione un Registro", vbInformation, Me.Caption
''        Exit Sub
''    End If
''    Select Case mo_HISDetalle.Opcion
''    Case sghAgregar
''        If UcHISCalidad.Registrado = 1 Then
''            MsgBox "Seleccione la opción Modificar(F3) para editar la doble digitación", vbInformation, Me.Caption
''            Exit Sub
''        End If
''        mo_HISDetalle.MostrarFormulario
''        UcHISCalidad.CargarListaGenerados
''    Case sghModificar, sghConsultar
''        If UcHISCalidad.Registrado = -1 Or UcHISCalidad.Registrado = 0 Then
''            MsgBox "Seleccione la opción Agregar(F2) para registrar la doble digitación", vbInformation, Me.Caption
''            Exit Sub
''        End If
''        mo_HISDetalle.MostrarFormulario
''        UcHISCalidad.CargarListaGenerados
''    Case sghEliminar
''        MsgBox "No puedes eliminar el registro para la doble digitación", vbInformation, Me.Caption
''        Exit Sub
''    End Select
''    'Frank HIS
''    Select Case sToolId
''        Case "ID_Agregar":
''        Case "ID_Modificar":
''        Case "ID_Consultar":
''        Case "ID_Eliminar":
''    End Select
''End Sub
'
'
''JVG - Muestra el formulario de edicion del los Lotes HIS en el sistema
'Sub EdicionHisLotesCE(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
''If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub
''Dim mo_HISLotes As New SIGHhisDigitacion.MantenimientoHISLotes
''
''mo_HISLotes.Opcion = SeleccionarOpcion(sToolId)
''mo_HISLotes.idUsuario = ml_IdUsuarioAuditoria
''mo_HISLotes.lcNombrePc = lc_NombrePc
''mo_HISLotes.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
''mo_HISLotes.IdEstablecimiento = Me.ucHISListaLotes.DevuelveIdEstablecimiento
''
''Select Case mo_HISLotes.Opcion
''Case sghAgregar
''Case sghModificar, sghConsultar, sghEliminar
''    mo_HISLotes.IdRegistroLote = Me.ucHISListaLotes.idRegistroSeleccionado
''    'mo_HISLotes.IdEstablecimiento = Me.ucHISListaLotes.IdEstablecimiento
''    If mo_HISLotes.IdRegistroLote = -1 Or mo_HISLotes.IdRegistroLote = 0 Then
''        MsgBox "Seleccion un Registro", vbInformation, Me.Caption
''        Exit Sub
''    End If
''End Select
''mo_HISLotes.MostrarFormulario
'''Frank HIS
''ucHISListaLotes.RealizarBusqueda
'''Unload mo_HISDetalle
''
''Select Case sToolId
''Case "ID_Agregar":
''Case "ID_Modificar":
''Case "ID_Consultar":
''Case "ID_Eliminar":
''End Select
'
'End Sub
'
'Sub EdicionProgramacionHIS(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
'If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub
'Select Case SeleccionarOpcion(sToolId)
'Case sghAgregar
'    Me.ucHISListaProgramacion.AgregarProgramacion
'Case sghEliminar
'    Me.ucHISListaProgramacion.EliminarProgramacion
'Case sghModificar
'    Me.ucHISListaProgramacion.ModificarProgramacion sghModificar
'Case sghConsultar
'    Me.ucHISListaProgramacion.ModificarProgramacion sghConsultar
'
''    mo_HISLotes.IdRegistroLote = Me.ucHISListaLotes.idRegistroSeleccionado
''    'mo_HISLotes.IdEstablecimiento = Me.ucHISListaLotes.IdEstablecimiento
''    If mo_HISLotes.IdRegistroLote = -1 Or mo_HISLotes.IdRegistroLote = 0 Then
''        MsgBox "Seleccion un Registro", vbInformation, Me.Caption
''        Exit Sub
''    End If
'End Select
'
'Select Case sToolId
'Case "ID_Agregar":
'Case "ID_Modificar":
'Case "ID_Consultar":
'Case "ID_Eliminar":
'End Select
'
'End Sub
'
Sub EdicionReceta(sToolId As String, lnIdListBarItems As Long, lnIdTipoServicio As Long)
    Dim oRecetaDetalle As New RecetaDetalle
    Select Case sToolId
    Case "ID_Agregar":
        oRecetaDetalle.Opcion = sghAgregar
    Case "ID_Modificar":
        oRecetaDetalle.Opcion = sghModificar
        If Me.ucRecetasLista1.TieneRegistros = False Then Exit Sub
        oRecetaDetalle.idCuentaAtencion = Me.ucRecetasLista1.idRegistroSeleccionado
        oRecetaDetalle.FechaReceta = Me.ucRecetasLista1.FechaReceta
    Case "ID_Consultar":
        oRecetaDetalle.Opcion = sghConsultar
        If Me.ucRecetasLista1.TieneRegistros = False Then Exit Sub
        oRecetaDetalle.idCuentaAtencion = Me.ucRecetasLista1.idRegistroSeleccionado
        oRecetaDetalle.FechaReceta = Me.ucRecetasLista1.FechaReceta
    Case "ID_Eliminar":
        oRecetaDetalle.Opcion = sghEliminar
        If Me.ucRecetasLista1.TieneRegistros = False Then Exit Sub
        oRecetaDetalle.idCuentaAtencion = Me.ucRecetasLista1.idRegistroSeleccionado
        oRecetaDetalle.FechaReceta = Me.ucRecetasLista1.FechaReceta
    End Select
    oRecetaDetalle.idUsuario = ml_IdUsuarioAuditoria
    oRecetaDetalle.lcNombrePc = lc_NombrePc
    oRecetaDetalle.lnIdTablaLISTBARITEMS = lnIdListBarItems
    oRecetaDetalle.idTipoServicio = lnIdTipoServicio
    oRecetaDetalle.Show 1
    Set oRecetaDetalle = Nothing
End Sub
'
''debb 26/7/12
Sub EdicionFua(sToolId As String)
        Dim oSisFua As New SIGHSis.clFUA
        Select Case sToolId
        Case "ID_Agregar":
           oSisFua.Opcion = sghAgregar
        Case "ID_Modificar":
           oSisFua.Opcion = sghModificar
           oSisFua.idCuentaAtencion = Me.UcSISfuaLista1.idRegistroSeleccionado
        Case "ID_Consultar":
           oSisFua.Opcion = sghConsultar
           oSisFua.idCuentaAtencion = Me.UcSISfuaLista1.idRegistroSeleccionado
        Case "ID_Eliminar":
           oSisFua.Opcion = sghEliminar
           oSisFua.idCuentaAtencion = Me.UcSISfuaLista1.idRegistroSeleccionado
        End Select
       oSisFua.idUsuario = ml_IdUsuarioAuditoria
       oSisFua.lcNombrePc = lc_NombrePc
       oSisFua.lnIdTablaLISTBARITEMS = 1345
       oSisFua.IdServicio = 0 'Al colocar cero el FUA seleccionado sera el del registro
       oSisFua.FuaVersionFormato = Me.UcSISfuaLista1.FuaVersionFormato
       oSisFua.FuaTipoAnexo2015 = IIf(Me.UcSISfuaLista1.FuaTipoAnexo2015 = "", 0, UcSISfuaLista1.FuaTipoAnexo2015)
       oSisFua.MostrarFormulario
       Set oSisFua = Nothing
End Sub
'
Sub EdicionTipoTarifa(sToolId As String)
    Dim mo_TiposTarifaDetalle As New SIGHCatalogos.clTiposTarifaDetalle
    mo_TiposTarifaDetalle.Opcion = SeleccionarOpcion(sToolId)
    mo_TiposTarifaDetalle.idUsuario = ml_IdUsuarioAuditoria
    mo_TiposTarifaDetalle.lnIdTablaLISTBARITEMS = 1337
    mo_TiposTarifaDetalle.lcNombrePc = lc_NombrePc
    Select Case mo_TiposTarifaDetalle.Opcion
    Case sghAgregar
    Case sghModificar, sghConsultar, sghEliminar
        mo_TiposTarifaDetalle.idTipoTarifa = Me.ucTiposTarifaLista1.idRegistroSeleccionado
        If mo_TiposTarifaDetalle.idTipoTarifa = -1 Or mo_TiposTarifaDetalle.idTipoTarifa = 0 Then
            MsgBox "Seleccione un registro", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select
    mo_TiposTarifaDetalle.MostrarFormulario
    Me.ucTiposTarifaLista1.RealizarBusqueda
    Set mo_TiposTarifaDetalle = Nothing
    Select Case sToolId
    Case "ID_Agregar":
    Case "ID_Modificar":
    Case "ID_Consultar":
    Case "ID_Eliminar":
    End Select

End Sub
'
'
''JVG - Muestra el formulario de edicion los Establecimientos de la MicroRed
'Sub EdicionHisEstablecimientos(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
''If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub
''Dim mo_HISEstabMR As New SIGHhisDigitacion.MantenimientoHISEstMR
''
''mo_HISEstabMR.Opcion = SeleccionarOpcion(sToolId)
''mo_HISEstabMR.idUsuario = ml_IdUsuarioAuditoria
''mo_HISEstabMR.lcNombrePc = lc_NombrePc
''mo_HISEstabMR.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
''
''Select Case mo_HISEstabMR.Opcion
''Case sghAgregar
''Case sghModificar, sghConsultar, sghEliminar
''    mo_HISEstabMR.IdEstablecimiento = Me.ucHISEstablecimientos.idRegistroSeleccionado
''    mo_HISEstabMR.NombreEstablecimiento = Me.ucHISEstablecimientos.NombreEstablecimiento
''    mo_HISEstabMR.CodigoEstablecimiento = Me.ucHISEstablecimientos.CodigoEstablecimiento
''
''    If mo_HISEstabMR.IdEstablecimiento = -1 Or mo_HISEstabMR.IdEstablecimiento = 0 Then
''        MsgBox "Seleccion un Registro", vbInformation, Me.Caption
''        Exit Sub
''    End If
''End Select
''mo_HISEstabMR.MostrarFormulario
'''Frank HIS
''ucHISEstablecimientos.RealizarBusqueda
'End Sub
'
'Sub EdicionPadronNominal(sToolId As String, lnIdTablaLISTBARITEMS As Long, ml_IdUsuarioAuditoria As Long, lc_NombrePc As String)
'
'If sToolId = "ID_ExportaHIS" Or sToolId = "ID_ExportaURENIS" Then Exit Sub
'
'Dim mo_DetallePadronInicial As New SIGHhisDigitacion.MantenimientoPN
'mo_DetallePadronInicial.Opcion = SeleccionarOpcion(sToolId)
'mo_DetallePadronInicial.idUsuario = ml_IdUsuarioAuditoria
'mo_DetallePadronInicial.lcNombrePc = lc_NombrePc
'mo_DetallePadronInicial.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS
'
'Select Case mo_DetallePadronInicial.Opcion
'Case sghAgregar
'Case sghModificar, sghConsultar, sghEliminar
'    mo_DetallePadronInicial.IdPadNominal = Me.UcHISPadronNominal.idRegistroSeleccionado
'    If mo_DetallePadronInicial.IdPadNominal = -1 Or mo_DetallePadronInicial.IdPadNominal = 0 Then
'        MsgBox "Seleccion un Registro", vbInformation, Me.Caption
'        Exit Sub
'    End If
'End Select
'mo_DetallePadronInicial.MostrarFormulario
'
'
'Select Case sToolId
'Case "ID_Agregar":
'Case "ID_Modificar":
'Case "ID_Consultar":
'Case "ID_Eliminar":
'End Select
'End Sub
'
Sub EdicionMantenedorFarmacia(sToolId As String)
        Dim mo_FarmAlmacen As New SighFarmacia.clAlmacen

        mo_FarmAlmacen.Opcion = SeleccionarOpcion(sToolId)
        mo_FarmAlmacen.idUsuario = ml_IdUsuarioAuditoria
        mo_FarmAlmacen.lnIdTablaLISTBARITEMS = 1355
        mo_FarmAlmacen.lcNombrePc = lc_NombrePc
        Select Case mo_FarmAlmacen.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_FarmAlmacen.IdDependenciaExt = ucFarmAlmacenes1.idRegistroSeleccionado
            If ucFarmAlmacenes1.idRegistroSeleccionado = -1 Or ucFarmAlmacenes1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_FarmAlmacen.MostrarFormulario
        ucFarmAlmacenes1.RealizarBusqueda
End Sub
'
''mgaray201411f
'Sub EdicionTipoModalidadSala(sToolId As String)
'    Dim mo_tipoModalidadSala As New SIGHCatalogos.clImagTipoModalidadSala
'
'    mo_tipoModalidadSala.Opcion = SeleccionarOpcion(sToolId)
'    mo_tipoModalidadSala.idUsuario = ml_IdUsuarioAuditoria
'    mo_tipoModalidadSala.lnIdTablaLISTBARITEMS = 1359
'    mo_tipoModalidadSala.lcNombrePc = lc_NombrePc
'    Select Case mo_tipoModalidadSala.Opcion
'    Case sghAgregar
'    Case sghModificar, sghConsultar, sghEliminar
'        mo_tipoModalidadSala.IdTipoModalidadSala = Me.ucImagTipoModalidadSala1.idRegistroSeleccionado
'        If mo_tipoModalidadSala.IdTipoModalidadSala = -1 Or mo_tipoModalidadSala.IdTipoModalidadSala = 0 Then
'            MsgBox "Seleccione un registro", vbInformation, Me.Caption
'            Exit Sub
'        End If
'    End Select
'
'    mo_tipoModalidadSala.MostrarFormulario
'    If mo_tipoModalidadSala.ResultadoOperacion = True Then
'        Me.ucImagTipoModalidadSala1.RealizarBusqueda
'    End If
'    Set mo_cajaDetalle = Nothing
'
'    Select Case sToolId
'    Case "ID_Agregar":
'    Case "ID_Modificar":
'    Case "ID_Consultar":
'    Case "ID_Eliminar":
'    End Select
'
'End Sub
'
'Sub EdicionSala(sToolId As String)
'    'buscar ImagSala cuando se cree los controles
'    Dim mo_sala As New SIGHCatalogos.clImagSala
'
'    mo_sala.Opcion = SeleccionarOpcion(sToolId)
'    mo_sala.idUsuario = ml_IdUsuarioAuditoria
'    mo_sala.lnIdTablaLISTBARITEMS = 1360
'    mo_sala.lcNombrePc = lc_NombrePc
'    Select Case mo_sala.Opcion
'    Case sghAgregar
'    Case sghModificar, sghConsultar, sghEliminar
'        mo_sala.idSala = Me.ucImagSala1.idRegistroSeleccionado
'        If mo_sala.idSala = -1 Or mo_sala.idSala = 0 Then
'            MsgBox "Seleccione un registro", vbInformation, Me.Caption
'            Exit Sub
'        End If
'    End Select
'
'    mo_sala.MostrarFormulario
'    If mo_sala.ResultadoOperacion = True Then
'        Me.ucImagSala1.RealizarBusqueda
'    End If
'    Set mo_sala = Nothing
'
'    Select Case sToolId
'    Case "ID_Agregar":
'    Case "ID_Modificar":
'    Case "ID_Consultar":
'    Case "ID_Eliminar":
'    End Select
'
'End Sub
'
'Sub EdicionImagFactCatalogoServiciosDuracion(sToolId As String)
'    'buscar ImagSala cuando se cree los controles
'    Dim mo_FactCatalogoServicioDuracion As New SIGHCatalogos.clCatalgoServicioDuracion
'
'    mo_FactCatalogoServicioDuracion.Opcion = SeleccionarOpcion(sToolId)
'    mo_FactCatalogoServicioDuracion.idUsuario = ml_IdUsuarioAuditoria
'    mo_FactCatalogoServicioDuracion.lnIdTablaLISTBARITEMS = 1361
'    mo_FactCatalogoServicioDuracion.lcNombrePc = lc_NombrePc
'    Select Case mo_FactCatalogoServicioDuracion.Opcion
'    Case sghAgregar
'        MsgBox "No se Pueden agregar Servicios desde esta interfaz", vbInformation, "Imagenológia"
'        Exit Sub
'    Case sghModificar, sghConsultar:
'        mo_FactCatalogoServicioDuracion.idProducto = Me.ucImagCatalgoServicioDuracion1.idRegistroSeleccionado
'        If mo_FactCatalogoServicioDuracion.idProducto = -1 Or mo_FactCatalogoServicioDuracion.idProducto = 0 Then
'            MsgBox "Seleccione un registro", vbInformation, Me.Caption
'            Exit Sub
'        End If
'    Case sghEliminar:
'        MsgBox "No se Pueden eliminar Servicios desde esta interfaz", vbInformation, "Imagenológia"
'        Exit Sub
'    End Select
'
'    mo_FactCatalogoServicioDuracion.MostrarFormulario
'    If mo_FactCatalogoServicioDuracion.ResultadoOperacion = True Then
'        Me.ucImagCatalgoServicioDuracion1.RealizarBusqueda
'    End If
'    Set mo_FactCatalogoServicioDuracion = Nothing
'
'    Select Case sToolId
'    Case "ID_Agregar":
'    Case "ID_Modificar":
'    Case "ID_Consultar":
'    Case "ID_Eliminar":
'    End Select
'End Sub
'
'Sub EdicionIntegracionSistema(sToolId As String)
'    'buscar ImagSala cuando se cree los controles
'    Dim mo_InteoIntegracionSistema As New SIGHCatalogos.clInteoIntegracionSistema
'
'    mo_InteoIntegracionSistema.Opcion = SeleccionarOpcion(sToolId)
'    mo_InteoIntegracionSistema.idUsuario = ml_IdUsuarioAuditoria
'    mo_InteoIntegracionSistema.lnIdTablaLISTBARITEMS = 1362
'    mo_InteoIntegracionSistema.lcNombrePc = lc_NombrePc
'    Select Case mo_InteoIntegracionSistema.Opcion
'    Case sghAgregar
'    Case sghModificar, sghConsultar, sghEliminar
'        mo_InteoIntegracionSistema.IdIntegracionSistema = Me.ucInteoIntegracionSistema1.idRegistroSeleccionado
'        If mo_InteoIntegracionSistema.IdIntegracionSistema = -1 Or mo_InteoIntegracionSistema.IdIntegracionSistema = 0 Then
'            MsgBox "Seleccione un registro", vbInformation, Me.Caption
'            Exit Sub
'        End If
'    End Select
'
'    mo_InteoIntegracionSistema.MostrarFormulario
'    If mo_InteoIntegracionSistema.ResultadoOperacion = True Then
'        Me.ucInteoIntegracionSistema1.RealizarBusqueda
'    End If
'    Set mo_InteoIntegracionSistema = Nothing
'
'    Select Case sToolId
'    Case "ID_Agregar":
'    Case "ID_Modificar":
'    Case "ID_Consultar":
'    Case "ID_Eliminar":
'    End Select
'End Sub
'
'
''debb2014b
Sub EdicionMantenedorHistoricoPrecios(sToolId As String)
        Dim mo_FarmHistPrecio As New SighFarmacia.clFarmHistPrecios

        mo_FarmHistPrecio.Opcion = SeleccionarOpcion(sToolId)
        mo_FarmHistPrecio.idUsuario = ml_IdUsuarioAuditoria
        mo_FarmHistPrecio.lnIdTablaLISTBARITEMS = 1359
        mo_FarmHistPrecio.lcNombrePc = lc_NombrePc
        Select Case mo_FarmHistPrecio.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_FarmHistPrecio.IdFarmHistPrecio = ucFarmHpreciosLista1.idRegistroSeleccionado
            If ucFarmHpreciosLista1.idRegistroSeleccionado = -1 Or ucFarmHpreciosLista1.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_FarmHistPrecio.MostrarFormulario
        ucFarmHpreciosLista1.RealizarBusqueda
End Sub
'
'Sub OcultarOpcionesPacticularesMenuPorEstablecimiento()
''toolbar.Index
'End Sub
'
''mgaray201504
'Private Function UsuarioEsCajero(mb_UsuarioAccesoGestionCaja As Boolean) As Boolean
'    UsuarioEsCajero = False
'
'    If mb_UsuarioAccesoGestionCaja = True Then
'        Dim oRsPermisos As New Recordset
'        Dim lbUsuarioRealizaApertura As Boolean
'
'        Set oRsPermisos = mo_AdminSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_IdUsuarioAuditoria)
'        If oRsPermisos.RecordCount > 0 Then
'           Do While Not oRsPermisos.EOF
'              Select Case oRsPermisos.Fields!IdPermiso
'              Case 201    'Caja - Realizar Apertura
'                   UsuarioEsCajero = True
'              End Select
'              oRsPermisos.MoveNext
'           Loop
'
'        End If
'        Set oRsPermisos = Nothing
'    End If
'
'End Function
'
''FRANK MAYO
Sub EdicionCajaNotaCredito(sToolId As String)
        Dim mo_CajaApruebaNotaCredito As New CajaApruebaNotaCredito
        Dim orsNotasCredito As New Recordset
        mo_CajaApruebaNotaCredito.idUsuario = ml_IdUsuarioAuditoria
        mo_CajaApruebaNotaCredito.Opcion = SeleccionarOpcion(sToolId)
        mo_CajaApruebaNotaCredito.lnIdTablaLISTBARITEMS = 1206
        mo_CajaApruebaNotaCredito.lcNombrePc = lc_NombrePc
        mo_CajaApruebaNotaCredito.idTipoNota = 2 'NOTA CREDITO
        Select Case mo_CajaApruebaNotaCredito.Opcion
        Case sghAgregar
            'mo_AdmisionHospDetalle.TipoServicio = sghEmergenciaConsultorios
        Case sghModificar, sghConsultar, sghEliminar
            Set orsNotasCredito = ucCajaNotaCredito1.DataSource
            If orsNotasCredito.State = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
            If orsNotasCredito.RecordCount = 0 Then
                MsgBox "No existen registros", vbInformation, Me.Caption
                Exit Sub
            End If
            If ucCajaNotaCredito1.idRegistroSeleccionado = 0 Then
                Exit Sub
            End If
            Set orsNotasCredito = Nothing
            mo_CajaApruebaNotaCredito.idRegistroSeleccionado = Me.ucCajaNotaCredito1.idRegistroSeleccionado
            If mo_CajaApruebaNotaCredito.idRegistroSeleccionado = -1 Or mo_CajaApruebaNotaCredito.idRegistroSeleccionado = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select
        mo_CajaApruebaNotaCredito.Show 1
End Sub
'
'
'Sub AgregaAfiliadosSIS(lcParametro313 As String, lcConexionExterna As String)
'        Dim nitem As String
'        Const lcArchivoExcel As String = "Afiliados.xlsx"
'        nitem = InputBox("Afiliados SIS en arhivo Excel: " & lcParametro313 & lcArchivoExcel & "  (Hoja1)", "")
'
'
'        On Error GoTo Error_AgregaAfiliadosSIS
'
'        Me.MousePointer = 11
'        Dim oConexionExterna As New Connection
'        Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
'        Dim ms_MensajeError As String
'        oConexionExterna.CommandTimeout = 300
'        oConexionExterna.CursorLocation = adUseClient
'        oConexionExterna.Open lcConexionExterna
'
'
'        Dim lcRango As String, lnFila As Long
'        Dim lcIdSiaSis As Long, lcCodigo As String, lcCdisa As String, lcCformato As String, lcCnumero As String
'        Dim lcAfiliacionNroIntegrante As String, lcTipoDocumento As String, lcCodigoEstablAdscripcion As String
'        Dim lcAfiliacionFecha As String, lcApPaterno As String, lcApMaterno As String, lcPnombre As String
'        Dim lcSnombre As String, lcSexo As String, lcFnacimiento As String, lcDistritoDomicilio As String
'        Dim lcEstadoSis As String, lcFbajaok As String, lcDNI As String, lcMotivoBaja As String
'        Dim ldAfiliacionFecha As Date, ldFNacimiento As Date, ldFbajaOk As Date
'        ' crea el objeto Excel
'        Dim ObjExcel As Object
'        Set ObjExcel = CreateObject("Excel.Application")
'        'abre el libro
'        ObjExcel.Workbooks.Open lcParametro313 & lcArchivoExcel
'        With ObjExcel
'            Dim SheetO As Object
'            Set SheetO = .Sheets("Hoja1")
'            '
'            lnFila = 2
'            Do While True
'                    lcRango = "A" & Trim(Str(lnFila))
'                    lcIdSiaSis = SheetO.range(lcRango).Value
'
'                    If Not lcIdSiaSis > 0 Then
'                       Exit Do
'                    End If
'                    lcRango = "B" & Trim(Str(lnFila))
'                    lcCodigo = SheetO.range(lcRango).Value
'                    lcRango = "C" & Trim(Str(lnFila))
'                    lcCdisa = SheetO.range(lcRango).Value
'                    lcRango = "D" & Trim(Str(lnFila))
'                    lcCformato = SheetO.range(lcRango).Value
'                    lcRango = "E" & Trim(Str(lnFila))
'                    lcCnumero = SheetO.range(lcRango).Value
'                    lcRango = "F" & Trim(Str(lnFila))
'                    lcAfiliacionNroIntegrante = SheetO.range(lcRango).Value
'                    lcRango = "G" & Trim(Str(lnFila))
'                    lcTipoDocumento = SheetO.range(lcRango).Value
'                    lcRango = "H" & Trim(Str(lnFila))
'                    lcCodigoEstablAdscripcion = SheetO.range(lcRango).Value
'                    lcRango = "I" & Trim(Str(lnFila))
'                    lcAfiliacionFecha = SheetO.range(lcRango).Value
'                    ldAfiliacionFecha = 0
'                    If IsDate(lcAfiliacionFecha) Then
'                       ldAfiliacionFecha = CDate(lcAfiliacionFecha)
'                    End If
'                    lcRango = "J" & Trim(Str(lnFila))
'                    lcApPaterno = SheetO.range(lcRango).Value
'                    lcRango = "K" & Trim(Str(lnFila))
'                    lcApMaterno = SheetO.range(lcRango).Value
'                    lcRango = "L" & Trim(Str(lnFila))
'                    lcPnombre = SheetO.range(lcRango).Value
'                    lcRango = "M" & Trim(Str(lnFila))
'                    lcSnombre = SheetO.range(lcRango).Value
'                    lcRango = "N" & Trim(Str(lnFila))
'                    lcSexo = SheetO.range(lcRango).Value
'                    lcRango = "O" & Trim(Str(lnFila))
'                    lcFnacimiento = SheetO.range(lcRango).Value
'                    ldFNacimiento = 0
'                    If IsDate(lcFnacimiento) Then
'                       ldFNacimiento = CDate(lcFnacimiento)
'                    End If
'                    lcRango = "P" & Trim(Str(lnFila))
'                    lcDistritoDomicilio = SheetO.range(lcRango).Value
'                    lcRango = "Q" & Trim(Str(lnFila))
'                    lcEstadoSis = SheetO.range(lcRango).Value
'                    lcRango = "R" & Trim(Str(lnFila))
'                    lcFbajaok = SheetO.range(lcRango).Value
'                    ldFbajaOk = 0
'                    If IsDate(lcFbajaok) Then
'                       ldFbajaOk = CDate(lcFbajaok)
'                    End If
'                    lcRango = "S" & Trim(Str(lnFila))
'                    lcDNI = SheetO.range(lcRango).Value
'                    lcRango = "T" & Trim(Str(lnFila))
'                    lcMotivoBaja = SheetO.range(lcRango).Value
'                    ms_MensajeError = mo_ReglasSISgalenhos.SisFiliacionesBuscaYactualizaDatosXafiliado(oConexionExterna, _
'                                                    Val(lcIdSiaSis), _
'                                                    lcCodigo, _
'                                                    lcCdisa, _
'                                                    lcCformato, _
'                                                    lcCnumero, _
'                                                    lcAfiliacionNroIntegrante, _
'                                                    lcTipoDocumento, _
'                                                    lcCodigoEstablAdscripcion, _
'                                                    ldAfiliacionFecha, _
'                                                    lcApPaterno, _
'                                                    lcApMaterno, _
'                                                    lcPnombre, _
'                                                    lcSnombre, _
'                                                    lcSexo, _
'                                                    ldFNacimiento, _
'                                                    lcDistritoDomicilio, _
'                                                    lcEstadoSis, _
'                                                    ldFbajaOk, _
'                                                    lcDNI, _
'                                                    lcMotivoBaja)
'                   lnFila = lnFila + 1
'         Loop
'      End With
'      MsgBox "Proceso en forma correcta", vbInformation, ""
'Error_AgregaAfiliadosSIS:
'      If Err.Number <> 0 Then
'         MsgBox Err.Description
'      End If
'      oConexionExterna.Close
'      Set oConexionExterna = Nothing
'      Set mo_ReglasSISgalenhos = Nothing
'      Me.MousePointer = 1
'      Exit Sub
'      Resume
'End Sub
'

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lcOpcionGalenhos As String, ml_IdListItem As Long, lbSiges As Boolean
    Dim rsPermisos As New Recordset
    lcOpcionGalenhos = ""
    Select Case Button.Key
    Case "Agregar", "Modificar", "Eliminar", "Consultar", "AltaMedica", "reembolsoXcuenta", "VisitaEnfermera"
         lbSiges = True
         ml_IdListItem = 0
         Set rsPermisos = mo_AdminSeguridad.ListbarItemsSeleccionarPorClave(ms_ListBarItemClave)
         If rsPermisos.RecordCount > 0 Then
            ml_IdListItem = rsPermisos!IdListItem
         End If
         rsPermisos.Close
         Set rsPermisos = mo_AdminSeguridad.RolesItemsSeleccionarPermisosPorEmpleadoYListItem(sighentidades.USUARIO, ml_IdListItem)
         If rsPermisos.RecordCount = 0 Then
            lbSiges = False
         Else
            Select Case Button
            Case "Agregar"
                If IsNull(rsPermisos!Agregar) Then
                   lbSiges = False
                ElseIf rsPermisos!Agregar = 0 Then
                   lbSiges = False
                End If
            Case "Modificar"
                If IsNull(rsPermisos!Modificar) Then
                   lbSiges = False
                ElseIf rsPermisos!Modificar = 0 Then
                   lbSiges = False
                End If
            Case "Eliminar"
                If IsNull(rsPermisos!Eliminar) Then
                   lbSiges = False
                ElseIf rsPermisos!Eliminar = 0 Then
                   lbSiges = False
                End If
            Case "Consultar"
                If IsNull(rsPermisos!Consultar) Then
                   lbSiges = False
                ElseIf rsPermisos!Consultar = 0 Then
                   lbSiges = False
                End If
            End Select
         End If
         rsPermisos.Close
         '
         If lbSiges Then
            lcOpcionGalenhos = IIf(Button.Key = "Agregar", "ID_Agregar", _
                               IIf(Button.Key = "Modificar", "ID_Modificar", _
                               IIf(Button.Key = "Eliminar", "ID_Eliminar", _
                               IIf(Button.Key = "reembolsoXcuenta", "ID_HospitalizacionAlojamientoConjunto", _
                               IIf(ms_ModuloSeleccionado = "AdmisionConsultorioEmerg" And Button.Key = "VisitaEnfermera", "ID_HospitalizacionVisitaEnfermera", _
                               IIf(ms_ModuloSeleccionado = "AdmisionHospitalizacion" And Button.Key = "VisitaEnfermera", "ID_HospitalizacionVisitaEnfermera", _
                               IIf(ms_ModuloSeleccionado = "AdmisionConsultorioEmerg" And Button.Key = "AltaMedica", "ID_HospitalizacionAlojamientoConjunto", _
                               IIf(ms_ModuloSeleccionado = "AdmisionHospitalizacion" And Button.Key = "AltaMedica", "ID_HospitalizacionAlojamientoConjunto", _
                               "ID_Consultar"))))))))
            Select Case ms_ModuloSeleccionado
            Case "PacienteCE"
                 EdicionPaciente lcOpcionGalenhos, sghConsultaExterna, 101
            Case "AdmisionCE"
                 EdicionCitas lcOpcionGalenhos
            Case "AtencionesTriaje"
                EdicionTriaje lcOpcionGalenhos
            Case "AtencionesCE"
                EdicionAdmisionCE lcOpcionGalenhos, sghConsultaExterna, 103
            Case "RecetasCE"
                EdicionReceta lcOpcionGalenhos, 1366, sghConsultaExterna
            Case "idConsultorioAsignado"
                EdicionArchiveroServicio lcOpcionGalenhos

            Case "AdmisionConsultorioEmerg"
                EdicionAdmisionEmergencia lcOpcionGalenhos
            Case "CamasEmergencia"
                EdicionCamas lcOpcionGalenhos, True
            Case "RecetasE"
                EdicionReceta lcOpcionGalenhos, 1343, sghEmergenciaConsultorios
                
            Case "AdmisionHospitalizacion"
                EdicionAdmisionHospitalizacion lcOpcionGalenhos
            Case "CamasHospitalizacion"
                EdicionCamas lcOpcionGalenhos, False

            Case "Programacion"
                EdicionProgMedica lcOpcionGalenhos
            Case "Turno"
                EdicionTurno lcOpcionGalenhos
            Case "Medico"
                EdicionMedico lcOpcionGalenhos

            Case "HistoriaClinica"
                EdicionHistoriaClinica lcOpcionGalenhos
            Case "MovimientoHistoria"
                EdicionMovimientoHistorias lcOpcionGalenhos
            Case "SolicitudHistorias"
               'EdicionSolicitudHistorias lcOpcionGalenhos
            Case "Archivero"
                EdicionArchiveroServicio lcOpcionGalenhos
                 
                
                 
            Case "EstadoCuenta"
            Case "FacturacionGeneral"
                 EdicionOrdenesServicio lcOpcionGalenhos
            Case "FactReembolsos"
                EdicionReembolsos lcOpcionGalenhos
            Case "PacExtConSeguro"
                EdicionPacExtConSeguro lcOpcionGalenhos
                 
                 
            Case "CatalogoBienes"
                 EdicionCatalogoBaseBienesInsumos lcOpcionGalenhos
            Case "TipoTarifa"
                EdicionTipoTarifa lcOpcionGalenhos
            Case "FacturacionCatalogoServicios"
                EdicionCatalogoBaseServicios lcOpcionGalenhos
            Case "TiposFinanciamiento"
                EdicionTiposFinanciamiento lcOpcionGalenhos
            Case "FacturacionPartidas"
                EdicionPartidaPresupuestal lcOpcionGalenhos
            Case "PqteServicio"
                EdicionPaqueteServicio lcOpcionGalenhos
            Case "ConfiguracionResLab"
                EdicionConfiguracionResLab lcOpcionGalenhos
            Case "FuentesFinanciamiento"
                EdicionFuentesFinanciamiento lcOpcionGalenhos
            
            Case "Servicios"
                EdicionServicio lcOpcionGalenhos
            Case "Diagnosticos"
                EdicionDiagnosticos lcOpcionGalenhos
            Case "Especialidades"
                EdicionEspecialidades lcOpcionGalenhos
            Case "EstablecimientosNoMinsa"
                EdicionEstablecimientosNoMinsa lcOpcionGalenhos
                
            Case "OrdenesLaboratorio"
                EdicionLaboratorio lcOpcionGalenhos
            Case "OrdenesPatologia"
                EdicionOrdenesServicioAnatomiaPatologia_ lcOpcionGalenhos
            Case "BS"
                EdicionOrdenesBS_ lcOpcionGalenhos
                
                
            Case "ImagRayosX"
                EdicionRayosX lcOpcionGalenhos
            Case "ImagEcografiaG"
                EdicionImagEcografiaGen lcOpcionGalenhos
            Case "ImagTomografia"
                EdicionImagTomografia lcOpcionGalenhos
            Case "ImagEcografiaO"
                EdicionImagEcografiaObs lcOpcionGalenhos
            Case "prgImagen"
                EdicionSiProgramacion lcOpcionGalenhos
            Case "FacturacionImagenologia"
                EdicionSiCitas lcOpcionGalenhos
                
            
            
            
            
            Case "Cajas"
                 EdicionCaja lcOpcionGalenhos
            Case "NotaCredito"
                 EdicionCajaNotaCredito lcOpcionGalenhos
            Case "GestionCaja"
                ucGestionCaja1.idUsuario = ml_IdUsuarioAuditoria
                ucGestionCaja1.NombreCajero = mo_AdminCaja.SeleccionaDatosCajero(sighentidades.USUARIO, sghUsuario)
                ucGestionCaja1.lnIdTablaLISTBARITEMS = 702
                ucGestionCaja1.lcNombrePc = lc_NombrePc
                ConfigurarControl ucGestionCaja1
            Case "FarmAlmacen"
                 EdicionMantenedorFarmacia lcOpcionGalenhos
            Case "Inventario"
                 EdicionInventario lcOpcionGalenhos
            Case "Ventas"
                 EdicionVentas lcOpcionGalenhos
            Case "NS", "NSF"
                EdicionNS lcOpcionGalenhos, IIf(ms_ModuloSeleccionado = "NS", False, True)
            Case "NI", "NIF", "FARMADOP"
                EdicionNI lcOpcionGalenhos, IIf(ms_ModuloSeleccionado = "NI", False, True)
            Case "IntervencionS"
                EdicionIntervencionS lcOpcionGalenhos
            Case "DependenciaExt"
                EdicionDependenciaExt lcOpcionGalenhos
            Case "DespachoDonaciones"
                EdicionDespachoDonaciones lcOpcionGalenhos
            Case "FarmPrecios"                                     'debb2014b
                EdicionMantenedorHistoricoPrecios lcOpcionGalenhos          'debb2014b
                 
                 
                 
            Case "Roles"
                 EdicionRoles lcOpcionGalenhos
            Case "Empleado"
                 EdicionEmpleado lcOpcionGalenhos
            
            
            Case "Constancias"
              EdicionConstancias lcOpcionGalenhos
              
            Case "Fua"
              EdicionFua lcOpcionGalenhos
            
            Case "HcElectronica"
                 
            
            End Select
         
         
         
         End If
    Case "SALIR"
         Me.Visible = False
    End Select
    Set rsPermisos = Nothing
End Sub
'
Sub EdicionPaciente(sToolId As String, lTipoServicio As sghTipoServicio, lnIdTablaLISTBARITEMS As Long)
Dim mo_PacienteDetalle As New PacienteDetalle


        mo_PacienteDetalle.Opcion = SeleccionarOpcion(sToolId)
        mo_PacienteDetalle.idUsuario = sighentidades.USUARIO
        mo_PacienteDetalle.TipoServicio = lTipoServicio
        mo_PacienteDetalle.lcNombrePc = lc_NombrePc
        mo_PacienteDetalle.lnIdTablaLISTBARITEMS = lnIdTablaLISTBARITEMS

        Select Case mo_PacienteDetalle.Opcion
        Case sghAgregar
        Case sghModificar, sghConsultar, sghEliminar
            mo_PacienteDetalle.idPaciente = Me.ucPacientesLista1.idRegistroSeleccionado
            If mo_PacienteDetalle.idPaciente = -1 Or mo_PacienteDetalle.idPaciente = 0 Then
                MsgBox "Seleccione un registro", vbInformation, Me.Caption
                Exit Sub
            End If
        End Select

        mo_PacienteDetalle.Icon = Me.Icon
        mo_PacienteDetalle.Show 1
        Unload mo_PacienteDetalle

        Select Case sToolId
        Case "ID_Agregar":
        Case "ID_Modificar":
            Dim doPaciente As New doPaciente
        Case "ID_Consultar":
        Case "ID_Eliminar":
        End Select

End Sub
'
'

Sub EdicionSiCitas(sToolId As String)
        ucSIlistasCitas1.lnIdTablaLISTBARITEMS = IIf(ucSIlistasCitas1.Area = sghImageneología, 604, 603)
        ucSIlistasCitas1.lcNombrePc = lc_NombrePc
        Select Case sToolId
        Case "ID_Agregar":
            ucSIlistasCitas1.mnuDiarioAgregarProgramacion_Click
        Case "ID_Modificar":
            ucSIlistasCitas1.mnuDiarioModificarProgramacion_Click
        Case "ID_Consultar":
            ucSIlistasCitas1.mnuDiarioConsultarProgramacion_Click
        Case "ID_Eliminar":
            ucSIlistasCitas1.mnuDiarioEliminarProgramacion_Click
        End Select
End Sub
Sub EdicionSiProgramacion(sToolId As String)
        ucSIlistasCitas1.lnIdTablaLISTBARITEMS = IIf(ucSIlistasCitas1.Area = sghImageneología, 604, 603)
        ucSIlistasCitas1.lcNombrePc = lc_NombrePc
        Select Case sToolId
        Case "ID_Agregar":
            ucHISListaLotes.mnuDiarioAgregarProgramacion_Click
        Case "ID_Modificar":
            ucHISListaLotes.mnuDiarioModificarProgramacion_Click
        Case "ID_Consultar":
            ucHISListaLotes.mnuDiarioConsultarProgramacion_Click
        Case "ID_Eliminar":
            ucHISListaLotes.mnuDiarioEliminarProgramacion_Click
        End Select
End Sub



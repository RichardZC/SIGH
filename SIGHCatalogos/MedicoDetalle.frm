VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form MedicoDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MedicoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
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
      Height          =   6930
      Left            =   5535
      TabIndex        =   31
      Top             =   -30
      Width           =   6315
      Begin VB.ComboBox cmbIdEspecialidad 
         Height          =   330
         Left            =   1590
         TabIndex        =   15
         Top             =   630
         Width           =   4620
      End
      Begin VB.ComboBox cmbIdDepartamento 
         Height          =   330
         Left            =   1590
         TabIndex        =   14
         Top             =   270
         Width           =   4620
      End
      Begin VB.CommandButton btnQuitar 
         DisabledPicture =   "MedicoDetalle.frx":0CCA
         DownPicture     =   "MedicoDetalle.frx":1055
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
         Left            =   2670
         Picture         =   "MedicoDetalle.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1005
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregar 
         DisabledPicture =   "MedicoDetalle.frx":1779
         DownPicture     =   "MedicoDetalle.frx":1B62
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
         Left            =   1605
         Picture         =   "MedicoDetalle.frx":1F6E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1005
         Width           =   1005
      End
      Begin UltraGrid.SSUltraGrid grdEspecialidades 
         Height          =   5400
         Left            =   105
         TabIndex        =   18
         Top             =   1395
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   9525
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Especialidades"
      End
      Begin VB.Label Label8 
         Caption         =   "Especialidad"
         Height          =   255
         Left            =   195
         TabIndex        =   33
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   180
         TabIndex        =   32
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   45
      TabIndex        =   28
      Top             =   6900
      Width           =   11805
      Begin VB.CommandButton Command1 
         Caption         =   "BUSCAR_EESSXCODIGO"
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   180
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BUSCAR_MEDICAMENTOSXCODIGO"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   540
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BUSCAR_INSUMOSXCODIGO"
         Height          =   375
         Left            =   8760
         TabIndex        =   46
         Top             =   180
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "BUSCAR_FFAMACEUTICAXCODIGO"
         Height          =   375
         Left            =   8160
         TabIndex        =   45
         Top             =   540
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MedicoDetalle.frx":237A
         DownPicture     =   "MedicoDetalle.frx":283E
         Height          =   700
         Left            =   6045
         Picture         =   "MedicoDetalle.frx":2D2A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MedicoDetalle.frx":3216
         DownPicture     =   "MedicoDetalle.frx":3676
         Height          =   700
         Left            =   4485
         Picture         =   "MedicoDetalle.frx":3AEB
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   216
         Width           =   1365
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
      Height          =   6930
      Left            =   45
      TabIndex        =   21
      Top             =   -30
      Width           =   5385
      Begin VB.ComboBox cmbIdPais 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   2010
         Width           =   3165
      End
      Begin VB.ComboBox cmbIdTipoSexo 
         Height          =   330
         Left            =   4050
         TabIndex        =   5
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CheckBox chkMedicoEgresado 
         Alignment       =   1  'Right Justify
         Caption         =   "Egresado (FUA)"
         Height          =   435
         Left            =   90
         TabIndex        =   51
         Top             =   5730
         Width           =   2175
      End
      Begin VB.TextBox txtRNE 
         Height          =   345
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   13
         Top             =   5445
         Width           =   3180
      End
      Begin VB.CheckBox chkEsActivoMedico 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   435
         Left            =   90
         TabIndex        =   49
         Top             =   6150
         Width           =   2190
      End
      Begin VB.TextBox txtSupervisor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   43
         Top             =   5040
         Width           =   2250
      End
      Begin VB.CommandButton cmdSupervisorDel 
         DisabledPicture =   "MedicoDetalle.frx":3F60
         DownPicture     =   "MedicoDetalle.frx":42EB
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
         Left            =   4830
         Picture         =   "MedicoDetalle.frx":467E
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5040
         Width           =   435
      End
      Begin VB.CommandButton cmdSupervisorAdd 
         DisabledPicture =   "MedicoDetalle.frx":4A0F
         DownPicture     =   "MedicoDetalle.frx":4DF8
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
         Left            =   4380
         Picture         =   "MedicoDetalle.frx":5204
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5040
         Width           =   435
      End
      Begin VB.CommandButton cmdBuscaEnTablasSIS 
         Caption         =   "..."
         Height          =   315
         Left            =   4965
         TabIndex        =   40
         ToolTipText     =   "Busca en Tablas del SIS: Apellidos y nombres, Colegiatura"
         Top             =   165
         Width           =   315
      End
      Begin VB.ComboBox cmbIdDocIdentidad 
         Height          =   330
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   180
         Width           =   1725
      End
      Begin VB.ComboBox cmbColegioHIS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   12
         Top             =   4665
         Width           =   3165
      End
      Begin VB.ComboBox cmbTipoDestacado 
         Height          =   330
         Left            =   2100
         TabIndex        =   10
         Top             =   3900
         Width           =   3165
      End
      Begin VB.TextBox txtLoteHis 
         Height          =   345
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   11
         Top             =   4275
         Width           =   1095
      End
      Begin VB.ComboBox cmbIdCondicionTrabajo 
         Height          =   330
         Left            =   2100
         TabIndex        =   9
         Top             =   3510
         Width           =   3165
      End
      Begin VB.ComboBox cmbIdTipoEmpleado 
         Height          =   330
         Left            =   2100
         TabIndex        =   8
         Top             =   3120
         Width           =   3165
      End
      Begin VB.TextBox txtCodigoPlanilla 
         Height          =   315
         Left            =   2100
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2385
         Width           =   1425
      End
      Begin VB.TextBox txtDNI 
         Height          =   345
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   0
         Top             =   165
         Width           =   1125
      End
      Begin VB.TextBox txtApellidoPaterno 
         Height          =   315
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   1
         Top             =   555
         Width           =   3165
      End
      Begin VB.TextBox txtApellidoMaterno 
         Height          =   315
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   2
         Top             =   915
         Width           =   3165
      End
      Begin VB.TextBox txtNombres 
         Height          =   315
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1290
         Width           =   3165
      End
      Begin VB.TextBox txtColegiatura 
         Height          =   345
         Left            =   2100
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2730
         Width           =   1425
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   1650
         Width           =   1440
         _ExtentX        =   2540
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "País"
         Height          =   210
         Left            =   120
         TabIndex        =   53
         Top             =   2025
         Width           =   300
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3645
         TabIndex        =   52
         Top             =   1665
         Width           =   405
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   2
         X1              =   15
         X2              =   5325
         Y1              =   5415
         Y2              =   5415
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   30
         X2              =   5340
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   5370
         Y1              =   2340
         Y2              =   2355
      End
      Begin VB.Label Label10 
         Caption         =   "RNE (FUA)"
         Height          =   285
         Left            =   105
         TabIndex        =   50
         Top             =   5475
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor (HisGalenhos)"
         Height          =   210
         Left            =   120
         TabIndex        =   44
         Top             =   5070
         Width           =   2010
      End
      Begin VB.Label lblHalladosEnSis 
         Alignment       =   2  'Center
         Caption         =   "Apellidos, Nombres, Colegiatura hallados en tablas SIS"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   39
         Top             =   6555
         Visible         =   0   'False
         Width           =   4965
      End
      Begin VB.Label Label7 
         Caption         =   "Colegio Profes.(HIS)"
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   4725
         Width           =   1725
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Destacado"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Nacimiento"
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Lote (HIS)"
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   4365
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Código planilla"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   2445
         Width           =   1500
      End
      Begin VB.Label DNI 
         Caption         =   "Documento"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Colegiatura"
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   2820
         Width           =   1425
      End
      Begin VB.Label lblIdTipoEmpleado 
         AutoSize        =   -1  'True
         Caption         =   "Tipo empleado"
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label lblIdCondicionTrabajo 
         Caption         =   "Condición trabajo"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   3570
         Width           =   1500
      End
      Begin VB.Label lblNombres 
         Caption         =   "Nombres"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1350
         Width           =   1500
      End
      Begin VB.Label lblApellidoMaterno 
         Caption         =   "Apellido materno"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label lblApellidoPaterno 
         Caption         =   "Apellido paterno"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1500
      End
   End
End
Attribute VB_Name = "MedicoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Médicos
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Dim ml_IdMedico As Long
Dim ml_IdEmpleado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_Empleado As New dOEmpleado
Dim mo_Medico As New DOMedico
Dim mo_CollMedicoEspecialidad As New Collection
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_SIGHSisConsumoWeb As New SIGHNegocios.SisConsumoWeb
'SCCQ 25/03/2020 Cambio2 Inicio
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
'SCCQ 25/03/2020 Cambio2 Fin
Dim mrs_Especialidades As New Recordset
Dim vcolegio As Integer
Dim mo_CmbIdTipoSexo As New sighentidades.ListaDespleglable
Dim mo_cmbIdDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoEmpleado As New sighentidades.ListaDespleglable
Dim mo_cmbIdCondicionTrabajo As New sighentidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New sighentidades.ListaDespleglable
Dim mo_cmbTipoDestacado As New sighentidades.ListaDespleglable
Dim mo_cmbColegioHIS As New sighentidades.ListaDespleglable
Dim mo_cmbIdDocIdentidad As New sighentidades.ListaDespleglable
'SCCQ 25/03/2020 Cambio2 Inicio
Dim mo_CmbIdPais As New sighentidades.ListaDespleglable
'SCCQ 25/03/2020 Cambio2 Fin
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_loginEstado As Long
Dim lnIdEstablecimientoExterno As Long
Dim lbReniecAutorizado As Boolean

'<(Inicio) Añadido Por: WABG el: 26/01/2021-12:05:07 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Dim mo_Reniec As New ReniecGalenhosNegocios
Dim lbBuscaDNIenReniec As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
'</(Fin) Añadido Por: WABG el: 26/01/2021-12:05:07 p.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Property Let idMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_IdMedico
End Property


Sub CargarComboBoxes()
Dim sSQL As String
       
       mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
       mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
       Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
       
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdTipoEmpleado.BoundColumn = "IdTipoEmpleado"
       mo_cmbIdTipoEmpleado.ListField = "DescripcionLarga"
       'SCCQ 16/04/2020 Cambio2 Inicio -->Se agregó and tipoempleadoHIS IS NOT NULL and tipoempleadoHIS<>''
       Set mo_cmbIdTipoEmpleado.RowSource = mo_AdminServiciosComunes.TiposEmpleadosSeleccionarSegunFiltro("where esProgramado=1 and tipoempleadoHIS IS NOT NULL and tipoempleadoHIS<>''")
       'SCCQ 16/04/2020 Cambio2 Fin
       mo_cmbIdCondicionTrabajo.BoundColumn = "IdCondicionTrabajo"
       mo_cmbIdCondicionTrabajo.ListField = "DescripcionLarga"
       Set mo_cmbIdCondicionTrabajo.RowSource = mo_AdminServiciosComunes.TiposCondicionTrabajoSeleccionarTodos
       
       mo_cmbTipoDestacado.BoundColumn = "idDestacado"
       mo_cmbTipoDestacado.ListField = "Destacado"
       Set mo_cmbTipoDestacado.RowSource = mo_AdminServiciosComunes.TiposDestacadosSeleccionarTodos
       mo_cmbTipoDestacado.BoundText = "3"
       '
       mo_cmbColegioHIS.BoundColumn = "cod_col"
       mo_cmbColegioHIS.ListField = "des_col"
       Set mo_cmbColegioHIS.RowSource = mo_AdminServiciosComunes.ColegiosHISseleccionarTodos
       
       mo_cmbIdDocIdentidad.BoundColumn = "IdDocIdentidad"
       mo_cmbIdDocIdentidad.ListField = "DescripcionLarga"
       Set mo_cmbIdDocIdentidad.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodos()
       mo_cmbIdDocIdentidad.BoundText = "1"
       
       'SCCQ 25/03/2020 Cambio2 Inicio
       mo_CmbIdPais.BoundColumn = "IdPais"
       mo_CmbIdPais.ListField = "Nombre"
       Set mo_CmbIdPais.RowSource = mo_AdminServiciosGeograficos.PaisesSeleccionarTodos()
       mo_CmbIdPais.BoundText = "166"
       '       cmbIdPais.Enabled = False
       'GLCC Combo debe estar activo - 20-07-2020
       cmbIdPais.Enabled = True
       'SCCQ 25/03/2020 Cambio2 Fin
End Sub

Private Sub btnEliminar_Click()

End Sub


Private Sub chkMedicoEgresado_Click()
    If chkMedicoEgresado.Value = 1 Then
        txtRNE.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtRNE, False
    Else
        mo_Formulario.HabilitarDeshabilitar txtRNE, True
    End If
End Sub

Private Sub chkMedicoEgresado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, chkMedicoEgresado
   AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbColegioHIS_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbColegioHIS
   AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDepartamento_Click()
Dim sMensaje As String
       mo_cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       mo_cmbIdEspecialidad.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
        MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdDepartamento_LostFocus()
   If cmbIdDepartamento.Text <> "" Then
       mo_cmbIdDepartamento.BoundText = Val(Split(cmbIdDepartamento.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdDepartamento
End Sub



Private Sub cmbIdDocIdentidad_LostFocus()
    Select Case mo_cmbIdDocIdentidad.BoundText
    Case 1    'dni
         txtDNI.MaxLength = 8
         'SCCQ 25/03/2020 Cambio2 Inicio
         mo_CmbIdPais.BoundText = "166"
         cmbIdPais.Enabled = False
         'SCCQ 25/03/2020 Cambio2 Inicio
    Case Else
         txtDNI.MaxLength = 20
         'SCCQ 25/03/2020 Cambio2 Inicio
         cmbIdPais.Enabled = True
         'SCCQ 25/03/2020 Cambio2 Inicio
    End Select

End Sub
Private Sub cmbIdEspecialidad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdEspecialidad
    AdministrarKeyPreview KeyCode

End Sub
Private Sub cmbIdEspecialidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub cmbIdEspecialidad_LostFocus()
   If cmbIdEspecialidad.Text <> "" Then
       mo_cmbIdEspecialidad.BoundText = Val(Split(cmbIdEspecialidad.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdEspecialidad

End Sub

Private Sub cmbIdTipoEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoEmpleado
AdministrarKeyPreview KeyCode
End Sub



Private Sub cmbIdTipoEmpleado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdCondicionTrabajo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCondicionTrabajo
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdCondicionTrabajo_LostFocus()
   If cmbIdCondicionTrabajo.Text <> "" Then
       mo_cmbIdCondicionTrabajo.BoundText = Val(Split(cmbIdCondicionTrabajo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdCondicionTrabajo
End Sub

Private Sub cmbIdCondicionTrabajo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub cmbTipoDestacado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbTipoDestacado
   AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscaEnTablasSIS_Click()
    ActualizaDatosDesdeSIS
End Sub

Private Sub cmdSupervisorDel_Click()
    txtSupervisor.Tag = 0
    txtSupervisor.Text = ""

End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdTipoEmpleado.MiComboBox = cmbIdTipoEmpleado
    Set mo_cmbIdCondicionTrabajo.MiComboBox = cmbIdCondicionTrabajo
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Set mo_cmbTipoDestacado.MiComboBox = cmbTipoDestacado
    Set mo_cmbColegioHIS.MiComboBox = cmbColegioHIS
    Set mo_cmbIdDocIdentidad.MiComboBox = cmbIdDocIdentidad
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
     'SCCQ 25/03/2020 Cambio2 Inicio
    Set mo_CmbIdPais.MiComboBox = cmbIdPais
    'SCCQ 25/03/2020 Cambio2 Fin
End Sub


Private Sub grdEspecialidades_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
Dim Col As SSColumn
Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    grdEspecialidades.Bands(0).Columns("IdEspecialidad").Hidden = True
    grdEspecialidades.Bands(0).Columns("IdEspecialidad").Width = 1000
    grdEspecialidades.Bands(0).Columns("DescripcionLarga").Width = 4000
    
    Dim rsEspecialidades As New Recordset
    Set rsEspecialidades = mo_AdminServiciosHosp.EspecialidadesSeleccionarporMedico(Me.idMedico, oConexion)
    If rsEspecialidades.RecordCount > 0 Then
        Do While Not rsEspecialidades.EOF
             With mrs_Especialidades
                 .AddNew
                 .Fields!IdEspecialidad = rsEspecialidades!IdEspecialidad
                 .Fields!DescripcionLarga = rsEspecialidades!DescripcionLarga
             End With
             rsEspecialidades.MoveNext
        Loop
        mrs_Especialidades.MoveFirst
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub txtColegiatura_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtColegiatura
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtColegiatura_LostFocus()
  '************************* Inicio GalenHos V.3 *****************
  If mi_Opcion = sghAgregar Then
        If Trim(txtColegiatura.Text) <> "" Then 'Actualizado 07102014
            Dim rsTmp As ADODB.Recordset
            Set rsTmp = MedicosSeleccionarPorColegiatura(txtColegiatura.Text)
            If Not (rsTmp.EOF = True And rsTmp.BOF = True) Then
              MsgBox "Ya existe un médico con el número de colegiatura ingresada", vbInformation, "SIGH "
              txtColegiatura.Text = ""
              txtColegiatura.SetFocus
            End If
            Set rsTmp = Nothing
        End If
  End If
  '************************* Fin GalenHos V.3 *****************
  mo_Formulario.MarcarComoVacio txtColegiatura
End Sub

Private Sub txtColegiatura_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtCodigoPlanilla_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoPlanilla
  AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigoPlanilla_LostFocus()
  '************************* Inicio GalenHos V.3 *****************
  If mi_Opcion = sghAgregar Then
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = EmpleadosSeleccionarPorCodigoPlanilla(txtCodigoPlanilla.Text)
    If Not (rsTmp.EOF = True And rsTmp.BOF = True) Then
      MsgBox "Ya existe un empleado con el código de planilla ingresado", vbInformation, "SIGH "
      txtCodigoPlanilla.Text = ""
      txtCodigoPlanilla.SetFocus
    End If
    Set rsTmp = Nothing
  End If
  '************************* Fin GalenHos V.3 *****************
  mo_Formulario.MarcarComoVacio txtCodigoPlanilla
End Sub

Private Sub txtCodigoPlanilla_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDNI
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDni_LostFocus()
  '************************* Inicio GalenHos V.3 *****************
  If Len(Trim(txtDNI.Text)) > 0 Then
   If mo_cmbIdDocIdentidad.BoundText = "1" And Len(Trim(txtDNI.Text)) <> 8 Then
      MsgBox "Si el Documento es DNI debe tener 8 dígitos", vbInformation, "Mensaje"
      On Error Resume Next
      txtDNI.SetFocus
      Exit Sub
   End If
  End If
  If mi_Opcion = sghAgregar And Len(txtDNI.Text) = 8 And mo_cmbIdDocIdentidad.BoundText = "1" Then
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = EmpleadosSeleccionarPorDNI(txtDNI.Text)
    If Not (rsTmp.EOF = True And rsTmp.BOF = True) Then
        MsgBox "Ya existe un empleado con el número de DNI ingresado", vbInformation, "SIGH "
        txtDNI.Text = ""
        txtDNI.SetFocus
    Else
        ActualizaDatosDesdeSIS
        
'<(Inicio) Añadido Por: WABG el: 26/01/2021-12:06:04 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        lbBuscaDNIenReniec = IIf(lcBuscaParametro.SeleccionaFilaParametro(296) = "S", True, False)
        If mo_cmbIdDocIdentidad.BoundText = "1" And lbBuscaDNIenReniec = True And Len(Trim(txtApellidoPaterno.Text)) = 0 And Len(Trim$(txtApellidoMaterno.Text)) = 0 And mi_Opcion = sghAgregar Then
            
                mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
                mo_Reniec.Inicializar
            
               mo_Reniec.ConsultarDNIenReniec txtDNI.Text
            If mo_Reniec.ApellidoPaterno <> "" Then
                  
                  txtCodigoPlanilla.Text = txtDNI.Text
                  txtApellidoPaterno.Text = mo_Reniec.ApellidoPaterno
                  txtApellidoMaterno.Text = mo_Reniec.ApellidoMaterno
                  txtNombres.Text = Trim(mo_Reniec.PrimerNombre) + " " + Trim(mo_Reniec.SegundoNombre) + " " + Trim$(mo_Reniec.TercerNombre)
                  txtFechaNacimiento.Text = mo_Reniec.FechaNacimiento
                  
                  'CARGAR SEXO
                  If mo_Reniec.idTipoSexo = 1 Then
                    cmbIdTipoSexo.ListIndex = 0
                  Else
                    cmbIdTipoSexo.ListIndex = 1
                  End If
                  
            End If
        End If
'</(Fin) Añadido Por: WABG el: 26/01/2021-12:06:04 p.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>
    End If
    Set rsTmp = Nothing
  End If
  '************************* Fin GalenHos V.3 *****************
  mo_Formulario.MarcarComoVacio txtDNI
End Sub

Sub ActualizaDatosDesdeSIS()
        Dim rsTmp As New ADODB.Recordset
        lblHalladosEnSis.Visible = False
        Set rsTmp = mo_ReglasSISgalenhos.a_resatencionSeleccionarPorDNI(txtDNI.Text)
        If rsTmp.RecordCount > 0 Then
             ActualizaDatosUbicadoEnSis rsTmp
             lblHalladosEnSis.Visible = True
             lblHalladosEnSis.Caption = "Datos personales hallados en TABLAS SIS LOCAL"
        Else 'Modificado para busqueda web*************************************************
            If mo_SIGHSisConsumoWeb.ConsultarServicioPerSaludxNroDoc(Trim(txtDNI.Text), rsTmp) = True Then
                ActualizaDatosUbicadoEnSis rsTmp
                lblHalladosEnSis.Visible = True
                lblHalladosEnSis.Caption = "Datos personales hallados en WEB SIS"
            End If
        End If
        Set rsTmp = Nothing
End Sub

Sub ActualizaDatosUbicadoEnSis(rsTmp As Recordset)
             txtApellidoPaterno.Text = rsTmp!pers_apePaterno
             If Not IsNull(rsTmp!pers_apeMaterno) Then
                txtApellidoMaterno.Text = rsTmp!pers_apeMaterno
             End If
             If Not IsNull(rsTmp!pers_priNombre) Then
                txtNombres.Text = rsTmp!pers_priNombre & IIf(IsNull(rsTmp!pers_OtrNombre), "", " " & rsTmp!pers_OtrNombre)
             End If
             If Not IsNull(rsTmp!pers_Colegiatura) Then
                txtColegiatura.Text = rsTmp!pers_Colegiatura
             End If

End Sub


Private Sub txtDNI_KeyPress(KeyAscii As Integer)
       If Val(mo_cmbIdDocIdentidad.BoundText) <> 2 And Val(mo_cmbIdDocIdentidad.BoundText) <> 3 Then
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If
       Else
            If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
                KeyAscii = 0
            End If
       End If
End Sub




Private Sub txtFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaNacimiento
AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaNacimiento_LostFocus()
    If Not EsFecha(txtFechaNacimiento.Text, "DD/MM/AAAA") Then
        On Error Resume Next
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
        txtFechaNacimiento.SetFocus
        Exit Sub
    End If
    If Year(Date) - Val(Right(txtFechaNacimiento.Text, 4)) < 15 Then
        MsgBox "No debe existir empleados menores a 15 años", vbInformation, Me.Caption
        txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
        txtFechaNacimiento.SetFocus
        Exit Sub
    End If

End Sub

Private Sub txtLoteHis_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtLoteHis
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombres
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombres_LostFocus()
txtNombres.Text = UCase(txtNombres.Text)   'debb-02/05/2016
   mo_Formulario.MarcarComoVacio txtNombres
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaterno_LostFocus()
txtApellidoMaterno.Text = UCase(txtApellidoMaterno.Text)   'debb-02/05/2016
   mo_Formulario.MarcarComoVacio txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_LostFocus()
txtApellidoPaterno.Text = UCase(txtApellidoPaterno.Text)    'debb-02/05/2016
   mo_Formulario.MarcarComoVacio txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar txtSupervisor, False
 Select Case mi_Opcion
     Case sghAgregar
        chkEsActivoMedico = 1
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
 
 Select Case mi_Opcion
     Case sghAgregar
        chkEsActivoMedico = 1
     Case sghModificar
     Case sghConsultar
        Me.Frame1.Enabled = False
        Me.Frame3.Enabled = False
        Me.btnAceptar.Enabled = False
     Case sghEliminar
        Me.Frame1.Enabled = False
        Me.Frame3.Enabled = False
 End Select
 
 
 
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
        GenerarRecordsetTemporal
    
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Profesional de la Salud"
           chkEsActivoMedico = 1
       Case sghModificar
           Me.Caption = "Modificar Profesional de la Salud"
       Case sghConsultar
           Me.Caption = "Consultar Profesional de la Salud"
       Case sghEliminar
           Me.Caption = "Eliminar Profesional de la Salud"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                    Me.txtDNI.SetFocus
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            CargaDatosAlObjetosDeDatos
           'If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbExclamation, Me.Caption
               End If
           'End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String

   ValidarDatosObligatorios = False
   
   If Me.txtCodigoPlanilla.Text = "" Then sMensaje = sMensaje + "- Ingrese el código de planilla" + Chr(13)
   If Me.txtApellidoPaterno.Text = "" Then sMensaje = sMensaje + "- Ingrese el apellido paterno" + Chr(13)
   If Me.txtApellidoMaterno.Text = "" Then sMensaje = sMensaje + "- Ingrese el apellido materno" + Chr(13)
   If Me.txtNombres.Text = "" Then sMensaje = sMensaje + "- Ingrese el nombre" + Chr(13)
   If Me.cmbIdTipoSexo.Text = "" Then sMensaje = sMensaje + "- Elija el Sexo" + Chr(13)
   If mo_cmbIdTipoEmpleado.BoundText = "" Then sMensaje = sMensaje + "- Ingrese el tipo de empleado" + Chr(13)
   If mo_cmbIdCondicionTrabajo.BoundText = "" Then sMensaje = sMensaje + "- Ingrese la condición de trabajo" + Chr(13)
'   If Me.txtColegiatura.Text = "" Then sMensaje = sMensaje + "- Ingrese la colegiatura" + Chr(13)
   If Me.txtDNI.Text = "" Then sMensaje = sMensaje + "- Ingrese el número de DNI" + Chr(13)
   If cmbColegioHIS.Text = "" Then sMensaje = sMensaje + "- Elija Colegio Profesional" + Chr(13)
   'SCCQ 26/03/2020 Cambio2 Inicio
   If Me.cmbIdPais.Text = "" Then sMensaje = sMensaje + "- Elija el País" + Chr(13)
   'SCCQ 26/03/2020 Cambio2 Fin
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function

Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
AdministrarKeyPreview KeyCode
End Sub


Function ValidarReglas() As Boolean
Dim rsEmpleado As ADODB.Recordset

   ValidarReglas = False
    '
    If mo_AdminServiciosComunes.TiposEmpleadosSeleccionarSiSeProgramaPorId(Val(mo_cmbIdTipoEmpleado.BoundText)) = False Then
        MsgBox "El TIPO DE EMPLEADO elegido no se programa" & Chr(13) & "para registrarlo debe utilizar el módulo de EMPLEADOS", vbInformation, Me.Caption
        Exit Function
    End If
    '
    Set rsEmpleado = mo_AdminServiciosComunes.EmpleadosObtenerConElMismoCodigoPlanilla(mo_Empleado)
    If Not (rsEmpleado.EOF And rsEmpleado.BOF) Then
        MsgBox "Ya existe un empleado con el mismo CODIGO DE PLANILLA" + Chr(13) + rsEmpleado!ApellidoPaterno + " " + rsEmpleado!ApellidoMaterno + " " + rsEmpleado!Nombres, vbExclamation, Me.Caption
        rsEmpleado.Close
        Exit Function
    End If
    '
    If Trim(txtColegiatura.Text) <> "" Then
        If Val(mo_cmbColegioHIS.BoundText) = 0 Then     'debb-02/07/2015
            mo_Medico.Colegiatura = "000000"
        Else
            Set rsEmpleado = mo_AdminServiciosComunes.EmpleadosObtenerConLaMismaCOLEGIATURA(txtColegiatura.Text)
            Select Case mi_Opcion
            Case sghAgregar
                 If rsEmpleado.RecordCount > 0 Then
                     MsgBox "Ese N° COLEGIATURA ya esta Registrado para: " + Trim(rsEmpleado.Fields!ApellidoPaterno) & " " & Trim(rsEmpleado.Fields!ApellidoMaterno) & rsEmpleado.Fields!Nombres, vbInformation, Me.Caption
                     Exit Function
                 End If
            Case sghModificar
                 If rsEmpleado.RecordCount > 0 Then
                    rsEmpleado.MoveFirst
                    Do While Not rsEmpleado.EOF
                       If Trim(rsEmpleado.Fields!Colegiatura) = Trim(Me.txtColegiatura.Text) And rsEmpleado.Fields!idMedico <> ml_IdMedico Then
                          MsgBox "Ese N° COLEGIATURA ya esta Registrado para: " + Trim(rsEmpleado.Fields!ApellidoPaterno) & " " & Trim(rsEmpleado.Fields!ApellidoMaterno) & rsEmpleado.Fields!Nombres, vbInformation, Me.Caption
                          Exit Function
                       End If
                       rsEmpleado.MoveNext
                    Loop
                 End If
            End Select
        End If
    Else
        '******************************************************************************
        'A.Yañez Actualizado 16102014
        Select Case mo_AdminServiciosComunes.TiposEmpleadosSeleccionarId(IIf(mo_cmbIdTipoEmpleado.BoundText = "", 0, CLng(mo_cmbIdTipoEmpleado.BoundText)))
        Case 0
             Me.txtColegiatura.Text = ""
        Case 1
            MsgBox "Debe registrar el N° de COLEGIATURA (obligatorio)"
            Exit Function
        Case Else
            MsgBox "Debe registrar el N° de COLEGIATURA (obligatorio)"
            Exit Function
        End Select
        '******************************************************************************
    End If
    '
    If txtDNI.Text <> "" Then
        If mo_cmbIdDocIdentidad.BoundText = "1" And Len(Trim(txtDNI.Text)) <> 8 Then
            MsgBox "El DNI debe tener longitud 8", vbInformation, Me.Caption
            Exit Function
        End If
        Set rsEmpleado = mo_AdminServiciosComunes.EmpleadosObtenerConelMismoDNI(txtDNI.Text, Val(mo_cmbIdDocIdentidad.BoundText))
        Select Case mi_Opcion
        Case sghAgregar
             If rsEmpleado.RecordCount > 0 Then
                 MsgBox "Ese N° DOCUMENTO ya esta Registrado para: " + Trim(rsEmpleado.Fields!ApellidoPaterno) & " " & Trim(rsEmpleado.Fields!ApellidoMaterno) & rsEmpleado.Fields!Nombres, vbInformation, Me.Caption
                 Exit Function
             End If
        Case sghModificar
             If rsEmpleado.RecordCount > 0 Then
                rsEmpleado.MoveFirst
                Do While Not rsEmpleado.EOF
                   If Trim(rsEmpleado.Fields!DNI) = Trim(Me.txtDNI.Text) And rsEmpleado.Fields!IdEmpleado <> ml_IdEmpleado Then
                      MsgBox "Ese N° DOCUMENTO ya esta Registrado para: " + Trim(rsEmpleado.Fields!ApellidoPaterno) & " " & Trim(rsEmpleado.Fields!ApellidoMaterno) & rsEmpleado.Fields!Nombres, vbInformation, Me.Caption
                      Exit Function
                   End If
                   rsEmpleado.MoveNext
                Loop
             End If
        End Select
    End If
    If mi_Opcion = sghAgregar And lblHalladosEnSis.Visible = False Then
       MsgBox "No se encontró en las tablas SIS, tendrá problemas con el formato FUA", vbInformation, "Mensaje"
    End If
    '
    Set rsEmpleado = Nothing
    ValidarReglas = True
   
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Empleado
           .IdEmpleado = Me.IdEmpleado
           .IdTipoEmpleado = mo_cmbIdTipoEmpleado.BoundText
           .IdCondicionTrabajo = mo_cmbIdCondicionTrabajo.BoundText
           .Nombres = Me.txtNombres.Text
           .ApellidoMaterno = Me.txtApellidoMaterno.Text
           .ApellidoPaterno = Me.txtApellidoPaterno.Text
           .IdUsuarioAuditoria = Me.idUsuario
           .DNI = Me.txtDNI
           .CodigoPlanilla = Me.txtCodigoPlanilla
           .LoginPc = mo_lcNombrePc
           .LoginEstado = mo_loginEstado
           .FechaNacimiento = IIf(Me.txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY, 0, Me.txtFechaNacimiento.Text)
           .idTipoDestacado = Val(mo_cmbTipoDestacado.BoundText)
           .IdEstablecimientoExterno = lnIdEstablecimientoExterno
           .ReniecAutorizado = lbReniecAutorizado
           .idTipoDocumento = Val(mo_cmbIdDocIdentidad.BoundText)
           .IdSupervisor = Val(txtSupervisor.Tag)
           .EsActivo = IIf(Me.chkEsActivoMedico.Value = 1, True, False)
           If Me.chkEsActivoMedico.Value = 0 Then
              .Clave = ""
           End If
           If mi_Opcion = sghAgregar Then
              .fechaingreso = Date
           End If
           .idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
           'SCCQ 26/03/2020 Cambio2 Inicio
           .IdPais = Val(mo_CmbIdPais.BoundText)
           'SCCQ 26/03/2020 Cambio2 Fin
   End With
   
    With mo_Medico
        .Colegiatura = Me.txtColegiatura
        .IdEmpleado = Me.IdEmpleado
        .idMedico = Me.idMedico
        .IdUsuarioAuditoria = Me.idUsuario
        .LoteHis = Me.txtLoteHis.Text
        '***************************************************************
        'A.Yañez
        If Me.txtColegiatura <> "" Then
           .idColegioHis = Right("0" & mo_cmbColegioHIS.BoundText, 2)
        Else
           .idColegioHis = "00"
        End If
        '***************************************************************
'        .idColegioHis = Right("0" & mo_cmbColegioHIS.BoundText, 2)
        .Rne = Me.txtRNE.Text
        .Egresado = IIf(Me.chkMedicoEgresado.Value = 1, True, False)
    End With
   
    Dim oDoEspecialidadMedico  As DOMedicoEspecialidad
    Set mo_CollMedicoEspecialidad = New Collection
    Dim oRow As SSRow
    Set oRow = Me.grdEspecialidades.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        Set oDoEspecialidadMedico = New DOMedicoEspecialidad
        oDoEspecialidadMedico.IdMedicoEspecialidad = 0
        oDoEspecialidadMedico.idMedico = Me.idMedico
        oDoEspecialidadMedico.IdEspecialidad = Val(oRow.Cells("IdEspecialidad").Value)
        oDoEspecialidadMedico.IdUsuarioAuditoria = ml_idUsuario
        mo_CollMedicoEspecialidad.Add oDoEspecialidadMedico
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            Set oDoEspecialidadMedico = New DOMedicoEspecialidad
            oDoEspecialidadMedico.IdMedicoEspecialidad = 0
            oDoEspecialidadMedico.idMedico = Me.idMedico
            oDoEspecialidadMedico.IdEspecialidad = Val(oRow.Cells("IdEspecialidad").Value)
            oDoEspecialidadMedico.IdUsuarioAuditoria = ml_idUsuario
            mo_CollMedicoEspecialidad.Add oDoEspecialidadMedico
        Loop
    End If
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    Dim returnValue As Boolean
    returnValue = mo_AdminProgramacionMedica.MedicosAgregar(mo_Medico, mo_Empleado, mo_CollMedicoEspecialidad, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtNombres.Text))
    If returnValue = True Then
'        'mgaray201411f
'        Dim o_ReglasIntegracion As New ReglasIntegracion
'        Call o_ReglasIntegracion.EnviarDatosMedicoRisPacs(mo_Medico, mo_Empleado, mo_CollMedicoEspecialidad)
    End If
    AgregarDatos = returnValue
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    Dim returnValue As Boolean
    returnValue = mo_AdminProgramacionMedica.MedicosModificar(mo_Medico, mo_Empleado, mo_CollMedicoEspecialidad, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtNombres.Text))
    If returnValue = True Then
'        'mgaray201411f
'        Dim o_ReglasIntegracion As New ReglasIntegracion
'        Call o_ReglasIntegracion.EnviarDatosMedicoRisPacs(mo_Medico, mo_Empleado, mo_CollMedicoEspecialidad, False)
    End If
    ModificarDatos = returnValue
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    Dim lcMensaje As String
    lcMensaje = mo_AdminServiciosComunes.EmpleadosChequeaNoTengaMovimientos(mo_Empleado.IdEmpleado)
    If lcMensaje <> "" Then
       MsgBox "No se puede eliminar" & Chr(13) & lcMensaje, vbInformation, Me.Caption
    Else
    EliminarDatos = mo_AdminProgramacionMedica.MedicosEliminar(mo_Medico, mo_Empleado, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtNombres.Text))
    End If
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
       Dim oConexion As New Connection
       Dim oRsTmp1 As New Recordset
       oConexion.Open sighentidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       
       mb_ExistenDatos = mo_AdminProgramacionMedica.MedicosSeleccionarPorId(Me.idMedico, mo_Medico, mo_Empleado, mo_CollMedicoEspecialidad, oConexion)
       
       If mo_AdminProgramacionMedica.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbInformation, Me.Caption
            mb_ExistenDatos = False
            Exit Sub
       End If
       
       If mb_ExistenDatos Then
            With mo_Medico
                Me.idMedico = .idMedico
                Me.txtColegiatura.Text = Trim(.Colegiatura)
                Me.txtLoteHis.Text = .LoteHis
                mo_cmbColegioHIS.BoundText = Trim(Str(Val(.idColegioHis)))
                Me.txtRNE.Text = .Rne
                Me.chkMedicoEgresado.Value = IIf(.Egresado = True, 1, 0)
            End With
       
            With mo_Empleado
                Me.IdEmpleado = .IdEmpleado
                mo_cmbIdTipoEmpleado.BoundText = .IdTipoEmpleado
                mo_cmbIdCondicionTrabajo.BoundText = .IdCondicionTrabajo
                Me.txtNombres.Text = .Nombres
                Me.txtApellidoMaterno.Text = .ApellidoMaterno
                Me.txtApellidoPaterno.Text = .ApellidoPaterno
                Me.txtDNI = .DNI
                Me.txtCodigoPlanilla = .CodigoPlanilla
                mo_lcNombrePc = .LoginPc
                mo_loginEstado = .LoginEstado
                Me.chkEsActivoMedico.Value = IIf(.EsActivo = True, 1, 0)
                If .FechaNacimiento <> 0 Then
                   Me.txtFechaNacimiento.Text = .FechaNacimiento
                End If
                mo_cmbTipoDestacado.BoundText = .idTipoDestacado
                lnIdEstablecimientoExterno = .IdEstablecimientoExterno
                lbReniecAutorizado = .ReniecAutorizado
                mo_cmbIdDocIdentidad.BoundText = .idTipoDocumento
                'SCCQ 25/03/2020 Cambio2 Inicio
                 Select Case mo_cmbIdDocIdentidad.BoundText
                    Case 1    'dni
                    'Si es DNI el país debe ser siempre Perú
                        mo_CmbIdPais.BoundText = "166" '166 es para Perú
                         cmbIdPais.Enabled = True
                    Case Else
                         cmbIdPais.Enabled = True
                        mo_CmbIdPais.BoundText = .IdPais
                    End Select
                'SCCQ 25/03/2020 Cambio2 Fin
                BuscaEmpleadoYllenaDatosDelSupervisor .IdSupervisor
                mo_CmbIdTipoSexo.BoundText = .idTipoSexo
            End With
            
            'La carga de la grilla esta en el initialize
            mb_ExistenDatos = True
       
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       oConexion.Close
       Set oConexion = Nothing
End Sub


'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

            Me.IdEmpleado = 0
            Me.txtCodigoPlanilla = ""
            mo_cmbIdCondicionTrabajo.BoundText = ""
            Me.txtNombres.Text = ""
            Me.txtApellidoMaterno.Text = ""
            Me.txtApellidoPaterno.Text = ""
            Me.txtColegiatura = ""
            Me.txtDNI = ""
            Me.txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
            lnIdEstablecimientoExterno = 0
            lblHalladosEnSis.Visible = False

            mo_cmbIdDepartamento.BoundText = ""
            mo_cmbIdEspecialidad.BoundText = ""
            mo_cmbTipoDestacado.BoundText = "3"
            mo_cmbIdDocIdentidad.BoundText = "1"
            txtSupervisor.Tag = 0
            txtSupervisor.Text = ""
            Me.txtRNE.Text = ""
            mo_Formulario.HabilitarDeshabilitar txtRNE, True
            Me.chkMedicoEgresado.Value = 0
            
            Do While Not mrs_Especialidades.EOF
                mrs_Especialidades.Delete
                mrs_Especialidades.Update
                mrs_Especialidades.MoveNext
            Loop
            
            Set Me.grdEspecialidades.DataSource = mrs_Especialidades

End Sub

Sub GenerarRecordsetTemporal()
    
    With mrs_Especialidades
          .Fields.Append "IdEspecialidad", adInteger, 4
          .Fields.Append "DescripcionLarga", adVarChar, 150
          .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set Me.grdEspecialidades.DataSource = mrs_Especialidades
    mo_Apariencia.ConfigurarFilasBiColores Me.grdEspecialidades, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnAgregar_Click()
    
    On Error Resume Next
    
    If mo_cmbIdDepartamento.BoundText = "" Or mo_cmbIdEspecialidad.BoundText = "" Then
        MsgBox "Debe ingresar el departamento y la especialidad", vbInformation, Me.Caption
        Exit Sub
    End If
    
    mrs_Especialidades.MoveFirst
    Do While Not mrs_Especialidades.EOF
        If mo_cmbIdEspecialidad.BoundText = mrs_Especialidades!IdEspecialidad Then
            MsgBox "La especialidad seleccionada ya existe", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_Especialidades.MoveNext
    Loop
    
    With mrs_Especialidades
        .AddNew
        .Fields!IdEspecialidad = mo_cmbIdEspecialidad.BoundText
        .Fields!DescripcionLarga = Me.cmbIdEspecialidad.Text
    End With
    
    Set Me.grdEspecialidades.DataSource = mrs_Especialidades
    mrs_Especialidades.MoveFirst
End Sub

Private Sub btnQuitar_Click()
    
    On Error Resume Next
    With mrs_Especialidades
        If Not .EOF And Not .BOF Then
            'mgaray201411c
            If ValidarElinacionEspecialidadMedico(ml_IdMedico, .Fields!IdEspecialidad) = False Then
                Exit Sub
            End If
           .Delete
           .Update
        End If
    End With

    Set Me.grdEspecialidades.DataSource = mrs_Especialidades

End Sub

'*************************Inicio GalenHos V.3 *****************
Function EmpleadosSeleccionarPorCodigoPlanilla(CodigoPlanilla As String) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para Seleccionar/Averiguar si existe un empleado con el Codigo de Planilla
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "EmpleadosSeleccionarPorCodigoPlanilla"
    Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 8, CodigoPlanilla): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set EmpleadosSeleccionarPorCodigoPlanilla = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function EmpleadosSeleccionarPorDNI(DNI As String) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para Seleccionar/Averiguar si existe un empleado con el DNI
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "EmpleadosSeleccionarPorDNI"
    Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, DNI): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set EmpleadosSeleccionarPorDNI = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function MedicosSeleccionarPorColegiatura(Colegiatura As String) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para Seleccionar/Averiguar si existe un medico con la Colegiatura
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "MedicosSeleccionarPorColegiatura"
    Set oParameter = .CreateParameter("@Colegiatura", adVarChar, adParamInput, 6, Colegiatura): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set MedicosSeleccionarPorColegiatura = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'***************************Fin GalenHos v.3 ******************
Private Sub cmdSupervisorAdd_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        BuscaEmpleadoYllenaDatosDelSupervisor oBusqueda.idRegistroSeleccionado
    End If
    Set oBusqueda = Nothing

End Sub

Sub BuscaEmpleadoYllenaDatosDelSupervisor(lnIdEmpleado As Long)
    Dim oDOEmpleado As New dOEmpleado
    Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(lnIdEmpleado)
    txtSupervisor.Tag = 0
    txtSupervisor.Text = ""
    If Not oDOEmpleado Is Nothing Then
        txtSupervisor.Tag = oDOEmpleado.IdEmpleado
        txtSupervisor.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    Set oDOEmpleado = Nothing
End Sub



Private Sub Command1_Click()
    Dim a As New SIGHNegocios.SisConsumoWeb
    Dim cad As String
    Dim rsTmp As ADODB.Recordset
    
    cad = Trim(InputBox("ingresa codigo renaes de eess"))
    If a.ConsultarServicioBuscarEESSxCodigo(cad, rsTmp) Then
        MsgBox "ok ubicado en web e insertado en tablas sis"
    Else
        MsgBox "no esta en la web"
    End If
    Set rsTmp = Nothing
End Sub

Private Sub Command2_Click()
    Dim a As New SIGHNegocios.SisConsumoWeb
    Dim cad As String
    Dim rsTmp As ADODB.Recordset
    cad = Trim(InputBox("ingresa codigo medicamento"))
    If a.ConsultarServicioMedicamentosxCodigo(cad, rsTmp) Then
        MsgBox "ok ubicado en web e insertado en tablas sis"
    Else
        MsgBox "no esta en la web"
    End If
    Set rsTmp = Nothing
End Sub

Private Sub Command3_Click()
   Dim a As New SIGHNegocios.SisConsumoWeb
    Dim cad As String
Dim rsTmp As ADODB.Recordset
    cad = Trim(InputBox("ingresa codigo insumo"))
    If a.ConsultarServicioInsumosxCodigo(cad, rsTmp) Then
        MsgBox "ok ubicado en web e insertado en tablas sis"
    Else
        MsgBox "no esta en la web"
    End If
    Set rsTmp = Nothing
End Sub

Private Sub Command4_Click()
    Dim a As New SIGHNegocios.SisConsumoWeb
    Dim cad As String
    Dim rsTmp As ADODB.Recordset
    cad = Trim(InputBox("ingresa abreviatura de presentacion"))
    If a.ConsultarServicioFFamaceuticaxCodigo(cad, rsTmp) Then
        MsgBox "ok ubicado en web e insertado en tablas sis"
    Else
        MsgBox "no esta en la web"
    End If
    Set rsTmp = Nothing
End Sub

'mgaray201411c
Private Function ValidarElinacionEspecialidadMedico(lIdMedico As Long, lIdEspecialidad As Long) As Boolean
    Dim oDoMedicoEspecialidad As New DOMedicoEspecialidad
    oDoMedicoEspecialidad.idMedico = lIdMedico
    oDoMedicoEspecialidad.IdEspecialidad = lIdEspecialidad
    If lIdMedico > 0 Then
        ValidarElinacionEspecialidadMedico = mo_AdminProgramacionMedica.EspecialidadMedicoValidaEliminar(oDoMedicoEspecialidad)
    Else
        ValidarElinacionEspecialidadMedico = True
    End If
End Function

Private Sub txtRNE_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtRNE
   AdministrarKeyPreview KeyCode
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportarHIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar HIS"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   Icon            =   "frmExportarHIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1005
      Left            =   30
      TabIndex        =   1
      Top             =   2280
      Width           =   5535
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmExportarHIS.frx":000C
         DownPicture     =   "frmExportarHIS.frx":046C
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
         Left            =   1320
         Picture         =   "frmExportarHIS.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         DisabledPicture =   "frmExportarHIS.frx":0D56
         DownPicture     =   "frmExportarHIS.frx":121A
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
         Left            =   2880
         Picture         =   "frmExportarHIS.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame 
      Height          =   2295
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox cmbTurno 
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
         ItemData        =   "frmExportarHIS.frx":1BF2
         Left            =   1200
         List            =   "frmExportarHIS.frx":1BF4
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   2025
      End
      Begin MSComctlLib.ProgressBar pgbProcesoExportacion 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ComboBox cmbAnio 
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
         ItemData        =   "frmExportarHIS.frx":1BF6
         Left            =   1200
         List            =   "frmExportarHIS.frx":1BF8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2025
      End
      Begin VB.ComboBox cmbMes 
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
         ItemData        =   "frmExportarHIS.frx":1BFA
         Left            =   1200
         List            =   "frmExportarHIS.frx":1BFC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label4 
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   8
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Departamento 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmExportarHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica en donde se hara la programacion del HIS para los responsables de MR.
'        Programado por: Palomino Y
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim mo_Formulario As New SIGHEntidades.Formulario

'VARIABLES
Dim IdUsuario As Long
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAnio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTurno As New SIGHEntidades.ListaDespleglable

Private Sub Form_Load()
Set mo_cmbMes.MiComboBox = Me.cmbMes
Set mo_cmbAnio.MiComboBox = Me.cmbAnio
Set mo_cmbTurno.MiComboBox = Me.cmbTurno

'CARGA LOS LISTADOS DEL FORMULARIO
CargarComboBoxes
End Sub

Private Sub btnAceptar_Click()
'PROCESO DE EXPORTACION DE DATOS DEL HIS CON SUS PARAMETROS
If ProcesarHIS Then
    Call MsgBox("Los registros se exportaron exitosamente.", vbInformation Or vbSystemModal, App.Title)
    btnCancelar_Click
Else
    Call MsgBox("Los Registro no se Pudieron Exportar.", vbCritical Or vbSystemModal, App.Title)
End If
End Sub

Private Sub btnCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyEscape
        btnCancelar_Click
    End Select
End Sub

Sub CargarComboBoxes()
mo_cmbMes.BoundColumn = "IdMes"
mo_cmbMes.ListField = "NombreMes"
Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses

mo_Formulario.LlenaComboConAnios Me.cmbAnio

'mo_cmbAnio.BoundColumn = ""
'mo_cmbAnio.ListField = ""
'Set mo_cmbAnio.RowSource = mo_ReglasHIS.ObtenerListaEstablecimientosMR

mo_cmbTurno.BoundColumn = "IdHisTurno"
mo_cmbTurno.ListField = "Descripcion"
Set mo_cmbTurno.RowSource = mo_ReglasHIS.ListaTurnos

End Sub

'Proceso en Cliente - Debido a que el HIS MINSA esta solo en una Maquina y no en un Servidor
Private Function ProcesarHIS() As Boolean
'DECLARACION DE VARIABLES DE PROCESO
Dim oConexionSQL As New Connection
Dim oConexionFOX As New Connection

'CONTENEDORES DE DATOS MIGRADOS DEL HIS
Dim oRcs_HISA As New Recordset
Dim oRcs_HIS1 As New Recordset
Dim oRcs_HIS_LOTE As New Recordset

'CONTENEDORES DE DATOS DE HIS GALENHOS
Dim oRcs_Lote As New Recordset
Dim oRcs_Atenciones As New Recordset
Dim oRcs_Diagnosticos As New Recordset

Dim ms_sSQL As String

'DATOS DE DIAGNOSTICOS
Dim CodDxCIE As String
Dim DescTipoDx As String
Dim ClaseDx As String
Dim CodLAB As String

'INICIALIZACION DE VALORES DEL PROCESO
oConexionSQL.CommandTimeout = 300
oConexionSQL.Open SIGHEntidades.CadenaConexion

oConexionFOX.CommandTimeout = 300
oConexionFOX.Open "DSN=HIS"

'LIMPIANDO TABLAS Y OBTENIENDO ESTRUCTURAS
'ms_sSQL = "DELETE FROM HISA"
'oRcs_HISA.Open ms_sSQL, oConexionFOX, adOpenKeyset, adLockOptimistic
'oRcs_HISA.Open "SELECT * FROM HIS_HISA", oConexionFOX, adOpenKeyset, adLockOptimistic
'
'ms_sSQL = "DELETE FROM HIS1"
'oRcs_HIS1.Open ms_sSQL, oConexionFOX, adOpenKeyset, adLockOptimistic
'oRcs_HIS1.Open "SELECT * FROM HIS_HIS1", oConexionFOX, adOpenKeyset, adLockOptimistic
'
'ms_sSQL = "DELETE FROM HIS_LOTE"
'oRcs_HIS_LOTE.Open ms_sSQL, oConexionFOX, adOpenKeyset, adLockOptimistic
'oRcs_HIS_LOTE.Open "SELECT * FROM HIS_LOTE", oConexionFOX, adOpenKeyset, adLockOptimistic

'CONSULTAS PARA CADA TABLA DEPENDIENDO DEL FILTRO QUE SE PIDE
'Set oRcs_Lote = mo_ReglasHIS.ExportacionHIS_Lote(IdUsuario, Val(Me.cmbMes.ItemData(Me.cmbMes.ListIndex)), Val(Me.cmbAnio.ItemData(Me.cmbAnio.ListIndex)))
Set oRcs_Atenciones = mo_ReglasHIS.ExportacionHIS_Atenciones(IdUsuario, Val(Me.cmbMes.ItemData(Me.cmbMes.ListIndex)), Val(Me.cmbAnio.ItemData(Me.cmbAnio.ListIndex)))
Set oRcs_Diagnosticos = mo_ReglasHIS.ExportacionHIS_Diagnosticos(IdUsuario, Val(Me.cmbMes.ItemData(Me.cmbMes.ListIndex)), Val(Me.cmbAnio.ItemData(Me.cmbAnio.ListIndex)))

'INGRESO DE VALORES A LAS TABLAS

'TABLA HISA - EQUIVALENTE A LAS ATENCIONES + DIAGNOSTICOS

'Do While Not oRcs_Atenciones.EOF
'    oRcs_HISA.AddNew
'    oRcs_HISA.Fields!Cod_2000 = lcCod_2000
'    oRcs_HISA.Fields!Ano = Val(lcAnio)
'    oRcs_HISA.Fields!Mes = Val(lcMes)
'    oRcs_HISA.Fields!Nom_lote = lcNomLote
'    oRcs_HISA.Fields!Num_pag = lnNumPag
'    oRcs_HISA.Fields!num_reg = lnReg
'    oRcs_HISA.Fields!dia = Day(rsReporte.Fields!FechaIngreso)
'    oRcs_HISA.Fields!fichafam = lcFichaOhistoria
'    oRcs_HISA.Fields!Cod_Dpto = Left(lcDistritoDomicilio, 2)
'    oRcs_HISA.Fields!Cod_Prov = Mid(lcDistritoDomicilio, 3, 2)
'    oRcs_HISA.Fields!Cod_Dist = Right(lcDistritoDomicilio, 2)
'    oRcs_HISA.Fields!Edad = rsReporte.Fields!Edad
'    oRcs_HISA.Fields!tip_edad = rsReporte.Fields!EdadCodigo
'    oRcs_HISA.Fields!Sexo = IIf(rsReporte.Fields!IdTipoSexo = 1, "M", "F")
'    oRcs_HISA.Fields!establec = IIf(rsReporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
'    oRcs_HISA.Fields!Servicio = IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
'
'    oRcs_HISA.Fields!diagnost1 = lcTDx1
'    oRcs_HISA.Fields!labconf1 = lcLDx1
'    oRcs_HISA.Fields!codigo1 = lcDx1
'    oRcs_HISA.Fields!diagnost2 = lcTDx2
'    oRcs_HISA.Fields!labconf2 = lcLDx2
'    oRcs_HISA.Fields!codigo2 = lcDx2
'    oRcs_HISA.Fields!diagnost3 = lcTDx3
'    oRcs_HISA.Fields!labconf3 = lcLDx3
'    oRcs_HISA.Fields!codigo3 = lcDx3
'    oRcs_HISA.Fields!diagnost4 = lcTDx4
'    oRcs_HISA.Fields!labconf4 = lcLDx4
'    oRcs_HISA.Fields!codigo4 = lcDx4
'    oRcs_HISA.Fields!diagnost5 = lcTDx5
'    oRcs_HISA.Fields!labconf5 = lcLDx5
'    oRcs_HISA.Fields!codigo5 = lcDx5
'    oRcs_HISA.Fields!diagnost6 = lcTDx6
'    oRcs_HISA.Fields!labconf6 = lcLDx6
'    oRcs_HISA.Fields!codigo6 = lcDx6
'
'    oRcs_HISA.Fields!esta_reg = "2"
'    oRcs_HISA.Fields!mt = lcMt
'    oRcs_HISA.Fields!DNI = lcDNIpaciente
'    oRcs_HISA.Fields!FI = lcFuenteFinanciamientoPaciente
'    oRcs_HISA.Fields!et = lcEtniaPaciente
'    oRcs_HISA.Fields!st = "1"
'    oRcs_HISA.Update
'Loop

'TABLA HIS1 - CABECERA DE HOJA HIS

'    oRcs_HIS1.AddNew
'    oRcs_HIS1.Fields!Cod_2000 = lcCod_2000
'    oRcs_HIS1.Fields!Ano = Val(lcAnio)
'    oRcs_HIS1.Fields!Mes = Val(lcMes)
'    oRcs_HIS1.Fields!Nom_lote = lcNomLote
'    oRcs_HIS1.Fields!Num_pag = lnNumPag
'    oRcs_HIS1.Fields!Codif = lcCodif
'    oRcs_HIS1.Fields!cod_ServSa = lcCod_servsa
'    oRcs_HIS1.Fields!plaza = lcCodif
'    oRcs_HIS1.Fields!Esta_pag = "2"
'    oRcs_HIS1.Fields!Tot_reg = lnTot_reg
'    oRcs_HIS1.Fields!FlagEnvio = Space(1)
'    oRcs_HIS1.Fields!mt = lcMt
'    oRcs_HIS1.Fields!st = "1"
'    oRcs_HIS1.Update
    
'TABLA HISLOTE - DATOS DEL LOTE DE LA HOJA
'DO WHILE
'    oRcs_HIS_LOTE.AddNew
'    oRcs_HIS_LOTE.Fields!Cod_2000 = lcCod_2000
'    oRcs_HIS_LOTE.Fields!Ano = Val(lcAnio)
'    oRcs_HIS_LOTE.Fields!Mes = Val(lcMes)
'    oRcs_HIS_LOTE.Fields!Nom_lote = lcNomLote
'    oRcs_HIS_LOTE.Fields!cod_dig = lcCod_dig
'    oRcs_HIS_LOTE.Fields!esta_lote = "LC"
'    oRcs_HIS_LOTE.Fields!tot_pag = lnTTot_pag
'    oRcs_HIS_LOTE.Fields!Tot_reg = lnTTot_reg
'    oRcs_HIS_LOTE.Fields!Tot_rmalos = 0
'    oRcs_HIS_LOTE.Fields!tot_pgcarg = 0
'    oRcs_HIS_LOTE.Fields!esta_local = Space(1)
'    oRcs_HIS_LOTE.Fields!nomfile = Space(14)
'    oRcs_HIS_LOTE.Fields!flagReport = "3"
'    oRcs_HIS_LOTE.Fields!mt = lcMt
'    oRcs_HIS_LOTE.Update
'Loop

End Function

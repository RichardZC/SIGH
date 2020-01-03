VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEconRepConvenios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consumos Individuales"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmEconRepConvenios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSOption optTodos 
      Height          =   240
      Left            =   840
      TabIndex        =   2
      Top             =   390
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Todos los consumos"
      Value           =   -1
   End
   Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
      Caption         =   "..."
      Height          =   285
      Left            =   2550
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
      Top             =   60
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   4695
      Left            =   60
      TabIndex        =   13
      Top             =   3120
      Width           =   11535
      Begin TabDlg.SSTab sstReporte 
         Height          =   4215
         Left            =   60
         TabIndex        =   19
         Top             =   360
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   7435
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
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
         TabCaption(0)   =   "Medicamentos"
         TabPicture(0)   =   "frmEconRepConvenios.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ssMed"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Insumos"
         TabPicture(1)   =   "frmEconRepConvenios.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ssIns"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Laboratorio"
         TabPicture(2)   =   "frmEconRepConvenios.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ssLab"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Rayos X"
         TabPicture(3)   =   "frmEconRepConvenios.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "ssImag"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Procedimientos"
         TabPicture(4)   =   "frmEconRepConvenios.frx":0D3A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "ssProc"
         Tab(4).ControlCount=   1
         Begin UltraGrid.SSUltraGrid ssMed 
            Height          =   3405
            Left            =   120
            TabIndex        =   20
            Top             =   510
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6006
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BorderStyle     =   5
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Medicamentos que consumió"
         End
         Begin UltraGrid.SSUltraGrid ssIns 
            Height          =   3405
            Left            =   -74880
            TabIndex        =   21
            Top             =   480
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6006
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BorderStyle     =   5
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Insumos que consumió"
         End
         Begin UltraGrid.SSUltraGrid ssLab 
            Height          =   3405
            Left            =   -74880
            TabIndex        =   22
            Top             =   480
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6006
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BorderStyle     =   5
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Pruebas de Laboratorio que se realizó"
         End
         Begin UltraGrid.SSUltraGrid ssImag 
            Height          =   3405
            Left            =   -74880
            TabIndex        =   23
            Top             =   480
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6006
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BorderStyle     =   5
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Imágenes de Rayos X que se tomó"
         End
         Begin UltraGrid.SSUltraGrid ssProc 
            Height          =   3405
            Left            =   -74880
            TabIndex        =   24
            Top             =   480
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6006
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BorderStyle     =   5
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Procedimientos que consumió"
         End
      End
      Begin VB.ComboBox cboServicio 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtFA 
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
         Left            =   90
         TabIndex        =   7
         Top             =   3000
         Width           =   1740
      End
      Begin VB.TextBox txtD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1320
         Width           =   1740
      End
      Begin VB.TextBox txtNC 
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
         Left            =   90
         TabIndex        =   5
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
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
         Left            =   60
         TabIndex        =   17
         Top             =   1110
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cama"
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
         Left            =   60
         TabIndex        =   16
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Alta"
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
         Left            =   60
         TabIndex        =   15
         Top             =   2790
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Consumos"
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
         TabIndex        =   14
         Top             =   30
         Width           =   1710
      End
   End
   Begin VB.TextBox txtAN 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   5220
   End
   Begin VB.TextBox txtHC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   60
      Width           =   1260
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   60
      TabIndex        =   10
      Top             =   7920
      Width           =   11550
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Visualizar (F2)"
         DisabledPicture =   "frmEconRepConvenios.frx":0D56
         DownPicture     =   "frmEconRepConvenios.frx":11B6
         Enabled         =   0   'False
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
         Left            =   4365
         Picture         =   "frmEconRepConvenios.frx":162B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmEconRepConvenios.frx":1AA0
         DownPicture     =   "frmEconRepConvenios.frx":1F64
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
         Left            =   5895
         Picture         =   "frmEconRepConvenios.frx":2450
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdAtenciones 
      Height          =   2085
      Left            =   60
      TabIndex        =   18
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   3678
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BorderStyle     =   5
      ScrollBars      =   2
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Atenciones que tuvo el paciente"
   End
   Begin Threed.SSOption optFechas 
      Height          =   240
      Left            =   840
      TabIndex        =   3
      Top             =   630
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Consumos por Fechas"
   End
   Begin MSMask.MaskEdBox txtFechaInicio 
      Height          =   315
      Left            =   4560
      TabIndex        =   28
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFechaFin 
      Height          =   315
      Left            =   9000
      TabIndex        =   29
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
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
      Left            =   3240
      TabIndex        =   27
      Top             =   660
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
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
      Left            =   7440
      TabIndex        =   26
      Top             =   660
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Apellidos y Nombres"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmEconRepConvenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte para convenio
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim oPaciente As New Pacientes '
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_paciente As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idUsuario As Long
Dim ml_idPaciente As Long
Dim ml_idServicio As Long
Dim ml_idAtencion As Long
Dim ml_idCtaAtencion As Long
Dim ml_idCama As Long
Dim ml_FechaAlta As Date
Dim ml_Cama As String
Dim ml_Diagnostico As String
Dim ml_Historia As Long
Dim ml_Nombres As String
Dim ml_CAfiliacion As String
Dim ml_Servicio As String
Dim ml_Alta As String
Dim mo_cboServicio As New sighentidades.ListaDespleglable
Dim gridInfra As New GridInfragistic

Dim oDOPaciente As New DOPaciente
Dim oDOCama As New DOCama
Dim oServicios As ADODB.Recordset
Dim oAtenciones As ADODB.Recordset
Dim oDiagnosticos As ADODB.Recordset
Dim oLab As ADODB.Recordset
Dim oImag As ADODB.Recordset
Dim oMed As ADODB.Recordset
Dim oIns As ADODB.Recordset
Dim oProc As ADODB.Recordset
Dim oDevolucion As ADODB.Recordset

Dim rsTmp As New Recordset
Dim oConexion As New ADODB.Connection
  
Dim ml_idTipoServicio As Long
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idCuentaAtencion As Long
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Sub CargaDB_TextBox(Tabla As ADODB.Recordset, T As TextBox)
  Dim K As Integer
End Sub

Private Function BuscaPaciente(HCPaciente As Long)
  If HCPaciente = 0 Then Exit Function
  Set oDOPaciente = mo_paciente.PacientesSeleccionarPorHistoriaClinicaDefinitiva(Val(HCigualDNI_AgregaNUEVEaLaHistoria(Trim(Str(HCPaciente)))))
  ml_idPaciente = Val(oDOPaciente.IdPaciente)
  If ml_idPaciente <> 0 Then
    ml_Nombres = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno & ", " & oDOPaciente.PrimerNombre & " " & oDOPaciente.SegundoNombre
    Set oAtenciones = mo_ReglasLaboratorio.AtencionesQueTuvoElPacienteEnGeneral(ml_idPaciente)
    Set grdAtenciones.DataSource = oAtenciones
    grdAtenciones.Enabled = True
  Else
    ml_Nombres = ""
    ml_idPaciente = 0
    grdAtenciones.Enabled = False
  End If
  txtAN.Text = ml_Nombres
End Function

Private Sub btnAceptar_Click()
'  Dim oExcel As Excel.Application
'  Dim oWorkBookPlantilla As Workbook
'  Dim oWorkBook As Workbook
'  Dim oWorkSheet As Worksheet
'  Dim iFila As Long, iCol As Integer
'  Dim rsReporte As New Recordset
'  Dim II As Integer, Devueltos As Long
'  Dim TCant As Long, TPrec As Double
'  Dim TCant1 As Long, TotGen1 As Double, Cod As String, TPrec1 As Double
'
'  If optFechas.Value = True Then
'    If Trim(txtFechaInicio.Text) = "" Or Not IsDate(txtFechaInicio.Text) Then
'      MsgBox "Por favor ingrese la fecha inicial", vbInformation, Me.Caption
'      txtFechaInicio.SetFocus
'      Exit Sub
'    End If
'    If Trim(txtFechaFin.Text) = "" Or Not IsDate(txtFechaFin.Text) Then
'      MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
'      txtFechaFin.SetFocus
'      Exit Sub
'    End If
'  End If
'
'  MousePointer = 11
'  If ml_idTipoServicio = 1 Then
'    Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
'  Else
'    Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
'  End If
'  If rsReporte.RecordCount = 0 Then
'    MousePointer = 1
'    Exit Sub
'  End If
'
'  'Crea nueva hoja
'  Set oExcel = GalenhosExcelApplication()
'  Set oWorkBook = oExcel.Workbooks.Add
'  'Abre, copia y cierra la plantilla
'  Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\RepConvenios.xls")
'  oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
'  oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
'  oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
'  oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
'  oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
'  oWorkBookPlantilla.Close
'
'  '------- MEDICINAS
'  'Activa la primera hoja
'  Set oWorkSheet = oWorkBook.Sheets(1)
'  oWorkBook.Sheets(1).Name = "Medicinas"
'  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'  'Inicio de Impresion
'  oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE MEDICINAS DEL PACIENTE"
'  oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
'  oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
'  oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
'  oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'  oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
'  oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
'  oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
'  oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
'  iFila = 9
'  iCol = 2
'  If oMed.State = adStateOpen And Not (oMed.EOF = True And oMed.BOF = True) Then
'    oMed.MoveFirst
'    II = 0: TCant = 0: TPrec = 0
'    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
'    Do While Not oMed.EOF
'      TCant1 = 0
'      II = II + 1
'      oWorkSheet.Cells(iFila, iCol).Value = II
'      oWorkSheet.Cells(iFila, iCol + 1).Value = oMed!Codigo
'      Cod = oMed!Codigo
'      Devueltos = 0     'AveriguaDevueltos(Cod)  'debb-05/04/2011
'      TPrec1 = oMed!Precio
'      Do While Cod = oMed!Codigo And oMed.BOF = False And oMed.EOF = False
'        oWorkSheet.Cells(iFila, iCol + 2).Value = oMed!Nombre
'        TCant = TCant + oMed!Cantidad - Devueltos
'        TCant1 = TCant1 + oMed!Cantidad - Devueltos
'        oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
'        oWorkSheet.Cells(iFila, iCol + 4).Value = oMed!Precio
'        TotGen1 = TCant1 * oMed!Precio
'        oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
'        oMed.MoveNext
'        If oMed.EOF = True Then Exit Do
'      Loop
'      TPrec = TPrec + TotGen1
'      iFila = iFila + 1
'      If oMed.EOF = True Then Exit Do
'    Loop
'    oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
'    iFila = iFila + 1
'    oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
'    oWorkSheet.Cells(iFila, 5).Value = TCant
'    oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
'    oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
'    oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
'  End If
'  oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
'  If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
'
'  '-INSUMOS
'  'Activa la segunda hoja
'  Set oWorkSheet = oWorkBook.Sheets(2)
'  oWorkBook.Sheets(2).Name = "Insumos" 'oWorkBook.ActiveSheet.Name = "Insumos"
'  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'  'Inicio de Impresion
'  oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE INSUMOS DEL PACIENTE"
'  oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
'  oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
'  oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
'  oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'  oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
'  oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
'  oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
'  oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
'  iFila = 9
'  iCol = 2
'  If oIns.State = adStateOpen And Not (oIns.EOF = True And oIns.BOF = True) Then
'    oIns.MoveFirst
'    II = 0: TCant = 0: TPrec = 0
'    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
'    Do While Not oIns.EOF
'      TCant1 = 0
'      II = II + 1
'      oWorkSheet.Cells(iFila, iCol).Value = II
'      oWorkSheet.Cells(iFila, iCol + 1).Value = oIns!Codigo
'      Cod = oIns!Codigo
'      Devueltos = 0    'AveriguaDevueltos(Cod)   'debb-05/04/2011
'      TPrec1 = oIns!Precio
'      Do While Cod = oIns!Codigo And oIns.BOF = False And oIns.EOF = False
'        oWorkSheet.Cells(iFila, iCol + 2).Value = oIns!Nombre
'        TCant = TCant + oIns!Cantidad - Devueltos
'        TCant1 = TCant1 + oIns!Cantidad - Devueltos
'        oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
'        oWorkSheet.Cells(iFila, iCol + 4).Value = oIns!Precio
'        TotGen1 = TCant1 * oIns!Precio
'        oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
'        oIns.MoveNext
'        If oIns.EOF = True Then Exit Do
'      Loop
'      TPrec = TPrec + TotGen1
'      iFila = iFila + 1
'      If oIns.EOF = True Then Exit Do
'    Loop
'    oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
'    iFila = iFila + 1
'    oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
'    oWorkSheet.Cells(iFila, 5).Value = TCant
'    oWorkSheet.Cells(iFila, 7).Value = TPrec
'    oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
'    oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
'  End If
'  oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
'  If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
'
'  '-Laboratorio
'  'Activa la tercera hoja
'  Set oWorkSheet = oWorkBook.Sheets(3)
'  oWorkBook.Sheets(3).Name = "Laboratorio"
'  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'  'Inicio de Impresion
'  oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE LABORATORIO DEL PACIENTE"
'  oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
'  oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
'  oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
'  oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'  oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
'  oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
'  oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
'  oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
'  iFila = 9
'  iCol = 2
'  If oLab.State = adStateOpen And Not (oLab.EOF = True And oLab.BOF = True) Then
'    oLab.MoveFirst
'    II = 0: TCant = 0: TPrec = 0
'    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
'    Do While Not oLab.EOF
'      TCant1 = 0
'      II = II + 1
'      oWorkSheet.Cells(iFila, iCol).Value = II
'      oWorkSheet.Cells(iFila, iCol + 1).Value = oLab!Codigo
'      Cod = oLab!Codigo
'      TPrec1 = oLab!Precio
'      Do While Cod = oLab!Codigo And oLab.BOF = False And oLab.EOF = False
'        oWorkSheet.Cells(iFila, iCol + 2).Value = oLab!Nombre
'        TCant = TCant + oLab!Cantidad
'        TCant1 = TCant1 + oLab!Cantidad
'        oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
'        oWorkSheet.Cells(iFila, iCol + 4).Value = oLab!Precio
'        TotGen1 = TCant1 * oLab!Precio
'        oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
'        oLab.MoveNext
'        If oLab.EOF = True Then Exit Do
'      Loop
'      TPrec = TPrec + TotGen1
'      iFila = iFila + 1
'      If oLab.EOF = True Then Exit Do
'    Loop
'    oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
'    iFila = iFila + 1
'    oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
'    oWorkSheet.Cells(iFila, 5).Value = TCant
'    oWorkSheet.Cells(iFila, 7).Value = TPrec
'    oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
'    oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
'  End If
'  oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
'  If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
'
'  '-Rayos X
'  'Activa la cuarta hoja
'  Set oWorkSheet = oWorkBook.Sheets(4)
'  oWorkBook.Sheets(4).Name = "Imágenes"
'  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'  'Inicio de Impresion
'  oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE IMÁGENES DEL PACIENTE"
'  oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
'  oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
'  oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
'  oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'  oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
'  oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
'  oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
'  oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
'  iFila = 9
'  iCol = 2
'  If oImag.State = adStateOpen And Not (oImag.EOF = True And oImag.BOF = True) Then
'    oImag.MoveFirst
'    II = 0: TCant = 0: TPrec = 0
'    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
'    Do While Not oImag.EOF
'      TCant1 = 0
'      II = II + 1
'      oWorkSheet.Cells(iFila, iCol).Value = II
'      oWorkSheet.Cells(iFila, iCol + 1).Value = oImag!Codigo
'      Cod = oImag!Codigo
'      TPrec1 = oImag!Precio
'      Do While Cod = oImag!Codigo And oImag.BOF = False And oImag.EOF = False
'        oWorkSheet.Cells(iFila, iCol + 2).Value = oImag!Nombre
'        TCant = TCant + oImag!Cantidad
'        TCant1 = TCant1 + oImag!Cantidad
'        oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
'        oWorkSheet.Cells(iFila, iCol + 4).Value = oImag!Precio
'        TotGen1 = TCant1 * oImag!Precio
'        oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
'        oImag.MoveNext
'        If oImag.EOF = True Then Exit Do
'      Loop
'      TPrec = TPrec + TotGen1
'      iFila = iFila + 1
'      If oImag.EOF = True Then Exit Do
'    Loop
'    oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
'    iFila = iFila + 1
'    oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
'    oWorkSheet.Cells(iFila, 5).Value = TCant
'    oWorkSheet.Cells(iFila, 7).Value = TPrec
'    oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
'    oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
'  End If
'  oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
'  If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
'
'  '-Procedimientos
'  'Activa la quinta hoja
'  Set oWorkSheet = oWorkBook.Sheets(5)
'  oWorkBook.Sheets(5).Name = "Procedimientos"
'  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'  'Inicio de Impresion
'  oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE PROCEDIMIENTOS DEL PACIENTE"
'  oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
'  oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
'  oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
'  oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
'  oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
'  oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
'  oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!FechaEgreso), "", Format(rsReporte.Fields!FechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
'  oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
'  iFila = 9
'  iCol = 2
'  If oProc.State = adStateOpen And Not (oProc.EOF = True And oProc.BOF = True) Then
'    oProc.MoveFirst
'    II = 0: TCant = 0: TPrec = 0
'    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
'    Do While Not oProc.EOF
'      TCant1 = 0
'      II = II + 1
'      oWorkSheet.Cells(iFila, iCol).Value = II
'      oWorkSheet.Cells(iFila, iCol + 1).Value = oProc!Codigo
'      Cod = oProc!Codigo
'      TPrec1 = oProc!Precio
'      Do While Cod = oProc!Codigo And oProc.BOF = False And oProc.EOF = False
'        oWorkSheet.Cells(iFila, iCol + 2).Value = oProc!Nombre
'        TCant = TCant + oProc!Cantidad
'        TCant1 = TCant1 + oProc!Cantidad
'        oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
'        oWorkSheet.Cells(iFila, iCol + 4).Value = oProc!Precio
'        TotGen1 = TCant1 * oProc!Precio
'        oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
'        oProc.MoveNext
'        If oProc.EOF = True Then Exit Do
'      Loop
'      TPrec = TPrec + TotGen1
'      iFila = iFila + 1
'      If oProc.EOF = True Then Exit Do
'    Loop
'    oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
'    iFila = iFila + 1
'    oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
'    oWorkSheet.Cells(iFila, 5).Value = TCant
'    oWorkSheet.Cells(iFila, 7).Value = TPrec
'    oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
'    oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
'  End If
'  oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
'  If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
'  oExcel.Visible = True
'  oWorkSheet.PrintPreview
'    Set oWorkSheet = Nothing
'    Set oExcel = Nothing
'  MousePointer = 1
  Dim iFila As Long, iCol As Integer
  Dim rsReporte As New Recordset
  Dim II As Integer, Devueltos As Long
  Dim TCant As Long, TPrec As Double
  Dim TCant1 As Long, TotGen1 As Double, Cod As String, TPrec1 As Double
  Dim mo_ReporteUtil As New ReporteUtil
  Dim lcSql As String
  Dim lbEsOpenOffice As Boolean
  lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
  
  
    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
  If optFechas.Value = True Then
    If Trim(txtFechaInicio.Text) = "" Or Not IsDate(txtFechaInicio.Text) Then
      MsgBox "Por favor ingrese la fecha inicial", vbInformation, Me.Caption
      txtFechaInicio.SetFocus
      Exit Sub
    End If
    If Trim(txtFechaFin.Text) = "" Or Not IsDate(txtFechaFin.Text) Then
      MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
      txtFechaFin.SetFocus
      Exit Sub
    End If
  End If
  
  MousePointer = 11
  If ml_idTipoServicio = 1 Then
    Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraIngresosPorIdAtencion(ml_idAtencion)
  Else
    Set rsReporte = mo_ReglasFacturacion.AtencionesFiltraEgresosPorIdAtencion(ml_idAtencion)
  End If
  If rsReporte.RecordCount = 0 Then
    MousePointer = 1
    Exit Sub
  End If

If lbEsOpenOffice = True Then
    'Abre el archivo ExcelOpenOffice
    lcArchivoExcel = App.Path + "\Plantillas\RepConvenios.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
    Fichier = Format(Time, "hhmmss") & ".ods"
    FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
    lcArchivoExcel = Fichier
    Chemin = "file:///" & App.Path & "\Plantillas\"
    Chemin = Replace(Chemin, "\", "/")
    Fichier = Chemin & "/" & lcArchivoExcel
    '
    Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
    Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
    Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
    Set Feuille = Document.getSheets().getByIndex(0)

    mo_CabeceraReportes.CabeceraReportes Document, True
    ret = SetForegroundWindow(lnHwnd)
Else
    'Crea nueva hoja
    Set oExcel = GalenhosExcelApplication()
    Set oWorkBook = oExcel.Workbooks.Add

    'Abre, copia y cierra la plantilla
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\RepConvenios.xls")
    oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Worksheets("Consumos").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    Set oWorkSheet = oWorkBook.Sheets(1)
    mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
End If
    '------- MEDICINAS
    'Activa la primera hoja
    If lbEsOpenOffice = True Then
        Set Feuille = Document.getSheets().getByIndex(0)
    Else
        Set oWorkSheet = oWorkBook.Sheets(1)
        oWorkBook.Sheets(1).Name = "Medicinas"
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
    End If


  'Inicio de Impresion
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 0).setFormula("CONSUMOS DE MEDICINAS DEL PACIENTE")
        Call Feuille.getcellbyposition(1, 2).setFormula("Nro.Cuenta:  " & ml_idCtaAtencion)
        Call Feuille.getcellbyposition(4, 2).setFormula("Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL())
        Call Feuille.getcellbyposition(1, 3).setFormula("Paciente: " & txtAN.Text)
        Call Feuille.getcellbyposition(4, 3).setFormula("Nº Historia Clínica: " & Trim(txtHC.Text))
        Call Feuille.getcellbyposition(1, 4).setFormula("F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso)
        Call Feuille.getcellbyposition(4, 4).setFormula("Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio))
        Call Feuille.getcellbyposition(1, 5).setFormula("F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm")))
        Call Feuille.getcellbyposition(4, 5).setFormula("Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
    Else
        oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE MEDICINAS DEL PACIENTE"
        oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
        oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
        oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
        oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
        oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
        oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
        oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
        oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
    End If
  iFila = 9
  iCol = 2
  If oMed.State = adStateOpen And Not (oMed.EOF = True And oMed.BOF = True) Then
    oMed.MoveFirst
    II = 0: TCant = 0: TPrec = 0
    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
    Do While Not oMed.EOF
      TCant1 = 0
      II = II + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oMed!Codigo)
        Else
            oWorkSheet.Cells(iFila, iCol).Value = II
            oWorkSheet.Cells(iFila, iCol + 1).Value = oMed!Codigo
        End If
      Cod = oMed!Codigo
      Devueltos = 0     'AveriguaDevueltos(Cod)  'debb-05/04/2011
      TPrec1 = oMed!precio
      Do While Cod = oMed!Codigo And oMed.BOF = False And oMed.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oMed!Nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oMed!Nombre
        End If
        TCant = TCant + oMed!Cantidad - Devueltos
        TCant1 = TCant1 + oMed!Cantidad - Devueltos
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oMed!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oMed!precio
        End If
        TotGen1 = TCant1 * oMed!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
        End If
        oMed.MoveNext
        If oMed.EOF = True Then Exit Do
      Loop
      TPrec = TPrec + TotGen1
      iFila = iFila + 1
      If oMed.EOF = True Then Exit Do
    Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
    End If
    iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(Format(TPrec, "0.00"))
    Else
        oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 7).Value = Format(TPrec, "0.00")
        oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
    End If
  End If
  If lbEsOpenOffice = True Then
    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
    PrintArea(0).Sheet = 0
    PrintArea(0).startcolumn = 0
    PrintArea(0).StartRow = 0
    PrintArea(0).EndColumn = 8
    PrintArea(0).EndRow = iFila
    Call Feuille.SetPrintAreas(PrintArea())
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
  Else
    oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
    If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
  End If
  '-INSUMOS
  'Activa la segunda hoja
  If lbEsOpenOffice = True Then
    Set Feuille = Document.getSheets().getByIndex(1)
  Else
    Set oWorkSheet = oWorkBook.Sheets(2)
    oWorkBook.Sheets(2).Name = "Insumos" 'oWorkBook.ActiveSheet.Name = "Insumos"
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
  End If
  
  'Inicio de Impresion
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 0).setFormula("CONSUMOS DE INSUMOS DEL PACIENTE")
        Call Feuille.getcellbyposition(1, 2).setFormula("Nro.Cuenta:  " & ml_idCtaAtencion)
        Call Feuille.getcellbyposition(4, 2).setFormula("Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL())
        Call Feuille.getcellbyposition(1, 3).setFormula("Paciente: " & txtAN.Text)
        Call Feuille.getcellbyposition(4, 3).setFormula("Nº Historia Clínica: " & Trim(txtHC.Text))
        Call Feuille.getcellbyposition(1, 4).setFormula("F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso)
        Call Feuille.getcellbyposition(4, 4).setFormula("Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio))
        Call Feuille.getcellbyposition(1, 5).setFormula("F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm")))
        Call Feuille.getcellbyposition(4, 5).setFormula("Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
    Else
        oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE INSUMOS DEL PACIENTE"
        oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
        oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
        oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
        oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
        oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
        oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
        oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
        oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
    End If
  iFila = 9
  iCol = 2
  If oIns.State = adStateOpen And Not (oIns.EOF = True And oIns.BOF = True) Then
    oIns.MoveFirst
    II = 0: TCant = 0: TPrec = 0
    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
    Do While Not oIns.EOF
      TCant1 = 0
      II = II + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oIns!Codigo)
        Else
            oWorkSheet.Cells(iFila, iCol).Value = II
            oWorkSheet.Cells(iFila, iCol + 1).Value = oIns!Codigo
        End If
      Cod = oIns!Codigo
      Devueltos = 0    'AveriguaDevueltos(Cod)   'debb-05/04/2011
      TPrec1 = oIns!precio
      Do While Cod = oIns!Codigo And oIns.BOF = False And oIns.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oIns!Nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oIns!Nombre
        End If
        TCant = TCant + oIns!Cantidad - Devueltos
        TCant1 = TCant1 + oIns!Cantidad - Devueltos
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oIns!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oIns!precio
        End If
        TotGen1 = TCant1 * oIns!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
        End If
        oIns.MoveNext
        If oIns.EOF = True Then Exit Do
      Loop
      TPrec = TPrec + TotGen1
      iFila = iFila + 1
      If oIns.EOF = True Then Exit Do
    Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
    End If
    iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(TPrec)
    Else
        oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 7).Value = TPrec
        oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
    End If
  End If
  If lbEsOpenOffice = True Then
    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
    PrintArea(0).Sheet = 0
    PrintArea(0).startcolumn = 0
    PrintArea(0).StartRow = 0
    PrintArea(0).EndColumn = 8
    PrintArea(0).EndRow = iFila
    Call Feuille.SetPrintAreas(PrintArea())
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
  Else
    oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
    If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
  End If
  '-Laboratorio
  'Activa la tercera hoja
    If lbEsOpenOffice = True Then
        Set Feuille = Document.getSheets().getByIndex(2)
    Else
        Set oWorkSheet = oWorkBook.Sheets(3)
        oWorkBook.Sheets(3).Name = "Laboratorio"
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
    End If
  'Inicio de Impresion+
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 0).setFormula("CONSUMOS DE LABORATORIO DEL PACIENTE")
        Call Feuille.getcellbyposition(1, 2).setFormula("Nro.Cuenta:  " & ml_idCtaAtencion)
        Call Feuille.getcellbyposition(4, 2).setFormula("Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL())
        Call Feuille.getcellbyposition(1, 3).setFormula("Paciente: " & txtAN.Text)
        Call Feuille.getcellbyposition(4, 3).setFormula("Nº Historia Clínica: " & Trim(txtHC.Text))
        Call Feuille.getcellbyposition(1, 4).setFormula("F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso)
        Call Feuille.getcellbyposition(4, 4).setFormula("Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio))
        Call Feuille.getcellbyposition(1, 5).setFormula("F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm")))
        Call Feuille.getcellbyposition(4, 5).setFormula("Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
    Else
        oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE LABORATORIO DEL PACIENTE"
        oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
        oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
        oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
        oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
        oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
        oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
        oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
        oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
    End If
  iFila = 9
  iCol = 2
  If oLab.State = adStateOpen And Not (oLab.EOF = True And oLab.BOF = True) Then
    oLab.MoveFirst
    II = 0: TCant = 0: TPrec = 0
    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
    Do While Not oLab.EOF
      TCant1 = 0
      II = II + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oLab!Codigo)
        Else
            oWorkSheet.Cells(iFila, iCol).Value = II
            oWorkSheet.Cells(iFila, iCol + 1).Value = oLab!Codigo
        End If
      Cod = oLab!Codigo
          TPrec1 = oLab!precio
      Do While Cod = oLab!Codigo And oLab.BOF = False And oLab.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oLab!Nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oLab!Nombre
        End If
        TCant = TCant + oLab!Cantidad
        TCant1 = TCant1 + oLab!Cantidad
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oLab!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oLab!precio
        End If
        TotGen1 = TCant1 * oLab!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
        End If
        oLab.MoveNext
        If oLab.EOF = True Then Exit Do
      Loop
      TPrec = TPrec + TotGen1
      iFila = iFila + 1
      If oLab.EOF = True Then Exit Do
    Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
    End If
    iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(TPrec)
    Else
        oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 7).Value = TPrec
        oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
    End If
  End If
  If lbEsOpenOffice = True Then
    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
    PrintArea(0).Sheet = 0
    PrintArea(0).startcolumn = 0
    PrintArea(0).StartRow = 0
    PrintArea(0).EndColumn = 8
    PrintArea(0).EndRow = iFila
    Call Feuille.SetPrintAreas(PrintArea())
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
  Else
    oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
    If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
  End If
  
  '-Rayos X
  'Activa la cuarta hoja
    If lbEsOpenOffice = True Then
        Set Feuille = Document.getSheets().getByIndex(3)
    Else
        Set oWorkSheet = oWorkBook.Sheets(4)
        oWorkBook.Sheets(4).Name = "Imágenes"
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
    End If
  'Inicio de Impresion
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 0).setFormula("CONSUMOS DE IMÁGENES DEL PACIENTE")
        Call Feuille.getcellbyposition(1, 2).setFormula("Nro.Cuenta:  " & ml_idCtaAtencion)
        Call Feuille.getcellbyposition(4, 2).setFormula("Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL())
        Call Feuille.getcellbyposition(1, 3).setFormula("Paciente: " & txtAN.Text)
        Call Feuille.getcellbyposition(4, 3).setFormula("Nº Historia Clínica: " & Trim(txtHC.Text))
        Call Feuille.getcellbyposition(1, 4).setFormula("F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso)
        Call Feuille.getcellbyposition(4, 4).setFormula("Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio))
        Call Feuille.getcellbyposition(1, 5).setFormula("F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm")))
        Call Feuille.getcellbyposition(4, 5).setFormula("Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
    Else
        oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE IMÁGENES DEL PACIENTE"
        oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
        oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
        oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
        oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
        oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
        oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
        oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
        oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
    End If
  iFila = 9
  iCol = 2
  If oImag.State = adStateOpen And Not (oImag.EOF = True And oImag.BOF = True) Then
    oImag.MoveFirst
    II = 0: TCant = 0: TPrec = 0
    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
    Do While Not oImag.EOF
      TCant1 = 0
      II = II + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oImag!Codigo)
        Else
            oWorkSheet.Cells(iFila, iCol).Value = II
            oWorkSheet.Cells(iFila, iCol + 1).Value = oImag!Codigo
        End If
      Cod = oImag!Codigo
      TPrec1 = oImag!precio
      Do While Cod = oImag!Codigo And oImag.BOF = False And oImag.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oImag!Nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oImag!Nombre
        End If
        TCant = TCant + oImag!Cantidad
        TCant1 = TCant1 + oImag!Cantidad
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oImag!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oImag!precio
        End If
        TotGen1 = TCant1 * oImag!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TotGen1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
        End If
        oImag.MoveNext
        If oImag.EOF = True Then Exit Do
      Loop
      TPrec = TPrec + TotGen1
      iFila = iFila + 1
      If oImag.EOF = True Then Exit Do
    Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
    End If
    iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(TPrec)
    Else
        oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 7).Value = TPrec
        oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
    End If
  End If
  If lbEsOpenOffice = True Then
    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
    PrintArea(0).Sheet = 0
    PrintArea(0).startcolumn = 0
    PrintArea(0).StartRow = 0
    PrintArea(0).EndColumn = 8
    PrintArea(0).EndRow = iFila
    Call Feuille.SetPrintAreas(PrintArea())
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
  Else
    oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
    If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
  End If
  
  '-Procedimientos
  'Activa la quinta hoja
    If lbEsOpenOffice = True Then
        Set Feuille = Document.getSheets().getByIndex(4)
    Else
        Set oWorkSheet = oWorkBook.Sheets(5)
        oWorkBook.Sheets(5).Name = "Procedimientos"
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
    End If
  'Inicio de Impresion
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 0).setFormula("CONSUMOS DE PROCEDIMIENTOS DEL PACIENTE")
        Call Feuille.getcellbyposition(1, 2).setFormula("Nro.Cuenta:  " & ml_idCtaAtencion)
        Call Feuille.getcellbyposition(4, 2).setFormula("Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL())
        Call Feuille.getcellbyposition(1, 3).setFormula("Paciente: " & txtAN.Text)
        Call Feuille.getcellbyposition(4, 3).setFormula("Nº Historia Clínica: " & Trim(txtHC.Text))
        Call Feuille.getcellbyposition(1, 4).setFormula("F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso)
        Call Feuille.getcellbyposition(4, 4).setFormula("Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio))
        Call Feuille.getcellbyposition(1, 5).setFormula("F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm")))
        Call Feuille.getcellbyposition(4, 5).setFormula("Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama))
    Else
        oWorkSheet.Cells(1, 2).Value = "CONSUMOS DE PROCEDIMIENTOS DEL PACIENTE"
        oWorkSheet.Cells(3, 2).Value = "Nro.Cuenta:  " & ml_idCtaAtencion
        oWorkSheet.Cells(3, 5).Value = "Fecha de Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL()
        oWorkSheet.Cells(4, 2).Value = "Paciente: " & txtAN.Text
        oWorkSheet.Cells(4, 5).Value = "Nº Historia Clínica: " & Trim(txtHC.Text) '& "       Dom.Pac: " & Trim(txtDomicilioPacienteEnAtencion.Text)
        oWorkSheet.Cells(5, 2).Value = "F.Ingreso: " & rsReporte.Fields!FechaIngreso & " " & rsReporte.Fields!HoraIngreso
        oWorkSheet.Cells(5, 5).Value = "Servicio Egreso: " & IIf(IsNull(rsReporte.Fields!CodServicio), "", rsReporte.Fields!CodServicio & " - " & rsReporte.Fields!DServicio)
        oWorkSheet.Cells(6, 2).Value = "F.Alta Médica: " & IIf(IsNull(rsReporte.Fields!fechaEgreso), "", Format(rsReporte.Fields!fechaEgreso & " " & rsReporte.Fields!horaEgreso, "dd/mm/yyyy hh:mm"))
        oWorkSheet.Cells(6, 5).Value = "Cama: " & IIf(IsNull(rsReporte.Fields!codCama), "", rsReporte.Fields!codCama)
    End If
  iFila = 9
  iCol = 2
  If oProc.State = adStateOpen And Not (oProc.EOF = True And oProc.BOF = True) Then
    oProc.MoveFirst
    II = 0: TCant = 0: TPrec = 0
    TCant1 = 0: TotGen1 = 0: Cod = "": TPrec1 = 0
    Do While Not oProc.EOF
      TCant1 = 0
      II = II + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(oProc!Codigo)
        Else
            oWorkSheet.Cells(iFila, iCol).Value = II
            oWorkSheet.Cells(iFila, iCol + 1).Value = oProc!Codigo
        End If
      Cod = oProc!Codigo
      TPrec1 = oProc!precio
      Do While Cod = oProc!Codigo And oProc.BOF = False And oProc.EOF = False
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(oProc!Nombre)
        Else
            oWorkSheet.Cells(iFila, iCol + 2).Value = oProc!Nombre
        End If
        TCant = TCant + oProc!Cantidad
        TCant1 = TCant1 + oProc!Cantidad
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(oProc!precio)
        Else
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
            oWorkSheet.Cells(iFila, iCol + 4).Value = oProc!precio
        End If
        TotGen1 = TCant1 * oProc!precio
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TotGen1)
        Else
            oWorkSheet.Cells(iFila, iCol + 5).Value = TotGen1
        End If
        oProc.MoveNext
        If oProc.EOF = True Then Exit Do
      Loop
      TPrec = TPrec + TotGen1
      iFila = iFila + 1
      If oProc.EOF = True Then Exit Do
    Loop
    If lbEsOpenOffice = True Then
    Else
        oWorkSheet.range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 7)).borders.LineStyle = 1
    End If
    iFila = iFila + 1
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":G" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(TPrec)
    Else
        oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
        oWorkSheet.Cells(iFila, 5).Value = TCant
        oWorkSheet.Cells(iFila, 7).Value = TPrec
        oWorkSheet.range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).borders.LineStyle = 1
        oWorkSheet.Cells(iFila, 7).borders.LineStyle = 1
    End If
  End If
  If lbEsOpenOffice = True Then
    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
    PrintArea(0).Sheet = 0
    PrintArea(0).startcolumn = 0
    PrintArea(0).StartRow = 0
    PrintArea(0).EndColumn = 8
    PrintArea(0).EndRow = iFila
    Call Feuille.SetPrintAreas(PrintArea())
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
  Else
    oWorkSheet.PageSetup.PrintTitleRows = "$1:$8"
    If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$H$" & (iFila + 2)
    oExcel.Visible = True
    oWorkSheet.PrintPreview
  End If
  If lbEsOpenOffice = True Then
    'Liberar Memoria
    Set Plage = Nothing
    Set Feuille = Nothing
    Set Document = Nothing
    Set Desktop = Nothing
    Set ServiceManager = Nothing
    Set Style = Nothing
    Set Border = Nothing
    'encabezado de pagina
    Set PageStyles = Nothing
    Set Sheet = Nothing
    Set StyleFamilies = Nothing
    Set DefPage = Nothing
    Set Htext = Nothing
    Set Hcontent = Nothing
  Else
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
  End If
  MousePointer = 1

End Sub

Private Sub btnCancelar_Click()
  Unload Me
End Sub

Private Sub chkExcel_Click()

End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
  Dim oBusqueda As New SIGHNegocios.BuscaPacientes
  Dim oDOPaciente As New DOPaciente
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  oBusqueda.TipoFiltro = sghFiltrarTodos
  oBusqueda.MostrarFormulario
  If oBusqueda.BotonPresionado = sghAceptar Then
    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
    If Not oDOPaciente Is Nothing Then
      txtHC.Text = oDOPaciente.NroHistoriaClinica
      txtHC.SetFocus
      SendKeys "{TAB}"
    End If
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub Form_Initialize()
  Set mo_cboServicio.MiComboBox = cboServicio
End Sub

Private Sub Form_Load()
  ml_idServicio = 0
  ml_idPaciente = 0
  ml_idAtencion = 0
  mo_cboServicio.BoundColumn = "idServicio"
  mo_cboServicio.ListField = "NombreServicio"
End Sub

Private Sub grdAtenciones_Click()
  txtNC.Text = ""
  txtD.Text = ""
  txtFA.Text = ""
  Set ssLab.DataSource = Nothing
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
  'Busca servicios donde estuvo el paciente, por cada atención que tuvo
  If oAtenciones.EOF = True And oAtenciones.BOF = True Then Exit Sub
  ml_idAtencion = oAtenciones("idAtencion")
  ml_idCtaAtencion = oAtenciones("idCuentaAtencion")
  ml_Diagnostico = txtD.Text
  If Not (IsNull(oAtenciones("FechaEgreso"))) Then
    ml_Alta = oAtenciones("FechaEgreso")
  Else
    ml_Alta = ""
  End If
  txtFA.Text = ml_Alta
  If Not (IsDate(txtFechaInicio.Text)) Then txtFechaInicio.Text = IIf(IsNull(oAtenciones("FechaIngreso")), Now, oAtenciones("FechaIngreso")) & " " & IIf(IsNull(oAtenciones("HoraIngreso")), Now, Format(oAtenciones("HoraIngreso"), "hh:mm:ss"))
  If Not (IsDate(txtFechaFin.Text)) Then txtFechaFin.Text = IIf(IsNull(oAtenciones("FechaEgreso")), Format(Now, "dd/mm/yyyy"), oAtenciones("FechaEgreso")) & " " & IIf(IsNull(oAtenciones("HoraEgreso")), Format(Now, "hh:mm:ss"), Format(oAtenciones("HoraEgreso"), "hh:mm:ss"))
  Frame1.Enabled = True
  btnAceptar.Enabled = True
  If optTodos.Value = True Then
    Set oLab = mo_ReglasLaboratorio.SeleccionaLaboratorioPorCuenta(ml_idCtaAtencion)
    Set oImag = mo_ReglasLaboratorio.SeleccionaImagenologiaPorCuenta(ml_idCtaAtencion)
    Set oIns = mo_ReglasLaboratorio.SeleccionaInsumosPorCuenta(ml_idCtaAtencion)
    Set oMed = mo_ReglasLaboratorio.SeleccionaFarmaciaPorCuenta(ml_idCtaAtencion)
    Set oProc = mo_ReglasLaboratorio.SeleccionaProcedimientosPorCuenta(ml_idCtaAtencion)
    Set oDevolucion = mo_ReglasLaboratorio.SeleccionaDevolucionesPorCuenta(ml_idCtaAtencion)
  ElseIf optFechas.Value = True Then
    Set oLab = mo_ReglasLaboratorio.SeleccionaLaboratorioPorCuentaYFecha(ml_idCtaAtencion, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Set oImag = mo_ReglasLaboratorio.SeleccionaImagenologiaPorCuentaYFecha(ml_idCtaAtencion, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Set oIns = mo_ReglasLaboratorio.SeleccionaInsumosPorCuentaYFecha(ml_idCtaAtencion, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Set oMed = mo_ReglasLaboratorio.SeleccionaFarmaciaPorCuentaYFecha(ml_idCtaAtencion, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Set oProc = mo_ReglasLaboratorio.SeleccionaProcedimientosPorCuentaYFecha(ml_idCtaAtencion, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
    Set oDevolucion = mo_ReglasLaboratorio.SeleccionaDevolucionesPorCuentaYFecha(ml_idCtaAtencion, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text))
  Else
    Exit Sub
  End If
  Set ssLab.DataSource = oLab
  Set ssImag.DataSource = oImag
  Set ssIns.DataSource = oIns
  Set ssMed.DataSource = oMed
  Set ssProc.DataSource = oProc
End Sub

Private Sub grdAtenciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  grdAtenciones.Bands(0).Columns("idCuentaAtencion").Header.Caption = "Cuenta Atención"
  grdAtenciones.Bands(0).Columns("idCuentaAtencion").Width = 1300
  grdAtenciones.Bands(0).Columns("idAtencion").Hidden = True
  grdAtenciones.Bands(0).Columns("Edad").Hidden = True
  grdAtenciones.Bands(0).Columns("FechaIngreso").Header.Caption = "Fecha Ingreso"
  grdAtenciones.Bands(0).Columns("FechaIngreso").Width = 1300
  grdAtenciones.Bands(0).Columns("Descripcion").Header.Caption = "Plan"
  grdAtenciones.Bands(0).Columns("HoraIngreso").Header.Caption = "Hora Ingreso"
  grdAtenciones.Bands(0).Columns("HoraIngreso").Width = 1300
  grdAtenciones.Bands(0).Columns("FechaEgreso").Header.Caption = "Fecha Egreso"
  grdAtenciones.Bands(0).Columns("FechaEgreso").Width = 1300
  grdAtenciones.Bands(0).Columns("HoraEgreso").Header.Caption = "Hora Egreso"
  grdAtenciones.Bands(0).Columns("HoraEgreso").Width = 1300
  grdAtenciones.Bands(0).Columns("idFormaPago").Hidden = True
  grdAtenciones.Bands(0).Columns("idPaciente").Hidden = True
  grdAtenciones.Bands(0).Columns("idServicioIngreso").Hidden = True
  grdAtenciones.Bands(0).Columns("nombre").Header.Caption = "Servicio Ingreso"
  grdAtenciones.Bands(0).Columns("nombre").Width = 2500
  gridInfra.ConfigurarFilasBiColores grdAtenciones, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub grdServicios_Click()
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Label15.Caption = "Servicio ..."
  btnAceptar.Enabled = False
  If oServicios.EOF = True And oServicios.BOF = True Then Exit Sub
  Frame1.Enabled = True
  btnAceptar.Enabled = True
  ml_idServicio = oServicios!idServicio
  If Not (IsNull(oServicios!idCama)) Then
    ml_idCama = oServicios!idCama
  Else
    ml_idCama = 0
  End If
  If Val(ml_idCama) <> 0 Then
    Set oDOCama = mo_ReglasHoteleria.CamasSeleccionarPorId(ml_idCama, oConexion)
    ml_Cama = oDOCama.Codigo
  Else
    ml_Cama = ""
  End If
  txtNC.Text = ml_Cama
  ml_Servicio = oServicios!NombreServicio
  Label15.Caption = "Servicio: " & ml_Servicio
  If Val(ml_idServicio) <> 0 Then
    Set oLab = mo_ReglasLaboratorio.SeleccionaLaboratorioPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
    Set ssLab.DataSource = oLab
    Set oImag = mo_ReglasLaboratorio.SeleccionaImagenologiaPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
    Set ssImag.DataSource = oImag
    Set oIns = mo_ReglasLaboratorio.SeleccionaInsumosPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
    Set ssIns.DataSource = oIns
    Set oMed = mo_ReglasLaboratorio.SeleccionaFarmaciaPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
    Set ssMed.DataSource = oMed
    Set oProc = mo_ReglasLaboratorio.SeleccionaProcedimientosPorCuentaYServicio(ml_idCtaAtencion, ml_idServicio)
    Set ssProc.DataSource = oProc
  Else
    Set ssLab.DataSource = Nothing
    Set ssImag.DataSource = Nothing
    Set ssIns.DataSource = Nothing
    Set ssMed.DataSource = Nothing
    Set ssProc.DataSource = Nothing
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub Option1_Click()

End Sub

Private Sub optTodos_Click(Value As Integer)
  If optTodos.Value = True Then
    Label4.Visible = False
    Label5.Visible = False
    txtFechaInicio.Visible = False
    txtFechaFin.Visible = False
  End If
End Sub

Private Sub optFechas_Click(Value As Integer)
  If optFechas.Value = True Then
    Label4.Visible = True
    Label5.Visible = True
    txtFechaInicio.Visible = True
    txtFechaFin.Visible = True
  End If
End Sub

Private Sub ssImag_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssImag.Bands(0).Columns("tipo").Hidden = True
  ssImag.Bands(0).Columns("codigo").Width = 1000
  ssImag.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssImag.Bands(0).Columns("Nombre").Width = 6500
  ssImag.Bands(0).Columns("cantidad").Width = 800
  ssImag.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssImag.Bands(0).Columns("cantidad").Width = 800
  ssImag.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssImag.Bands(0).Columns("precio").Width = 800
  ssImag.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssImag.Bands(0).Columns("precio").Format = "0.00"
  ssImag.Bands(0).Columns("total").Width = 800
  ssImag.Bands(0).Columns("total").Header.Caption = "Total"
  ssImag.Bands(0).Columns("total").Format = "0.00"
  ssImag.Bands(0).Columns("idOrden").Hidden = True
  ssImag.Bands(0).Columns("idCuentaAtencion").Hidden = True
  ssImag.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssImag.Bands(0).Columns("FechaCreacion").Hidden = True
  ssImag.Bands(0).Columns("Descripcion").Hidden = True
  ssImag.Bands(0).Columns("Dfinanciamiento").Hidden = True
  ssImag.Bands(0).Columns("idEstadoFacturacion").Hidden = True
  ssImag.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssImag.Bands(0).Columns("idPuntoCarga").Hidden = True
  ssImag.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssImag.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssImag, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssIns_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssIns.Bands(0).Columns("tipo").Hidden = True
  ssIns.Bands(0).Columns("codigo").Width = 1000
  ssIns.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssIns.Bands(0).Columns("Nombre").Width = 6500
  ssIns.Bands(0).Columns("cantidad").Width = 800
  ssIns.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssIns.Bands(0).Columns("precio").Width = 800
  ssIns.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssIns.Bands(0).Columns("precio").Format = "0.00"
  ssIns.Bands(0).Columns("total").Width = 800
  ssIns.Bands(0).Columns("total").Header.Caption = "Total"
  ssIns.Bands(0).Columns("total").Format = "0.00"
  ssIns.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssIns.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssIns.Bands(0).Columns("idTipoFinanAtenciones").Hidden = True
  ssIns.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssIns.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssIns, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssLab_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssLab.Bands(0).Columns("tipo").Hidden = True
  ssLab.Bands(0).Columns("codigo").Width = 1000
  ssLab.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssLab.Bands(0).Columns("Nombre").Width = 6500
  ssLab.Bands(0).Columns("cantidad").Width = 800
  ssLab.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssLab.Bands(0).Columns("cantidad").Width = 800
  ssLab.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssLab.Bands(0).Columns("precio").Width = 800
  ssLab.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssLab.Bands(0).Columns("precio").Format = "0.00"
  ssLab.Bands(0).Columns("total").Width = 800
  ssLab.Bands(0).Columns("total").Header.Caption = "Total"
  ssLab.Bands(0).Columns("total").Format = "0.00"
  ssLab.Bands(0).Columns("idOrden").Hidden = True
  ssLab.Bands(0).Columns("idCuentaAtencion").Hidden = True
  ssLab.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssLab.Bands(0).Columns("FechaCreacion").Hidden = True
  ssLab.Bands(0).Columns("Descripcion").Hidden = True
  ssLab.Bands(0).Columns("Dfinanciamiento").Hidden = True
  ssLab.Bands(0).Columns("idEstadoFacturacion").Hidden = True
  ssLab.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssLab.Bands(0).Columns("idPuntoCarga").Hidden = True
  ssLab.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssLab.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssLab, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssMed_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssMed.Bands(0).Columns("tipo").Hidden = True
  ssMed.Bands(0).Columns("codigo").Width = 1000
  ssMed.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssMed.Bands(0).Columns("Nombre").Width = 6500
  ssMed.Bands(0).Columns("cantidad").Width = 800
  ssMed.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssMed.Bands(0).Columns("precio").Width = 800
  ssMed.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssMed.Bands(0).Columns("precio").Format = "0.00"
  ssMed.Bands(0).Columns("total").Width = 800
  ssMed.Bands(0).Columns("total").Header.Caption = "Total"
  ssMed.Bands(0).Columns("total").Format = "0.00"
  ssMed.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssMed.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssMed.Bands(0).Columns("idTipoFinanAtenciones").Hidden = True
  ssMed.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssMed.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssMed, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub ssProc_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  ssProc.Bands(0).Columns("tipo").Hidden = True
  ssProc.Bands(0).Columns("codigo").Width = 1000
  ssProc.Bands(0).Columns("codigo").Header.Caption = "Código"
  ssProc.Bands(0).Columns("Nombre").Width = 6500
  ssProc.Bands(0).Columns("cantidad").Width = 800
  ssProc.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssProc.Bands(0).Columns("cantidad").Width = 800
  ssProc.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  ssProc.Bands(0).Columns("precio").Width = 800
  ssProc.Bands(0).Columns("precio").Header.Caption = "Precio"
  ssProc.Bands(0).Columns("precio").Format = "0.00"
  ssProc.Bands(0).Columns("total").Width = 800
  ssProc.Bands(0).Columns("total").Header.Caption = "Total"
  ssProc.Bands(0).Columns("total").Format = "0.00"
  ssProc.Bands(0).Columns("idOrden").Hidden = True
  ssProc.Bands(0).Columns("idCuentaAtencion").Hidden = True
  ssProc.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
  ssProc.Bands(0).Columns("FechaCreacion").Hidden = True
  ssProc.Bands(0).Columns("Descripcion").Hidden = True
  ssProc.Bands(0).Columns("Dfinanciamiento").Hidden = True
  ssProc.Bands(0).Columns("idEstadoFacturacion").Hidden = True
  ssProc.Bands(0).Columns("idFuenteFinanciamiento").Hidden = True
  ssProc.Bands(0).Columns("idPuntoCarga").Hidden = True
  ssProc.Bands(0).Columns("idServicioPaciente").Hidden = True
  ssProc.Bands(0).Columns("idProducto").Hidden = True
  gridInfra.ConfigurarFilasBiColores ssProc, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub txtFechaFin_Change()
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssLab.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
End Sub

Private Sub txtFechaFin_GotFocus()
'  SeleccionaMask txtFechaFin
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY_HMS Then
        If Not IsDate(txtFechaFin) Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY_HMS
        End If
    End If
End Sub

Private Sub txtFechaInicio_Change()
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssLab.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
End Sub

Private Sub txtFechaInicio_GotFocus()
 ' SeleccionaMask txtFechaInicio
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY_HMS Then
        If Not IsDate(txtFechaInicio) Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY_HMS
        End If
    End If
End Sub

Private Sub txtHC_GotFocus()
  'SeleccionaTexto txtHC
End Sub

Private Sub txtHC_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 22 Or KeyAscii = 3) Then KeyAscii = 0
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHC_LostFocus()
  If Trim(txtHC.Text) = "" Then
    cmdBuscaCuentaPorApellidos_Click
    Exit Sub
  End If
  Set grdAtenciones.DataSource = Nothing
  Set ssImag.DataSource = Nothing
  Set ssIns.DataSource = Nothing
  Set ssLab.DataSource = Nothing
  Set ssMed.DataSource = Nothing
  Set ssProc.DataSource = Nothing
  txtAN.Text = ""
  txtNC.Text = ""
  txtD.Text = ""
  txtFA.Text = ""
  Frame1.Enabled = False
  ml_Historia = Val(txtHC.Text)
  BuscaPaciente ml_Historia
End Sub

Private Function AveriguaDevueltos(CodProducto) As Long
  AveriguaDevueltos = 0
  If oDevolucion.EOF = True And oDevolucion.BOF = True Then Exit Function
  oDevolucion.MoveFirst
  Do While Not oDevolucion.EOF
    If oDevolucion!Codigo = CodProducto Then AveriguaDevueltos = AveriguaDevueltos + oDevolucion!Cantidad
    oDevolucion.MoveNext
  Loop
End Function

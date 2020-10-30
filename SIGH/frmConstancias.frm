VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form frmConstancias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Constancias de Atención"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   Icon            =   "frmConstancias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNconstancia 
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
      Left            =   12090
      TabIndex        =   35
      Top             =   90
      Width           =   1260
   End
   Begin VB.ComboBox cmbFormato 
      Height          =   315
      ItemData        =   "frmConstancias.frx":0CCA
      Left            =   8340
      List            =   "frmConstancias.frx":0CD4
      TabIndex        =   33
      Top             =   105
      Width           =   1995
   End
   Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
      Caption         =   "..."
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
      Top             =   120
      Width           =   315
   End
   Begin UltraGrid.SSUltraGrid grdServicios 
      Height          =   1725
      Left            =   60
      TabIndex        =   16
      Top             =   2400
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   3043
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
      Caption         =   "Servicios donde estuvo el paciente en la Atención seleccionada"
   End
   Begin VB.TextBox txtAN 
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
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Width           =   4950
   End
   Begin VB.TextBox txtHC 
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
      Left            =   945
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   60
      TabIndex        =   8
      Top             =   7920
      Width           =   13350
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "Exportar a Excel"
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
         Left            =   2640
         Picture         =   "frmConstancias.frx":0D09
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Visualizar (F2)"
         DisabledPicture =   "frmConstancias.frx":101B
         DownPicture     =   "frmConstancias.frx":147B
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
         Left            =   4830
         Picture         =   "frmConstancias.frx":18F0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmConstancias.frx":1D65
         DownPicture     =   "frmConstancias.frx":2229
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
         Left            =   6368
         Picture         =   "frmConstancias.frx":2715
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdAtenciones 
      Height          =   1845
      Left            =   60
      TabIndex        =   15
      Top             =   480
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   3254
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
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3660
      Left            =   60
      TabIndex        =   10
      Top             =   4200
      Width           =   13335
      Begin MSMask.MaskEdBox txtNR 
         Height          =   345
         Left            =   120
         TabIndex        =   32
         Top             =   2580
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-########"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtFI 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1740
      End
      Begin VB.TextBox txtIdMedic 
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
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   1710
      End
      Begin VB.TextBox txtIdMedico 
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
         Height          =   315
         Left            =   11280
         TabIndex        =   26
         Top             =   3240
         Width           =   1500
      End
      Begin VB.TextBox txtNMedico 
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
         Height          =   315
         Left            =   2340
         TabIndex        =   24
         Top             =   3240
         Width           =   8460
      End
      Begin VB.TextBox txtO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   2340
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1770
         Width           =   8460
      End
      Begin VB.TextBox txtFA 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1740
      End
      Begin VB.TextBox txtD 
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
         Height          =   915
         Left            =   2340
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   480
         Width           =   8460
      End
      Begin VB.TextBox txtNC 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1740
      End
      Begin Threed.SSOption optTipoConstancia 
         Height          =   255
         Index           =   0
         Left            =   11280
         TabIndex        =   21
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   262144
         Caption         =   "Atención Médica"
      End
      Begin Threed.SSOption optTipoConstancia 
         Height          =   255
         Index           =   1
         Left            =   11280
         TabIndex        =   22
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   262144
         Caption         =   "Atención Psicológica"
      End
      Begin Threed.SSOption optTipoConstancia 
         Height          =   255
         Index           =   2
         Left            =   11280
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   262144
         Caption         =   "Hospitalización"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso"
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
         TabIndex        =   31
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Id Médico"
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
         TabIndex        =   29
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Colegiatura Médico"
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
         Left            =   11280
         TabIndex        =   27
         Top             =   3000
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Médico"
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
         Left            =   2340
         TabIndex        =   25
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Left            =   2340
         TabIndex        =   20
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Número de Recibo"
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
         Top             =   2280
         Width           =   1515
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
         Left            =   2340
         TabIndex        =   14
         Top             =   270
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
         Left            =   120
         TabIndex        =   13
         Top             =   270
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
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Servicio ..."
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
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "N° Constancia"
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
      Left            =   10920
      TabIndex        =   36
      Top             =   135
      Width           =   1140
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Formato"
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
      Left            =   7590
      TabIndex        =   34
      Top             =   165
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "H. Clínica"
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
      TabIndex        =   9
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmConstancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Constancias
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim oPaciente As New Pacientes
Dim mi_Opcion As sghOpciones
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_paciente As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim ml_IdPaciente As Long
Dim ml_IdServicio As Long
Dim ml_idAtencion As Long
Dim ml_idAtencion1 As Long
Dim ml_idTipoConstancia1 As Long
Dim ml_Recibo1 As String
Dim ml_idServicio1 As Long
Dim ml_Observaciones1 As String

Dim ml_IdTipoServicio As Long
Dim ml_TipoConstancia As Long
Dim ml_idCama As Long
Dim ml_FechaAlta As String
Dim ml_FechaIngreso As String
Dim ml_Cama As String
Dim ml_Diagnostico As String
Dim ml_Historia As Long
Dim ml_Nombres As String
Dim ml_CAfiliacion As String
Dim ml_servicio As String
Dim mo_cboServicio As New sighEntidades.ListaDespleglable
Dim gridInfra As New GridInfragistic
Dim ml_NombreMaquina As String
Dim ml_NConstancia As String
Dim ml_NColegiatura As String
Dim ml_MedicoIngreso As String
Dim ml_MedicoEgreso As String
Dim ml_Medico As String
Dim ml_Hospitaliza As Boolean
Dim ml_idConstancia As Long

Dim oDOPaciente As New doPaciente
Dim oDOCama As New DOCama
Dim oServicios As ADODB.Recordset
Dim oAtenciones As ADODB.Recordset
Dim oDiagnosticos As ADODB.Recordset
Dim oMedico As ADODB.Recordset

Dim rsTmp As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnIdDx As Long


Property Let lcNombrePc(lValue As String)
  mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
  mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
  mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let idConstancia(lValue As Long)
   ml_idConstancia = lValue
End Property

Property Let Historia(lValue As Long)
  ml_Historia = lValue
End Property

Property Let idAtencion(lValue As Long)
  ml_idAtencion1 = lValue
End Property

Property Let idTipoConstancia(lValue As Long)
  ml_idTipoConstancia1 = lValue
End Property

Property Let Recibo(lValue As String)
  ml_Recibo1 = lValue
End Property

Property Let Observaciones(lValue As String)
  ml_Observaciones1 = lValue
End Property

Property Let IdServicio(lValue As Long)
  ml_idServicio1 = lValue
End Property

Sub CargaDB_TextBox(Tabla As ADODB.Recordset, T As TextBox)
  Dim K As Integer
  T.Text = ""
  K = 0
  If Tabla.EOF = True And Tabla.BOF = True Then Exit Sub
  Tabla.MoveFirst
  Do While Not (Tabla.EOF)
    K = K + 1
    If K = 1 Then
      T.Text = Tabla!descripcion
    Else
      T.Text = T.Text & vbCrLf & Tabla!descripcion
    End If
    Tabla.MoveNext
  Loop
  Tabla.Close
End Sub

Private Function BuscaPaciente(HCPaciente As Long)
  If HCPaciente = 0 Then Exit Function
  Set oDOPaciente = mo_paciente.PacientesSeleccionarPorHistoriaClinicaDefinitiva(Val(HCigualDNI_AgregaNUEVEaLaHistoria(txtHC.Text)))
  ml_IdPaciente = Val(oDOPaciente.idPaciente)
  If ml_IdPaciente <> 0 Then
    ml_Nombres = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno & ", " & oDOPaciente.PrimerNombre & " " & oDOPaciente.SegundoNombre
    If mi_Opcion = sghAgregar Then
      Set oAtenciones = mo_ReglasLaboratorio.AtencionesQueTuvoElPacienteConstancias(ml_IdPaciente)
      Set grdAtenciones.DataSource = oAtenciones
      grdAtenciones.Enabled = True
    Else
      mo_Formulario.HabilitarDeshabilitar txtNconstancia, False
      txtNconstancia.Text = ml_idConstancia
      Set oAtenciones = mo_ReglasLaboratorio.AtencionesQueTuvoElPacienteConstanciasPorId(ml_IdPaciente, ml_idAtencion1)
      Set grdAtenciones.DataSource = oAtenciones
      Dim tmp As UltraGrid.SSRow
      grdAtenciones.Enabled = False
      Set tmp = grdAtenciones.GetRow(ssChildRowFirst)
      grdAtenciones_Click
      With grdAtenciones.Override.RowSelectorAppearance
        .BackColor = vbGreen
        .BorderColor = vbBlue
      End With
      On Error Resume Next
      txtNR.Text = Format(ml_Recibo1, "###-########")
      If mi_Opcion <> sghModificar Then txtNR.Enabled = False
      txtO.Text = ml_Observaciones1
      txtO.Enabled = False
    End If
  Else
    ml_Nombres = ""
    ml_IdPaciente = 0
    grdAtenciones.Enabled = False
  End If
  txtAN.Text = ml_Nombres
End Function

Function ValidaDatos() As Boolean
  Dim Mensaje As String, tipo As Boolean
  IngresaDx
  Mensaje = ""
  ValidaDatos = True
  tipo = True
  If optTipoConstancia(2).Value = True Then If Trim(txtNC.Text) = "" Then Mensaje = "- Falta número de Cama" & Chr(13): tipo = False
  If Trim(txtFI.Text) = "" Then Mensaje = Mensaje & "- Falta Fecha de Ingreso." & Chr(13): tipo = tipo And False
  If optTipoConstancia(2).Value = True Then If Trim(txtFA.Text) = "" Then Mensaje = Mensaje & "- Falta Fecha de Alta" & Chr(13): tipo = tipo And False
  If Trim(txtNR.Text) = "" Then Mensaje = Mensaje & "- Falta número de Recibo de Caja." & Chr(13): tipo = tipo And False
  
  If Trim(txtNMedico.Text) = "" Or Trim(txtIdMedico.Text) = "" Then Mensaje = Mensaje & "- Falta datos de Médico." & Chr(13): tipo = tipo And False
  If Trim(txtD.Text) = "" Then Mensaje = Mensaje & "- Falta diagnóstico ." & Chr(13): tipo = tipo And False
  If optTipoConstancia(0).Value = False And optTipoConstancia(1).Value = False And optTipoConstancia(2).Value = False Then Mensaje = Mensaje & "- Falta escoger tipo de Constancia." & Chr(13): tipo = tipo And False
  If Mensaje <> "" Then MsgBox Mensaje, vbInformation, "SIGH "
  ValidaDatos = tipo
End Function

Sub IngresaDx()
    lnIdDx = 0
    If (optTipoConstancia.Item(0).Value = True Or optTipoConstancia.Item(1).Value = True) And Trim(txtD.Text) = "" Then
        Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
        Dim oDODiagnostico As DODiagnostico
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDODiagnostico Is Nothing Then
                lnIdDx = oDODiagnostico.idDiagnostico
                txtD.Text = oDODiagnostico.CodigoCIE2004 & " " & oDODiagnostico.descripcion
            End If
        End If
    End If
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
    If Me.cmbFormato.ListIndex = -1 Then
        Call MsgBox("Debe Elijir un formato.", vbInformation Or vbSystemModal, Me.Caption)
    Else
        If mi_Opcion = sghAgregar Then
            Dim oRsTmp As New Recordset
            Set oRsTmp = ConstanciasSeleccionarPorId(Val(txtNconstancia.Text))
            If oRsTmp.RecordCount > 0 Then
               MsgBox "Ese NUMERO DE CONSTANCIA ya existe", vbInformation, "CONSTANCIAS"
               txtNconstancia.Text = 0
            End If
            Set oRsTmp = Nothing
        End If
        Select Case Me.cmbFormato.ListIndex
        Case 0
            Formato1
        Case 1
            Formato2
        End Select
        If mi_Opcion = sghAgregar Then
           btnCancelar_Click
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
  Unload Me
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
  Dim oBusqueda As New SIGHNegocios.BuscaPacientes
  Dim oDOPaciente As New doPaciente
  Dim oConexion As New Connection
  oConexion.Open sighEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  oBusqueda.TipoFiltro = sghFiltrarTodos
  oBusqueda.MostrarFormulario
  If oBusqueda.BotonPresionado = sghAceptar Then
    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
    If Not oDOPaciente Is Nothing Then
      txtHC.Text = oDOPaciente.NroHistoriaClinica
      txtHC.SetFocus
      SendKeys "{TAB}"
    End If
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub Form_Load()
  ml_IdServicio = 0
  ml_IdPaciente = 0
  ml_idAtencion = 0
  ml_NombreMaquina = sighEntidades.RetornaNombrePC
  
  Select Case mi_Opcion
    Case sghAgregar
      Me.Caption = "Agregar Constacias"
    Case sghModificar
      Me.Caption = "Modificar Constancias"
    Case sghConsultar
      Me.Caption = "Consultar Constancia"
    Case sghEliminar
      Me.Caption = "Eliminar Constancia"
  End Select
  If mi_Opcion = sghAgregar Then
     txtNconstancia.Text = ConstanciasMuestraUltimoIDENT
  Else
    txtHC.Text = ml_Historia
    SendKeys "{TAB}"
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdAtenciones_Click()
  Dim oDODiagnostico As New DODiagnostico
  txtNC.Text = ""
  txtD.Text = ""
  txtFA.Text = ""
  txtFI.Text = ""
  txtIdMedic.Text = ""
  optTipoConstancia(0).Value = False
  optTipoConstancia(1).Value = False
  optTipoConstancia(2).Value = False
    
  If oAtenciones.EOF = True And oAtenciones.BOF = True Then Exit Sub
  ml_idAtencion = oAtenciones("idAtencion")
  ml_IdTipoServicio = oAtenciones!idTipoServicio
  If IsNull(oAtenciones!IdMedicoIngreso) Then
    ml_MedicoIngreso = 0
  Else
    ml_MedicoIngreso = oAtenciones!IdMedicoIngreso
  End If
  If IsNull(oAtenciones!IdMedicoEgreso) Then
    ml_MedicoEgreso = 0
  Else
    ml_MedicoEgreso = oAtenciones!IdMedicoEgreso
  End If
  Dim tmp As UltraGrid.SSRow
  If ml_IdTipoServicio = 3 Then
  'Busca servicios donde estuvo el paciente, por cada atención que tuvo --> Hospitalización
    optTipoConstancia(0).Visible = False
    optTipoConstancia(1).Visible = False
    optTipoConstancia(2).Visible = True
    If mi_Opcion = sghAgregar Then
    
    Else
      If ml_idTipoConstancia1 = 3 Then
        optTipoConstancia(2).Value = True
        optTipoConstancia(2).Enabled = False
      End If
    End If
    If mi_Opcion = sghAgregar Then
      Set oServicios = mo_ReglasLaboratorio.ServiciosDondeEstuvoElPaciente(ml_IdPaciente, ml_idAtencion)
      grdServicios.Enabled = True
    Else
      Set oServicios = mo_ReglasLaboratorio.ServiciosDondeEstuvoElPacientePorId(ml_IdPaciente, ml_idServicio1, ml_idAtencion)
      grdServicios.Enabled = False
      Set tmp = grdServicios.GetRow(ssChildRowFirst)
      grdServicios_Click
      With grdServicios.Override.RowSelectorAppearance
        .BackColor = vbGreen
        .BorderColor = vbBlue
      End With
    End If
    '
    txtD.Text = mo_ReglasFacturacion.DevuelveDxAltaMedicaTodosDx(ml_idAtencion, ml_IdTipoServicio, "")
    
    ml_Diagnostico = txtD.Text
    'Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(ml_idAtencion, ml_IdTipoServicio)
    'txtD.Text = Trim(oDODiagnostico.CodigoCIE2004) & " " & oDODiagnostico.Descripcion
    'ml_Diagnostico = txtD.Text
'    Set oDiagnosticos = mo_ReglasLaboratorio.DiagnosticosSeleccionarPorIdAtencion(ml_idAtencion)
'    CargaDB_TextBox oDiagnosticos, txtD
'    ml_Diagnostico = txtD.Text
    '
    If Not (IsNull(oAtenciones("FechaEgreso"))) Then
      ml_FechaAlta = oAtenciones("FechaEgreso")
    Else
      ml_FechaAlta = ""
    End If
    txtFA.Text = ml_FechaAlta
    If Not (IsNull(oAtenciones("FechaIngreso"))) Then
      ml_FechaIngreso = oAtenciones("FechaIngreso")
    Else
      ml_FechaIngreso = ""
    End If
    txtFI.Text = ml_FechaIngreso
    Set grdServicios.DataSource = oServicios
  Else
    'Consultorios Externos y Emergencia
    optTipoConstancia(0).Visible = True
    optTipoConstancia(1).Visible = True
    optTipoConstancia(2).Visible = False
    Set oServicios = mo_ReglasLaboratorio.ServiciosDondeSeAtendioElPaciente(ml_IdPaciente, ml_idAtencion)
    
    If mi_Opcion = sghAgregar Then
      grdServicios.Enabled = True
    Else
      grdServicios.Enabled = False
    End If
    Set tmp = grdServicios.GetRow(ssChildRowFirst)
    grdServicios_Click
    With grdServicios.Override.RowSelectorAppearance
      .BackColor = vbGreen
      .BorderColor = vbBlue
    End With
    Set oDiagnosticos = mo_ReglasLaboratorio.DiagnosticosSeleccionarPorIdAtencion(ml_idAtencion)
    If mi_Opcion = sghAgregar Then
    
    Else
      If ml_idTipoConstancia1 = 1 Then
        optTipoConstancia(0).Value = True
        optTipoConstancia(0).Enabled = False
        optTipoConstancia(1).Enabled = False
      End If
      If ml_idTipoConstancia1 = 2 Then
        optTipoConstancia(1).Value = True
        optTipoConstancia(0).Enabled = False
        optTipoConstancia(1).Enabled = False
      End If
    End If
    '
    txtD.Text = mo_ReglasFacturacion.DevuelveDxAltaMedicaTodosDx(ml_idAtencion, ml_IdTipoServicio, "")
    ml_Diagnostico = txtD.Text
'    Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(ml_idAtencion, ml_IdTipoServicio)
'    txtD.Text = Trim(oDODiagnostico.CodigoCIE2004) & " " & oDODiagnostico.Descripcion
'    ml_Diagnostico = txtD.Text
'    CargaDB_TextBox oDiagnosticos, txtD
'    ml_Diagnostico = txtD.Text
    '
    If Not (IsNull(oAtenciones("FechaEgreso"))) Then
      ml_FechaAlta = oAtenciones("FechaEgreso")
    Else
      ml_FechaAlta = ""
    End If
    txtFA.Text = ml_FechaAlta
    If Not (IsNull(oAtenciones("FechaIngreso"))) Then
      ml_FechaIngreso = oAtenciones("FechaIngreso")
    Else
      ml_FechaIngreso = ""
    End If
    txtFI.Text = ml_FechaIngreso
    Set grdServicios.DataSource = oServicios
  End If
  Set oDODiagnostico = Nothing
End Sub

Private Sub grdAtenciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  grdAtenciones.Bands(0).Columns("idAtencion").Header.Caption = "Id Atencion"
  grdAtenciones.Bands(0).Columns("idAtencion").Width = 1000
  grdAtenciones.Bands(0).Columns("FechaIngreso").Header.Caption = "Fecha Ingreso"
  grdAtenciones.Bands(0).Columns("FechaIngreso").Width = 1500
  grdAtenciones.Bands(0).Columns("HoraIngreso").Header.Caption = "Hora Ingreso"
  grdAtenciones.Bands(0).Columns("HoraIngreso").Width = 1500
  grdAtenciones.Bands(0).Columns("FechaEgreso").Header.Caption = "Fecha Egreso"
  grdAtenciones.Bands(0).Columns("FechaEgreso").Width = 1500
  grdAtenciones.Bands(0).Columns("HoraEgreso").Header.Caption = "Hora Egreso"
  grdAtenciones.Bands(0).Columns("HoraEgreso").Width = 1500
  grdAtenciones.Bands(0).Columns("Descripcion").Header.Caption = "Tipo de Servicio"
  grdAtenciones.Bands(0).Columns("Descripcion").Width = 1500
  grdAtenciones.Bands(0).Columns("idFormaPago").Hidden = True
  grdAtenciones.Bands(0).Columns("idMedicoIngreso").Hidden = True
  grdAtenciones.Bands(0).Columns("idMedicoEgreso").Hidden = True
  grdAtenciones.Bands(0).Columns("idTipoServicio").Hidden = True
  gridInfra.ConfigurarFilasBiColores grdAtenciones, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grdAtenciones_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode

End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub grdServicios_Click()
  Dim oConexion As New Connection
  oConexion.Open sighEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Label15.Caption = "Servicio ..."
  'btnAceptar.Enabled = False
  If oServicios.EOF = True And oServicios.BOF = True Then Exit Sub
  Frame1.Enabled = True
  'btnAceptar.Enabled = True
  ml_IdServicio = oServicios!IdServicio
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
  ml_servicio = oServicios!NombreServicio
  Label15.Caption = "Servicio: " & ml_servicio
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  grdServicios.Bands(0).Columns("idEstanciaHospitalaria").Header.Caption = "Id Estancia"
  grdServicios.Bands(0).Columns("idEstanciaHospitalaria").Width = 900
  grdServicios.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
  grdServicios.Bands(0).Columns("NombreServicio").Width = 4000
  grdServicios.Bands(0).Columns("IdServicio").Header.Caption = "Id Servicio"
  grdServicios.Bands(0).Columns("IdServicio").Width = 1000
  grdServicios.Bands(0).Columns("FechaOcupacion").Header.Caption = "Fecha Ingreso"
  grdServicios.Bands(0).Columns("FechaOcupacion").Width = 1300
  grdServicios.Bands(0).Columns("FechaOcupacion").Format = sighEntidades.DevuelveFechaSoloFormato_DMY
  grdServicios.Bands(0).Columns("HoraOcupacion").Header.Caption = "Hora Ingreso"
  grdServicios.Bands(0).Columns("HoraOcupacion").Width = 1200
  grdServicios.Bands(0).Columns("HoraOcupacion").Format = sighEntidades.DevuelveHoraSoloFormato_HMS
  grdServicios.Bands(0).Columns("FechaDesocupacion").Header.Caption = "Fecha Salida"
  grdServicios.Bands(0).Columns("FechaDesocupacion").Width = 1300
  grdServicios.Bands(0).Columns("FechaDesocupacion").Format = sighEntidades.DevuelveFechaSoloFormato_DMY
  grdServicios.Bands(0).Columns("HoraDesocupacion").Header.Caption = "Hora Salida"
  grdServicios.Bands(0).Columns("HoraDesocupacion").Width = 1200
  grdServicios.Bands(0).Columns("HoraDesocupacion").Format = sighEntidades.DevuelveHoraSoloFormato_HMS
  grdServicios.Bands(0).Columns("IdCama").Header.Caption = "Id Cama"
  grdServicios.Bands(0).Columns("IdCama").Width = 900
  gridInfra.ConfigurarFilasBiColores grdServicios, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub grdServicios_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode     As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode

End Sub

Private Sub optTipoConstancia_Click(Index As Integer, Value As Integer)
  txtIdMedic.Text = ""
  If Index = 0 Then
    If Value = True Then
      ml_NConstancia = "ATENCIÓN MÉDICA"
      ml_NColegiatura = "C. M. P. : "
      txtIdMedic.Text = ml_MedicoIngreso
      ml_TipoConstancia = 1
      ml_Hospitaliza = False
    End If
  ElseIf Index = 1 Then
    If Value = True Then
      ml_NConstancia = "ATENCIÓN PSICOLÓGICA"
      ml_NColegiatura = "C. PS. P. : "
      ml_TipoConstancia = 2
      txtIdMedic.Text = ml_MedicoIngreso
      ml_Hospitaliza = False
    End If
  Else
    If Value = True Then
      ml_NConstancia = "HOSPITALIZACIÓN"
      ml_NColegiatura = "C. M. P. : "
      ml_TipoConstancia = 3
      txtIdMedic.Text = ml_MedicoEgreso
      ml_Hospitaliza = True
    End If
  End If
End Sub

Private Sub txtHC_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 22 Or KeyAscii = 3) Then KeyAscii = 0
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHC_LostFocus()
  If mo_Teclado.TextoEsSoloNumeros(txtHC.Text) = False Then
    cmdBuscaCuentaPorApellidos_Click
    Exit Sub
  End If
  Set grdAtenciones.DataSource = Nothing
  Set grdServicios.DataSource = Nothing
  txtAN.Text = ""
  txtNC.Text = ""
  txtD.Text = ""
  txtFA.Text = ""
  txtFI.Text = ""
  txtNR.Text = "___-________"
  txtIdMedic.Text = ""
  txtNMedico.Text = ""
  txtIdMedico.Text = ""
  optTipoConstancia(0).Value = False
  optTipoConstancia(1).Value = False
  optTipoConstancia(2).Value = False
  grdServicios.Enabled = False
  Frame1.Enabled = False
  ml_Historia = Val(txtHC.Text)
  BuscaPaciente ml_Historia
End Sub

Private Sub txtIdMedic_Change()
  txtNMedico.Text = ""
  txtIdMedico.Text = ""
  Set oMedico = ConstanciasSeleccionaMedico(Val(txtIdMedic.Text))
  If oMedico.EOF = True And oMedico.BOF = True Then Exit Sub
  txtNMedico.Text = oMedico!apnom
  txtIdMedico.Text = oMedico!Colegiatura
End Sub

Function ConstanciasSeleccionaMedico(idMedico As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar el Medico que atendió en una atención
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open sighEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ConstanciasSeleccionaMedico"
    Set oParameter = .CreateParameter("IdMedico", adInteger, adParamInput, 0, idMedico): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ConstanciasSeleccionaMedico = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function GeneraIdentNuevoParaTablaConstancias(lnIdentNuevo As Long, oConexion As Connection) As Boolean
  If lnIdentNuevo > 0 Then
        On Error GoTo ErrGIN
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "constanciasActualizaIDENT"
          Set oParameter = .CreateParameter("@IdentNuevo", adInteger, adParamInput, 0, lnIdentNuevo): .Parameters.Append oParameter
          .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
  End If
  GeneraIdentNuevoParaTablaConstancias = True
  Exit Function
ErrGIN:
'  If Err.Number Then
'     lnIdentNuevo = lnIdentNuevo + 1
'     Resume
'  End If
End Function

Public Function constanciasAgregar(idPaciente As Double, idAtencion As Double, idResponsable As Double, _
                                   IdServicio As Long, idMedico As Long, idTipoConstancia As Long, _
                                   fecha As Date, PC As String, Recibo As String, Observacion As String, _
                                   lnIdentNuevo As Long) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oRsTmp1 As New Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim mo_ReglasSeguridad As New ReglasDeSeguridad
  Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  Dim lnIdConstancia As Long
  Dim oDOMovimientoHistoriaClinica As New DOMovimientoHistoriaClinica
  Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
  Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico
  Dim oAtencionesDiagnosticos As New AtencionesDiagnosticos
  Dim oBuscaCodigoNombre As New SIGHNegocios.ReglasComunes
  constanciasAgregar = False
  ms_MensajeError = ""
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighEntidades.CadenaConexionIntegrada
  oConexion.BeginTrans
  '
  lnIdentNuevo = lnIdentNuevo - 1
  If GeneraIdentNuevoParaTablaConstancias(lnIdentNuevo, oConexion) = False Then
     GoTo ManejadorDeError
  End If
  '

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "constanciasAgregar"
    Set oParameter = .CreateParameter("@idConstancia", adInteger, adParamOutput, 0): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, idResponsable): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idMedico", adInteger, adParamInput, 0, idMedico): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idTipoConstancia", adInteger, adParamInput, 0, idTipoConstancia): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@fecha", adDBTimeStamp, adParamInput, 0, fecha): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@pc", adVarChar, adParamInput, 32, PC): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@recibo", adVarChar, adParamInput, 20, Recibo): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@observaciones", adLongVarChar, adParamInput, 2147483647, Observacion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    lnIdConstancia = .Parameters("@idConstancia")
    ml_idConstancia = lnIdConstancia
  End With
  '
  mo_ReglasSeguridad.AuditoriaAgregarV ml_idUsuario, "A", lnIdConstancia, "Constancias", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, ""
  '
  'Agrega MOVIMIENTO DE ARCHIVO si no esta la Historia ubicada en Archivo Clinico}
  Set oMovimientosHistoriaClinica.Conexion = oConexion
  Set oRsTmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(ml_IdPaciente)
  If oRsTmp1.RecordCount > 0 Then
        If oRsTmp1.Fields!idServicioDestino <> oBuscaCodigoNombre.ParametrosIdServicioArchivoClinico And oRsTmp1.Fields!idServicioDestino <> Val(lcBuscaParametro.SeleccionaFilaParametro(256)) Then
            oDOMovimientoHistoriaClinica.idPaciente = ml_IdPaciente
            oDOMovimientoHistoriaClinica.FechaMovimiento = lcBuscaParametro.RetornaFechaHoraServidorSQL()
            oDOMovimientoHistoriaClinica.idMotivo = 7    'Transferencia
            oDOMovimientoHistoriaClinica.IdServicioOrigen = oRsTmp1.Fields!idServicioDestino
            oDOMovimientoHistoriaClinica.idServicioDestino = Val(lcBuscaParametro.SeleccionaFilaParametro(256))  'Estadistica
            oDOMovimientoHistoriaClinica.IdEmpleadoArchivo = ml_idUsuario
            oDOMovimientoHistoriaClinica.IdEmpleadoTransporte = ml_idUsuario
            oDOMovimientoHistoriaClinica.IdEmpleadoRecepcion = ml_idUsuario
            oDOMovimientoHistoriaClinica.IdGrupoMovimiento = 1
            oDOMovimientoHistoriaClinica.idAtencion = ml_idAtencion
            If Not oMovimientosHistoriaClinica.Insertar(oDOMovimientoHistoriaClinica) Then
                GoTo ManejadorDeError
            End If
        End If
  End If
  oRsTmp1.Close
  '
  If lnIdDx > 0 Then
     Set oAtencionesDiagnosticos.Conexion = oConexion
     oDOAtencionDiagnostico.idAtencion = ml_idAtencion
     'oDOAtencionDiagnostico.IdAtencionDiagnostico
     oDOAtencionDiagnostico.IdClasificacionDx = 1
     oDOAtencionDiagnostico.idDiagnostico = lnIdDx
     oDOAtencionDiagnostico.IdSubclasificacionDx = 102
     oDOAtencionDiagnostico.IdUsuarioAuditoria = ml_idUsuario
     oDOAtencionDiagnostico.labConfHIS = " "
     If Not oAtencionesDiagnosticos.Insertar(oDOAtencionDiagnostico) Then
         GoTo ManejadorDeError
     End If
  End If
  '
  oConexion.CommitTrans
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Set oDOMovimientoHistoriaClinica = Nothing
  Set oMovimientosHistoriaClinica = Nothing
  Set oRsTmp1 = Nothing
  Set mo_ReglasArchivoClinico = Nothing
  Set oDOAtencionDiagnostico = Nothing
  Set oAtencionesDiagnosticos = Nothing
  constanciasAgregar = True
  Exit Function
  
ManejadorDeError:
  oConexion.RollbackTrans
  oConexion.Close
  Set oConexion = Nothing
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Public Function constanciasEliminar(idConstancia As Double) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRsTmp1 As New Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim mo_ReglasSeguridad As New ReglasDeSeguridad
  Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
  Dim ms_MensajeError As String
  Dim oDOMovimientoHistoriaClinica As New DOMovimientoHistoriaClinica
  Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
  
  constanciasEliminar = False
  ms_MensajeError = ""
  oConexion.Open sighEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "constanciasEliminar"
    Set oParameter = .CreateParameter("@idConstancia", adInteger, adParamInput, 0, idConstancia): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  '
  mo_ReglasSeguridad.AuditoriaAgregarV ml_idUsuario, "E", CLng(idConstancia), "Constancias", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, ""
      
  'Elimina MOVIMIENTO DE ARCHIVO (si corresponde al Servicio de ESTADISTICA}
  Set oMovimientosHistoriaClinica.Conexion = oConexion
  Set oRsTmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionaUltimoMovimientoPorPaciente(ml_IdPaciente)
  If oRsTmp1.RecordCount > 0 Then
        If oRsTmp1.Fields!idServicioDestino = Val(lcBuscaParametro.SeleccionaFilaParametro(256)) Then
            oDOMovimientoHistoriaClinica.IdMovimiento = oRsTmp1.Fields!IdMovimiento
            oDOMovimientoHistoriaClinica.IdUsuarioAuditoria = ml_idUsuario
            If Not oMovimientosHistoriaClinica.Eliminar(oDOMovimientoHistoriaClinica) Then
               GoTo ManejadorDeError
            End If
        End If
  End If
  oRsTmp1.Close
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Set oDOMovimientoHistoriaClinica = Nothing
  Set oMovimientosHistoriaClinica = Nothing
  Set mo_ReglasArchivoClinico = Nothing
  Set oRsTmp1 = Nothing
  constanciasEliminar = True
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Public Function constanciasModificar(idConstancia As Double, idResponsable As Double, Recibo As String) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim mo_ReglasSeguridad As New ReglasDeSeguridad
  Dim ms_MensajeError As String
  
  constanciasModificar = False
  ms_MensajeError = ""
  oConexion.Open sighEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "constanciasModificar"
    Set oParameter = .CreateParameter("@idConstancia", adInteger, adParamInput, 0, idConstancia): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@recibo", adVarChar, adParamInput, 20, Recibo): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idResponsable", adInteger, adParamInput, 0, idResponsable): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  mo_ReglasSeguridad.AuditoriaAgregarV ml_idUsuario, "M", CLng(idConstancia), "Constancias", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, ""
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  constanciasModificar = True
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function




Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_paciente = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set mo_ReglasHoteleria = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_cboServicio = Nothing
    Set gridInfra = Nothing
    Set oDOPaciente = Nothing
    Set oDOCama = Nothing
    Set rsTmp = Nothing
    Set oConexion = Nothing
End Sub

'INICIO JVG - FORMATO 1
Private Sub Formato1()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
  If ValidaDatos = False Then Exit Sub
  Dim ml_Responsable As Long
  If mi_Opcion = sghEliminar Then
    If constanciasEliminar(CDbl(ml_idConstancia)) = True Then
      MsgBox "La constancia de Atención fue anulada correctamente", vbInformation, "SIGH "
      Exit Sub
    End If
  End If
  Me.MousePointer = 11
  Dim oRptConstanciaAM As New RptConstanciaAM
  oRptConstanciaAM.Consultorio = UCase(ml_servicio)
  oRptConstanciaAM.DIAGNOSTICO = txtD.Text
  If optTipoConstancia(2).Value = True Then
    If ml_FechaAlta <> "" Then
      oRptConstanciaAM.FechaAt = ml_FechaIngreso & "  al  " & ml_FechaAlta
    Else
      oRptConstanciaAM.FechaAt = ml_FechaIngreso & "  al  Continúa"
    End If
  Else
    oRptConstanciaAM.FechaAt = ml_FechaIngreso
  End If
  oRptConstanciaAM.fecha = ", " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " del " & Format(Now, "yyyy")
  oRptConstanciaAM.HC = "H.C. " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtHC.Text, False)
  oRptConstanciaAM.idAtencion = ml_idAtencion
  oRptConstanciaAM.idPaciente = ml_IdPaciente
  ml_Responsable = sighEntidades.Usuario
  oRptConstanciaAM.idResponsable = ml_Responsable
  oRptConstanciaAM.Medico = txtNMedico.Text
  oRptConstanciaAM.idMedico = ml_NColegiatura & txtIdMedico.Text
  oRptConstanciaAM.NConstancia = ml_NConstancia
  oRptConstanciaAM.Observacion = txtO.Text
  oRptConstanciaAM.Paciente = txtAN.Text
  oRptConstanciaAM.PC = ml_NombreMaquina
  oRptConstanciaAM.Recibo = "Recibo de Caja Nº " & txtNR.Text
  oRptConstanciaAM.Tabla = oMedico
  oRptConstanciaAM.cama = Trim(txtNC.Text)
  oRptConstanciaAM.Hospitaliza = ml_Hospitaliza
  If mi_Opcion = sghAgregar Then
     constanciasAgregar Val(ml_IdPaciente), Val(ml_idAtencion), Val(ml_Responsable), ml_IdServicio, _
                        Val(txtIdMedic.Text), ml_TipoConstancia, _
                        Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY_HM), ml_NombreMaquina, _
                        txtNR.Text, txtO.Text, Val(txtNconstancia.Text)
     oRptConstanciaAM.NConstancia = ml_idConstancia
  End If
  If mi_Opcion = sghModificar Then
    If constanciasModificar(CDbl(ml_idConstancia), Val(ml_Responsable), txtNR.Text) = True Then
      If MsgBox("Constancia de Atención modificada correctamente" & Chr(13) & Chr(13) & "Desea visualizar la constancia", vbInformation + vbOKCancel, "SIGH ") = vbCancel Then Me.MousePointer = 1: Exit Sub
    End If
  End If
  oRptConstanciaAM.CrearReporte
  Me.MousePointer = 1
End Sub

'INICIO JVG - FORMATO 2
Private Sub Formato2()
If btnAceptar.Enabled = False Then
      Exit Sub
   End If
  If ValidaDatos = False Then Exit Sub
  Dim ml_Responsable As Long
  If mi_Opcion = sghEliminar Then
    If constanciasEliminar(CDbl(ml_idConstancia)) = True Then
      MsgBox "La constancia de Atención fue anulada correctamente", vbInformation, "SIGH "
      Exit Sub
    End If
  End If
  Me.MousePointer = 11
  
  Dim oRptConstanciaAMAlternativa As New RptConstanciaAMAlternativa
  oRptConstanciaAMAlternativa.DescripcionServicio = UCase(ml_servicio)
  oRptConstanciaAMAlternativa.DIAGNOSTICO = txtD.Text
  
  If optTipoConstancia(2).Value = True Then
    If ml_FechaAlta <> "" Then
      oRptConstanciaAMAlternativa.FechasEstancia = "Desde el Dia " & ml_FechaIngreso & "  al  " & ml_FechaAlta
    Else
      oRptConstanciaAMAlternativa.FechasEstancia = "Desde el dia " & ml_FechaIngreso & " hasta la fecha"
    End If
  Else
    If ml_FechaIngreso = ml_FechaAlta Then
      oRptConstanciaAMAlternativa.FechasEstancia = "Solo el dia " & ml_FechaIngreso
    Else
      oRptConstanciaAMAlternativa.FechasEstancia = "Desde el Dia " & ml_FechaIngreso & " al " & ml_FechaAlta
    End If
  End If
  oRptConstanciaAMAlternativa.FechaDescriptivaActual = ", " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " del " & Format(Now, "yyyy")
  oRptConstanciaAMAlternativa.NroHC = "H.C. " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtHC.Text, False)
  oRptConstanciaAMAlternativa.NroConstancia = ml_idConstancia & " - " & Year(Now)
  oRptConstanciaAMAlternativa.EdadPaciente = Year(Now) - Year(oDOPaciente.FechaNacimiento)
  ml_Responsable = sighEntidades.Usuario
  oRptConstanciaAMAlternativa.NombreMedico = txtNMedico.Text
  oRptConstanciaAMAlternativa.CodigoMedico = ml_NColegiatura & txtIdMedico.Text
  oRptConstanciaAMAlternativa.TipoConstancia = ml_NConstancia
  oRptConstanciaAMAlternativa.NombrePaciente = txtAN.Text
  oRptConstanciaAMAlternativa.Datos = oMedico

  If mi_Opcion = sghAgregar Then
     constanciasAgregar Val(ml_IdPaciente), Val(ml_idAtencion), Val(ml_Responsable), ml_IdServicio, _
                        Val(txtIdMedic.Text), ml_TipoConstancia, _
                        Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY_HM), ml_NombreMaquina, _
                        txtNR.Text, txtO.Text, Val(txtNconstancia.Text)
     oRptConstanciaAMAlternativa.NroConstancia = ml_idConstancia & " - " & Year(Now)
  End If
  If mi_Opcion = sghModificar Then
    If constanciasModificar(CDbl(ml_idConstancia), Val(ml_Responsable), txtNR.Text) = True Then
      If MsgBox("Constancia de Atención modificada correctamente" & Chr(13) & Chr(13) & "Desea visualizar la constancia", vbInformation + vbOKCancel, "SIGH ") = vbCancel Then Me.MousePointer = 1: Exit Sub
    End If
  End If
  oRptConstanciaAMAlternativa.CrearReporte
  Me.MousePointer = 1
End Sub



Function ConstanciasMuestraUltimoIDENT() As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "constanciasMuestraUltimoIDENT"
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
       ml_idConstancia = oRecordset.Fields!UltimaConstancia
       ConstanciasMuestraUltimoIDENT = ml_idConstancia + 1
   Else
       GoTo ManejadorDeError
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Private Sub txtNconstancia_LostFocus()
    If mo_Teclado.TextoEsSoloNumeros(txtNconstancia.Text) Then
       Dim oRsTmp As New Recordset
       Set oRsTmp = ConstanciasSeleccionarPorId(Val(txtNconstancia.Text))
       If oRsTmp.RecordCount > 0 Then
          MsgBox "Ese NUMERO DE CONSTANCIA ya existe", vbInformation, "CONSTANCIAS"
          txtNconstancia.Text = ml_idConstancia
       End If
       Set oRsTmp = Nothing
    Else
       MsgBox "Solo debe contener NUMEROS", vbInformation, "CONSTANCIAS"
       txtNconstancia.Text = ml_idConstancia
    End If
End Sub


Function ConstanciasSeleccionarPorId(lnIdConstancia As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "constanciasSeleccionarPorId"
        Set oParameter = .CreateParameter("@idConstancia", adInteger, adParamInput, 0, lnIdConstancia): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ConstanciasSeleccionarPorId = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   oConexion.Close
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

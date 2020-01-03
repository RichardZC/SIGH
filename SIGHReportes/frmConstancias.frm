VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmConstancias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Constancias de Atención"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   Icon            =   "frmConstancias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
      Caption         =   "..."
      Height          =   315
      Left            =   2160
      TabIndex        =   18
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
      Left            =   4455
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   5220
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
      TabIndex        =   7
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
         Picture         =   "frmConstancias.frx":0CCA
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Visualizar (F2)"
         DisabledPicture =   "frmConstancias.frx":0FDC
         DownPicture     =   "frmConstancias.frx":143C
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
         Left            =   4838
         Picture         =   "frmConstancias.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmConstancias.frx":1D26
         DownPicture     =   "frmConstancias.frx":21EA
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
         Picture         =   "frmConstancias.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   6
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
         TabIndex        =   32
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
         TabIndex        =   30
         Top             =   3240
         Width           =   1500
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
         TabIndex        =   28
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
         TabIndex        =   26
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
         TabIndex        =   21
         Top             =   1770
         Width           =   8460
      End
      Begin VB.TextBox txtNR 
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
         Top             =   2520
         Width           =   1500
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   480
         Width           =   1740
      End
      Begin Threed.SSOption optTipoConstancia 
         Height          =   255
         Index           =   0
         Left            =   11280
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   29
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
         TabIndex        =   27
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
         TabIndex        =   22
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
         TabIndex        =   20
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
      Left            =   2760
      TabIndex        =   9
      Top             =   150
      Width           =   1635
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
      TabIndex        =   8
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmConstancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oPaciente As New Pacientes
Dim mi_Opcion As sghOpciones
Dim mo_Teclado As SIGHComun.Teclado
Dim mo_paciente As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idUsuario As Long
Dim ml_idPaciente As Long
Dim ml_idServicio As Long
Dim ml_idAtencion As Long
Dim ml_idAtencion1 As Long
Dim ml_idTipoConstancia1 As Long
Dim ml_Recibo1 As String
Dim ml_idServicio1 As Long
Dim ml_Observaciones1 As String

Dim ml_idTipoServicio As Long
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
'Dim ml_Alta As String
Dim mo_cboServicio As New SIGHComun.ListaDespleglable
Dim gridInfra As New GridInfragistic
Dim ml_NombreMaquina As String
Dim ml_NConstancia As String
Dim ml_NColegiatura As String
Dim ml_MedicoIngreso As String
Dim ml_MedicoEgreso As String
Dim ml_Medico As String
Dim ml_Hospitaliza As Boolean

Dim oDOPaciente As New doPaciente
Dim oDOCama As New DOCama
Dim oServicios As ADODB.Recordset
Dim oAtenciones As ADODB.Recordset
Dim oDiagnosticos As ADODB.Recordset
Dim oMedico As ADODB.Recordset

Dim rsTmp As New Recordset
Dim oConexion As New ADODB.Connection
  
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
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

Property Let idServicio(lValue As Long)
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
      T.Text = Tabla!Descripcion
    Else
      T.Text = T.Text & vbCrLf & Tabla!Descripcion
    End If
    Tabla.MoveNext
  Loop
  Tabla.Close
End Sub

Private Function BuscaPaciente(HCPaciente As Long)
  If HCPaciente = 0 Then Exit Function
  Set oDOPaciente = mo_paciente.PacientesSeleccionarPorHistoriaClinicaDefinitiva(HCPaciente)
  ml_idPaciente = Val(oDOPaciente.IdPaciente)
  If ml_idPaciente <> 0 Then
    ml_Nombres = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno & ", " & oDOPaciente.PrimerNombre & " " & oDOPaciente.segundoNombre
    If mi_Opcion = sghAgregar Then
      Set oAtenciones = mo_ReglasLaboratorio.AtencionesQueTuvoElPacienteConstancias(ml_idPaciente)
      Set grdAtenciones.DataSource = oAtenciones
      grdAtenciones.Enabled = True
    Else
      Set oAtenciones = mo_ReglasLaboratorio.AtencionesQueTuvoElPacienteConstanciasPorId(ml_idPaciente, ml_idAtencion1)
      Set grdAtenciones.DataSource = oAtenciones
      Dim tmp As UltraGrid.SSRow
      grdAtenciones.Enabled = False
      Set tmp = grdAtenciones.GetRow(ssChildRowFirst)
      grdAtenciones_Click
      With grdAtenciones.Override.RowSelectorAppearance
        .BackColor = vbGreen
        .BorderColor = vbBlue
      End With
      txtNR.Text = ml_Recibo1
      txtNR.Enabled = False
      txtO.Text = ml_Observaciones1
      txtO.Enabled = False
    End If
  Else
    ml_Nombres = ""
    ml_idPaciente = 0
    grdAtenciones.Enabled = False
  End If
  txtAN.Text = ml_Nombres
End Function

Function ValidaDatos() As Boolean
  Dim Mensaje As String, Tipo As Boolean
  Mensaje = ""
  ValidaDatos = True
  Tipo = True
  If optTipoConstancia(2).Value = True Then If Trim(txtNC.Text) = "" Then Mensaje = "- Falta número de Cama" & Chr(13): Tipo = False
  If Trim(txtFI.Text) = "" Then Mensaje = Mensaje & "- Falta Fecha de Ingreso." & Chr(13): Tipo = Tipo And False
  If optTipoConstancia(2).Value = True Then If Trim(txtFA.Text) = "" Then Mensaje = Mensaje & "- Falta Fecha de Alta" & Chr(13): Tipo = Tipo And False
  If Trim(txtNR.Text) = "" Then Mensaje = Mensaje & "- Falta número de Recibo de Caja." & Chr(13): Tipo = Tipo And False
  
  If Trim(txtNMedico.Text) = "" Or Trim(txtIdMedico.Text) = "" Then Mensaje = Mensaje & "- Falta datos de Médico." & Chr(13): Tipo = Tipo And False
  If Trim(txtD.Text) = "" Then Mensaje = Mensaje & "- Falta diagnóstico ." & Chr(13): Tipo = Tipo And False
  If optTipoConstancia(0).Value = False And optTipoConstancia(1).Value = False And optTipoConstancia(2).Value = False Then Mensaje = Mensaje & "- Falta escoger tipo de Constancia." & Chr(13): Tipo = Tipo And False
  If Mensaje <> "" Then MsgBox Mensaje, vbCritical, "SIGH GalenHos"
  ValidaDatos = Tipo
End Function

Private Sub btnAceptar_Click()
  If ValidaDatos = False Then Exit Sub
  Dim ml_Responsable As Long
  Me.MousePointer = 11
  Dim oRptConstanciaAM As New RptConstanciaAM
  oRptConstanciaAM.Consultorio = UCase(ml_servicio)
  oRptConstanciaAM.Diagnostico = txtD.Text
  If optTipoConstancia(2).Value = True Then
    If ml_FechaAlta <> "" Then
      oRptConstanciaAM.FechaAt = ml_FechaIngreso & "  al  " & ml_FechaAlta
    Else
      oRptConstanciaAM.FechaAt = ml_FechaIngreso & "  al  Continúa"
    End If
  Else
    oRptConstanciaAM.FechaAt = ml_FechaIngreso
  End If
  oRptConstanciaAM.Fecha = "Ayacucho, " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " del " & Format(Now, "yyyy")
  oRptConstanciaAM.HC = "H.C. " & txtHC.Text
  oRptConstanciaAM.idAtencion = ml_idAtencion
  oRptConstanciaAM.IdPaciente = ml_idPaciente
  ml_Responsable = SIGHComun.Usuario
  oRptConstanciaAM.IdResponsable = ml_Responsable
  oRptConstanciaAM.Medico = txtNMedico.Text
  oRptConstanciaAM.idMedico = ml_NColegiatura & txtIdMedico.Text
  oRptConstanciaAM.NConstancia = ml_NConstancia
  oRptConstanciaAM.Observacion = txtO.Text
  oRptConstanciaAM.Paciente = txtAN.Text
  oRptConstanciaAM.PC = ml_NombreMaquina
  oRptConstanciaAM.Recibo = "Recibo de Caja Nº " & txtNR.Text
  oRptConstanciaAM.Tabla = oMedico
  oRptConstanciaAM.Cama = Trim(txtNC.Text)
  oRptConstanciaAM.Hospitaliza = ml_Hospitaliza
  If mi_Opcion = sghAgregar Then AgregaConstancia Val(ml_idPaciente), Val(ml_idAtencion), Val(ml_Responsable), ml_idServicio, Val(txtIdMedic.Text), ml_TipoConstancia, Format(Now, "DD/MM/YYYY HH:MM"), ml_NombreMaquina, txtNR.Text, txtO.Text
  oRptConstanciaAM.CrearReporte
  Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
  Unload Me
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
  Dim oBusqueda As New SIGHNegocios.BuscaPacientes
  Dim oDOPaciente As New doPaciente
  oBusqueda.TipoFiltro = sghFiltrarTodos
  oBusqueda.MostrarFormulario
  If oBusqueda.BotonPresionado = sghAceptar Then
    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
    If Not oDOPaciente Is Nothing Then
      txtHC.Text = oDOPaciente.nroHistoriaClinica
      txtHC.SetFocus
      SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub Form_Load()
  ml_idServicio = 0
  ml_idPaciente = 0
  ml_idAtencion = 0
  ml_NombreMaquina = SIGHComun.RetornaNombrePC
  
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
    
  Else
'    grdAtenciones.Visible = False
'    grdServicios.Visible = False
'    Frame1.Top = grdAtenciones.Top
'    Frame3.Top = Frame1.Top + Frame1.Height + 60
'    Me.Height = Frame3.Top + Frame3.Height + 435
    txtHC.Text = ml_Historia
    SendKeys "{TAB}"
  End If
End Sub


Private Sub grdAtenciones_Click()
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
  ml_idTipoServicio = oAtenciones!idTipoServicio
  If IsNull(oAtenciones!idMedicoIngreso) Then
    ml_MedicoIngreso = 0
  Else
    ml_MedicoIngreso = oAtenciones!idMedicoIngreso
  End If
  If IsNull(oAtenciones!idMedicoEgreso) Then
    ml_MedicoEgreso = 0
  Else
    ml_MedicoEgreso = oAtenciones!idMedicoEgreso
  End If
  Dim tmp As UltraGrid.SSRow
  If ml_idTipoServicio = 3 Then
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
      Set oServicios = mo_ReglasLaboratorio.ServiciosDondeEstuvoElPaciente(ml_idPaciente, ml_idAtencion)
    Else
      Set oServicios = mo_ReglasLaboratorio.ServiciosDondeEstuvoElPacientePorId(ml_idPaciente, ml_idServicio1, ml_idAtencion)
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
    End If
    Set oDiagnosticos = mo_ReglasLaboratorio.DiagnosticosSeleccionarPorIdAtencion(ml_idAtencion)
    CargaDB_TextBox oDiagnosticos, txtD
    ml_Diagnostico = txtD.Text
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
    Set oServicios = mo_ReglasLaboratorio.ServiciosDondeSeAtendioElPaciente(ml_idPaciente, ml_idAtencion)
    
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
    CargaDB_TextBox oDiagnosticos, txtD
    ml_Diagnostico = txtD.Text
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
  grdAtenciones.Bands(0).Columns("Descripcion").Width = 3000
  grdAtenciones.Bands(0).Columns("idFormaPago").Hidden = True
  grdAtenciones.Bands(0).Columns("idMedicoIngreso").Hidden = True
  grdAtenciones.Bands(0).Columns("idMedicoEgreso").Hidden = True
  grdAtenciones.Bands(0).Columns("idTipoServicio").Hidden = True
  gridInfra.ConfigurarFilasBiColores grdAtenciones, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub grdServicios_Click()
  Label15.Caption = "Servicio ..."
  'btnAceptar.Enabled = False
  If oServicios.EOF = True And oServicios.BOF = True Then Exit Sub
  Frame1.Enabled = True
  'btnAceptar.Enabled = True
  ml_idServicio = oServicios!idServicio
  If Not (IsNull(oServicios!idCama)) Then
    ml_idCama = oServicios!idCama
  Else
    ml_idCama = 0
  End If
  If Val(ml_idCama) <> 0 Then
    Set oDOCama = mo_ReglasHoteleria.CamasSeleccionarPorId(ml_idCama)
    ml_Cama = oDOCama.codigo
  Else
    ml_Cama = ""
  End If
  'If Not (IsNull(oServicios!FechaDesocupacion)) Then
  '  ml_FechaAlta = oServicios!FechaDesocupacion
  'Else
  '  ml_FechaAlta = ""
  'End If
  'If Not (IsNull(oServicios!FechaOcupacion)) Then
  '  ml_FechaIngreso = oServicios!FechaOcupacion
  'Else
  '  ml_FechaIngreso = ""
  'End If
  'txtFI.Text = ml_FechaIngreso
  'txtFA.Text = ml_FechaAlta
  txtNC.Text = ml_Cama
  ml_servicio = oServicios!NombreServicio
  Label15.Caption = "Servicio: " & ml_servicio
  If Val(ml_idServicio) <> 0 Then
    
  Else
    
  End If
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
  grdServicios.Bands(0).Columns("FechaOcupacion").Format = "dd/mm/yyyy"
  grdServicios.Bands(0).Columns("HoraOcupacion").Header.Caption = "Hora Ingreso"
  grdServicios.Bands(0).Columns("HoraOcupacion").Width = 1200
  grdServicios.Bands(0).Columns("HoraOcupacion").Format = "hh:mm:ss"
  grdServicios.Bands(0).Columns("FechaDesocupacion").Header.Caption = "Fecha Salida"
  grdServicios.Bands(0).Columns("FechaDesocupacion").Width = 1300
  grdServicios.Bands(0).Columns("FechaDesocupacion").Format = "dd/mm/yyyy"
  grdServicios.Bands(0).Columns("HoraDesocupacion").Header.Caption = "Hora Salida"
  grdServicios.Bands(0).Columns("HoraDesocupacion").Width = 1200
  grdServicios.Bands(0).Columns("HoraDesocupacion").Format = "hh:mm:ss"
  grdServicios.Bands(0).Columns("IdCama").Header.Caption = "Id Cama"
  grdServicios.Bands(0).Columns("IdCama").Width = 900
  gridInfra.ConfigurarFilasBiColores grdServicios, SIGHComun.GrillaConFilasBicolor
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
  If Trim(txtHC.Text) = "" Then
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
  txtNR.Text = ""
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
  txtIdMedico.Text = oMedico!colegiatura
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
  oConexion.Open SIGHComun.CadenaConexion
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

Public Function AgregaConstancia(IdPaciente As Double, idAtencion As Double, IdResponsable As Double, idServicio As Long, idMedico As Long, idTipoConstancia As Long, Fecha As Date, PC As String, Recibo As String, Observacion As String)
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  Busca = "INSERT INTO constancias (idPaciente, idAtencion, idResponsable, idServicio, idMedico, idTipoConstancia, Fecha, PC, recibo, observaciones) VALUES ('" & IdPaciente & "', '" & idAtencion & "', '" & IdResponsable & "', '" & idServicio & "', '" & idMedico & "', '" & idTipoConstancia & "', getdate(), '" & PC & "', '" & Recibo & "', '" & Observacion & "')"
  oTabla.CursorLocation = adUseClient
  oTabla.LockType = adLockOptimistic
  oTabla.Open Busca, oConexSQL
  oConexSQL.Close
End Function


VERSION 5.00
Begin VB.UserControl ucTriaje 
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   LockControls    =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   11055
   Begin VB.Frame frDatosTriaje 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      Begin VB.TextBox txtSaturacionOxigeno 
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
         Left            =   9060
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1965
         Width           =   975
      End
      Begin VB.TextBox txtPermAbdominal 
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtPresionDiast 
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
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtPresionSist 
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1080
         Width           =   455
      End
      Begin VB.TextBox txtFrespiratoria 
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
         Left            =   9030
         MaxLength       =   2
         TabIndex        =   7
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox txtPulso 
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox txtTemperatura 
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   2
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtTalla 
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
         Left            =   9030
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1110
         Width           =   975
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
         Left            =   9030
         MaxLength       =   6
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPerimetroCefalico 
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
         Left            =   9030
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtFrecuenciaCardiaca 
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "95 a 100"
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
         Left            =   10170
         TabIndex        =   34
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Saturación de Oxígeno"
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
         Left            =   7140
         TabIndex        =   33
         Top             =   1980
         Width           =   1860
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "cm."
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
         Left            =   3195
         TabIndex        =   32
         Top             =   2010
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Perímetro abdominal"
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
         Left            =   135
         TabIndex        =   31
         Top             =   1965
         Width           =   1665
      End
      Begin VB.Label lblIMC 
         AutoSize        =   -1  'True
         Caption         =   "..."
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
         Left            =   7500
         TabIndex        =   30
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   29
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label lblAlertaTemperatura 
         Caption         =   "(..........FIEBRE...........)"
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
         Left            =   4230
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Pulso"
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
         Left            =   150
         TabIndex        =   27
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Frecuencia Respiratoria"
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
         Left            =   7110
         TabIndex        =   26
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label lblNormalFrecuenciaResp 
         AutoSize        =   -1  'True
         Caption         =   "10 a 20"
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
         Left            =   10140
         TabIndex        =   25
         Top             =   225
         Width           =   705
      End
      Begin VB.Label lblNormalPulso 
         AutoSize        =   -1  'True
         Caption         =   "60 a 100"
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
         Left            =   3240
         TabIndex        =   24
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblNormalTemperatura 
         AutoSize        =   -1  'True
         Caption         =   "° C"
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
         Left            =   3240
         TabIndex        =   23
         Top             =   720
         Width           =   270
      End
      Begin VB.Label lblNormalPeso 
         AutoSize        =   -1  'True
         Caption         =   "Kg."
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
         Left            =   10140
         TabIndex        =   22
         Top             =   720
         Width           =   300
      End
      Begin VB.Label lblNormalTall 
         AutoSize        =   -1  'True
         Caption         =   "cm."
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
         Left            =   10140
         TabIndex        =   21
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label lblNormalPresionArterial 
         AutoSize        =   -1  'True
         Caption         =   "Sistólica/Diastólica"
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
         Left            =   3195
         TabIndex        =   20
         Top             =   1140
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peso"
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
         Left            =   8580
         TabIndex        =   19
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Temperatura"
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
         Left            =   150
         TabIndex        =   18
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Presión Arterial"
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
         Left            =   150
         TabIndex        =   17
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
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
         Left            =   8610
         TabIndex        =   16
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Perímetro Cefálico"
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
         Left            =   7500
         TabIndex        =   15
         Top             =   1590
         Width           =   1470
      End
      Begin VB.Label lblNormalPerimetroCefalico 
         AutoSize        =   -1  'True
         Caption         =   "cm."
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
         Left            =   10140
         TabIndex        =   14
         Top             =   1590
         Width           =   315
      End
      Begin VB.Label lblNormalFrecuenciaCardiaca 
         AutoSize        =   -1  'True
         Caption         =   "10 a 20"
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
         Left            =   3150
         TabIndex        =   13
         Top             =   1575
         Width           =   705
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Frecuencia Cardiaca"
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
         Top             =   1575
         Width           =   1590
      End
   End
End
Attribute VB_Name = "ucTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para el mantenimiento de Triaje
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros
'Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
'variables standares para el mantenimiento y auditoria
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mb_ExistenDatos As Boolean

Dim ml_idAtencion As Long
Dim ml_idPaciente As Long
Dim ml_edadPacienteEnDias As Long
'despierto, dormido
Dim ml_EstadoPaciente As Long

Dim mo_DOAtencion As New DOAtencion
Dim mo_DOAtencionCE As New DOAtencionesCE
Dim mo_DOAtencionCEAnterior As DOAtencionesCE
Dim mo_DoPaciente As New DOPaciente
Dim mb_CambioDatosTriaje As Boolean
Dim md_fechaTriaje As Date
Dim mb_EsAtencionCRED As Boolean
Dim lnTallaDeUltimaAtencion As Double
Dim oReglasTriaje As New ReglasTriaje


Dim ml_triajeOrigen As sightriajeorigen
Dim rsValoresNormalesTriaje As ADODB.Recordset
Dim rsDatosTriaje As ADODB.Recordset

Dim rsTriajeValiable As ADODB.Recordset
'variables usadas para la verificacion de la coherencia en el ingreso de datos
Dim cNumberError As New Collection
Dim cLimitesCoherentes As New Collection
'mgaray20141013
Dim oRsTriajeExcepciones As ADODB.Recordset

'=============================================================
'EVENTOS DEL CONTROL
'=============================================================
Public Event SePresionoTeclaEspecial(KeyCode As Integer)

'=============================================================
'EVENTOS
'=============================================================

Public Sub FocusPulso()
    If txtPulso.Visible = True Then
        UserControl.txtPulso.SetFocus
    End If
End Sub

'=============================================================
'CONSTRUCTOR
'=============================================================

'Private Sub Class_Initialize()
'    Call setIdentificadorVariable
'End Sub
'=============================================================
'METODOS DE LECTURA
'=============================================================
Property Get SeCambioDatosTriaje() As Long
   SeCambioDatosTriaje = mb_CambioDatosTriaje
End Property

Property Get DOAtencion() As DOAtencion
   Set DOAtencion = mo_DOAtencion
End Property

Property Get DOPaciente() As DOPaciente
   Set DOPaciente = mo_DoPaciente
End Property

Property Get EdadPacienteEnDias() As Long
   EdadPacienteEnDias = ml_edadPacienteEnDias
End Property


'=============================================================
'METODOS DE LECTURA Y ESCRITURA
'=============================================================
Property Let FechaTriaje(lValue As Date)
   md_fechaTriaje = lValue
End Property

Property Get FechaTriaje() As Date
   FechaTriaje = md_fechaTriaje
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property

Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property


Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
   setIdAtencion (lValue)
End Property

Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property

Property Set DOAtencionCE(lValue As DOAtencionesCE)
    Set mo_DOAtencionCE = lValue
End Property

Property Get DOAtencionCE() As Long
   DOAtencionCE = mo_DOAtencionCE
End Property

Property Let TriajeOrigen(lValue As sightriajeorigen)
   ml_triajeOrigen = lValue
End Property

Property Get TriajeOrigen() As sightriajeorigen
   TriajeOrigen = ml_triajeOrigen
End Property

Property Let IdPaciente(lValue As Long)
    setIdPaciente (lValue)
End Property

Property Get IdPaciente() As Long
   IdPaciente = ml_idPaciente
End Property

Property Let EstadoPaciente(lValue As Long)
    ml_EstadoPaciente = lValue
End Property

Property Get EstadoPaciente() As Long
   IdPaciente = ml_EstadoPaciente
End Property

'=============================================================
'METODOS DE ESCRITURA
'=============================================================

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Private Sub txtFrecuenciaCardiaca_GotFocus()
    mo_Formulario.controlSelectText txtFrecuenciaCardiaca
End Sub

Property Let EsAtencionCRED(bValue As Boolean)
   mb_EsAtencionCRED = bValue
End Property

'=============================================================
'EVENTOS DE CONTROLES
'=============================================================
Private Sub txtFrecuenciaCardiaca_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFrecuenciaCardiaca
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFrecuenciaCardiaca_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtFrespiratoria_GotFocus()
    mo_Formulario.controlSelectText txtFrespiratoria
End Sub

Private Sub txtFrespiratoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFrespiratoria
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFrespiratoria_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPerimetroCefalico_GotFocus()
    mo_Formulario.controlSelectText txtPerimetroCefalico
End Sub

Private Sub txtPerimetroCefalico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPerimetroCefalico
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPerimetroCefalico_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtPermAbdominal_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPermAbdominal
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub txtPermAbdominal_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtPeso_GotFocus()
    mo_Formulario.controlSelectText txtPeso
End Sub

Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPeso
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

'Private Sub txtPresion_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtPresion
'    RaiseEvent SePresionoTeclaEspecial(KeyCode)
'End Sub


Private Sub txtPresionDiast_GotFocus()
    mo_Formulario.controlSelectText txtPresionDiast
End Sub

Private Sub txtPresionDiast_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPresionDiast
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPresionDiast_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPresionSist_GotFocus()
    mo_Formulario.controlSelectText txtPresionSist
End Sub

Private Sub txtPresionSist_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPresionSist
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPresionSist_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPulso_GotFocus()
    mo_Formulario.controlSelectText txtPulso
End Sub

Private Sub txtPulso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPulso
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPulso_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub





Private Sub txtSaturacionOxigeno_GotFocus()
    mo_Formulario.controlSelectText txtSaturacionOxigeno
End Sub

Private Sub txtSaturacionOxigeno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSaturacionOxigeno
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtSaturacionOxigeno_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSaturacionOxigeno_LostFocus()
    If Not (Val(txtSaturacionOxigeno.Text) >= 0 And Val(txtSaturacionOxigeno.Text) <= 100) Then
       MsgBox "La SATURACION DE OXIGENO debe de estar entre 0 y 100", vbInformation, ""
       txtSaturacionOxigeno.Text = ""
    End If
End Sub

Private Sub txtTalla_GotFocus()
    mo_Formulario.controlSelectText txtTalla
End Sub

Private Sub txtTalla_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtTalla
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtTalla_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

'debb-09/06/2016
Private Sub txtTalla_LostFocus()
   
    If Val(txtTalla.Text) > 0 Then
       If Val(txtTalla.Text) < lnTallaDeUltimaAtencion Then
          If MsgBox("La nueva TALLA no debería ser menor a " & Trim(Str(lnTallaDeUltimaAtencion)) & " de la última atención" & _
                    Chr(13) + "¿Es correcta la nueva TALLA?", vbQuestion + vbYesNo, "") = vbNo Then
             txtTalla.Text = lnTallaDeUltimaAtencion
             CalculaIMC
             Exit Sub
          End If
       End If
       If Val(txtTalla.Text) < 10 Then
          MsgBox "Debe registrar una talla mayor a 10cm", vbInformation, ""
          txtTalla.Text = ""
          CalculaIMC
          Exit Sub
       End If
       
    End If
    txtTalla.Text = Format(txtTalla, "##0.0")
    CalculaIMC
End Sub

Private Sub txtTemperatura_GotFocus()
    mo_Formulario.controlSelectText txtTemperatura
End Sub

Private Sub txtTemperatura_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtTemperatura
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtTemperatura_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Public Sub txtTemperatura_LostFocus()
'    lblAlertaTemperatura.Visible = False
'    If txtTemperatura.Text <> "" Then
'        If buscarValorNormalVariable(Temperatura) = True Then
            lblAlertaTemperatura.Visible = oReglasTriaje.RetornaTieneFiebre(txtTemperatura.Text, rsValoresNormalesTriaje)
'            If Val(txtTemperatura.Text) > rsValoresNormalesTriaje!ValorNormalMaximo Then
'                If IsNull(rsValoresNormalesTriaje!ValorCoherenteMaximo) Then
'                    lblAlertaTemperatura.Visible = True
'                ElseIf Val(txtTemperatura.Text) <= rsValoresNormalesTriaje!ValorCoherenteMaximo Then
'                    lblAlertaTemperatura.Visible = True
'                End If
'            End If
'        End If
'    End If
End Sub

'===================================================================
'METODOS PUBLICOS
'===================================================================

Private Function setIdAtencion(lValue As Long) As Boolean
    Dim bReturnValue As Boolean
    bReturnValue = False
    
    ml_idAtencion = lValue
    If ml_idAtencion = 0 Then
        Call BloqueoTodosLosControles
        ms_MensajeError = "Id atención indicado no valido (" & ml_idAtencion & ")"
    Else
        Dim oReglasAdmision As New ReglasAdmision
        Dim oConexion As New ADODB.Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
    
        Set mo_DOAtencion = oReglasAdmision.AtencionesSeleccionarPorId(ml_idAtencion, oConexion)
        If mo_DOAtencion Is Nothing Then
            ms_MensajeError = "Id Atención Indicado No Existe"
        Else
            setIdPaciente (mo_DOAtencion.IdPaciente)
            bReturnValue = True
        End If
        oConexion.Close
        Set oConexion = Nothing
    End If
    setIdAtencion = bReturnValue
End Function

Private Function setIdPaciente(lValue As Long) As Boolean
On Error GoTo miError
    Dim bReturnValue As Boolean
    bReturnValue = False
    
    ml_idPaciente = lValue
    If ml_idPaciente = 0 Then
        Call BloqueoTodosLosControles
        ms_MensajeError = "Id de paciente no valido (" & ml_idPaciente & ")"
        'Err.Raise Number:=vbObjectError + 1234 ',  "Id de Paciente no valido(" & ml_idPaciente & ")"
    Else
        Dim oReglasAdmision As New ReglasAdmision
        Dim oConexion As New ADODB.Connection
        
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
    
        Set mo_DoPaciente = oReglasAdmision.PacientesSeleccionarPorId(ml_idPaciente, oConexion)
        If mo_DoPaciente Is Nothing Then
            ms_MensajeError = "Id paciente indicado no existe"
        Else
            Dim oFechaHora As New FechaHora
            ml_edadPacienteEnDias = oFechaHora.EdadActualEnDias(mo_DoPaciente.FechaNacimiento, md_fechaTriaje)
            bReturnValue = True
        End If
        oConexion.Close
        Set oConexion = Nothing
    End If
miError:
    If Err Then
        ms_MensajeError = Err.Number & " " & Err.Description
    End If
    setIdPaciente = bReturnValue
End Function

Public Sub LimpiarControles()
    ml_idAtencion = 0
    ml_idPaciente = 0
    
    UserControl.txtPeso.Text = ""
'    UserControl.txtPresion.Text = SIGHEntidades.PresionDevuelveVacia
    UserControl.txtTalla.Text = ""
    UserControl.txtTemperatura.Text = ""
    UserControl.txtFrespiratoria.Text = ""
    UserControl.txtPulso.Text = ""
    UserControl.txtPresionDiast = ""
    UserControl.txtPresionSist = ""
    UserControl.txtPerimetroCefalico.Text = ""
    UserControl.txtFrecuenciaCardiaca = ""
    UserControl.txtPermAbdominal.Text = ""
    lblAlertaTemperatura.Visible = False
    lblIMC.Caption = ""
End Sub

Private Function ValidarCoherenciaDatos() As Variant
    Dim listError() As Integer
    
End Function

Private Function devuelveMensajeError()
    Dim listError() As String
    
    
'    sighTriajeVariable.Peso
    
End Function

Private Sub setIdentificadorVariable()
    txtPulso.Tag = sighTriajeVariable.Pulso
    txtTemperatura.Tag = sighTriajeVariable.Temperatura
    
    'txtPresion.Tag = sighTriajeVariable.PresArtSistolica
'    txtPresion.Visible = False
    
    txtPresionDiast.Tag = sighTriajeVariable.PresArtDiastolica
    txtPresionSist.Tag = sighTriajeVariable.PresArtSistolica
    
    txtFrecuenciaCardiaca.Tag = sighTriajeVariable.FrecCardiaca
    txtFrespiratoria.Tag = sighTriajeVariable.FrecRespiratoria
    txtPeso.Tag = sighTriajeVariable.Peso
    txtTalla.Tag = sighTriajeVariable.Talla
    txtPerimetroCefalico.Tag = sighTriajeVariable.PerimCefalico
    
    '=====================
    lblNormalPulso.Tag = sighTriajeVariable.Pulso
    lblNormalTemperatura.Tag = sighTriajeVariable.Temperatura
    lblNormalPresionArterial.Tag = sighTriajeVariable.PresArtSistolica
    
    lblNormalFrecuenciaCardiaca.Tag = sighTriajeVariable.FrecCardiaca
    lblNormalFrecuenciaResp.Tag = sighTriajeVariable.FrecRespiratoria
    
'    lblNormalPeso.Tag = sighTriajeVariable.Peso
'    lblNormalTall.Tag = sighTriajeVariable.Talla
'    lblNormalPerimetroCefalico.Tag = sighTriajeVariable.PerimCefalico
End Sub

Public Function Inicializar()
    'Call LimpiarControles
    Call setIdentificadorVariable
    Dim oReglasTriaje As New ReglasTriaje
    Set rsTriajeValiable = oReglasTriaje.ListaVariableTriajeTodos()
    'mgaray20141013
    Set oRsTriajeExcepciones = oReglasTriaje.ListaTriajeExcepcionesTodos()
    Set rsValoresNormalesTriaje = oReglasTriaje.ListarValorNormalesSegunParametros(cargarDatosATriajeValorNormal())
    
    Call bloqueoControlesSegunEdad
    Call mostrarLimitesNormales
    lblAlertaTemperatura.Visible = oReglasTriaje.RetornaTieneFiebre(txtTemperatura.Text, rsValoresNormalesTriaje)
End Function

Public Function cargarDatosATriajeValorNormal() As DOTriajeValorNormal

    Set cargarDatosATriajeValorNormal = oReglasTriaje.RetornaObjetoValorNormalParaBusqueda(mo_DoPaciente, _
                                    mo_DOAtencion, ml_EstadoPaciente)
                                    
    
'    Dim oDOTriajeValorNormal As DOTriajeValorNormal
'    Set oDOTriajeValorNormal = New DOTriajeValorNormal
'    oDOTriajeValorNormal.EdadInicialEnDia = ml_edadPacienteEnDias
'    oDOTriajeValorNormal.SexoPaciente = mo_DoPaciente.idTipoSexo
'    oDOTriajeValorNormal.FechaVigencia = md_fechaTriaje
'    oDOTriajeValorNormal.EstadoPaciente = ml_EstadoPaciente
'
'    Set cargarDatosATriajeValorNormal = oDOTriajeValorNormal
End Function

Public Sub BloqueoTodosLosControles()
    Dim mo_Control As Control
    
    For Each mo_Control In UserControl.Controls
        'If UCase(Left(mo_Control.Name, Len(getAcronimoNombreControlIngresoDatos()))) = UCase(getAcronimoNombreControlIngresoDatos()) And mo_Control.Tag <> "" Then
        If UCase(Left(mo_Control.Name, Len(getAcronimoNombreControlIngresoDatos()))) = UCase(getAcronimoNombreControlIngresoDatos()) Then
            mo_Formulario.HabilitarDeshabilitar mo_Control, False
        End If
    Next
End Sub

Public Sub bloqueoControlesSegunEdad()
    Dim mo_Control As Control
    
    For Each mo_Control In UserControl.Controls
    If Left(mo_Control.Name, 9) = "txtPerime" Then
ms_MensajeError = ""
End If
    
        If UCase(Left(mo_Control.Name, Len(getAcronimoNombreControlIngresoDatos()))) = UCase(getAcronimoNombreControlIngresoDatos()) And mo_Control.Tag <> "" Then
        

           mo_Formulario.HabilitarDeshabilitar mo_Control, False

            If Not (rsTriajeValiable.EOF = True And rsTriajeValiable.BOF = True) Then
                rsTriajeValiable.MoveFirst
                rsTriajeValiable.Find "IdTriajeVariable=" & mo_Control.Tag
                If rsTriajeValiable.BOF = False Then
                    If rsTriajeValiable!TieneLimiteMedicion = True Then
                        If ml_edadPacienteEnDias >= rsTriajeValiable!EdadDiaLimiteMinima _
                                    And ml_edadPacienteEnDias <= rsTriajeValiable!EdadDiaLimiteMaxima Then
                            mo_Formulario.HabilitarDeshabilitar mo_Control, True
                        End If
                    Else
                        mo_Formulario.HabilitarDeshabilitar mo_Control, True
                    End If
                End If
            Else
                ms_MensajeError = "No se Habilita el ingreso de datos por que no hay variables de medición configuradas"
            End If
        End If
    Next
'    Call oReglasTriaje.OcultarControlesCRED(mb_EsAtencionCRED, Label14, txtPerimetroCefalico, lblNormalPerimetroCefalico)
End Sub


Private Function mostrarLimitesNormales() As Boolean
    Dim mo_Control As Control
    Dim sAcronimoLbl As String
    sAcronimoLbl = "lblNormal"
    
    For Each mo_Control In UserControl.Controls
        If UCase(Left(mo_Control.Name, Len(sAcronimoLbl))) = UCase(sAcronimoLbl) _
                And mo_Control.Tag <> "" Then
            'Call muestraValoresNormalesTriaje(mo_Control.Tag)
            mo_Control.Caption = muestraValoresNormalesTriaje(mo_Control.Tag)
        End If
    Next
End Function

Private Function muestraValoresNormalesTriaje(IdTriajeVariable As Long, _
                Optional AddUnidadMedida As Boolean = True, _
                Optional esPresion As Boolean = False) As String
    muestraValoresNormalesTriaje = oReglasTriaje.muestraValoresNormalesTriaje(IdTriajeVariable, _
                                    rsValoresNormalesTriaje, AddUnidadMedida, esPresion)
'    Exit Function
'    Dim sValorNormal As String
'    sValorNormal = ""
'
'    If esPresion = False And (IdTriajeVariable = sighTriajeVariable.PresArtSistolica _
'                        Or IdTriajeVariable = sighTriajeVariable.PresArtDiastolica) Then
'
'        sValorNormal = "Sist/Diast : " & muestraValoresNormalesTriaje(sighTriajeVariable.PresArtSistolica, False, True)
'        sValorNormal = sValorNormal & delimitardorPresion & muestraValoresNormalesTriaje(sighTriajeVariable.PresArtDiastolica, False, True)
'    Else
'        If Not (rsValoresNormalesTriaje.BOF = True And rsValoresNormalesTriaje.EOF = True) Then
'            rsValoresNormalesTriaje.MoveFirst
'            rsValoresNormalesTriaje.Find "IdTriajeVariable=" & IdTriajeVariable
'            If rsValoresNormalesTriaje.EOF = False Then
'                If Not IsNull(rsValoresNormalesTriaje!ValorNormalMinimo) Or Not IsNull(rsValoresNormalesTriaje!ValorNormalMaximo) Then
'                    If Not IsNull(rsValoresNormalesTriaje!ValorNormalMinimo) And _
'                                Not IsNull(rsValoresNormalesTriaje!ValorNormalMaximo) Then
'
'                        sValorNormal = rsValoresNormalesTriaje!ValorNormalMinimo & " a " & _
'                                        rsValoresNormalesTriaje!ValorNormalMaximo
'
'                    ElseIf Not IsNull(rsValoresNormalesTriaje!ValorNormalMinimo) Then
'                        sValorNormal = "min. " & rsValoresNormalesTriaje!ValorNormalMinimo
'                    Else
'                        sValorNormal = "max." & rsValoresNormalesTriaje!ValorNormalMaximo
'                    End If
'                End If
'            End If
'        End If
'
'    End If
'    If AddUnidadMedida = True Then
'        sValorNormal = sValorNormal & unidadMedidaTriaje(IdTriajeVariable)
'    End If
'    muestraValoresNormalesTriaje = sValorNormal
'miError:
'    If Err Then
'        MsgBox Err.Number & " " & Err.Description
'    End If
End Function

Public Function limpiarErroresValidacion() As Boolean
    Dim i As Integer
    If cNumberError.Count > 0 Then
        For i = cNumberError.Count To 1 Step -1
            cNumberError.Remove i
        Next
    End If
    If cLimitesCoherentes.Count > 0 Then
        For i = cLimitesCoherentes.Count To 1 Step -1
            cLimitesCoherentes.Remove i
        Next
    End If
    limpiarErroresValidacion = True
End Function

Public Function validarTodosValoresTriaje() As Boolean
    Dim mo_Control As Control
    Dim i As Integer
    Dim numberError As Integer
    
    validarTodosValoresTriaje = True
    
    limpiarErroresValidacion
    
    For Each mo_Control In UserControl.Controls
        If UCase(Left(mo_Control.Name, Len(getAcronimoNombreControlIngresoDatos()))) = UCase(getAcronimoNombreControlIngresoDatos()) _
                And mo_Control.Tag <> "" And mo_Formulario.ControlEstaHabilitado(mo_Control) Then
                numberError = validarValorNormalTriaje(mo_Control.Tag, mo_Control.Text)
        End If
    Next
    
    If cNumberError.Count > 0 Then
        validarTodosValoresTriaje = False
    End If
End Function

Private Function getAcronimoNombreControlIngresoDatos() As String
    getAcronimoNombreControlIngresoDatos = "txt"
End Function

'Private Function unidadMedidaTriaje(IdTriajeVariable As Long)
'    Select Case IdTriajeVariable
'        Case sighTriajeVariable.Temperatura
'            unidadMedidaTriaje = " °C"
'        Case sighTriajeVariable.Peso
'            unidadMedidaTriaje = " Kg"
'        Case sighTriajeVariable.Talla
'            unidadMedidaTriaje = " Cm"
'        Case Else
'            unidadMedidaTriaje = ""
'    End Select
'End Function

Public Function validarValorNormalTriaje(IdTriajeVariable As sighTriajeVariable, _
                valor As String, Optional esPresion As Boolean = False) As Integer
'On Error GoTo miError
    Dim numberError As Integer
    'magaray20141013
    Dim lb_EsDatoObligatorio As Boolean
'    Dim valor As String
    numberError = 0
    
    valor = Trim(valor)
    
    
    If esPresion = False And (IdTriajeVariable = sighTriajeVariable.PresArtSistolica _
                        Or IdTriajeVariable = sighTriajeVariable.PresArtDiastolica) Then
                        
        Dim sArrayPresion() As String
        
        valor = Replace(valor, "_", "")
        sArrayPresion = Split(valor, delimitardorPresion)
        If UBound(sArrayPresion) = 1 Then
            numberError = validarValorNormalTriaje(PresArtSistolica, sArrayPresion(0), True)
            numberError = validarValorNormalTriaje(PresArtDiastolica, sArrayPresion(1), True)
            validarValorNormalTriaje = numberError
            Exit Function
        End If
    End If
    
    If Not (rsValoresNormalesTriaje.BOF = True And rsValoresNormalesTriaje.EOF = True) Then
        'verificar si el dato es obligatorio
        rsTriajeValiable.MoveFirst
        rsTriajeValiable.Find "IdTriajeVariable=" & IdTriajeVariable
        If rsTriajeValiable.EOF = False Then
            'magaray20141013
            lb_EsDatoObligatorio = BuscarExcepcionesDatoObligatorioSegunEdad(IdTriajeVariable, _
                                    rsTriajeValiable!EsDatoObligatorio, ml_edadPacienteEnDias, _
                                    oRsTriajeExcepciones)
            If lb_EsDatoObligatorio = True Then
                If valor = "" Then
                    numberError = IdTriajeVariable * -1
                End If
            End If
        End If
            
        rsValoresNormalesTriaje.MoveFirst
        rsValoresNormalesTriaje.Find "IdTriajeVariable=" & IdTriajeVariable
        If rsValoresNormalesTriaje.EOF = False Then
            
            If numberError = 0 Then
                If Not IsNull(rsValoresNormalesTriaje!ValorCoherenteMinimo) Or Not IsNull(rsValoresNormalesTriaje!ValorCoherenteMaximo) Then
                    If Not IsNull(rsValoresNormalesTriaje!ValorCoherenteMinimo) And _
                            Not IsNull(rsValoresNormalesTriaje!ValorCoherenteMaximo) Then
                        
                        cLimitesCoherentes.Add rsValoresNormalesTriaje!ValorCoherenteMinimo _
                                                & "-" & rsValoresNormalesTriaje!ValorCoherenteMaximo, Str(IdTriajeVariable)
                        If valor <> "" Then
                            If Not (Val(valor) >= rsValoresNormalesTriaje!ValorCoherenteMinimo _
                                        And Val(valor) <= rsValoresNormalesTriaje!ValorCoherenteMaximo) Then
                                numberError = IdTriajeVariable
                            End If
                        End If
                                        
                    ElseIf Not IsNull(rsValoresNormalesTriaje!ValorCoherenteMinimo) Then
                    
                        cLimitesCoherentes.Add "min. " & rsValoresNormalesTriaje!ValorCoherenteMinimo, Str(IdTriajeVariable)
                        If valor <> "" Then
                            If Val(valor) < rsValoresNormalesTriaje!ValorCoherenteMinimo Then
                                numberError = IdTriajeVariable
                            End If
                        End If
                    Else
                        cLimitesCoherentes.Add "max. " & rsValoresNormalesTriaje!ValorCoherenteMaximo, Str(IdTriajeVariable)
                        If valor <> "" Then
                            If Val(valor) > rsValoresNormalesTriaje!ValorCoherenteMaximo Then
                                numberError = IdTriajeVariable
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    If numberError <> 0 Then
        cNumberError.Add numberError
    End If
    validarValorNormalTriaje = numberError
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description
    End If
End Function

Public Function RetornaMensageDeErrorDatosObligatorios(Optional separadorMensages As String = vbCrLf)
    Dim arrayMessage() As String
    arrayMessage = RetornarArrayMensagesErrorDatosObligatorios
    RetornaMensageDeErrorDatosObligatorios = Join(arrayMessage, separadorMensages)
End Function


Public Function RetornaMensageDeErrorReglas(Optional separadorMensages As String = vbCrLf)
    Dim arrayMessage() As String
    arrayMessage = RetornarArrayMensagesErrorReglas
    RetornaMensageDeErrorReglas = Join(arrayMessage, separadorMensages)
End Function


Public Function RetornaMensageDeError(Optional separadorMensages As String = vbCrLf)
    Dim arrayMessage() As String
    arrayMessage = RetornarArrayMensagesError
    RetornaMensageDeError = Join(arrayMessage, separadorMensages)
End Function

Public Function RetornarArrayMensagesError() As Variant
    RetornarArrayMensagesError = RetornarArrayMensagesErrorPorTipo(0)
End Function

Public Function RetornarArrayMensagesErrorReglas() As Variant
    RetornarArrayMensagesErrorReglas = RetornarArrayMensagesErrorPorTipo(2)
End Function

Public Function RetornarArrayMensagesErrorDatosObligatorios() As Variant
    RetornarArrayMensagesErrorDatosObligatorios = RetornarArrayMensagesErrorPorTipo(1)
End Function

Private Function RetornarArrayMensagesErrorPorTipo(tipoError As Integer) As Variant
    Dim arrayMessage() As String
    Dim numberError As Variant
    Dim bAgregarMensage As Boolean
    
    ReDim Preserve arrayMessage(0)
    For Each numberError In cNumberError
        bAgregarMensage = False
        Select Case tipoError
            Case 1:
                If numberError < 0 Then
                    bAgregarMensage = True
                End If
            Case 2:
                If numberError > 0 Then
                    bAgregarMensage = True
                End If
            Case Else
                bAgregarMensage = True
            End Select
            
            If bAgregarMensage = True Then
                If arrayMessage(UBound(arrayMessage)) <> "" Then
                    ReDim Preserve arrayMessage(UBound(arrayMessage) + 1)
                End If
                arrayMessage(UBound(arrayMessage)) = mapearMensageError(Val(numberError))
            End If
    Next
    RetornarArrayMensagesErrorPorTipo = arrayMessage
End Function

Public Function mapearMensageError(numberError As Integer)

    Dim message As String
    Dim emptyMessage As String, limitMessage As String
    
    emptyMessage = ""
    'emptyMessage = "(vacio)"
    If cNumberError.Count > 0 Then
        If numberError > 0 And cLimitesCoherentes.Count > 0 Then
            limitMessage = "(" & cLimitesCoherentes(Str(numberError)) & ")"
        End If
    
        Select Case numberError
            Case sighTriajeVariable.FrecCardiaca:
                message = Label17.Caption & limitMessage
            Case sighTriajeVariable.FrecRespiratoria:
                message = Label12.Caption & limitMessage
            Case sighTriajeVariable.PerimCefalico:
                message = Label14.Caption & limitMessage
            Case sighTriajeVariable.Peso:
                message = Label3.Caption & limitMessage
            Case sighTriajeVariable.PresArtDiastolica:
                message = "Pres. Art. Diastólica" & limitMessage
            Case sighTriajeVariable.PresArtSistolica:
                message = "Pres. Art. Sistólica" & limitMessage
            Case sighTriajeVariable.Pulso:
                message = Label13.Caption & limitMessage
            Case sighTriajeVariable.Talla:
                message = Label5.Caption & limitMessage
            Case sighTriajeVariable.Temperatura:
                message = Label2.Caption & limitMessage
                'GLCC-Validar Perímetro Cefálico-21/07/2020 Inicio
'             Case sighTriajeVariable.PerimCefalico:
'                message = "Perimetro Cefálicos" & limitMessage
'
'            Case sighTriajeVariable.PresArtDiastolica:
'                message = "Pres. Art. Diastólica" & limitMessage
                'GLCC-Validar Perímetro Cefálico-21/07/2020 Fin
                
            'valores obligatorios vacios
            Case sighTriajeVariable.FrecCardiaca * -1:
                message = Label17.Caption & emptyMessage
            Case sighTriajeVariable.FrecRespiratoria * -1:
                message = Label12.Caption & emptyMessage
            Case sighTriajeVariable.PerimCefalico * -1:
                message = Label14.Caption & emptyMessage
            Case sighTriajeVariable.Peso * -1:
                message = Label3.Caption & emptyMessage
            Case sighTriajeVariable.PresArtDiastolica * -1:
                message = "Pres. Art. Diastólica" & emptyMessage
            Case sighTriajeVariable.PresArtSistolica * -1:
                message = "Pres. Art. Sistólica" & emptyMessage
            Case sighTriajeVariable.Pulso * -1:
                message = Label13.Caption & emptyMessage
            Case sighTriajeVariable.Talla * -1:
                message = Label5.Caption & emptyMessage
            Case sighTriajeVariable.Temperatura * -1:
                message = Label2.Caption & emptyMessage
            Case Else
                message = "Número de error no encontrado"
        End Select
    End If
    mapearMensageError = message
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description
    End If
End Function

Public Function CargaDatosAlObjetosDeDatos(ByRef mo_DOAtencionesCE As DOAtencionesCE) As DOAtencionesCE
    Dim lTriajeOrigen As Long
    lTriajeOrigen = ml_triajeOrigen
    'se puede haber llamado desde un formulario distinto al que registro el triaje
    If Not (mo_DOAtencionCEAnterior Is Nothing) Then
        If mo_DOAtencionCEAnterior.TriajeOrigen <> 0 Then
            lTriajeOrigen = mo_DOAtencionCEAnterior.TriajeOrigen
        End If
    End If
    With mo_DOAtencionesCE
        '.idAtencion = ml_idAtencion
        .TriajePulso = Val(txtPulso.Text)
        .TriajeTemperatura = txtTemperatura.Text
        .TriajePresion = Trim(txtPresionSist.Text) & delimitardorPresion & Trim(txtPresionDiast.Text)
        .TriajeFrecCardiaca = Val(txtFrecuenciaCardiaca.Text)
        
        .TriajeFrecRespiratoria = Val(txtFrespiratoria.Text)
        .TriajePeso = txtPeso.Text
        .TriajeTalla = txtTalla.Text
'        .TriajePerimCefalico = Val(txtPerimetroCefalico.Text)
        .TriajeOrigen = lTriajeOrigen
        .TriajePerimAbdominal = UserControl.txtPermAbdominal.Text
        .TriajeSaturacionOxigeno = UserControl.txtSaturacionOxigeno.Text
   End With
   Set CargaDatosAlObjetosDeDatos = mo_DOAtencionesCE
End Function

Public Function CargarDatosALosControles(ByVal mo_DOAtencionesCE As DOAtencionesCE)
    Dim arrayPresion() As String
    
    With mo_DOAtencionesCE
        txtPulso.Text = IIf(.TriajePulso = 0, "", .TriajePulso)
        txtTemperatura.Text = .TriajeTemperatura
        If .TriajePresion <> "" Then
            arrayPresion = Split(.TriajePresion, delimitardorPresion)
            If UBound(arrayPresion) = 1 Then
                txtPresionSist.Text = arrayPresion(0)
                txtPresionDiast.Text = arrayPresion(1)
            End If
        End If
        txtFrecuenciaCardiaca.Text = IIf(.TriajeFrecCardiaca = 0, "", .TriajeFrecCardiaca)
        
        txtFrespiratoria.Text = IIf(.TriajeFrecRespiratoria = 0, "", .TriajeFrecRespiratoria)
        txtPeso.Text = .TriajePeso
        txtTalla.Text = .TriajeTalla
'        txtPerimetroCefalico.Text = .TriajePerimCefalico
 '       txtPerimetroCefalico.Text = IIf(.TriajePerimCefalico = 0, "", .TriajePerimCefalico)
        UserControl.txtPermAbdominal.Text = .TriajePerimAbdominal
        UserControl.txtSaturacionOxigeno.Text = .TriajeSaturacionOxigeno
    End With
    CalculaIMC
    Set mo_DOAtencionCEAnterior = mo_DOAtencionesCE
End Function

Private Function delimitardorPresion()
    delimitardorPresion = "/"
End Function

Function PresionVerificaSiTieneDatosYsiEstaOK() As Boolean
    PresionVerificaSiTieneDatosYsiEstaOK = True
    If txtPresionSist.Text <> "" Or txtPresionDiast.Text <> "" Then
        If txtPresionSist.Text = "" Or txtPresionDiast.Text = "" Then
            MsgBox "Debe Ingresar valores para Presión Sistólica y Diastólica", vbInformation, "Mensaje"
            PresionVerificaSiTieneDatosYsiEstaOK = False
        Else
            If Val(txtPresionSist.Text) <= Val(txtPresionDiast.Text) Then
               MsgBox "Si registra la PRESION: SISTOLICA/DIASTOLICA, el valor SISTOLICA debe ser mayor a la DIASTOLICA", vbInformation, "mensaje"
               PresionVerificaSiTieneDatosYsiEstaOK = False
            End If
        End If
    End If
End Function


Private Function buscarValorNormalVariable(IdTriajeVariable As sighTriajeVariable) As Boolean
    buscarValorNormalVariable = oReglasTriaje.buscarValorNormalVariable(IdTriajeVariable, rsValoresNormalesTriaje)
'    buscarValorNormalVariable = False
'    If Not (rsValoresNormalesTriaje Is Nothing) Then
'        If Not (rsValoresNormalesTriaje.BOF = True And rsValoresNormalesTriaje.EOF = True) Then
'            rsValoresNormalesTriaje.MoveFirst
'            rsValoresNormalesTriaje.Find "IdTriajeVariable=" & IdTriajeVariable
'            If rsValoresNormalesTriaje.EOF = False Then
'                buscarValorNormalVariable = True
'            End If
'        End If
'    End If
End Function

Public Function setValorTalla(sTalla As String, sPeso As String)
    If txtTalla.Text = "" Then txtTalla.Text = sTalla
    lnTallaDeUltimaAtencion = Val(sTalla)
    If txtPeso.Text = "" Then txtPeso.Text = sPeso
End Function
'mgaray20141013
Private Function BuscarExcepcionesDatoObligatorioSegunEdad(lIdTriajeVariable As Long, _
        bEsDatoObligatorio As Boolean, lEdadPacienteEnDias As Long, _
        oRsTriajeExcepciones As ADODB.Recordset) As Boolean
On Error GoTo miError
    Dim lb_EsDatoObigatorio As Boolean
    
    lb_EsDatoObigatorio = bEsDatoObligatorio
    If Not (oRsTriajeExcepciones Is Nothing) Then
'        oRsTriajeExcepciones.Filter
        oRsTriajeExcepciones.Filter = "IdTriajeVariable=" & lIdTriajeVariable
        If oRsTriajeExcepciones.RecordCount > 0 Then
            oRsTriajeExcepciones.MoveFirst
            While oRsTriajeExcepciones.EOF = False
                If lEdadPacienteEnDias >= oRsTriajeExcepciones.Fields!EdadInicialEnDia _
                        And lEdadPacienteEnDias <= oRsTriajeExcepciones.Fields!EdadFinalEnDia Then
                    lb_EsDatoObigatorio = oRsTriajeExcepciones.Fields!EsDatoObligatorio
                    oRsTriajeExcepciones.MoveLast
                End If
                oRsTriajeExcepciones.MoveNext
            Wend
        End If
    End If
    BuscarExcepcionesDatoObligatorioSegunEdad = lb_EsDatoObigatorio
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation, "Triaje"
    End If
End Function



'debb-09/06/2016
Private Sub txtFrecuenciaCardiaca_LostFocus()
    If ml_triajeOrigen = sightriajeorigen.Triaje And Val(txtFrecuenciaCardiaca.Text) > 0 Then
        txtPulso.Text = txtFrecuenciaCardiaca.Text
    End If
End Sub
'debb-09/06/2016
Private Sub txtPulso_LostFocus()
    If ml_triajeOrigen = sightriajeorigen.Triaje And Val(txtPulso.Text) > 0 Then
       txtFrecuenciaCardiaca.Text = txtPulso.Text
    End If
End Sub

'debb-10/08/2016
Public Function ValidarReglas() As Boolean
       ValidarReglas = True
       If Val(txtTalla.Text) > 0 Then
       If CCur(txtTalla.Text) < 20 Then
          MsgBox "El Paciente debe tener una TALLA mayor a 20 cm", vbInformation, "Reglas"
          ValidarReglas = False
       End If
       'GLCC 21/07/20 CAMBIO42 INICIO
       ' If Val(txtTalla.Text) = " " Then
       ' MsgBox "El Paciente debe tener una TALLA", vbInformation, "Reglas"
         'GLCC 21/07/20 CAMBIO42 FIN
      ' End If
       End If
End Function
Private Sub txtPeso_LostFocus()
    If Val(txtPeso.Text) > 300 Then
       MsgBox "El PESO no puede pasar de 300 Kg", vbInformation, "TRIAJE"
       txtPeso.Text = " "
    End If
    CalculaIMC
    'GLCC 21/07/20 CAMBIO42 INICIO
'        If (txtPeso.Text) = " " Then
'        MsgBox "El Paciente debe tener un PESO", vbInformation, "Reglas"
'       End If
        'GLCC 21/07/20 CAMBIO42 FIN
End Sub

Sub CalculaIMC()
        'debb-29/03/2017
        lblIMC.Caption = ""
        If Val(txtPeso.Text) > 0 And Val(txtTalla.Text) > 20 Then
           lblIMC.Caption = "IMC: " & Trim(Str(Round(CStr(txtPeso.Text) / (CStr(txtTalla.Text) * CStr(txtTalla.Text) * 0.0001), 0)))
        End If
End Sub


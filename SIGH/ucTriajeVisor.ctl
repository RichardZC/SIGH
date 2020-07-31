VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ucTriajeVisor 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ScaleHeight     =   3285
   ScaleWidth      =   3105
   Begin VB.VScrollBar vsTriaje 
      Height          =   495
      Left            =   11880
      TabIndex        =   0
      Top             =   0
      Value           =   20
      Width           =   255
   End
   Begin VB.HScrollBar hsTriaje 
      Height          =   255
      Left            =   2550
      TabIndex        =   1
      Top             =   690
      Width           =   555
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   3240
      Left            =   0
      ScaleHeight     =   3240
      ScaleWidth      =   3030
      TabIndex        =   3
      Top             =   0
      Width           =   3030
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
         Left            =   2025
         MaxLength       =   30
         TabIndex        =   31
         Top             =   2925
         Width           =   480
      End
      Begin VB.CommandButton btnBuscaHistoricos 
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
         Left            =   2655
         Picture         =   "ucTriajeVisor.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1110
         Width           =   315
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
         Left            =   1605
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2205
         Width           =   495
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
         Left            =   1620
         MaxLength       =   30
         TabIndex        =   9
         Top             =   2565
         Width           =   495
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
         Left            =   630
         TabIndex        =   8
         Top             =   360
         Width           =   600
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
         Left            =   630
         TabIndex        =   7
         Top             =   765
         Width           =   600
      End
      Begin VB.TextBox txtTalla 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   6
         Top             =   1095
         Width           =   600
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
         Left            =   630
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   600
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
         Left            =   1605
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1845
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtPresion 
         Height          =   315
         Left            =   630
         TabIndex        =   12
         Top             =   0
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saturación de oxígeno"
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
         TabIndex        =   33
         Top             =   2940
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0-100"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2565
         TabIndex        =   32
         Top             =   2940
         Width           =   345
      End
      Begin VB.Label lblIMC 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1740
         TabIndex        =   30
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Frec. Cardiaca"
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
         TabIndex        =   29
         Top             =   2220
         Width           =   1125
      End
      Begin VB.Label lblNormalFrecuenciaCardiaca 
         AutoSize        =   -1  'True
         Caption         =   "10 a 20"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2190
         TabIndex        =   28
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label lblNormalPerimetroCefalico 
         AutoSize        =   -1  'True
         Caption         =   "cm."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2205
         TabIndex        =   27
         Top             =   2580
         Width           =   255
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
         Left            =   60
         TabIndex        =   26
         Top             =   2580
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Presión"
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
         TabIndex        =   25
         Top             =   45
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temp"
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
         TabIndex        =   24
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label9 
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
         Left            =   60
         TabIndex        =   23
         Top             =   765
         Width           =   390
      End
      Begin VB.Label Label10 
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
         Left            =   60
         TabIndex        =   22
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kg."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1260
         TabIndex        =   21
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "cm."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1260
         TabIndex        =   20
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "° C"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1260
         TabIndex        =   19
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Sist/Diast: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1740
         TabIndex        =   18
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "600-250"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1260
         TabIndex        =   17
         Top             =   1455
         Width           =   495
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "0-70"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2190
         TabIndex        =   16
         Top             =   1845
         Width           =   270
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Frec.Respiratoria"
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
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label Label31 
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
         Left            =   60
         TabIndex        =   14
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label lblAlertaTemperatura 
         Caption         =   "(..FIEBRE..)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1710
         TabIndex        =   13
         Top             =   405
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.PictureBox picEsquina 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10755
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   495
      Width           =   255
   End
End
Attribute VB_Name = "ucTriajeVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mostrar Signos Vitales
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_DOAtencionesCE As DOAtencionesCE
Dim mo_DOAtencionesCENew As DOAtencionesCE
Dim mo_DoAtencion As New DOAtencion
Dim mo_DoPaciente As New doPaciente
Dim ml_idAtencion As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdPaciente As Long
Dim md_FechaAtencion As Date
Dim ml_EstadoPaciente As Long
Dim ml_Origen As Long
Dim mi_Opcion As sghOpciones
Dim mi_OpcionFormulario As sghOpciones
Dim ml_edadPacienteEnDias As Long
Dim mb_EsAtencionCRED As Boolean

Dim mo_Formulario As New Formulario
Dim lcBuscaParametro As New Parametros

Dim rsTriajeValiable As New ADODB.Recordset
Dim rsValoresNormalesTriaje  As New ADODB.Recordset

Public Event changeDataControl(mo_DOAtencionesCE As DOAtencionesCE, _
    mo_DOAtencionesCENew As DOAtencionesCE)
    
'=================================================================
'=================================================================

Property Let Origen(lValue As sightriajeorigen)
   ml_Origen = lValue
End Property

Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property

Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property

Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property

Property Get idPaciente() As Long
   idPaciente = ml_IdPaciente
End Property

Property Let EstadoPaciente(lValue As Long)
   ml_EstadoPaciente = lValue
End Property

Property Get EstadoPaciente() As Long
   EstadoPaciente = ml_EstadoPaciente
End Property

Property Set DOAtencionCE(lValue As DOAtencionesCE)
    Set mo_DOAtencionesCE = lValue
End Property


Property Set DOAtencion(oValue As DOAtencion)
    Set mo_DoAtencion = oValue
End Property



Property Get DOAtencionCE() As DOAtencionesCE
   Set DOAtencionCE = mo_DOAtencionesCE
End Property

Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property

Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property

Property Let FechaAtencion(lValue As Date)
   md_FechaAtencion = lValue
End Property

Property Get FechaAtencion() As Date
   FechaAtencion = md_FechaAtencion
End Property

Property Let EsAtencionCRED(bValue As Boolean)
   mb_EsAtencionCRED = bValue
End Property

Property Let OpcionFormulario(bValue As sghOpciones)
   mi_OpcionFormulario = bValue
End Property

'Public Function cargarDatosATriajeValorNormal() As DOTriajeValorNormal
'    Set cargarDatosATriajeValorNormal = oReglasTriaje.RetornaObjetoValorNormalParaBusqueda(mo_DoPaciente, mo_DOAtencionesCE, ml_EstadoPaciente)
'    Dim oDOTriajeValorNormal As DOTriajeValorNormal
'    Set oDOTriajeValorNormal = New DOTriajeValorNormal
'    oDOTriajeValorNormal.EdadInicialEnDia = ml_edadPacienteEnDias
'    oDOTriajeValorNormal.SexoPaciente = mo_DoPaciente.idTipoSexo
'    oDOTriajeValorNormal.FechaVigencia = md_FechaAtencion
'    oDOTriajeValorNormal.EstadoPaciente = ml_EstadoPaciente
'
'
'    Set cargarDatosATriajeValorNormal = oDOTriajeValorNormal
'End Function

'antes de llamar a este metodo debe haber seteado FechaAtencion, IdPaciente, EstadoPacinte y Origen
Public Function Inicializar()
    Dim oReglasTriaje As New ReglasTriaje
    Set mo_DoPaciente = RetornaPaciente(ml_IdPaciente)
    
    Set rsTriajeValiable = oReglasTriaje.ListaVariableTriajeTodos()
    Set rsValoresNormalesTriaje = oReglasTriaje.ListarValorNormalesSegunParametros( _
                                oReglasTriaje.RetornaObjetoValorNormalParaBusqueda(mo_DoPaciente, _
                                mo_DoAtencion, ml_EstadoPaciente))
    
    Label22.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.PresArtDiastolica, rsValoresNormalesTriaje, True)
    Label21.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.Temperatura, rsValoresNormalesTriaje, True)
    Label28.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.Pulso, rsValoresNormalesTriaje, True)

    Label19.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.Peso, rsValoresNormalesTriaje, True)
    Label20.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.Talla, rsValoresNormalesTriaje, True)
    Label29.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.FrecRespiratoria, rsValoresNormalesTriaje, True)
    lblNormalFrecuenciaCardiaca.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.FrecCardiaca, rsValoresNormalesTriaje, True)
    lblNormalPerimetroCefalico.Caption = oReglasTriaje.muestraValoresNormalesTriaje(sighTriajeVariable.PerimCefalico, rsValoresNormalesTriaje, True)
    
    picContainer.Top = 0
    leftTopWidthOfScroll
    mostrarOcultarScroll
    configuracionScrolls
    Call bloqueoControles
    Call oReglasTriaje.OcultarControlesCRED(mb_EsAtencionCRED, Label14, txtPerimetroCefalico, lblNormalPerimetroCefalico)
End Function

Private Sub bloqueoControles()
    mo_Formulario.HabilitarDeshabilitar txtPeso, False
    mo_Formulario.HabilitarDeshabilitar txtPresion, False
    mo_Formulario.HabilitarDeshabilitar txtTalla, False
    mo_Formulario.HabilitarDeshabilitar txtTemperatura, False
    mo_Formulario.HabilitarDeshabilitar txtFrespiratoria, False
    mo_Formulario.HabilitarDeshabilitar txtPulso, False
    mo_Formulario.HabilitarDeshabilitar txtFrecuenciaCardiaca, False
    mo_Formulario.HabilitarDeshabilitar txtPerimetroCefalico, False
    mo_Formulario.HabilitarDeshabilitar txtSaturacionOxigeno, False
End Sub

'necesario idPaciente, fechaAtencion
Public Function AsignarIdAtencionYLlenarControles(mlIdAtencion As Long)
On Error GoTo ErrJamo
    ml_idAtencion = mlIdAtencion
    
    Dim oReglasAdmision As New ReglasAdmision
    Dim oAtencionesCE As New AtencionesCE
    Dim oConexionExterna As New Connection
    Dim oConexion As New Connection
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    
    
    Set mo_DoAtencion = oReglasAdmision.AtencionesSeleccionarPorId(mlIdAtencion, oConexion)
    
    'datos paciente
    ml_IdPaciente = mo_DoAtencion.idPaciente
    md_FechaAtencion = mo_DoAtencion.FechaIngreso
    ml_idCuentaAtencion = mo_DoAtencion.idCuentaAtencion
    
    'buscar datos de consulta externa
    If mo_DOAtencionesCE Is Nothing Then
        Set mo_DOAtencionesCE = New DOAtencionesCE
    End If
    mo_DOAtencionesCE.idAtencion = ml_idAtencion
    Set oAtencionesCE.Conexion = oConexionExterna
    If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
        Set mo_DOAtencionesCE = Nothing
    End If
    Call Inicializar
    Call CargarDatosAlosControles(mo_DOAtencionesCE)
    
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Exit Function
ErrJamo:
'    Resume
End Function

Public Function CargarDatosAlosControles(moDOAtencionesCE As DOAtencionesCE)
    If Not (moDOAtencionesCE Is Nothing) Then
        txtPresion.Text = moDOAtencionesCE.TriajePresion
        txtTemperatura.Text = moDOAtencionesCE.TriajeTemperatura
        txtPeso.Text = moDOAtencionesCE.triajePeso
        txtTalla.Text = moDOAtencionesCE.triajeTalla
        txtPulso.Text = IIf(moDOAtencionesCE.TriajePulso = 0, "", moDOAtencionesCE.TriajePulso)
        txtFrespiratoria.Text = IIf(moDOAtencionesCE.TriajeFrecRespiratoria = 0, "", moDOAtencionesCE.TriajeFrecRespiratoria)
        txtFrecuenciaCardiaca.Text = IIf(moDOAtencionesCE.TriajeFrecCardiaca = 0, "", moDOAtencionesCE.TriajeFrecCardiaca)
        txtPerimetroCefalico.Text = IIf(moDOAtencionesCE.TriajePerimCefalico = 0, "", moDOAtencionesCE.TriajePerimCefalico)
        txtSaturacionOxigeno.Text = moDOAtencionesCE.TriajeSaturacionOxigeno
        
        mostrarAlertas
        
        Dim lcBuscaParametro As New SIGHDatos.Parametros
                
        mi_Opcion = sghModificar
        
'        If moDOAtencionesCE.TriajeOrigen <> ml_Origen And moDOAtencionesCE.TriajeOrigen <> 0 Then
'            mi_Opcion = sghConsultar
'        ElseIf md_FechaAtencion <> lcBuscaParametro.RetornaFechaServidorSQL And moDOAtencionesCE.TriajeOrigen <> 0 Then
'            mi_Opcion = sghConsultar
'        End If
        If md_FechaAtencion <> lcBuscaParametro.RetornaFechaServidorSQL _
                    And ml_Origen = sightriajeorigen.ConsultaExterna Then
            mi_Opcion = sghConsultar
        End If
        'debb-29/03/2017
        lblIMC.Caption = ""
        If Val(txtPeso.Text) > 0 And Val(txtTalla.Text) > 20 Then
           lblIMC.Caption = "IMC: " & Trim(Str(Round(CStr(txtPeso.Text) / (CStr(txtTalla.Text) * CStr(txtTalla.Text) * 0.0001), 0)))
        End If
    Else
        LimpiarControles
        mi_Opcion = sghAgregar
    End If
    If mi_OpcionFormulario = sghConsultar Then
        mi_Opcion = sghConsultar
    End If
    Set mo_DOAtencionesCE = moDOAtencionesCE
End Function

Private Function mostrarAlertas() As Boolean
    Dim oReglasTriaje As New ReglasTriaje
        
    lblAlertaTemperatura.Visible = oReglasTriaje.RetornaTieneFiebre(txtTemperatura.Text, rsValoresNormalesTriaje)
    
    mostrarAlertas = True
End Function

Private Function RetornaPaciente(mlIdPaciente As Long) As doPaciente
'    Dim moDoPaciente As New DOPaciente
    Dim oReglasAdmision As New ReglasAdmision
    Dim oFechaHOra As New FechaHora
    
    Set mo_DoPaciente = oReglasAdmision.RetornaPacientesSeleccionarPorId(mlIdPaciente)
    If Not (mo_DoPaciente Is Nothing) Then
        ml_edadPacienteEnDias = oFechaHOra.EdadActualEnDias(mo_DoPaciente.FechaNacimiento, md_FechaAtencion)
    End If
    Set RetornaPaciente = mo_DoPaciente
End Function

Private Function LimpiarControles()
    txtPresion.Text = ""
    txtTemperatura.Text = ""
    txtPeso.Text = ""
    txtTalla.Text = ""
    txtPulso.Text = ""
    txtFrespiratoria.Text = ""
    txtFrecuenciaCardiaca.Text = ""
    txtPerimetroCefalico.Text = ""
    lblIMC.Caption = ""
End Function

Private Sub btnBuscaHistoricos_Click()
    Dim oTriaje As New clTriaje
    oTriaje.idAtencion = ml_idAtencion
    oTriaje.idUsuario = sighEntidades.Usuario
    oTriaje.Opcion = mi_Opcion
    oTriaje.lcNombrePc = sighEntidades.RetornaNombrePC
    Call oTriaje.MostrarFormularioDesdeAtenciones(ml_Origen, ml_idCuentaAtencion)
    Call verificarCambiosEnTriaje
End Sub

Private Function verificarCambiosEnTriaje() As Boolean
    Dim oReglasAdmision As New ReglasAdmision
    Dim oDOAtencionesCE As DOAtencionesCE
    Dim SeRealizoCambios As Boolean
    
    SeRealizoCambios = True
    
    
    Set oDOAtencionesCE = oReglasAdmision.AtencionCESeleccionarPorId(ml_idAtencion)
    
    If oDOAtencionesCE Is Nothing And mo_DOAtencionesCE Is Nothing Then
        SeRealizoCambios = False
    ElseIf mo_DOAtencionesCE Is Nothing Then
        SeRealizoCambios = True
    Else
        SeRealizoCambios = objetosSonIguales(mo_DOAtencionesCE, oDOAtencionesCE)
    End If
    If SeRealizoCambios = True Then
        Call CambiarDatosControl(mo_DOAtencionesCE, oDOAtencionesCE)
    End If
End Function

Private Function CambiarDatosControl(oDOAtencionesCE As DOAtencionesCE, oDOAtencionesCENew As DOAtencionesCE) As Boolean
    Dim oDoAtencionesCEActual As DOAtencionesCE
    Set oDoAtencionesCEActual = oDOAtencionesCE
    Call CargarDatosAlosControles(oDOAtencionesCENew)
    RaiseEvent changeDataControl(oDoAtencionesCEActual, oDOAtencionesCENew)
End Function


Private Function objetosSonIguales(oDOAtencionesCE As DOAtencionesCE, _
            oDOAtencionesCENew As DOAtencionesCE) As Boolean
    Dim bSonIguales As Boolean
    
    bSonIguales = False
    
    
    If oDOAtencionesCE.TriajeFrecCardiaca <> oDOAtencionesCENew.TriajeFrecCardiaca Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajeFrecRespiratoria <> oDOAtencionesCENew.TriajeFrecRespiratoria Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajePerimCefalico <> oDOAtencionesCENew.TriajePerimCefalico Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.triajePeso <> oDOAtencionesCENew.triajePeso Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajePresion <> oDOAtencionesCENew.TriajePresion Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajePulso <> oDOAtencionesCENew.TriajePulso Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.triajeTalla <> oDOAtencionesCENew.triajeTalla Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajeTemperatura <> oDOAtencionesCENew.TriajeTemperatura Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajeSaturacionOxigeno <> oDOAtencionesCENew.TriajeSaturacionOxigeno Then
        bSonIguales = True
        GoTo Final
    End If
    If oDOAtencionesCE.TriajePerimAbdominal <> oDOAtencionesCENew.TriajePerimAbdominal Then
        bSonIguales = True
        GoTo Final
    End If
    
Final:
    objetosSonIguales = bSonIguales
End Function


'=================================================================
'comportamiento de barras de desplazamiento
'=================================================================

Private Sub hsTriaje_Change()
     picContainer.Left = -hsTriaje.Value
End Sub

Private Sub leftTopWidthOfScroll()
    'horizontal
    hsTriaje.Top = UserControl.Height - hsTriaje.Height
    hsTriaje.Left = 0
    hsTriaje.Width = UserControl.Width - vsTriaje.Width
    
    'vertical
    vsTriaje.Top = 0
    vsTriaje.Left = UserControl.Width - vsTriaje.Width
    vsTriaje.Height = UserControl.Height - hsTriaje.Height
    
    'esquina de los scroll
    picEsquina.Height = UserControl.Height - vsTriaje.Height
    picEsquina.Width = UserControl.Width - hsTriaje.Width
    picEsquina.Left = hsTriaje.Left + hsTriaje.Width
    picEsquina.Top = vsTriaje.Top + vsTriaje.Height
End Sub

Private Sub mostrarOcultarScroll()
    'si el tamaño del la instancia supera al control entonces se ocultan los scrolls
    If UserControl.Width >= picContainer.Width And UserControl.Height >= picContainer.Height Then
        hsTriaje.Visible = False
        vsTriaje.Visible = False
        picEsquina.Visible = False
        'si el tamaño de la instancia es menor que el control en alto y ancho se muestran los scrolls
    ElseIf UserControl.Width < picContainer.Width And UserControl.Height < picContainer.Height Then
        hsTriaje.Visible = True
        vsTriaje.Visible = True
        picEsquina.Visible = True
        'ancho de la instancia supera al control y el alto no entonces ocultar barra horizontal y mostrar barra vertical
    ElseIf UserControl.Width >= picContainer.Width And UserControl.Height < picContainer.Height Then
        vsTriaje.Visible = True
        hsTriaje.Visible = False
        vsTriaje.Height = vsTriaje.Height + picEsquina.Height
        picEsquina.Visible = False
    'ancho de la instancia no supera al control y el alto si coultar scroll vertical y mostrar horizontal
    ElseIf UserControl.Width < picContainer.Width And UserControl.Height >= picContainer.Height Then
        vsTriaje.Visible = False
        hsTriaje.Visible = True
        hsTriaje.Width = hsTriaje.Width + picEsquina.Width
        picEsquina.Visible = False
    End If
End Sub

Private Sub configuracionScrolls()
    With vsTriaje
        .Min = 0
        .Max = picContainer.Height - hsTriaje.Top
        .SmallChange = Abs(.Max / 16) + 1
        .LargeChange = Abs(.Max / 4) + 1
        .ZOrder 0
        .Value = 0
    End With
    
    With hsTriaje
        .Min = 0
        .Max = picContainer.Width - vsTriaje.Left
        .SmallChange = Abs(.Max / 16) + 1
        .LargeChange = Abs(.Max / 4) + 1
        .ZOrder 0
        .Value = 0
    End With
End Sub


Private Sub vsTriaje_Change()
    picContainer.Top = -vsTriaje.Value
End Sub



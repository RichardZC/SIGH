VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.UserControl ucHCelectronicaLista 
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12585
   ScaleHeight     =   9465
   ScaleWidth      =   12585
   Begin TabDlg.SSTab SSTab 
      Height          =   3195
      Left            =   0
      TabIndex        =   36
      Top             =   6165
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   5636
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "CPT/FARMACIA"
      TabPicture(0)   =   "ucHCelectronicaLista.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PDF/JPG generados"
      TabPicture(1)   =   "ucHCelectronicaLista.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ucPacientesPDF1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Procedimientos de Apoyo al Diagnóstico realizados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   15
         TabIndex        =   37
         Top             =   315
         Width           =   12525
         Begin UltraGrid.SSUltraGrid grdApoyoDx 
            Height          =   2535
            Left            =   30
            TabIndex        =   38
            Top             =   210
            Width           =   12360
            _ExtentX        =   21802
            _ExtentY        =   4471
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "grdApoyoDx"
         End
         Begin VB.Label Label12 
            Caption         =   "<Enter> ó <Doble Clic> = Detalle del RESULTADO de LABORATORIO, los que tengan la palabra SI en columna RESULTADOS"
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
            Height          =   255
            Left            =   75
            TabIndex        =   39
            Top             =   2850
            Width           =   11625
         End
      End
      Begin SISGalenPlus.ucPacientesPDF ucPacientesPDF1 
         Height          =   2595
         Left            =   -74775
         TabIndex        =   40
         Top             =   525
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   4577
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Historia individual electrónica resumida del Establecimiento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   30
      TabIndex        =   30
      Top             =   3465
      Width           =   12525
      Begin UltraGrid.SSUltraGrid grdAtenciones 
         Height          =   2100
         Left            =   90
         TabIndex        =   31
         Top             =   270
         Width           =   12360
         _ExtentX        =   21802
         _ExtentY        =   3704
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdAtenciones"
      End
      Begin VB.Label Label11 
         Caption         =   "<ENTER> ó <Doble Clic> = Detalle de la Cuenta"
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
         Height          =   255
         Left            =   105
         TabIndex        =   32
         Top             =   2400
         Width           =   6105
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   30
      TabIndex        =   2
      Top             =   570
      Width           =   6615
      Begin VB.CommandButton cmdSinApellidoMaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3045
         Picture         =   "ucHCelectronicaLista.ctx":0038
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1095
         Width           =   345
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1305
         Picture         =   "ucHCelectronicaLista.ctx":05C2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1110
         Width           =   345
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1110
         Width           =   1185
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1110
         Width           =   1200
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   1830
         MaxLength       =   9
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtDNI 
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
         MaxLength       =   8
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5220
         Picture         =   "ucHCelectronicaLista.ctx":0B4C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   5220
         Picture         =   "ucHCelectronicaLista.ctx":3795
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1305
      End
      Begin UltraGrid.SSUltraGrid grdPacientes 
         Height          =   1320
         Left            =   90
         TabIndex        =   7
         Top             =   1500
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   2328
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de pacientes"
      End
      Begin VB.Label Label50 
         Caption         =   "Ap.Paterno              Ap.Materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   3105
      End
      Begin VB.Label Label6 
         Caption         =   "DNI                         N° Historia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6705
      TabIndex        =   1
      Top             =   570
      Width           =   5865
      Begin VB.CommandButton bntReporte 
         Height          =   390
         Left            =   150
         Picture         =   "ucHCelectronicaLista.ctx":3DBE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Imprimir LISTA"
         Top             =   2445
         Width           =   525
      End
      Begin VB.TextBox txtGrupoS 
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
         TabIndex        =   29
         Top             =   1860
         Width           =   1680
      End
      Begin VB.TextBox txtFactorRH 
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
         Left            =   3300
         TabIndex        =   27
         Top             =   1260
         Width           =   900
      End
      Begin VB.TextBox txtPadres 
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
         Left            =   960
         TabIndex        =   25
         Top             =   1245
         Width           =   1290
      End
      Begin VB.TextBox txtAlergias 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   960
         TabIndex        =   22
         Top             =   1575
         Width           =   1305
      End
      Begin VB.TextBox txtDomicilio 
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
         Left            =   960
         TabIndex        =   21
         Top             =   900
         Width           =   4845
      End
      Begin VB.TextBox txtSexo 
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
         Left            =   960
         TabIndex        =   20
         Top             =   1890
         Width           =   1320
      End
      Begin VB.TextBox txtEdad 
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
         Left            =   960
         TabIndex        =   19
         Top             =   570
         Width           =   1770
      End
      Begin VB.TextBox txtPaciente 
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
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   4845
      End
      Begin VB.Image pi_ImagSeleccionada 
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   4230
         MouseIcon       =   "ucHCelectronicaLista.ctx":4297
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   1245
         Width           =   1545
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO SANGUINEO"
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
         Left            =   2535
         TabIndex        =   28
         Top             =   1620
         Width           =   1650
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FACTOR RH"
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
         TabIndex        =   26
         Top             =   1290
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N.Padres"
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
         TabIndex        =   24
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Alergias"
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
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
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
         TabIndex        =   16
         Top             =   1635
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
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
         TabIndex        =   15
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   150
         TabIndex        =   14
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Edad"
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
         TabIndex        =   13
         Top             =   615
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paciente"
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
         TabIndex        =   12
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Historia Clínica electrónica individual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12525
   End
End
Attribute VB_Name = "ucHCelectronicaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar atenciones Historicas del Paciente
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_idTipoSexo As Long
Dim lbFrame3Visible As Boolean
Private Const PM_POSITIONCTRL = 1
Private Const PM_MOVEPREVCELL = 2
Private Const PM_MOVENEXTCELL = 3
Private Const PM_EXITEDITMODE = 4
Private Const PM_PROCESSKEY = 5


Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdPacientes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdPacientes.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoFiltro(lValue As sghTipoFiltroPacientes)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroPacientes
    TipoFiltro = ml_TipoFiltro
End Property




Public Sub RealizarBusqueda()
Dim oPaciente As New doPaciente
        
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
           oPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text))
           
        End If
        oPaciente.nrodocumento = txtDNI.Text
        oPaciente.IdDocIdentidad = 1
        oPaciente.FichaFamiliar = ""
        Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrarConHistoriasDefinitivas(oPaciente, wxSinApellido)
        
        
        Dim rsRespuesta As New Recordset
        Set rsRespuesta = grdPacientes.DataSource
        On Error Resume Next
        If rsRespuesta.RecordCount = 0 Then
            MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
        End If
        
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbInformation, "Filtro Pacientes"
        End If
        
        
        If rsRespuesta.RecordCount = 1 Then
           
           grdPacientes_DblClick
        End If
End Sub



Private Sub bntReporte_Click()
    On Error GoTo errores
    UserControl.MousePointer = 11
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Set oRsTmp1 = grdAtenciones.DataSource
    mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp1, Frame2.Caption, txtPaciente.Text, "", 1, False, True
    If grdApoyoDx.Visible = True Then
        Set oRsTmp2 = grdApoyoDx.DataSource
        mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp2, Frame3.Caption, txtPaciente.Text, "", 1, False, True
    End If
    Set mo_ReglasReportes = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
errores:
    UserControl.MousePointer = 1
End Sub

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
   
    If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
         UserControl.txtNroHistoria = "" And UserControl.txtDNI.Text = "") Then
        MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, DNI  o Nro Historia)", vbInformation, "Filtro de pacientes"
        Exit Sub
    End If
    If UserControl.txtNroHistoria = "" And txtDNI.Text = "" Then
        If UserControl.txtApellidoPaterno = "" Then
            MsgBox "Por favor ingrese Ap. Paterno", vbInformation, "Filtro de pacientes"
            Exit Sub
        End If
    End If
    RealizarBusqueda
    Screen.MousePointer = vbDefault

End Sub
Public Sub DesdeHistorico(lnHistoriaClinica As Long)
    txtNroHistoria.Text = lnHistoriaClinica
    btnBuscar_Click
    'grdPacientes_DblClick
    fraBusqueda.Enabled = False
    grdApoyoDx.Height = 1800
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
    Set grdPacientes.DataSource = Nothing
    Set grdAtenciones.DataSource = Nothing
    Set grdApoyoDx.DataSource = Nothing
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNroHistoria = ""
        txtDNI.Text = ""
        UserControl.txtPaciente = ""
        UserControl.txtEdad = ""
        UserControl.txtSexo = ""
        UserControl.txtDomicilio = ""
        UserControl.txtTelefono = ""
        UserControl.txtPadres = ""
        UserControl.txtFactorRH = ""
        UserControl.txtGrupoS = ""
        UserControl.txtAlergias = ""
End Sub





Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub grdApoyoDx_DblClick()
  On Error GoTo Fin
  
  
  Dim ml_IdPruebaSeleccionada As String
  Dim ml_NombrePruebaSeleccionada As String
  Dim ml_nombrePaciente As String
  Dim ml_idOrden As Long
  Dim ml_IdProducto As Long
  Dim ml_NombreMedico As String
  Dim ml_areaTrabajo As Long
  Dim ml_idOrdenLab As Long
  Dim oRsTmp As New Recordset
  Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
  

  
  
  'Cargar los formularios para el resultado
  Set oRsTmp = grdApoyoDx.DataSource
  If oRsTmp.Fields!resultado <> "SI" Then
     Set oRsTmp = Nothing
     Set mo_ReglasLaboratorio = Nothing
     Exit Sub
  End If
  If Len(oRsTmp!resultado1) > 0 And oRsTmp!resultado1 <> "SI" Then
     MsgBox oRsTmp!resultado1, vbInformation, "Resultado"
     Set oRsTmp = Nothing
     Set mo_ReglasLaboratorio = Nothing
     Exit Sub
  End If
  
  '*********************Imagen********************
  If Left(oRsTmp!ServicioApDx, 1) = "I" Then
    Dim mo_reglasImagen As New SIGHNegocios.ReglasImagenes
    Dim oResultadosImg As New SIGHImagen.ResultadosImg
    Dim rsResultados As New Recordset
    Dim oRsTmp9 As New Recordset
    Set oRsTmp9 = mo_ReglasComunes.ImagMovimientoSeleccionarIdOrden(oRsTmp!IdOrden)
    If oRsTmp9.RecordCount > 0 Then
        Set rsResultados = mo_reglasImagen.ImagMovimientoResultadosSeleccionarPorId(oRsTmp9!IdMovimiento)
        oResultadosImg.Producto = oRsTmp!Codigo & " " & oRsTmp!Item
        oResultadosImg.SoloEsConsulta = True
        oResultadosImg.idProductoCpt = oRsTmp!idProducto
        oResultadosImg.IdMovimiento = oRsTmp9!IdMovimiento
        Set oResultadosImg.rsResultados = rsResultados
        oResultadosImg.Paciente = txtPaciente.Text
        oResultadosImg.PuntoCarga = oRsTmp9!idPuntoCarga
        oResultadosImg.MostrarFormulario
    End If
    Set mo_ReglasLaboratorio = Nothing
    Set oRsTmp = Nothing
    Set mo_reglasImagen = Nothing
    Set oResultadosImg = Nothing
    Set rsResultados = Nothing
    Set oRsTmp9 = Nothing
    Exit Sub
  End If
  
  
  ml_IdPruebaSeleccionada = oRsTmp("codigo")
  ml_NombrePruebaSeleccionada = oRsTmp("item")
  ml_nombrePaciente = txtPaciente.Text
  ml_idOrden = oRsTmp("idOrden")
  ml_IdProducto = oRsTmp("idProducto")
  
  'debb-10/07/2018
  Dim ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
  ReglasArchivoClinico.ActualizaIDenTablaLabResultadoPorItems Nothing, ml_idOrden
  Set ReglasArchivoClinico = Nothing
  
  If mo_ReglasLaboratorio.UsaNuevaVentanaResultadosLaboratorio(oRsTmp!IdOrden, oRsTmp!Codigo) = True Then
      '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
    
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & ml_IdPruebaSeleccionada & "'")
     ' If oRsTmp1.RecordCount > 0 Then
            Dim lcHistoria As String, lcEdadEnAtencion As String, lcServicioActualPaciente As String
            mo_ReglasLaboratorio.LlenaItemsConResultadosParaImpresion oRsTmp1, ml_idOrden, lcEdadEnAtencion, _
                                                                           lcHistoria, lcServicioActualPaciente, ml_IdPruebaSeleccionada
            mo_ReglasLaboratorio.Imprimir_LabResultadosItems oRsTmp1, lcEdadEnAtencion, lcHistoria, lcServicioActualPaciente
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            Exit Sub
     ' End If
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
      '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
  Else
    Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
    oMuestraResultado.MuestraResultadoDelExamen ml_IdPruebaSeleccionada, ml_NombrePruebaSeleccionada, _
                                                ml_nombrePaciente, ml_idOrden, ml_idRegistroSeleccionado, ml_NombreMedico, _
                                                ml_areaTrabajo, ml_idOrdenLab, ml_idTipoSexo, True
    Set oMuestraResultado = Nothing
  End If
  Set oRsTmp = Nothing
  Set mo_ReglasLaboratorio = Nothing
Fin:
End Sub

Private Sub grdApoyoDx_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdApoyoDx.Override.RowSizing = ssRowSizingFree
    ConfiguraGrillas 3

End Sub

Private Sub grdApoyoDx_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdApoyoDx_DblClick
    End If
End Sub

Private Sub grdAtenciones_DblClick()

        If lbFrame3Visible = False Then
           Exit Sub
        End If
        
        Dim oConsultaCta As New ReembolsosCta
        Dim rsRecordset As ADODB.Recordset
        Set rsRecordset = grdAtenciones.DataSource
        If rsRecordset.Fields!idCuentaAtencion > 0 Then
            oConsultaCta.idCuentaAtencion = rsRecordset.Fields!idCuentaAtencion
            oConsultaCta.Show 1
        End If
        Set oConsultaCta = Nothing
        Set rsRecordset = Nothing

End Sub

Private Sub grdAtenciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdAtenciones.Override.RowSizing = ssRowSizingFree
    ConfiguraGrillas 1
End Sub

Sub ConfiguraGrillas(lnGrilla As Long)
    On Error Resume Next
    If lnGrilla = 1 Then
        grdAtenciones.Bands(0).Columns("IdAtencion").Hidden = True
        grdAtenciones.Bands(0).Columns("idCuentaAtencion").Hidden = True
        grdAtenciones.Bands(0).Columns("FechaIngreso").Width = 800
        grdAtenciones.Bands(0).Columns("HoraIngreso").Width = 500
        grdAtenciones.Bands(0).Columns("TipoServicio").Width = 1000
        grdAtenciones.Bands(0).Columns("Especialidad").Width = 1000
        If lbFrame3Visible = True Then
            grdAtenciones.Bands(0).Columns("Motivo").Width = 2400
            grdAtenciones.Bands(0).Columns("Dx").Width = 2900
            grdAtenciones.Bands(0).Columns("Tratamiento").Width = 1600
            
            grdAtenciones.Bands(0).Columns("Motivo").CellMultiLine = ssCellMultiLineTrue
            grdAtenciones.Bands(0).Columns("Tratamiento").CellMultiLine = ssCellMultiLineTrue
            grdAtenciones.Bands(0).Columns("Dx").CellMultiLine = ssCellMultiLineTrue
            grdAtenciones.Bands(0).Columns("Medico").Width = 1500
        Else
            grdAtenciones.Bands(0).Columns("Medico").Width = 3500
            grdAtenciones.Bands(0).Columns("TipoServicio").Width = 3500
            grdAtenciones.Bands(0).Columns("Especialidad").Width = 3500
        End If
        grdAtenciones.Caption = ""
    ElseIf lnGrilla = 2 Then
        grdPacientes.Bands(0).Columns("IdPaciente").Hidden = True
        grdPacientes.Bands(0).Columns("IdTipoNumeracion").Hidden = True
        
        grdPacientes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "N°Historia"
        grdPacientes.Bands(0).Columns("NroHistoriaClinica").Width = 700
        
        grdPacientes.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
        grdPacientes.Bands(0).Columns("ApellidoPaterno").Width = 1000
        
        grdPacientes.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
        grdPacientes.Bands(0).Columns("ApellidoMaterno").Width = 1000
        
        grdPacientes.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
        grdPacientes.Bands(0).Columns("PrimerNombre").Width = 800
    
        grdPacientes.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
        grdPacientes.Bands(0).Columns("SegundoNombre").Width = 600
    
        grdPacientes.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
        grdPacientes.Bands(0).Columns("FechaNacimiento").Width = 1500
    
        grdPacientes.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
        grdPacientes.Bands(0).Columns("TipoNumeracion").Width = 1500
        grdPacientes.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight
    
        'On Error Resume Next
        grdPacientes.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
        grdPacientes.Bands(0).Columns("TipoServicio").Width = 2000
    
        grdPacientes.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult. Fec Ing."
        grdPacientes.Bands(0).Columns("FechaIngreso").Width = 1500
    
        grdPacientes.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult. Fec Egr."
        grdPacientes.Bands(0).Columns("FechaEgreso").Width = 1500
    
        grdPacientes.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
        grdPacientes.Bands(0).Columns("ServicioIngreso").Width = 1500
        grdPacientes.Caption = ""
    
    ElseIf lnGrilla = 3 Then
        grdApoyoDx.Bands(0).Columns("idCuentaAtencion").Hidden = True
        grdApoyoDx.Bands(0).Columns("resultado1").Hidden = True
        grdApoyoDx.Bands(0).Columns("IdOrden").Hidden = True
        grdApoyoDx.Bands(0).Columns("idProducto").Hidden = True
        grdApoyoDx.Bands(0).Columns("Codigo").Hidden = True
        grdApoyoDx.Bands(0).Columns("Fecha").Width = 800
        grdApoyoDx.Bands(0).Columns("hora").Width = 500
        grdApoyoDx.Bands(0).Columns("servicioApDx").Width = 1000
        grdApoyoDx.Bands(0).Columns("item").Width = 3100
        grdApoyoDx.Bands(0).Columns("cantidad").Width = 400
        grdApoyoDx.Bands(0).Columns("resultado").Width = 2400
        grdApoyoDx.Bands(0).Columns("especialista").Width = 1500
        grdApoyoDx.Bands(0).Columns("resultado").CellMultiLine = ssCellMultiLineTrue
        grdApoyoDx.Bands(0).Columns("item").CellMultiLine = ssCellMultiLineTrue
        grdApoyoDx.Bands(0).Columns("Receta").Width = 800
        grdApoyoDx.Bands(0).Columns("UsuarioDespacho").Width = 1200
        
        grdApoyoDx.Caption = ""
    End If
End Sub



Private Sub grdAtenciones_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdAtenciones_DblClick
    End If
End Sub

Private Sub grdPacientes_DblClick()
    Dim rsRecordset As ADODB.Recordset
    Dim oRsTmp1 As New ADODB.Recordset
    Dim oEdad As Edad
    Dim ml_idCuentaAtencion As Long
    Set rsRecordset = grdPacientes.DataSource
    ml_idRegistroSeleccionado = rsRecordset("IdPaciente")
    txtPaciente.Text = Trim(rsRecordset("ApellidoPaterno")) & " " & _
                       Trim(rsRecordset("ApellidoMaterno")) & " " & _
                       Trim(rsRecordset("PrimerNombre")) & _
                       IIf(IsNull(rsRecordset("SegundoNombre")), "", " " & rsRecordset("SegundoNombre"))
    oEdad = sighEntidades.CalcularEdad(rsRecordset("fechaNacimiento"), Date)
    txtEdad.Text = oEdad.Edad & " " & IIf(oEdad.TipoEdad = sghTipoEdades.sghAño, "Años", _
                                      IIf(oEdad.TipoEdad = sghTipoEdades.sghMeses, "Meses", _
                                      IIf(oEdad.TipoEdad = sghTipoEdades.sghDias, "Días", "Horas")))
                                      
    Set oRsTmp1 = mo_AdminAdmision.PacientesDatosAdicionalesSeleccionarPorIdPaciente(ml_idRegistroSeleccionado)
    If oRsTmp1.RecordCount > 0 Then
        ucPacientesPDF1.Inicializar rsRecordset("IdPaciente"), oRsTmp1!NroHistoriaClinica
        ml_idTipoSexo = oRsTmp1.Fields!idTipoSexo
        txtSexo.Text = IIf(oRsTmp1.Fields!idTipoSexo = 1, "Masculino", "Femenino")
        txtDomicilio.Text = IIf(IsNull(oRsTmp1.Fields!DireccionDomicilio), "", oRsTmp1.Fields!DireccionDomicilio)
        txtTelefono.Text = IIf(IsNull(oRsTmp1.Fields!TELEFONO), "", oRsTmp1.Fields!TELEFONO)
        txtPadres.Text = IIf(IsNull(oRsTmp1.Fields!NombrePadre), "", "(Pad: " & Trim(oRsTmp1.Fields!NombrePadre) & ")") & _
                         IIf(IsNull(oRsTmp1.Fields!Nombremadre), "", "(Mad: " & Trim(oRsTmp1.Fields!Nombremadre) & ")")
        txtFactorRH.Text = IIf(IsNull(oRsTmp1.Fields!FactorRh), "", oRsTmp1.Fields!FactorRh)
        txtGrupoS.Text = IIf(IsNull(oRsTmp1.Fields!GrupoSanguineo), "", oRsTmp1.Fields!GrupoSanguineo)
        txtAlergias.Text = IIf(IsNull(oRsTmp1.Fields!antecedAlergico), "", oRsTmp1.Fields!antecedAlergico)
        'carga Imagen..........si demora mucho al cargar, cambiar en parametros la ruta
        Dim lcRutaImg As String
        lcRutaImg = lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & Trim(Str(oRsTmp1!NroHistoriaClinica)) & ".jpg"
        If sighEntidades.ArchivoExiste(lcRutaImg) Then
           pi_ImagSeleccionada.Picture = LoadPicture(lcRutaImg)
        Else
           pi_ImagSeleccionada.Picture = LoadPicture("")
        End If
        '
        lcRutaImg = mo_AdminFacturacion.DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(rsRecordset("IdPaciente"))
        If lcRutaImg <> "" Then
            txtAlergias.Text = lcRutaImg & Chr(13) & Chr(10) & txtAlergias.Text
        End If
        
        
    End If
    Set grdAtenciones.DataSource = mo_AdminAdmision.AtencionesSeleccionarPorPaciente(ml_idRegistroSeleccionado, lbFrame3Visible)
    If lbFrame3Visible = True Then
       Set grdApoyoDx.DataSource = mo_AdminAdmision.ServiciosIntermediosSeleccionarPorPaciente(ml_idRegistroSeleccionado, False, _
                                     True, 0, True, False)
    End If
    Set oRsTmp1 = Nothing
    Set rsRecordset = Nothing
End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    ConfiguraGrillas 2

End Sub






Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDNI
    AdministrarKeyPreview KeyCode
End Sub





Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
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
    'AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   
'    On Error Resume Next
'
'   fraBusqueda.Width = UserControl.Width - 110
'   lblNombre.Width = UserControl.Width
'
'   grdPacientes.Width = fraBusqueda.Width
'   grdPacientes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 330)
   
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighEntidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, "99"
        mo_Apariencia.ConfigurarFilasBiColores grdAtenciones, "99"
        mo_Apariencia.ConfigurarFilasBiColores grdApoyoDx, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighEntidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores grdAtenciones, sighEntidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores grdApoyoDx, sighEntidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub


Public Sub Inicializar()
      SkinConfigura
      mo_Formulario.HabilitarDeshabilitar txtPaciente, False
      mo_Formulario.HabilitarDeshabilitar txtEdad, False
      mo_Formulario.HabilitarDeshabilitar txtSexo, False
      mo_Formulario.HabilitarDeshabilitar txtDomicilio, False
      mo_Formulario.HabilitarDeshabilitar txtTelefono, False
      mo_Formulario.HabilitarDeshabilitar txtPadres, False
      mo_Formulario.HabilitarDeshabilitar txtFactorRH, False
      mo_Formulario.HabilitarDeshabilitar txtGrupoS, False
      mo_Formulario.HabilitarDeshabilitar txtAlergias, False
'      mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighentidades.GrillaConFilasBicolor
'      mo_Apariencia.ConfigurarFilasBiColores grdAtenciones, sighentidades.GrillaConFilasBicolor
'      mo_Apariencia.ConfigurarFilasBiColores grdApoyoDx, sighentidades.GrillaConFilasBicolor
      
      Dim oRsPermisosTabs As New Recordset
      Dim ms_ReglasSeguridad As New ReglasDeSeguridad
      Set oRsPermisosTabs = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(sighEntidades.Usuario)
      oRsPermisosTabs.Filter = "IdPermiso=600"
      lbFrame3Visible = True
      UserControl.Frame3.Visible = True
      Label11.Visible = True
      SSTab.Visible = True
      If oRsPermisosTabs.RecordCount > 0 Then
         SSTab.Visible = False
         lbFrame3Visible = False
         Frame3.Visible = False
         Label11.Visible = False
         Frame2.Height = Frame2.Height + 2000
         grdAtenciones.Height = Frame2.Height - 500
      End If
      oRsPermisosTabs.Close
      Set oRsPermisosTabs = Nothing
      Set ms_ReglasSeguridad = Nothing
End Sub

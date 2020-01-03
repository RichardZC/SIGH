VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucPacientesLista 
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10005
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
      Height          =   885
      Left            =   75
      TabIndex        =   5
      Top             =   540
      Width           =   9930
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9165
         Picture         =   "ucPacientesLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7800
         Picture         =   "ucPacientesLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1305
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
         Left            =   150
         MaxLength       =   9
         TabIndex        =   0
         Top             =   465
         Width           =   1845
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
         Left            =   2070
         MaxLength       =   30
         TabIndex        =   1
         Top             =   465
         Width           =   1845
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
         Left            =   3975
         MaxLength       =   30
         TabIndex        =   2
         Top             =   465
         Width           =   1845
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   5895
         MaxLength       =   30
         TabIndex        =   3
         Top             =   465
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Historia clínica         Apellido paterno         Apellido materno         Primer nombre                     "
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
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdPacientes 
      Height          =   4590
      Left            =   75
      TabIndex        =   4
      Top             =   1515
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   8096
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de pacientes"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Pacientes"
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
      TabIndex        =   7
      Top             =   15
      Width           =   9975
   End
End
Attribute VB_Name = "ucPacientesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHCOmun.Teclado
Dim mo_Apariencia As New SIGHCOmun.GridInfragistic
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdPacientes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdPacientes.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
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


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
Dim oPaciente As New doPaciente
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, Nombres o Nro Historia)", vbInformation, "Filtro de pacientes"
            Exit Sub
        End If
            
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.PrimerNombre = UserControl.txtPrimerNombre
        oPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
        
'        Select Case ml_TipoFiltro
'        Case sghFiltrarTodos
            Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrar(oPaciente)
'        Case sghFiltrarConHistoriasTemporales
'            Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrarConHistoriasTemporales(oPaciente)
'        Case sghFiltrarConHistoriasDefinitivas
'            Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrarConHistoriasDefinitivas(oPaciente)
'        Case Else
'            MsgBox "Opcion no implementada", vbExclamation, Me.Titulo
'        End Select
        
        Dim rsRespuesta As New Recordset
        Set rsRespuesta = grdPacientes.DataSource
        On Error Resume Next
        If rsRespuesta.RecordCount = 0 Then
            MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
        Else
            UserControl.txtApellidoMaterno = ""
            UserControl.txtApellidoPaterno = ""
            UserControl.txtPrimerNombre = ""
            UserControl.txtNroHistoria = ""
        End If
        
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, SIGHCOmun.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtPrimerNombre = ""
        UserControl.txtNroHistoria = ""
End Sub

Private Sub grdPacientes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdPaciente")
    Debug.Print rsRecordset("NroHistoriaClinica")
End Sub

Private Sub grdPacientes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdPaciente")
    
End Sub

Private Sub grdPacientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdPaciente")
    
End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdPacientes.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientes.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Width = 1300
    
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientes.Bands(0).Columns("PrimerNombre").Width = 1500

    grdPacientes.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientes.Bands(0).Columns("SegundoNombre").Width = 1500

    grdPacientes.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientes.Bands(0).Columns("FechaNacimiento").Width = 1500

    grdPacientes.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientes.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientes.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    On Error Resume Next
    grdPacientes.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    grdPacientes.Bands(0).Columns("TipoServicio").Width = 2000

    grdPacientes.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult. Fec Ing."
    grdPacientes.Bands(0).Columns("FechaIngreso").Width = 1500

    grdPacientes.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult. Fec Egr."
    grdPacientes.Bands(0).Columns("FechaEgreso").Width = 1500

    grdPacientes.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientes.Bands(0).Columns("ServicioIngreso").Width = 1500

End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
End Sub


Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
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
     Case vbKeyF8
    End Select
       
End Sub
Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdPacientes.Width = fraBusqueda.Width
   grdPacientes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 330)
   
End Sub




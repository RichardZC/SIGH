VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucExoneracionLista 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   ScaleHeight     =   5790
   ScaleWidth      =   9840
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
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   9705
      Begin VB.TextBox txtIdCuentaAtencion 
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
         MaxLength       =   30
         TabIndex        =   4
         Top             =   450
         Width           =   1845
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
         Left            =   120
         MaxLength       =   9
         TabIndex        =   3
         Top             =   465
         Width           =   1845
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9165
         Picture         =   "ucExoneracionesLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7800
         Picture         =   "ucExoneracionesLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Historia clínica             N° de cuenta"
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
         TabIndex        =   5
         Top             =   225
         Width           =   8775
      End
   End
   Begin UltraGrid.SSUltraGrid grdAdmision 
      Height          =   4200
      Left            =   60
      TabIndex        =   6
      Top             =   1515
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7408
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de admisiones"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Exoneración"
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
      Top             =   30
      Width           =   9765
   End
End
Attribute VB_Name = "ucExoneracionLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_Teclado As New SIGHComun.Teclado
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroAdmision
Public Event OnClick(oRecordset As Recordset)

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdAdmision.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdAdmision.DataSource
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

Private Sub btnBuscar_Click()
   
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault

End Sub

Public Sub RealizarBusqueda()
Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
        
        oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
        oDOCuentaAtencion.IdCuentaAtencion = Val(UserControl.txtIdCuentaAtencion)
        Set grdAdmision.DataSource = mo_AdminAdmision.AtencionesFiltrarPacientesParaIngresarProcedimientos(oDOPaciente, oDOCuentaAtencion)
        
        Dim rsRespuesta As New Recordset
        Set rsRespuesta = grdAdmision.DataSource
        On Error Resume Next
        If rsRespuesta.RecordCount = 0 Then
            MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
        Else
            UserControl.txtIdCuentaAtencion = ""
            UserControl.txtNroHistoria = ""
        End If
        
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdAdmision, SIGHComun.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtIdCuentaAtencion = ""
        UserControl.txtNroHistoria = ""
End Sub
Private Sub grdAdmision_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdAdmision.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
    RaiseEvent OnClick(rsRecordset)
End Sub

Private Sub grdAdmision_Click()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdAdmision.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
    
    RaiseEvent OnClick(rsRecordset)
    
End Sub


Private Sub grdAdmision_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdAdmision.Bands(0).Columns("IdPaciente").Hidden = True
    grdAdmision.Bands(0).Columns("IdAtencion").Hidden = True
    grdAdmision.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    
    grdAdmision.Bands(0).Columns("FechaIngreso").Header.Caption = "Fecha Ing."
    grdAdmision.Bands(0).Columns("FechaIngreso").Width = 1300
    
    grdAdmision.Bands(0).Columns("HoraIngreso").Header.Caption = "Hora Ing"
    grdAdmision.Bands(0).Columns("HoraIngreso").Width = 1000
    
    grdAdmision.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdAdmision.Bands(0).Columns("TipoNumeracion").Width = 1500
      
    grdAdmision.Bands(0).Columns("ServicioIngreso").Header.Caption = "Servicio Ing"
    grdAdmision.Bands(0).Columns("ServicioIngreso").Width = 2000
    
    grdAdmision.Bands(0).Columns("Edad").Header.Caption = "Edad"
    grdAdmision.Bands(0).Columns("Edad").Width = 500
    
    grdAdmision.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdAdmision.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdAdmision.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdAdmision.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdAdmision.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdAdmision.Bands(0).Columns("PrimerNombre").Width = 1500

    grdAdmision.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdAdmision.Bands(0).Columns("SegundoNombre").Width = 1500

    grdAdmision.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdAdmision.Bands(0).Columns("NroHistoriaClinica").Width = 1200

    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
        grdAdmision.Bands(0).Columns("IdCita").Hidden = True
    Case sghFiltrarEmergencia
        
        grdAdmision.Bands(0).Columns("IdTipoServicio").Hidden = True
        
        grdAdmision.Bands(0).Columns("TipoServicio").Header.Caption = "Tipo servicio"
        grdAdmision.Bands(0).Columns("TipoServicio").Width = 3000
        
        grdAdmision.Bands(0).Columns("FechaEgreso").Header.Caption = "Fecha Egreso"
        grdAdmision.Bands(0).Columns("FechaEgreso").Width = 1300
        
        grdAdmision.Bands(0).Columns("HoraEgreso").Header.Caption = "Hora Egr"
        grdAdmision.Bands(0).Columns("HoraEgreso").Width = 1000
    
    Case sghFiltrarHospitalizacion
        grdAdmision.Bands(0).Columns("FechaEgreso").Header.Caption = "Fecha Egreso"
        grdAdmision.Bands(0).Columns("FechaEgreso").Width = 1000
        
        grdAdmision.Bands(0).Columns("HoraEgreso").Header.Caption = "Hora Egr"
        grdAdmision.Bands(0).Columns("HoraEgreso").Width = 1000
    
        grdAdmision.Bands(0).Columns("DxPrincipal").Header.Caption = "Dx Prin."
        grdAdmision.Bands(0).Columns("DxPrincipal").Width = 600
    
        grdAdmision.Bands(0).Columns("TipoAlta").Header.Caption = "Tipo Alta"
        grdAdmision.Bands(0).Columns("TipoAlta").Width = 2500
    
        grdAdmision.Bands(0).Columns("CondicionAlta").Header.Caption = "Cond. Alta"
        grdAdmision.Bands(0).Columns("CondicionAlta").Width = 1000
    
    End Select


End Sub

Private Sub txtIdCuentaAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtIdCuentaAtencion
End Sub

Private Sub txtIdCuentaAtencion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
        
    Case vbKeyF2
        
    Case vbKeyF3
         btnBuscar_Click
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
     Case vbKeyF7
     Case vbKeyF8
    End Select
       
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
    fraBusqueda.Width = UserControl.Width - 110
    lblNombre.Width = UserControl.Width
   
    grdAdmision.Width = fraBusqueda.Width
    grdAdmision.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 750)

End Sub



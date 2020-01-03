VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucFarmaciaRecetasLista 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10050
   ScaleHeight     =   6210
   ScaleWidth      =   10050
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
      Height          =   945
      Left            =   45
      TabIndex        =   0
      Top             =   525
      Width           =   9930
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
         TabIndex        =   5
         Top             =   465
         Width           =   1845
      End
      Begin VB.TextBox txtNroCuenta 
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
         MaxLength       =   9
         TabIndex        =   4
         Top             =   465
         Width           =   1845
      End
      Begin VB.TextBox txtNroReceta 
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
         Left            =   3990
         MaxLength       =   30
         TabIndex        =   3
         Top             =   465
         Width           =   1845
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7275
         Picture         =   "ucFarmaciaRecetasLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5880
         Picture         =   "ucFarmaciaRecetasLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Historia clínica         Nro Cuenta                 Nro Receta"
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
         TabIndex        =   7
         Top             =   240
         Width           =   7635
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Left            =   150
         TabIndex        =   6
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdProcedimientos 
      Height          =   4650
      Left            =   60
      TabIndex        =   8
      Top             =   1500
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   8202
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
      Caption         =   "Lista de procedimientos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Pre-Facturacion Recetas"
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucFarmaciaRecetasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdProcedimientos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdProcedimientos.DataSource
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
        
        If (UserControl.txtNroCuenta = "" And UserControl.txtNroHistoria = "" And _
            UserControl.txtNroReceta = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Nro Historia, Nro cuenta o Nro Orden)", vbInformation, "Filtro de ordenes de procedimientos"
            Exit Sub
        End If
            
        Dim oFarmaciaRecetas As New DOFarmaciaRecetas
        oFarmaciaRecetas.IdCuentaAtencion = Val(UserControl.txtNroCuenta)
        oFarmaciaRecetas.NroReceta = UserControl.txtNroReceta
        
        Dim oDOPaciente As New doPaciente
        oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
        
        Set grdProcedimientos.DataSource = mo_AdminFacturacion.FarmaciaRecetasFiltrar(oFarmaciaRecetas, oDOPaciente)
        
        If mo_AdminFacturacion.MensajeError <> "" Then
            MsgBox mo_AdminFacturacion.MensajeError, vbCritical, "Filtro órdenes de procedimientos"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdProcedimientos, SIGHComun.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNroCuenta = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtNroReceta = ""
End Sub

Private Sub grdProcedimientos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProcedimientos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdReceta")
    
End Sub

Private Sub grdProcedimientos_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProcedimientos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdReceta")
    
End Sub


Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProcedimientos.Bands(0).Columns("IdReceta").Hidden = True
    
    grdProcedimientos.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "Nro Cuenta"
    grdProcedimientos.Bands(0).Columns("IdCuentaAtencion").Width = 1500
    
    grdProcedimientos.Bands(0).Columns("NroReceta").Header.Caption = "Nro Receta"
    grdProcedimientos.Bands(0).Columns("NroReceta").Width = 1200
    
    grdProcedimientos.Bands(0).Columns("FechaReceta").Header.Caption = "Fecha Receta"
    grdProcedimientos.Bands(0).Columns("FechaReceta").Width = 1200
    
    grdProcedimientos.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdProcedimientos.Bands(0).Columns("NroHistoriaClinica").Width = 1500

    grdProcedimientos.Bands(0).Columns("TipoHistoria").Header.Caption = "Tipo Historia"
    grdProcedimientos.Bands(0).Columns("TipoHistoria").Width = 3000

End Sub

Private Sub txtNroCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroCuenta
End Sub

Private Sub txtNroCuenta_KeyPress(KeyAscii As Integer)
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
   
   grdProcedimientos.Width = fraBusqueda.Width
   grdProcedimientos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub








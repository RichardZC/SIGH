VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucDiagnosticosLista 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   10110
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
      Width           =   10005
      Begin VB.CheckBox ChkSoloDxGalenhos 
         Caption         =   "Solo Dx Galenhos"
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
         Left            =   8145
         TabIndex        =   8
         ToolTipText     =   "Al estar marcado sólo muestra los diagnosticos que tiene un nombre Minsa asignado"
         Top             =   540
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5310
         Picture         =   "ucDiagnosticosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   6675
         Picture         =   "ucDiagnosticosLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   0
         Top             =   480
         Width           =   1065
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   1
         Top             =   480
         Width           =   3915
      End
      Begin VB.Label Label2 
         Caption         =   "     Código                               Descripción"
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
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   3795
      End
   End
   Begin UltraGrid.SSUltraGrid grdDiagnosticos 
      Height          =   4050
      Left            =   75
      TabIndex        =   4
      Top             =   1515
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   7144
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
      Caption         =   "Lista de diagnósticos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Diagnósticos (CIE-10)"
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
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   10080
   End
End
Attribute VB_Name = "ucDiagnosticosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar diagnósticos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim ml_idRegistroSeleccionado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
'mgaray09
Dim mb_mostrarSoloActivos As Boolean

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdDiagnosticos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdDiagnosticos.DataSource
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
'mgaray09
Property Let MostrarSoloActivos(bValue As Boolean)
    mb_mostrarSoloActivos = bValue
End Property

Property Get MostrarSoloActivos() As Boolean
    MostrarSoloActivos = mb_mostrarSoloActivos
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
Dim oDODiagnostico As New DODiagnostico

        oDODiagnostico.CodigoCIE2004 = UserControl.txtCodigo
        oDODiagnostico.Descripcion = UserControl.txtDescripcion
        'mgaray09
        If mb_mostrarSoloActivos = False Then
            Set grdDiagnosticos.DataSource = mo_AdminServiciosComunes.DiagnosticosFiltrar(oDODiagnostico, ChkSoloDxGalenhos.Value, False)
        Else
            Set grdDiagnosticos.DataSource = mo_AdminServiciosComunes.DiagnosticosFiltrarSoloActivos(oDODiagnostico, ChkSoloDxGalenhos.Value, False)
        End If
'        Set grdDiagnosticos.DataSource = mo_AdminServiciosComunes.DiagnosticosFiltrar(oDODiagnostico, True, False)
        
     '   mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtCodigo = ""
        UserControl.txtDescripcion = ""
End Sub

Private Sub ChkSoloDxGalenhos_Click()
    btnBuscar_Click
End Sub

Private Sub grdDiagnosticos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdDiagnosticos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdDiagnostico")
 
End Sub

Private Sub grdDiagnosticos_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdDiagnosticos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdDiagnostico")
    
End Sub


Private Sub grdDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdDiagnosticos.Bands(0).Columns("IdDiagnostico").Hidden = True
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Header.Caption = "CIE-10"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Width = 1000
    
    grdDiagnosticos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 8400
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE10").Hidden = True
    grdDiagnosticos.Bands(0).Columns("CodigoCIE10").Header.Caption = "CIE10"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE10").Width = 1000
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE9").Header.Caption = "CIE-9"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE9").Width = 1000

    grdDiagnosticos.Bands(0).Columns("EsActivo").Header.Caption = "Habilitado"
    grdDiagnosticos.Bands(0).Columns("EsActivo").Width = 500
    
    grdDiagnosticos.Bands(0).Columns("FechaInicioVigencia").Header.Caption = "F. Vigencia"
    grdDiagnosticos.Bands(0).Columns("FechaInicioVigencia").Width = 1100
    grdDiagnosticos.Bands(0).Columns("FechaInicioVigencia").CellAppearance.TextAlign = ssAlignCenter
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(txtCodigo)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Initialize()
    'mo_Formulario.ConfigurarTipoLetraDeControles UserControl.Controls
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdDiagnosticos.Width = fraBusqueda.Width
   grdDiagnosticos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
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
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub inicializar()
    SkinConfigura
End Sub

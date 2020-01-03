VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucEstablecNoMinsaLista 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LockControls    =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   10140
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
      Height          =   915
      Left            =   75
      TabIndex        =   7
      Top             =   510
      Width           =   10050
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8880
         Picture         =   "ucEstablecimientoNoMinsa.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   10245
         Picture         =   "ucEstablecimientoNoMinsa.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.ComboBox cmbIdDistrito 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6585
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Width           =   2250
      End
      Begin VB.ComboBox cmbIdProvincia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Width           =   2085
      End
      Begin VB.ComboBox cmbIdDepartamento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2295
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   2085
      End
      Begin VB.TextBox txtNombre 
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
         TabIndex        =   0
         Top             =   450
         Width           =   2145
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre                                Departamento                     Provincia                         Distrito"
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
         Left            =   270
         TabIndex        =   9
         Top             =   210
         Width           =   9075
      End
   End
   Begin UltraGrid.SSUltraGrid grdEstablecimientosNoMinsa 
      Height          =   4110
      Left            =   75
      TabIndex        =   6
      Top             =   1500
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   7250
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
      Caption         =   "Lista de establecimientos no minsa"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Establecimientos no Minsa"
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
      TabIndex        =   8
      Top             =   30
      Width           =   10095
   End
End
Attribute VB_Name = "ucEstablecNoMinsaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Establecimientos NO MINSA
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminReglasCOmunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbIdProvincia As New sighentidades.ListaDespleglable
Dim mo_cmbIdDistrito As New sighentidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdEstablecimientosNoMinsa.DataSource = oValue
    mo_Apariencia.ConfigurarFilasBiColores grdEstablecimientosNoMinsa, sighentidades.GrillaConFilasBicolor
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdEstablecimientosNoMinsa.DataSource
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
Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNombre = ""
        mo_cmbIdDepartamento.BoundText = ""
        mo_cmbIdProvincia.BoundText = ""
        mo_cmbIdDistrito.BoundText = ""
End Sub
Sub RealizarBusqueda()
Dim oEstablecimiento As New DOEstablecimientoNoMinsa
        
        If (UserControl.txtNombre = "" And _
            UserControl.cmbIdDepartamento = "" And UserControl.cmbIdProvincia = "" _
            And UserControl.cmbIdDistrito = "") Then
        End If
        
        oEstablecimiento.nombre = UserControl.txtNombre
        oEstablecimiento.IdDistrito = Val(mo_cmbIdDistrito.BoundText)
        
        Set grdEstablecimientosNoMinsa.DataSource = mo_AdminReglasCOmunes.EstablecimientosNoMinsaFiltrar(oEstablecimiento, Val(mo_cmbIdDepartamento.BoundText), Val(mo_cmbIdProvincia.BoundText))
        
        If mo_AdminReglasCOmunes.MensajeError <> "" Then
            MsgBox mo_AdminReglasCOmunes.MensajeError, vbInformation, "Filtro PrestamosHC"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdEstablecimientosNoMinsa, sighentidades.GrillaConFilasBicolor

End Sub
Private Sub cmbIdDepartamento_Click()
        
       mo_cmbIdProvincia.BoundColumn = "IdProvincia"
       mo_cmbIdProvincia.ListField = "Nombre"
       On Error Resume Next
       Set mo_cmbIdProvincia.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdProvincia.BoundText = ""
       mo_cmbIdDistrito.BoundText = ""
        
End Sub
Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDepartamento_LostFocus()
   'If cmbIdDepartamento.Text <> "" Then
   '    mo_cmbIdDepartamento.BoundText = Val(Split(cmbIdDepartamento.Text, " = ")(0))
   'End If
End Sub

Private Sub cmbIdDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistrito
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdDistrito_LostFocus()
   'If cmbIdDistrito.Text <> "" Then
   '    mo_cmbIdDistrito.BoundText = Val(Split(cmbIdDistrito.Text, " = ")(0))
   'End If
   
End Sub

Private Sub cmbIdDistrito_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdProvincia_Click()
        
       mo_cmbIdDistrito.BoundColumn = "IdDistrito"
       mo_cmbIdDistrito.ListField = "Nombre"
       Set mo_cmbIdDistrito.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(mo_cmbIdProvincia.BoundText))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Lista de establecimientos"
       End If
       
       mo_cmbIdDistrito.BoundText = ""
       
End Sub

Private Sub cmbIdProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvincia
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdProvincia_LostFocus()
   
   'If cmbIdProvincia.Text <> "" Then
   '    mo_cmbIdProvincia.BoundText = Val(Split(cmbIdProvincia.Text, " = ")(0))
   'End If
   
   
End Sub

Private Sub cmbIdProvincia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub grdEstablecimientosNoMinsa_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdEstablecimientosNoMinsa.DataSource
    ml_IdRegistroSeleccionado = rsRecordset("IdEstablecimientoNoMinsa")
End Sub

Private Sub grdEstablecimientosNoMinsa_Click()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdEstablecimientosNoMinsa.DataSource
    ml_IdRegistroSeleccionado = rsRecordset("IdEstablecimientoNoMinsa")
    
End Sub

Private Sub grdEstablecimientosNoMinsa_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdEstablecimientosNoMinsa.Bands(0).Columns("IdEstablecimientoNoMinsa").Hidden = True
    
    grdEstablecimientosNoMinsa.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdEstablecimientosNoMinsa.Bands(0).Columns("Nombre").Width = 3000
    
    grdEstablecimientosNoMinsa.Bands(0).Columns("SubSector").Header.Caption = "SubSector"
    grdEstablecimientosNoMinsa.Bands(0).Columns("SubSector").Width = 3000
    
    grdEstablecimientosNoMinsa.Bands(0).Columns("Departamento").Header.Caption = "Departamento"
    grdEstablecimientosNoMinsa.Bands(0).Columns("Departamento").Width = 2000
    
    grdEstablecimientosNoMinsa.Bands(0).Columns("Provincia").Header.Caption = "Provincia"
    grdEstablecimientosNoMinsa.Bands(0).Columns("Provincia").Width = 2000
    
    grdEstablecimientosNoMinsa.Bands(0).Columns("Distrito").Header.Caption = "Distrito"
    grdEstablecimientosNoMinsa.Bands(0).Columns("Distrito").Width = 2000

    
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Public Function Inicializar()
    Set mo_cmbIdProvincia.MiComboBox = cmbIdProvincia
    Set mo_cmbIdDistrito.MiComboBox = cmbIdDistrito
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
End Function


Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdEstablecimientosNoMinsa.Width = fraBusqueda.Width
   grdEstablecimientosNoMinsa.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Sub ConfigurarEstablecimientos()
    
    mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
    mo_cmbIdDepartamento.ListField = "Nombre"
    Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosGeograficos.DepartamentosSeleccionarTodos()
    mo_cmbIdDepartamento.BoundText = Trim(Str(Val(Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2))))
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


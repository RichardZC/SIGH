VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucEspecialidadesLista 
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   ScaleHeight     =   6165
   ScaleWidth      =   10215
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
      TabIndex        =   4
      Top             =   525
      Width           =   10035
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
         Left            =   120
         TabIndex        =   0
         Top             =   450
         Width           =   4125
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
         Left            =   4320
         TabIndex        =   1
         Top             =   450
         Width           =   3330
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8550
         Picture         =   "ucEspecialidadesLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "                           Nombre                                                 Departamento"
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
         Top             =   240
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdEspecialidades 
      Height          =   4590
      Left            =   75
      TabIndex        =   3
      Top             =   1500
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8096
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
      Caption         =   "Lista de Especialidades"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Especialidades"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   6
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "ucEspecialidadesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Especialidades
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idRegistroSeleccionado As Long
Dim mo_AdminServHosp As New ReglasServiciosHosp
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdDepartamento As New ListaDespleglable
Dim mo_Teclado As New sighentidades.Teclado

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdEspecialidades.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdEspecialidades.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let IdDepartamento(lValue As Long)
   mo_cmbIdDepartamento.BoundText = lValue
End Property
Property Get IdDepartamento() As Long
   IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
End Property



Sub ConfigurarDepartamento()
    
    mo_cmbIdDepartamento.ListField = "DescripcionLarga"
    mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
    Set mo_cmbIdDepartamento.RowSource = mo_AdminServHosp.DepartamentosSeleccionarTodos()

End Sub

Private Sub btnBuscar_Click()
Me.RealizarBusqueda
End Sub

Private Sub cmbIdDepartamento_Click()
    ActualizarGrilla
    'ActualizarJerarquia
End Sub

Public Sub ActualizarGrilla()
If (mo_cmbIdDepartamento.BoundText <> "") Then
    Set grdEspecialidades.DataSource = mo_AdminServHosp.EspecialidadesSeleccionarporDepartamentoV2(Val(mo_cmbIdDepartamento.BoundText))
Else
    Dim oDOEspecialidades As New DOEspecialidades
    Set grdEspecialidades.DataSource = mo_AdminServHosp.EspecialidadesFiltrar(oDOEspecialidades)
End If
    mo_Apariencia.ConfigurarFilasBiColores grdEspecialidades, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbIdDepartamento_LostFocus()

   If cmbIdDepartamento.Text <> "" Then
       mo_cmbIdDepartamento.BoundText = Val(Split(cmbIdDepartamento.Text, " = ")(0))
   End If

End Sub

Private Sub grdEspecialidades_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdEspecialidades.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdEspecialidad")
End Sub

Private Sub grdEspecialidades_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdEspecialidades.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdEspecialidad")
    
End Sub


Private Sub grdEspecialidades_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    FormatearGrilla
End Sub

Sub FormatearGrilla()

    grdEspecialidades.Bands(0).Columns("IdEspecialidad").Hidden = True
    grdEspecialidades.Bands(0).Columns("IdDepartamento").Hidden = True
    
    grdEspecialidades.Bands(0).Columns("DescripcionLarga").Width = 3500
    grdEspecialidades.Bands(0).Columns("DescripcionLarga").Header.Caption = "Nombre"

    grdEspecialidades.Bands(0).Columns("TiempoPromedioAtencion").Width = 2000
    grdEspecialidades.Bands(0).Columns("TiempoPromedioAtencion").Header.Caption = "Tiempo promedio"

    grdEspecialidades.Bands(0).Columns("ProductoConsulta").Width = 4500
    grdEspecialidades.Bands(0).Columns("ProductoConsulta").Header.Caption = "Producto consulta"

    grdEspecialidades.Bands(0).Columns("ProductoInterconsulta").Width = 4500
    grdEspecialidades.Bands(0).Columns("ProductoInterconsulta").Header.Caption = "Producto interconsulta"

End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
'        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
'        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdEspecialidades, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdEspecialidades, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Public Function Inicializar()
    SkinConfigura
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    ConfigurarDepartamento
End Function


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub

Private Sub UserControl_Resize()
   
     On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = UserControl.Width
   
   grdEspecialidades.Width = fraBusqueda.Width
   grdEspecialidades.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Sub RealizarBusqueda()
Dim oDOEspecialidades As New DOEspecialidades

        oDOEspecialidades.nombre = UserControl.txtNombre
        oDOEspecialidades.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
        
        Set grdEspecialidades.DataSource = mo_AdminServHosp.EspecialidadesFiltrar(oDOEspecialidades)
        
       ' mo_Apariencia.ConfigurarFilasBiColores grdEspecialidades, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNombre = ""
        UserControl.cmbIdDepartamento.ListIndex = 0
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

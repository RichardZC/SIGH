VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ucServiciosLista 
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   12105
   Begin VB.Frame fraResultado 
      Height          =   6975
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   7245
      Begin VB.ComboBox cmbIdTipoServicio 
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
         Left            =   1485
         TabIndex        =   0
         Top             =   240
         Width           =   2730
      End
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   6135
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   10821
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
         Caption         =   "Lista De Servicios"
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Servicio"
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
         Left            =   240
         TabIndex        =   6
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame fraTree 
      Height          =   6975
      Left            =   7350
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      Begin MSComctlLib.TreeView treeServicios 
         Height          =   6585
         Left            =   150
         TabIndex        =   2
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   11615
         _Version        =   393217
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Servicios"
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
      TabIndex        =   5
      Top             =   15
      Width           =   11985
   End
End
Attribute VB_Name = "ucServiciosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Servicios/Consultorios
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idRegistroSeleccionado As Long
Dim mo_AdminServHosp As New ReglasServiciosHosp
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdTipoServicio As New ListaDespleglable

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdServicios.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdServicios.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let idTipoServicio(lValue As Long)
   mo_cmbIdTipoServicio.BoundText = lValue
End Property
Property Get idTipoServicio() As Long
   idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
End Property

Sub ConfigurarTipoServicio()
    
    mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
    mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
    Set mo_cmbIdTipoServicio.RowSource = mo_AdminServHosp.TiposServicioSeleccionarTodos()

End Sub

Private Sub cmbIdTipoServicio_Click()
    ActualizarGrilla
    ActualizarJerarquia
End Sub

Public Sub ActualizarGrilla()
    Set grdServicios.DataSource = mo_AdminServHosp.ServiciosSeleccionarPorTipo(Val(mo_cmbIdTipoServicio.BoundText))
   ' mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
End Sub
Public Sub ActualizarJerarquia()
Dim rsDpto As ADODB.Recordset
Dim rsEspecialidad As ADODB.Recordset
Dim rsServicio As ADODB.Recordset
Dim oNodeDpto As Node
Dim oNodeEsp As Node
Dim oNodeServ As Node
Dim oNode As Node

    treeServicios.Nodes.Clear
    
    Set rsDpto = mo_AdminServHosp.DepartamentoSeleccionarPorTipoServicio(Val(mo_cmbIdTipoServicio.BoundText))
    Do While Not rsDpto.EOF
        Set oNodeDpto = treeServicios.Nodes.Add(, , "D" & rsDpto!IdDepartamento, rsDpto!nombre)
        oNodeDpto.Expanded = True
        Set rsEspecialidad = mo_AdminServHosp.EspecialidadSeleccionarPorTipoServicioYDpto(Val(mo_cmbIdTipoServicio.BoundText), rsDpto!IdDepartamento)
        Do While Not rsEspecialidad.EOF
            Set oNodeEsp = treeServicios.Nodes.Add("D" & rsDpto!IdDepartamento, tvwChild, "E" & rsEspecialidad!IdEspecialidad, rsEspecialidad!nombre)
            oNodeEsp.Expanded = True
            Set rsServicio = mo_AdminServHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(Val(mo_cmbIdTipoServicio.BoundText), rsDpto!IdDepartamento, rsEspecialidad!IdEspecialidad)
            Do While Not rsServicio.EOF
                Set oNodeServ = treeServicios.Nodes.Add("E" & rsEspecialidad!IdEspecialidad, tvwChild, "S" & rsServicio!IdServicio, rsServicio!nombre)
                oNodeServ.Expanded = True
                rsServicio.MoveNext
            Loop
            rsServicio.Close
            rsEspecialidad.MoveNext
        Loop
        rsEspecialidad.Close
        rsDpto.MoveNext
    Loop

    

End Sub


Private Sub cmbIdTipoServicio_LostFocus()

   If cmbIdTipoServicio.Text <> "" Then
       mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If

End Sub

Private Sub grdServicios_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdServicios.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdServicio")
End Sub

Private Sub grdServicios_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdServicios.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdServicio")
    
End Sub


Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdServicios.Bands(0).Columns("IdServicio").Hidden = True
    
    grdServicios.Bands(0).Columns("Codigo").Width = 1000
    grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"

    grdServicios.Bands(0).Columns("Servicio").Width = 3000
    grdServicios.Bands(0).Columns("Especialidad").Width = 3000
    
    grdServicios.Bands(0).Columns("Departamento").Width = 2500
    grdServicios.Bands(0).Columns("Departamento").Header.Caption = "Departamento"
    
    
End Sub
Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
'        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
'        btnBuscar.Caption = ""
'        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
'        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Public Function Inicializar()
    SkinConfigura
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    ConfigurarTipoServicio
End Function

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   lblNombre.Width = UserControl.Width
   
   fraResultado.Width = UserControl.Width - 200 - fraTree.Width
   fraTree.Left = fraResultado.Left + fraResultado.Width + 100
   
   grdServicios.Width = fraResultado.Width - 350
   fraResultado.Height = UserControl.Height - (lblNombre.Height + 150)
   grdServicios.Height = fraResultado.Height - 830
   
   fraTree.Height = fraResultado.Height
   treeServicios.Height = fraTree.Height - 400
   
End Sub

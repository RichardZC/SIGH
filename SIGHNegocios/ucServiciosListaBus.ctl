VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.UserControl ucServiciosListaBus 
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   ScaleHeight     =   6120
   ScaleWidth      =   10140
   Begin VB.Frame fraBusqueda 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   60
      TabIndex        =   6
      Top             =   600
      Width           =   9975
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7980
         Picture         =   "ucServiciosListaBus.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1305
      End
      Begin MSDataListLib.DataCombo cmbIdTipoServicio 
         Height          =   315
         Left            =   4350
         TabIndex        =   2
         Top             =   450
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   450
         Width           =   3345
      End
      Begin VB.TextBox txtIdServicio 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Código     Nombre                                                              Tipo de Servicio"
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
   End
   Begin VB.Frame fraResultado 
      Height          =   4545
      Left            =   60
      TabIndex        =   5
      Top             =   1470
      Width           =   9975
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7435
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
         Caption         =   "Lista de servicios"
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00000000&
      Caption         =   "Búsqueda de servicios"
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
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   9975
   End
End
Attribute VB_Name = "ucServiciosListaBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Servicios del Establecimiento
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_SoloIdTipoServicio As Long

Property Let SoloIdTipoServicio(lValue As Long)
    ml_SoloIdTipoServicio = lValue
    If ml_SoloIdTipoServicio > 0 Then
       cmbIdTipoServicio.BoundText = Trim(Str(ml_SoloIdTipoServicio))
       'cmbIdTipoServicio.Enabled = False
    End If
End Property


Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdServicios.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdServicios.DataSource
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
Property Let idTipoServicio(lValue As Long)
    cmbIdTipoServicio.BoundText = lValue
    btnBuscar_Click
End Property
Property Get idTipoServicio() As Long
    idTipoServicio = Val(cmbIdTipoServicio.BoundText)
End Property
Property Let HabilitarTipoServicio(lValue As Boolean)
    cmbIdTipoServicio.Enabled = lValue
End Property
Property Get HabilitarTipoServicio() As Boolean
    HabilitarTipoServicio = cmbIdTipoServicio.Enabled
End Property


Private Sub btnBuscar_Click()
Dim oServicio As New DOServicio
        
        oServicio.Codigo = Val(UserControl.txtIdServicio)
        oServicio.nombre = UserControl.txtNombre
        oServicio.idTipoServicio = Val(UserControl.cmbIdTipoServicio)
        
        Set grdServicios.DataSource = mo_AdminServiciosHosp.ServiciosFiltrar(oServicio, 0, sghFiltraSoloActivos)
        
        If mo_AdminServiciosHosp.MensajeError <> "" Then
            MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, "Filtro Servicios"
        End If
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
        'debb2014b
        On Error Resume Next
        grdServicios.SetFocus
End Sub

Public Sub ColocarFocoEnGrillaServicio()
    grdServicios.SetFocus
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdServicios_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdServicios.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdServicio")
    
End Sub

Private Sub grdServicios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdServicios.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdServicio")
    
End Sub

Private Sub grdServicios_DblClick()
     grdServicios_Click
     RaiseEvent SeleccionaRegistro(ml_IdRegistroSeleccionado)
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdServicios.Bands(0).Columns("IdServicio").Hidden = True
    
    grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(0).Columns("Codigo").Width = 750
    
    grdServicios.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(0).Columns("Nombre").Width = 5000
    
    grdServicios.Bands(0).Columns("Especialidad").Header.Caption = "Especialidad"
    grdServicios.Bands(0).Columns("Especialidad").Width = 1500
    

End Sub

Private Sub grdServicios_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdServicios_DblClick
    End If
End Sub



Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = fraBusqueda.Width
   
   fraResultado.Width = UserControl.Width - 150
   grdServicios.Width = fraResultado.Width - 250
   
   fraResultado.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 250)
   grdServicios.Height = fraResultado.Height - 320
   
End Sub

Public Sub ConfigurarTiposServicio()
    
    UserControl.cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
    UserControl.cmbIdTipoServicio.ListField = "DescripcionLarga"
    Set UserControl.cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
    
    
End Sub

Property Let NombreServicio(lValue As String)
    txtNombre = lValue
    btnBuscar_Click
    
End Property

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

'debb2014b
Public Sub FocusEnDescripcion()
    On Error Resume Next
    If txtIdServicio.Text = "" And txtNombre.Text = "" Then
       txtNombre.SetFocus
    Else
       grdServicios.SetFocus
    End If
End Sub

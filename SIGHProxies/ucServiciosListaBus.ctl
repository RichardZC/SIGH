VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucServiciosListaBus 
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10140
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      Height          =   885
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   9975
      Begin MSDataListLib.DataCombo cmbIdTipoServicio 
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Top             =   450
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   450
         Width           =   4125
      End
      Begin VB.TextBox txtIdServicio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   450
         Width           =   735
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9120
         Picture         =   "ucServiciosListaBus.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Código       Nombre                                                                                Tipo de Servicio"
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   7635
      End
   End
   Begin VB.Frame fraResultado 
      Height          =   4545
      Left            =   60
      TabIndex        =   0
      Top             =   1470
      Width           =   9975
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7435
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Lista de servicios"
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Búsqueda de servicios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   9975
   End
End
Attribute VB_Name = "ucServiciosListaBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New SIGHComun.GridInfragistic
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
Property Let IdTipoServicio(lValue As Long)
    cmbIdTipoServicio.BoundText = lValue
End Property
Property Get IdTipoServicio() As Long
    IdTipoServicio = Val(cmbIdTipoServicio.BoundText)
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
        oServicio.Nombre = UserControl.txtNombre
        oServicio.IdTipoServicio = Val(UserControl.cmbIdTipoServicio)
        
        Set grdServicios.DataSource = mo_AdminServiciosHosp.ServiciosFiltrar(oServicio, 0)
        
        If mo_AdminServiciosHosp.MensajeError <> "" Then
            MsgBox mo_AdminServiciosHosp.MensajeError, vbCritical, "Filtro Servicios"
        End If
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, SIGHComun.GrillaConFilasBicolor
        
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

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdServicios.Bands(0).Columns("IdServicio").Hidden = True
    
    grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(0).Columns("Codigo").Width = 750
    
    grdServicios.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(0).Columns("Nombre").Width = 3000
    
    grdServicios.Bands(0).Columns("Especialidad").Header.Caption = "Especialidad"
    grdServicios.Bands(0).Columns("Especialidad").Width = 2000
    

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

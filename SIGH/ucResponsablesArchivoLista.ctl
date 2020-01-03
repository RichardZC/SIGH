VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucArchivadoresLista 
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   LockControls    =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   10245
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      Height          =   915
      Left            =   90
      TabIndex        =   0
      Top             =   570
      Width           =   10035
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   5325
         TabIndex        =   5
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtApellidoMaterno 
         Height          =   315
         Left            =   3420
         TabIndex        =   4
         Top             =   465
         Width           =   1845
      End
      Begin VB.TextBox txtApellidoPaterno 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   465
         Width           =   1845
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7290
         Picture         =   "ucResponsablesArchivoLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   465
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Código Planilla        Apellido paterno                 Apellido materno                Nombre                   "
         Height          =   225
         Left            =   165
         TabIndex        =   6
         Top             =   240
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdArchiveroServicio 
      Height          =   4860
      Left            =   90
      TabIndex        =   7
      Top             =   1560
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8573
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "Lista de responsables"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Asignación de servicios "
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
      Left            =   60
      TabIndex        =   8
      Top             =   45
      Width           =   10200
   End
End
Attribute VB_Name = "ucArchivadoresLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdArchiveroServicio.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdArchiveroServicio.DataSource
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
Property Let TipoBusqueda(lValue As sghTipoBusquedaPrestamoHistoria)
    ml_TipoBusqueda = lValue
End Property
Property Get TipoBusqueda() As sghTipoBusquedaPrestamoHistoria
    TipoBusqueda = ml_TipoBusqueda
End Property


Private Sub btnBuscar_Click()
Dim oEmpleado As New DOEmpleado
Dim oArchivero As New DOArchiveroServicio
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtNombre = "" And UserControl.txtCodigo = "") Then
        End If
            
        
        oEmpleado.ApellidoMaterno = UserControl.txtApellidoMaterno
        oEmpleado.ApellidoPaterno = UserControl.txtApellidoPaterno
        oEmpleado.Nombres = UserControl.txtNombre
        oEmpleado.CodigoPlanilla = UserControl.txtCodigo
        
        Set grdArchiveroServicio.DataSource = mo_AdminArchivoClinico.ArchiveroServicioFiltrar(oEmpleado)
        
        mo_Apariencia.ConfigurarFilasBiColores grdArchiveroServicio, SIGHComun.GrillaConFilasBicolor
        
End Sub

Private Sub grdArchiveroServicio_Click()
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdArchiveroServicio.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdEmpleado")
    
End Sub

Private Sub grdArchiveroServicio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    
End Sub

Private Sub grdArchiveroServicio_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdArchiveroServicio.Bands(0).Columns("IdEmpleado").Hidden = True
    
    grdArchiveroServicio.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdArchiveroServicio.Bands(0).Columns("ApellidoPaterno").Width = 2000
    
    grdArchiveroServicio.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdArchiveroServicio.Bands(0).Columns("ApellidoMaterno").Width = 2000
    
    grdArchiveroServicio.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdArchiveroServicio.Bands(0).Columns("Nombres").Width = 2000
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdArchiveroServicio.Width = fraBusqueda.Width
   grdArchiveroServicio.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub












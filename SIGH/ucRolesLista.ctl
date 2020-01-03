VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucRolesLista 
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   10110
   Begin UltraGrid.SSUltraGrid grdRoles 
      Height          =   5265
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9287
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
      Caption         =   "Relación de roles"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Roles"
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
      TabIndex        =   1
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "ucRolesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Roles
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim ml_idRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdRoles.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdRoles.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Public Sub Refrescar()
        Set grdRoles.DataSource = mo_AdminSeguridad.RolesSeleccionarTodos()
        'mo_Apariencia.ConfigurarFilasBiColores grdRoles, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub grdRoles_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdRoles.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdRol")
End Sub

Private Sub grdRoles_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdRoles.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdRol")
    
End Sub


Private Sub grdRoles_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdRoles.Bands(0).Columns("IdRol").Hidden = True
    
    grdRoles.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdRoles.Bands(0).Columns("Nombre").Width = 4000
    
    'mo_Apariencia.ConfigurarFilasBiColores grdRoles, sighentidades.GrillaConFilasBicolor
    
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   lblNombre.Width = UserControl.Width
   grdRoles.Width = UserControl.Width - 150
   grdRoles.Height = UserControl.Height - (lblNombre.Height + 100)
   
End Sub
Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        mo_Apariencia.ConfigurarFilasBiColores grdRoles, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdRoles, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub inicializar()
    SkinConfigura
End Sub

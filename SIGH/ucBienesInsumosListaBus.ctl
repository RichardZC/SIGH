VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucBienesInsumosListaBus 
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   KeyPreview      =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   10125
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
      Top             =   525
      Width           =   10035
      Begin VB.TextBox txtIdProducto 
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
         TabIndex        =   0
         Top             =   450
         Width           =   735
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
         Left            =   960
         TabIndex        =   1
         Top             =   450
         Width           =   4125
      End
      Begin VB.ComboBox cmbIdClasificacionBienInsumo 
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
         Left            =   5145
         TabIndex        =   2
         Top             =   450
         Width           =   3330
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8550
         Picture         =   "ucBienesInsumosListaBus.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Código                                Nombre                                                 Tipo"
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
         TabIndex        =   6
         Top             =   240
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdBienesInsumos 
      Height          =   4590
      Left            =   75
      TabIndex        =   4
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
      Caption         =   "Lista de Bienes e Insumos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Búsqueda de Bienes e Insumos"
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
      TabIndex        =   7
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "ucBienesInsumosListaBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Bienes Insumos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminComunes As New SIGHNegocios.ReglasComunes
Dim ml_idRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdClasificacionBienInsumo As New ListaDespleglable
Dim ml_IdDepartamentoHospital As Long
Dim mo_Teclado As New sighentidades.Teclado

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdBienesInsumos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdBienesInsumos.DataSource
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
Property Let IdClasificacionBienInsumo(lValue As Long)
    mo_cmbIdClasificacionBienInsumo.BoundText = lValue
End Property
Property Get IdClasificacionBienInsumo() As Long
    IdClasificacionBienInsumo = Val(mo_cmbIdClasificacionBienInsumo.BoundText)
End Property
Property Let HabilitarTipoBienInsumo(lValue As Boolean)
    cmbIdClasificacionBienInsumo.Enabled = lValue
End Property
Property Get HabilitarTipoBienInsumo() As Boolean
    HabilitarTipoBienInsumo = cmbIdClasificacionBienInsumo.Enabled
End Property
Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim oBienesInsumos As New DOCatalogoBienesInsumos
                
        oBienesInsumos.Codigo = Val(UserControl.txtIdProducto)
        oBienesInsumos.nombre = UserControl.txtNombre
        oBienesInsumos.IdClasificacionBienInsumo = Me.IdClasificacionBienInsumo
        
        Set grdBienesInsumos.DataSource = mo_AdminComunes.CatalogoBienesInsumosFiltrar(oBienesInsumos, 0)
        
        If mo_AdminComunes.MensajeError <> "" Then
            MsgBox mo_AdminComunes.MensajeError, vbInformation, "Filtro Servicios"
        End If
        mo_Apariencia.ConfigurarFilasBiColores grdBienesInsumos, sighentidades.GrillaConFilasBicolor
        
End Sub



Private Sub cmbIdClasificacionBienInsumo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdClasificacionBienInsumo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub grdBienesInsumos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdBienesInsumos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdProducto")
End Sub

Private Sub grdBienesInsumos_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdBienesInsumos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdProducto")
    
End Sub


Private Sub grdBienesInsumos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdBienesInsumos.Bands(0).Columns("IdProducto").Hidden = True
    
    grdBienesInsumos.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdBienesInsumos.Bands(0).Columns("Codigo").Width = 750
    
    grdBienesInsumos.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdBienesInsumos.Bands(0).Columns("Nombre").Width = 3000
    
    grdBienesInsumos.Bands(0).Columns("NombreComercial").Header.Caption = "Nombre Comercial"
    grdBienesInsumos.Bands(0).Columns("NombreComercial").Width = 2000
    
    grdBienesInsumos.Bands(0).Columns("DescTiposDeBienesEInsumos").Header.Caption = "Tipo"
    grdBienesInsumos.Bands(0).Columns("DescTiposDeBienesEInsumos").Width = 1000

End Sub

Public Function inicializar()
    Set mo_cmbIdClasificacionBienInsumo.MiComboBox = cmbIdClasificacionBienInsumo
End Function



Private Sub txtIdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdProducto
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdBienesInsumos.Width = fraBusqueda.Width
   grdBienesInsumos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Sub ConfigurarTiposDeBienesEInsumos()
    
'    mo_cmbIdClasificacionBienInsumo.BoundColumn = "IdTipoBienInsumo"
'    mo_cmbIdClasificacionBienInsumo.ListField = "Descripcion"
'    Set mo_cmbIdClasificacionBienInsumo.RowSource = mo_AdminComunes.TiposDeBienesEInsumosSeleccionarTodos()
    
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

VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFarmAlmacenes 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   ScaleHeight     =   6210
   ScaleWidth      =   9960
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   5550
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   9790
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
      Caption         =   "Lista de Farmacias"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00000000&
      Caption         =   "Farmacias"
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
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucFarmAlmacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Almacenes
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdLista.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdLista.DataSource
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
Property Let TipoFiltro(lValue As sghTipoFiltroPacientes)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroPacientes
    TipoFiltro = ml_TipoFiltro
End Property

Public Sub RealizarBusqueda()
    Set grdLista.DataSource = mo_ReglasFarmacia.FarmAlmacenConsultaXtipoFarmacia
    If mo_ReglasFarmacia.MensajeError <> "" Then
        MsgBox mo_ReglasFarmacia.MensajeError, vbInformation, "Lista Inventarios"
    End If
    'mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub grdLista_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idAlmacen")
End Sub

Private Sub grdLista_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idAlmacen")
End Sub


Private Sub grdLista_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdLista.Bands(0).Columns("idTipoLocales").Hidden = True
    grdLista.Bands(0).Columns("idEstado").Hidden = True
    grdLista.Bands(0).Columns("idAlmacen").Header.Caption = "Código"
    grdLista.Bands(0).Columns("idAlmacen").Width = 1000
    grdLista.Bands(0).Columns("Descripcion").Header.Caption = "Almacén"
    grdLista.Bands(0).Columns("Descripcion").Width = 8000
    grdLista.Bands(0).Columns("Estado").Header.Caption = "Estado"
    grdLista.Bands(0).Columns("Estado").Width = 1500
End Sub
Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        mo_Apariencia.ConfigurarFilasBiColores grdLista, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
    
Public Function inicializar()
    SkinConfigura
    RealizarBusqueda
End Function

Private Sub grdLista_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstado").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
            'Row.Appearance.Font.Strikethrough = True
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
     Case vbKeyF7
     Case vbKeyF8
    End Select
       
End Sub
Private Sub UserControl_Resize()
   On Error Resume Next
   lblNombre.Width = UserControl.Width
   
   grdLista.Width = UserControl.Width
   grdLista.Height = UserControl.Height - (lblNombre.Height + 150)
   
End Sub







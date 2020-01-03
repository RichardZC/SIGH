VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucListaProductos 
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
      TabIndex        =   3
      Top             =   540
      Width           =   10005
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5310
         Picture         =   "ucListaProductos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   6675
         Picture         =   "ucListaProductos.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   6
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
         TabIndex        =   4
         Top             =   240
         Width           =   3795
      End
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   4050
      Left            =   75
      TabIndex        =   2
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
      Caption         =   "Lista de Productos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Productos "
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
      Width           =   10080
   End
End
Attribute VB_Name = "ucListaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para buscar Items
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event OnClick(IdSeleccionado As Long, lcCodigoSeleccionado As String, lcNombreSeleccionado As String)
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_CodigoSeleccionado As String
Dim ml_NombreSeleccionado As String
Dim lb_MuestraTodosItems As Boolean

Property Let MuestraTodosItems(lValue As Boolean)
    lb_MuestraTodosItems = lValue
End Property

Property Let NombreSeleccionado(lValue As String)
    ml_NombreSeleccionado = lValue
End Property
Property Get NombreSeleccionado() As String
    NombreSeleccionado = ml_NombreSeleccionado
End Property
Property Let CodigoSeleccionado(lValue As String)
    ml_CodigoSeleccionado = lValue
End Property
Property Get CodigoSeleccionado() As String
    CodigoSeleccionado = ml_CodigoSeleccionado
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdProductos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdProductos.DataSource
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

Public Sub RealizarBusqueda()
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
        If UserControl.txtCodigo.Text <> "" Or UserControl.txtDescripcion.Text <> "" Then
            If lb_MuestraTodosItems = True Then
                Set grdProductos.DataSource = mo_ReglasFarmacia.FactCatalogoBienesInsumosSeleccionarXDescripYcodigo(UserControl.txtCodigo.Text, UserControl.txtDescripcion.Text)
            Else
                If UserControl.txtDescripcion <> "" Then
                   Set grdProductos.DataSource = mo_ReglasCaja.BienesFiltrarParaCajero(Trim(UserControl.txtDescripcion), 1, 5)
                Else
                   Set grdProductos.DataSource = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(Trim(UserControl.txtCodigo), 1, 5)
                End If
            End If
            mo_Apariencia.ConfigurarFilasBiColores grdProductos, sighentidades.GrillaConFilasBicolor
        End If
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtCodigo = ""
        UserControl.txtDescripcion = ""
End Sub

Private Sub grdProductos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProductos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
 
End Sub



Private Sub grdProductos_Click()
    Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProductos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
    ml_CodigoSeleccionado = rsRecordset("codigo")
    If UserControl.txtDescripcion <> "" Then
       ml_NombreSeleccionado = rsRecordset("nombre")
    Else
       ml_NombreSeleccionado = rsRecordset("NombreProducto")
    End If
End Sub

Private Sub grdProductos_DblClick()
     Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProductos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
    ml_CodigoSeleccionado = rsRecordset("codigo")
    If UserControl.txtDescripcion <> "" Then
       ml_NombreSeleccionado = rsRecordset("nombre")
    Else
       ml_NombreSeleccionado = rsRecordset("NombreProducto")
    End If
    RaiseEvent OnClick(ml_IdRegistroSeleccionado, ml_CodigoSeleccionado, ml_NombreSeleccionado)
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProductos.Bands(0).Columns("IdProducto").Hidden = True
    grdProductos.Bands(0).Columns("PrecioUnitario").Hidden = True
    grdProductos.Bands(0).Columns("Activo").Hidden = True
    
    grdProductos.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdProductos.Bands(0).Columns("Codigo").Width = 1000
    If UserControl.txtDescripcion <> "" Then
        grdProductos.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
        grdProductos.Bands(0).Columns("Nombre").Width = 11700
    Else
        grdProductos.Bands(0).Columns("NombreProducto").Header.Caption = "Descripción"
        grdProductos.Bands(0).Columns("NombreProducto").Width = 11700
    End If
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdProductos_DblClick
    End If
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



Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdProductos.Width = fraBusqueda.Width
   grdProductos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
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

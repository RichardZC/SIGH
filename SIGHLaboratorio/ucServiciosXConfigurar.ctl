VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucServiciosXConfigurar 
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13755
   ScaleHeight     =   6915
   ScaleWidth      =   13755
   Begin VB.Frame fraResultado 
      Height          =   5460
      Left            =   30
      TabIndex        =   5
      Top             =   1440
      Width           =   13665
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   5280
         Left            =   15
         TabIndex        =   7
         Top             =   105
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   9313
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
         Caption         =   "Lista de Servicios"
      End
   End
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
      Left            =   15
      TabIndex        =   0
      Top             =   510
      Width           =   13695
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9345
         Picture         =   "ucServiciosXConfigurar.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtIdServicio 
         Height          =   315
         Left            =   195
         TabIndex        =   3
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Top             =   450
         Width           =   6315
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7980
         Picture         =   "ucServiciosXConfigurar.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Código              Nombre                                                          "
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
         TabIndex        =   4
         Top             =   240
         Width           =   7635
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00000000&
      Caption         =   " Búsqueda de servicios"
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
      Left            =   45
      TabIndex        =   6
      Top             =   15
      Width           =   13665
   End
End
Attribute VB_Name = "ucServiciosXConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Procedimientos sin Resultados
'        Programado por: Madrid S
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
Dim mo_AdminServicios As New SIGHNegocios.ReglasConfiguarcionReslab

Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado

Dim ml_IdTipoCatalogo As Long

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


Private Sub btnBuscar_Click()
Dim oDOCatalogoServicio As New DOCatalogoServicio
        
        oDOCatalogoServicio.Codigo = UserControl.txtIdServicio
        oDOCatalogoServicio.Nombre = UserControl.txtNombre
        
        Set grdServicios.DataSource = mo_AdminServicios.FiltrarCatalogoSC(oDOCatalogoServicio)
        ConfigurarGrilla ml_IdTipoCatalogo = 0
        If mo_AdminServicios.MensajeError <> "" Then
            MsgBox mo_AdminServicios.MensajeError, vbInformation, "Filtro Servicios"
        End If
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
        
End Sub

Private Sub btnLimpiar_Click()
    UserControl.txtIdServicio = ""
    UserControl.txtNombre = ""
End Sub

Private Sub grdServicios_Click()
Dim rsRecordset As ADODB.Recordset

'    Set rsRecordset = grdServicios.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
    
End Sub

Private Sub grdServicios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdServicios.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
    
End Sub
Private Sub grdServicios_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
        ml_IdRegistroSeleccionado = Row.Cells(1).Value
End Sub
Private Sub grdServicios_DblClick()
     grdServicios_Click
     RaiseEvent SeleccionaRegistro(ml_IdRegistroSeleccionado)
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    ConfigurarGrilla False
End Sub

Sub ConfigurarGrilla(lCatalogoBase As Boolean)
    Dim lnFilaProductos As Integer
    If ml_IdTipoCatalogo = 0 Then
        lnFilaProductos = 1
        grdServicios.Bands(0).Columns("IdServicioSubGrupo").Hidden = True
        grdServicios.Bands(0).Columns("IdProducto").Hidden = True
        grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
        grdServicios.Bands(0).Columns("Codigo").Width = 1000
        grdServicios.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
        grdServicios.Bands(0).Columns("Descripcion").Header.Caption = "Nombre"
        grdServicios.Bands(0).Columns("Descripcion").Width = 7000
        grdServicios.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
    Else
        lnFilaProductos = 0
    End If
    grdServicios.Bands(lnFilaProductos).Columns("IdServicioSubGrupo").Hidden = True
    grdServicios.Bands(lnFilaProductos).Columns("IdProducto").Hidden = True
    grdServicios.Bands(lnFilaProductos).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(lnFilaProductos).Columns("Codigo").Width = 1000
    grdServicios.Bands(lnFilaProductos).Columns("Codigo").Activation = ssActivationActivateNoEdit
    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Width = 7000
    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Activation = ssActivationActivateNoEdit
    grdServicios.Bands(lnFilaProductos).Columns("NombreMInsa").Header.Caption = "Nombre Minsa"
    grdServicios.Bands(lnFilaProductos).Columns("NombreMinsa").Width = 7000
    grdServicios.Bands(lnFilaProductos).Columns("nombreMinsa").Activation = ssActivationActivateNoEdit
    grdServicios.Bands(lnFilaProductos).Columns("Descripcion").Hidden = True

    grdServicios.Bands(0).CollapseAll
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


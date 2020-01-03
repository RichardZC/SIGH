VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucProcedimientosLista 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10050
   LockControls    =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   10050
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
      Left            =   60
      TabIndex        =   6
      Top             =   510
      Width           =   9975
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   6675
         Picture         =   "ucProcedimientosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5310
         Picture         =   "ucProcedimientosLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1305
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
         MaxLength       =   255
         TabIndex        =   1
         Top             =   480
         Width           =   3915
      End
      Begin VB.Label Label2 
         Caption         =   "    Código                                 Descripción"
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
         TabIndex        =   7
         Top             =   240
         Width           =   4500
      End
   End
   Begin UltraGrid.SSUltraGrid grdProcedimientos 
      Height          =   4140
      Left            =   60
      TabIndex        =   4
      Top             =   1485
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   7303
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
      Caption         =   "Lista de procedimientos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Procedimientos (CPT2000)"
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
      Top             =   0
      Width           =   10020
   End
End
Attribute VB_Name = "ucProcedimientosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar procedimientos cPT
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim ml_idRegistroSeleccionado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mi_IdDiferenciacion As Integer 'WCG20060322
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdProcedimientos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdProcedimientos.DataSource
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
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
Dim oDOProcedimiento As New DOProcedimiento

        oDOProcedimiento.CodigoCPT2004 = UserControl.txtCodigo
        oDOProcedimiento.Descripcion = UserControl.txtDescripcion
        oDOProcedimiento.IdDiferenciacion = mi_IdDiferenciacion
        Set grdProcedimientos.DataSource = mo_AdminServiciosComunes.ProcedimientosFiltrar(oDOProcedimiento)
        mo_Apariencia.ConfigurarFilasBiColores grdProcedimientos, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtCodigo = ""
        UserControl.txtDescripcion = ""
End Sub

Private Sub grdProcedimientos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProcedimientos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdProcedimiento")
End Sub

Private Sub grdProcedimientos_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdProcedimientos.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdProcedimiento")
    
End Sub



Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProcedimientos.Bands(0).Columns("IdProcedimiento").Hidden = True
    
    grdProcedimientos.Bands(0).Columns("CodigoCPT2004").Header.Caption = "CPT"
    grdProcedimientos.Bands(0).Columns("CodigoCPT2004").Width = 1000
    
    grdProcedimientos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdProcedimientos.Bands(0).Columns("Descripcion").Width = 9000
    

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdProcedimientos.Width = fraBusqueda.Width
   grdProcedimientos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub


Public Property Get IdDiferenciacion() As Integer
    IdDiferenciacion = mi_IdDiferenciacion
End Property

Public Property Let IdDiferenciacion(ByVal iValue As Integer)
    mi_IdDiferenciacion = iValue
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
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub

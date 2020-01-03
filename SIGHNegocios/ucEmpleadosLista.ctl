VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucEmpleadosLista 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   ScaleHeight     =   5850
   ScaleWidth      =   10065
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
      TabIndex        =   7
      Top             =   510
      Width           =   9975
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   6870
         Picture         =   "ucEmpleadosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8235
         Picture         =   "ucEmpleadosLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox txtCodigoPlanilla 
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
         MaxLength       =   8
         TabIndex        =   0
         Top             =   465
         Width           =   1065
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   1290
         TabIndex        =   1
         Top             =   465
         Width           =   1785
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   3150
         TabIndex        =   2
         Top             =   465
         Width           =   1770
      End
      Begin VB.TextBox txtNombres 
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
         Left            =   4965
         TabIndex        =   3
         Top             =   465
         Width           =   1845
      End
      Begin VB.Label Label2 
         Caption         =   "Cod. planilla      Apellido paterno         Apellido materno             Nombres                            "
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
         TabIndex        =   8
         Top             =   240
         Width           =   7335
      End
   End
   Begin UltraGrid.SSUltraGrid grdEmpleados 
      Height          =   4350
      Left            =   75
      TabIndex        =   6
      Top             =   1470
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7673
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
      Caption         =   "Relación de empleados"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Empleados"
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
      TabIndex        =   9
      Top             =   15
      Width           =   10110
   End
End
Attribute VB_Name = "ucEmpleadosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Empleados
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdEmpleados.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdEmpleados.DataSource
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
Dim oDOEmpleado As New DOEmpleado

        oDOEmpleado.ApellidoPaterno = Trim(UserControl.txtApellidoPaterno)
        oDOEmpleado.ApellidoMaterno = Trim(UserControl.txtApellidoMaterno)
        oDOEmpleado.Nombres = Trim(UserControl.txtNombres)
        oDOEmpleado.CodigoPlanilla = UserControl.txtCodigoPlanilla
        
        Set grdEmpleados.DataSource = mo_AdminServiciosComunes.EmpleadosFiltrar(oDOEmpleado)

        mo_Apariencia.ConfigurarFilasBiColores grdEmpleados, sighentidades.GrillaConFilasBicolor
        'mo_Apariencia.ConfigurarFilasAlphaBlending grdEmpleados, "D:\SIGH_VB6\Imagenes\Logo.jpg"


End Sub
Private Sub Set_Appearance_UseAlpha(App As SSAppearance, AlphaLevel As Long, Use As Constants_Alpha)
    With App
        .AlphaLevel = AlphaLevel
        .BackColorAlpha = Use
        .BorderAlpha = Use
        .ForegroundAlpha = Use
        .PictureAlpha = Use
        .PictureBackgroundAlpha = Use
    End With
    
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNombres = ""
        UserControl.txtCodigoPlanilla.Text = ""
End Sub

Private Sub grdEmpleados_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdEmpleados.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdEmpleado")
End Sub

Private Sub grdEmpleados_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdEmpleados.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdEmpleado")
    
End Sub

Private Sub grdEmpleados_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_IdRegistroSeleccionado = -1
    Set rsRecordset = grdEmpleados.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdEmpleado")
    
End Sub


Private Sub grdEmpleados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdEmpleados.Bands(0).Columns("IdEmpleado").Hidden = True
    
    grdEmpleados.Bands(0).Columns("CodigoPlanilla").Header.Caption = "Cod. Planilla"
    grdEmpleados.Bands(0).Columns("CodigoPlanilla").Width = 750
    
    grdEmpleados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Apellido Paterno"
    grdEmpleados.Bands(0).Columns("ApellidoPaterno").Width = 2000
    
    grdEmpleados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Apellido Materno"
    grdEmpleados.Bands(0).Columns("ApellidoMaterno").Width = 2000
    
    grdEmpleados.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdEmpleados.Bands(0).Columns("Nombres").Width = 2500

    grdEmpleados.Bands(0).Columns("TipoEmpleado").Header.Caption = "Tipo empleado"
    grdEmpleados.Bands(0).Columns("TipoEmpleado").Width = 2500

    grdEmpleados.Bands(0).Columns("CondicionTrabajo").Header.Caption = "Condición trabajo"
    grdEmpleados.Bands(0).Columns("CondicionTrabajo").Width = 2500

End Sub

Private Sub txtCodigoPlanilla_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoPlanilla
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodigoPlanilla_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombres
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombres_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 150
   lblNombre.Width = UserControl.Width
   
   grdEmpleados.Width = fraBusqueda.Width
   grdEmpleados.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
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

VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form ImagEquipos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14325
   Icon            =   "ImagEquipos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNuevo 
      Caption         =   "Nuevo Equipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   10635
      TabIndex        =   10
      Top             =   495
      Width           =   3690
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
         Height          =   345
         Left            =   810
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   2805
      End
      Begin VB.TextBox txtRuta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   855
         TabIndex        =   5
         Top             =   3375
         Width           =   2805
      End
      Begin VB.TextBox txtTipo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   840
         TabIndex        =   4
         Top             =   2430
         Width           =   2805
      End
      Begin VB.TextBox txtModelo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   840
         TabIndex        =   3
         Top             =   1545
         Width           =   2805
      End
      Begin VB.TextBox txtMarca 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   825
         TabIndex        =   2
         Top             =   615
         Width           =   2805
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ImagEquipos.frx":0CCA
         DownPicture     =   "ImagEquipos.frx":118E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2250
         Picture         =   "ImagEquipos.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4365
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Grabar"
         DisabledPicture =   "ImagEquipos.frx":1B66
         DownPicture     =   "ImagEquipos.frx":1FC6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         Picture         =   "ImagEquipos.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4380
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   16
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   3390
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   14
         Top             =   2475
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   1530
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   630
         Width           =   465
      End
   End
   Begin VB.CommandButton btnQuitar 
      DisabledPicture =   "ImagEquipos.frx":28B0
      DownPicture     =   "ImagEquipos.frx":2C3B
      Height          =   480
      Left            =   10125
      Picture         =   "ImagEquipos.frx":2FCE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar un Item"
      Top             =   1020
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton btnAgregar 
      DisabledPicture =   "ImagEquipos.frx":335F
      DownPicture     =   "ImagEquipos.frx":3748
      Height          =   480
      Left            =   10125
      Picture         =   "ImagEquipos.frx":3B54
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Agregar un Item"
      Top             =   540
      Width           =   450
   End
   Begin UltraGrid.SSUltraGrid grdEquipos 
      Height          =   5100
      Left            =   30
      TabIndex        =   7
      Top             =   540
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   8996
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
      Caption         =   "Lista de Equipos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Equipos"
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
      TabIndex        =   8
      Top             =   0
      Width           =   14325
   End
End
Attribute VB_Name = "ImagEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca procedimientos
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim oInteroperaEquipos As New InteroperaEquipos
Dim oDoInteroperaEquipos As New DoInteroperaEquipos
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes

Private Sub btnAceptar_Click()
   If ValidaDatosObligatorios Then
        Dim oConexion As New Connection
        With oDoInteroperaEquipos
            .codigo = Me.txtCodigo.Text
            .IdUsuarioAuditoria = sighentidades.Usuario
            .marca = Me.txtMarca.Text
            .modelo = Me.txtModelo.Text
            .ruta = Me.txtRuta.Text
            .tipo = Me.txtTipo.Text
        End With
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        Set oInteroperaEquipos.Conexion = oConexion
        If oInteroperaEquipos.Insertar(oDoInteroperaEquipos) = False Then
           MsgBox oInteroperaEquipos.MensajeError
        End If
        oConexion.Close
        Set oConexion = Nothing
        If oInteroperaEquipos.MensajeError = "" Then
            mo_Formulario.HabilitarDeshabilitar Me.fraNuevo, False
            CargaEquipos
            Limpiar
        End If
   End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    ValidaDatosObligatorios = False
    If txtCodigo.Text = "" Then
       MsgBox "Debe ingresar el CODIGO ", vbInformation, ""
       Exit Function
    End If
    
    If txtMarca.Text = "" Then
       MsgBox "Debe ingresar la MARCA ", vbInformation, ""
       Exit Function
    End If
    If txtModelo.Text = "" Then
       MsgBox "Debe ingresar el MODELO ", vbInformation, ""
       Exit Function
    End If
    If txtTipo.Text = "" Then
       MsgBox "Debe ingresar el TIPO ", vbInformation, ""
       Exit Function
    End If
    If txtRuta.Text = "" Then
       MsgBox "Debe ingresar la RUTA", vbInformation, ""
       Exit Function
    End If
    
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = Me.grdEquipos.DataSource
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       oRsTmp1.Find "codigo='" & txtCodigo.Text & "'"
       If Not oRsTmp1.EOF Then
            MsgBox "Ese CODIGO ya existe", vbInformation, ""
            Exit Function
       End If
    End If
    Set oRsTmp1 = Nothing
    
    ValidaDatosObligatorios = True
End Function

Private Sub btnAgregar_Click()
    Limpiar
    mo_Formulario.HabilitarDeshabilitar Me.fraNuevo, True
    On Error Resume Next
    txtCodigo.SetFocus
End Sub


Sub Limpiar()
    txtCodigo.Text = ""
    txtMarca.Text = ""
    txtModelo.Text = ""
    txtTipo.Text = ""
    txtRuta.Text = ""
End Sub

Private Sub btnCancelar_Click()
    mo_Formulario.HabilitarDeshabilitar Me.fraNuevo, False
End Sub

Private Sub Form_Load()
    CargaEquipos
    mo_Formulario.HabilitarDeshabilitar Me.fraNuevo, False
End Sub

Sub CargaEquipos()
    Set grdEquipos.DataSource = mo_AdminComun.InteroperaEquiposSeleccionarTodos()
    mo_Apariencia.ConfigurarFilasBiColores grdEquipos, sighentidades.GrillaConFilasBicolor
End Sub



Private Sub grdEquipos_DblClick()
    On Error Resume Next
    mo_Formulario.HabilitarDeshabilitar Me.fraNuevo, False
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = Me.grdEquipos.DataSource
    txtCodigo.Text = oRsTmp1!codigo
    txtMarca.Text = oRsTmp1!marca
    txtModelo.Text = oRsTmp1!modelo
    txtTipo.Text = oRsTmp1!tipo
    txtRuta.Text = oRsTmp1!ruta
    
End Sub

Private Sub grdEquipos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdEquipos.Bands(0).Columns("codigo").Width = 500
    grdEquipos.Bands(0).Columns("ruta").Width = 4000
    grdEquipos.Bands(0).Columns("Equipo").Hidden = True

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtMarca
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtModelo
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtRuta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtRuta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtTipo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtTipo
    AdministrarKeyPreview KeyCode
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           'btnCancelar_Click
       Case vbKeyF2
           'btnAceptar_Click
       End Select
End Sub

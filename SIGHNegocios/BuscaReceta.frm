VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form BuscaReceta 
   Caption         =   "Busqueda de Receta"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "BuscaReceta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
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
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   10650
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7170
         Picture         =   "BuscaReceta.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8535
         Picture         =   "BuscaReceta.frx":2C55
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtApaterno 
         Height          =   315
         Left            =   180
         MaxLength       =   40
         TabIndex        =   5
         Top             =   480
         Width           =   1395
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1650
         TabIndex        =   9
         Top             =   480
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   3570
         TabIndex        =   10
         Top             =   480
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Paterno                      Fecha de Recetas"
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
         Top             =   270
         Width           =   5295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   0
      Top             =   7200
      Width           =   10665
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BuscaReceta.frx":5831
         DownPicture     =   "BuscaReceta.frx":5CF5
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5430
         Picture         =   "BuscaReceta.frx":61E1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BuscaReceta.frx":66CD
         DownPicture     =   "BuscaReceta.frx":6B2D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3900
         Picture         =   "BuscaReceta.frx":6FA2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   6225
      Left            =   30
      TabIndex        =   3
      Top             =   900
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   10980
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
      Caption         =   ".."
   End
End
Attribute VB_Name = "BuscaReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Receta
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim oRsLista As New Recordset
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim lcSql As String
Dim ml_idReceta As Long
Dim ml_IdPuntoCarga As sghPuntosCargaBasicos
Dim ml_idCuentaAtencion As Long

Property Let IdPuntoCarga(lValue As sghPuntosCargaBasicos)
    ml_IdPuntoCarga = lValue
End Property
Property Get IdPuntoCarga() As Long
    IdPuntoCarga = ml_IdPuntoCarga
End Property

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
    If ml_idCuentaAtencion > 0 Then
       Me.txtFdesde.Text = Format(Date - 60, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
       'Me.txtFhasta.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
       btnBuscar_Click
    End If
End Property


Property Let idReceta(lValue As Long)
    ml_idReceta = lValue
End Property
Property Get idReceta() As Long
    idReceta = ml_idReceta
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

'debb-18/05/2016
Private Sub btnBuscar_Click()
    If CDate(Me.txtFdesde.Text) > CDate(Me.txtFhasta.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
    End If
     
     Set oRsLista = mo_reglasComunes.RecetaCabeceraPorRangoFechas(CDate(Me.txtFdesde.Text), _
                                      CDate(Me.txtFhasta.Text), Trim(UCase(Me.txtApaterno.Text)), IIf(ml_idCuentaAtencion = 0, 1, 0))
     If ml_IdPuntoCarga = sghPtoCargaCaja Then
        oRsLista.Filter = "IdFormaPago = 1 AND IdPuntoCarga <> 5"
     ElseIf ml_idCuentaAtencion > 0 Then  'carga todas las recetas (farmacia, Imágenes, Laboratorio) de una CUENTA DEL Paciente
        oRsLista.Filter = "idCuentaAtencion=" & ml_idCuentaAtencion
     Else
        oRsLista.Filter = "idPuntoCarga=" & ml_IdPuntoCarga
     End If
     Set grdLista.DataSource = oRsLista
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False

End Sub

Private Sub btnLimpiar_Click()
    txtFdesde.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY) & " 00:01"
    txtFhasta.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY) & " 23:59"
    Me.txtApaterno.Text = ""
End Sub

Private Sub Form_Load()
    btnLimpiar_Click
    btnBuscar_Click
    mo_Apariencia.ConfigurarFilasBiColores Me.grdLista, sighentidades.GrillaConFilasBicolor
    
End Sub


Private Sub grdLista_Click()
    If Not (oRsLista.BOF = True And oRsLista.EOF = True) Then
        ml_idReceta = oRsLista.Fields!idReceta
        ml_IdPuntoCarga = oRsLista!IdPuntoCarga
    End If
End Sub

Private Sub grdLista_DblClick()
    If Not (oRsLista.BOF = True And oRsLista.EOF = True) Then
        ml_idReceta = oRsLista.Fields!idReceta
        ml_IdPuntoCarga = oRsLista!IdPuntoCarga
        btnAceptar_Click
    End If
End Sub

Private Sub grdLista_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdLista_DblClick
    End If
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            btnBuscar_Click
        Case vbKeyF7
            btnLimpiar_Click
        Case vbKeyEscape
           
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub txtApaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApaterno
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFdesde_LostFocus()
If Not IsDate(txtFdesde.Text) Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde.Text = sighentidades.FECHA_VACIA_DMY_HM
        End If
End Sub

Private Sub txtFhasta_LostFocus()
If Not IsDate(txtFhasta.Text) Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFhasta.Text = sighentidades.FECHA_VACIA_DMY_HM
        End If

End Sub

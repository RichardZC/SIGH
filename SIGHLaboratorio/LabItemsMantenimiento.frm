VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form LabItemsMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "LabItemsMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
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
      Left            =   60
      TabIndex        =   2
      Top             =   510
      Width           =   5835
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   4485
         Picture         =   "LabItemsMantenimiento.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   165
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4485
         Picture         =   "LabItemsMantenimiento.frx":38A6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   495
         Width           =   1305
      End
      Begin VB.TextBox txtGrupo 
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
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtIdgrupo 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   " Código           Nombre de Items"
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
         Left            =   135
         TabIndex        =   7
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton btnAgregarDx 
      DisabledPicture =   "LabItemsMantenimiento.frx":64EF
      DownPicture     =   "LabItemsMantenimiento.frx":68D8
      Height          =   480
      Left            =   5400
      Picture         =   "LabItemsMantenimiento.frx":6CE4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Agregar un Item"
      Top             =   1455
      Width           =   450
   End
   Begin VB.CommandButton btnQuitarDx 
      DisabledPicture =   "LabItemsMantenimiento.frx":70F0
      DownPicture     =   "LabItemsMantenimiento.frx":747B
      Height          =   480
      Left            =   5400
      Picture         =   "LabItemsMantenimiento.frx":780E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Eliminar un Item"
      Top             =   1935
      Width           =   450
   End
   Begin UltraGrid.SSUltraGrid grdGrupos 
      Height          =   4290
      Left            =   75
      TabIndex        =   8
      Top             =   1455
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   7567
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
      Caption         =   "Items Disponibles"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Items de Pruebas de Laboratorio"
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
      TabIndex        =   9
      Top             =   0
      Width           =   6165
   End
End
Attribute VB_Name = "LabItemsMantenimiento"
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
Dim mo_AdminComun As New ReglasConfiguarcionReslab
Dim mo_Labitems As New SIGHDatos.LabItems
Dim oDoLabItems As New SIGHComun.DoLabItems
Dim rst As New ADODB.Recordset


Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_IdRegistroSeleccionado As Long

Dim conta As Integer
Dim lb_Switch As Boolean


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = rst.Fields!idItem
End Property

Private Sub btnAgregarDx_Click()
    Dim lb_check As Boolean
    Dim nitem As String
    
    lb_check = False
    nitem = InputBox("ingrese el nuevo Item", "INGRESO DE DATOS")
    If nitem = "" Then Exit Sub
    rst.MoveFirst
    Do While (Not rst.EOF)
        If rst.Fields!Item = nitem Then
            lb_check = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    If lb_check Then
        MsgBox "El item ya esta registrado", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    Else
        oDoLabItems.IdUsuarioAuditoria = ml_idUsuario
        oDoLabItems.idItem = mo_AdminComun.LabMayor("LabItems", "IdItem")
        oDoLabItems.Item = nitem
        oDoLabItems.idProductoCPT = 0
        
        If mo_AdminComun.LabItemAgregar(oDoLabItems, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
            Me.IdRegistroSeleccionado = oDoLabItems.idItem
            Set rst = mo_AdminComun.LabItemsSeleccionarTodos("")
            Set grdGrupos.DataSource = rst
            configuraGrilla
        End If
    End If
    
End Sub

Private Sub btnBuscar_Click()
    Dim filtro As String
    
    filtro = ""
    If txtIdgrupo.Text <> "" Then filtro = "idItem='" & txtIdgrupo & "'"
    If txtGrupo <> "" Then filtro = "item Like '%" & txtGrupo & "%'"
    
    Set rst = mo_AdminComun.LabItemsSeleccionarTodos(filtro)
    Set grdGrupos.DataSource = rst
    configuraGrilla
End Sub

Private Sub btnLimpiar_Click()
    txtGrupo.Text = ""
    txtIdgrupo.Text = ""
    Set rst = mo_AdminComun.LabItemsSeleccionarTodos("")
    Set grdGrupos.DataSource = rst
End Sub

Private Sub btnQuitarDx_Click()
    If mo_AdminComun.ValidarReglasItem(rst.Fields!idItem) Then
        oDoLabItems.IdUsuarioAuditoria = ml_idUsuario
        oDoLabItems.idItem = rst.Fields!idItem
        oDoLabItems.Item = rst.Fields!Item
        If mo_AdminComun.LabItemEliminar(oDoLabItems, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
            Set rst = mo_AdminComun.LabItemsSeleccionarTodos("")
            Set grdGrupos.DataSource = rst
            configuraGrilla
        End If
    Else
        MsgBox "Este Item esta en uso, no puede ser eliminado", vbOKOnly + vbCritical, Me.Caption
    End If
        
End Sub

Private Sub Form_Load()
    Set rst = mo_AdminComun.LabItemsSeleccionarTodos("")
    Set grdGrupos.DataSource = rst
    configuraGrilla
    mo_Apariencia.ConfigurarFilasBiColores grdGrupos, sighentidades.GrillaConFilasBicolor
    
End Sub

Private Sub configuraGrilla()
    grdGrupos.Bands(0).Columns("IdItem").Header.Caption = "Còdigo"
    grdGrupos.Bands(0).Columns("Item").Header.Caption = "Grupo de Items"
    grdGrupos.Bands(0).Columns("idItem").Width = 900
    grdGrupos.Bands(0).Columns("Item").Width = 3800
    grdGrupos.Bands(0).Columns("IdProductoCPT").Hidden = True
End Sub

Private Sub grdGrupos_DblClick()
    Me.IdRegistroSeleccionado = rst.Fields!idItem
    Me.Hide
End Sub


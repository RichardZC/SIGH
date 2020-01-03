VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form LabItemsGrupoMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "LabItemsGrupoMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnQuitarDx 
      DisabledPicture =   "LabItemsGrupoMantenimiento.frx":0CCA
      DownPicture     =   "LabItemsGrupoMantenimiento.frx":1055
      Height          =   480
      Left            =   5400
      Picture         =   "LabItemsGrupoMantenimiento.frx":13E8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar un Item"
      Top             =   1935
      Width           =   450
   End
   Begin VB.CommandButton btnAgregarDx 
      DisabledPicture =   "LabItemsGrupoMantenimiento.frx":1779
      DownPicture     =   "LabItemsGrupoMantenimiento.frx":1B62
      Height          =   480
      Left            =   5400
      Picture         =   "LabItemsGrupoMantenimiento.frx":1F6E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Agregar un Item"
      Top             =   1455
      Width           =   450
   End
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
      TabIndex        =   0
      Top             =   510
      Width           =   5835
      Begin VB.TextBox txtIdgrupo 
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   480
         Width           =   1155
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
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4485
         Picture         =   "LabItemsGrupoMantenimiento.frx":237A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   495
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   4485
         Picture         =   "LabItemsGrupoMantenimiento.frx":4FC3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   165
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   " Código           Nombre de Grupo"
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
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
   End
   Begin UltraGrid.SSUltraGrid grdGrupos 
      Height          =   4290
      Left            =   75
      TabIndex        =   5
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
      Caption         =   "Grupos de Items Disponibles"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   " Grupos de Items de laboratorios"
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
      TabIndex        =   6
      Top             =   0
      Width           =   6165
   End
End
Attribute VB_Name = "LabItemsGrupoMantenimiento"
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
Dim mo_Labitemsgrupo As New SIGHDatos.LabItemsGrupos
Dim oDoLabItemsGrupos As New SIGHComun.DoLabItemsGrupos
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
    IdRegistroSeleccionado = rst.Fields!idItemGrupo
End Property

Private Sub btnAgregarDx_Click()
    Dim lb_check As Boolean
    Dim nitem As String
    
    lb_check = False
    nitem = InputBox("ingrese el nuevo Item", "INGRESO DE DATOS")
    If nitem = "" Then Exit Sub
    rst.MoveFirst
    Do While (Not rst.EOF)
        If rst.Fields!Grupo = nitem Then
            lb_check = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    If lb_check Then
        MsgBox "El item ya esta registrado", vbInformation + vbInformation, "INGRESO DE DATOS"
        Exit Sub
    Else
        oDoLabItemsGrupos.IdUsuarioAuditoria = ml_idUsuario
        oDoLabItemsGrupos.idItemGrupo = mo_AdminComun.LabMayor("LabItemsGrupos", "IdItemGrupo")
        oDoLabItemsGrupos.Grupo = nitem
        If mo_AdminComun.LabItemGrupoAgregar(oDoLabItemsGrupos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
            Me.IdRegistroSeleccionado = oDoLabItemsGrupos.idItemGrupo
            Set rst = mo_AdminComun.LabItemsGruposSeleccionarTodos("")
            Set grdGrupos.DataSource = rst
            configuraGrilla
        End If
    End If
    
End Sub

Private Sub btnLimpiar_Click()
    txtGrupo.Text = ""
    txtIdgrupo.Text = ""
    Set rst = mo_AdminComun.LabItemsGruposSeleccionarTodos("")
    Set grdGrupos.DataSource = rst
End Sub

Private Sub btnQuitarDx_Click() 'modificado 05/08 Samuel
    If mo_AdminComun.ValidarReglasGrupoItem(rst.Fields!idItemGrupo) Then
        oDoLabItemsGrupos.IdUsuarioAuditoria = ml_idUsuario
        oDoLabItemsGrupos.idItemGrupo = rst.Fields!idItemGrupo
        oDoLabItemsGrupos.Grupo = rst.Fields!Grupo
        If mo_AdminComun.LabItemGrupoEliminar(oDoLabItemsGrupos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
            Set rst = mo_AdminComun.LabItemsGruposSeleccionarTodos("")
            Set grdGrupos.DataSource = rst
            configuraGrilla
        End If
    Else
        MsgBox "Este Grupo de Items esta en uso, no puede ser eliminado", vbOKOnly + vbCritical, Me.Caption
    End If
End Sub

Private Sub Form_Load()
    Set rst = mo_AdminComun.LabItemsGruposSeleccionarTodos("")
    Set grdGrupos.DataSource = rst
    configuraGrilla
    mo_Apariencia.ConfigurarFilasBiColores grdGrupos, sighentidades.GrillaConFilasBicolor
    
End Sub

Private Sub configuraGrilla()
    grdGrupos.Bands(0).Columns("IdItemGrupo").Header.Caption = "Còdigo"
    grdGrupos.Bands(0).Columns("grupo").Header.Caption = "Grupo de Items"
    grdGrupos.Bands(0).Columns("idItemGrupo").Width = 1000
    grdGrupos.Bands(0).Columns("grupo").Width = 3800
End Sub

Private Sub grdGrupos_DblClick()
    Me.IdRegistroSeleccionado = rst.Fields!idItemGrupo
    Me.Hide
End Sub


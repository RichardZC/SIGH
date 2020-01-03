VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form mTablitas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas: LolCli"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "mTablitas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuarios SISGALENPLUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7620
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   10485
      Begin Threed.SSOption optOpciones 
         Height          =   300
         Left            =   5010
         TabIndex        =   5
         Top             =   6975
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   262144
         Caption         =   "Opciones"
      End
      Begin Threed.SSOption optApellidos 
         Height          =   300
         Left            =   2655
         TabIndex        =   4
         Top             =   6975
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   529
         _Version        =   262144
         Caption         =   "Apellidos y nombres"
      End
      Begin Threed.SSOption optRol 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   6975
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "Rol"
         Value           =   -1
      End
      Begin UltraGrid.SSUltraGrid SSUltraGrid 
         Height          =   6435
         Left            =   150
         TabIndex        =   1
         Top             =   345
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   11351
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Lista de Usuarios y opciones"
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por"
         Height          =   195
         Left            =   255
         TabIndex        =   2
         Top             =   7020
         Width           =   975
      End
   End
End
Attribute VB_Name = "mTablitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Tablas Lolcli
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Dim oRsTmp1 As New Recordset

Private Sub Form_Load()
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "UsuariosRolesSeleccionarTodasOpciones"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set SSUltraGrid.DataSource = oRsTmp1
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
End Sub


Private Sub optApellidos_Click(Value As Integer)
   oRsTmp1.Sort = "apellidoPaterno,apellidoMaterno,nombres"
   SSUltraGrid.Refresh
End Sub

Private Sub optOpciones_Click(Value As Integer)
    oRsTmp1.Sort = "opcion"
    SSUltraGrid.Refresh
End Sub

Private Sub optRol_Click(Value As Integer)
    oRsTmp1.Sort = "rol"
    SSUltraGrid.Refresh
End Sub

Private Sub SSUltraGrid_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti

End Sub

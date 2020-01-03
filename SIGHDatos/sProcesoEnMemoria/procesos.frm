VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form Procesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes del Administrador"
   ClientHeight    =   9045
   ClientLeft      =   5460
   ClientTop       =   4365
   ClientWidth     =   13785
   ControlBox      =   0   'False
   Icon            =   "procesos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13785
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraCuposLibres 
      Height          =   7575
      Left            =   30
      TabIndex        =   30
      Top             =   270
      Width           =   13575
      Begin UltraGrid.SSUltraGrid grdCupos1 
         Height          =   7335
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   12938
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CONSULTORIOS EXTERNOS"
      End
      Begin UltraGrid.SSUltraGrid grdCupos2 
         Height          =   6855
         Left            =   6960
         TabIndex        =   33
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   12091
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CONSULTORIOS EXTERNOS"
      End
      Begin VB.Label lblTextoCabecera 
         Alignment       =   2  'Center
         Caption         =   "Cupos Hasta : XX/XX/XXXXX xx:xx:xx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   6960
         TabIndex        =   31
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Frame FraProcesos 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   375
      Begin MSDataGridLib.DataGrid grdHospitalizados 
         Height          =   7095
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   13305
         _ExtentX        =   23469
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   16
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cuentas que serán Anuladas en: Hospitalización"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "nroHistoria"
            Caption         =   "Nro Historia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Paciente"
            Caption         =   "Paciente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "NroCuenta"
            Caption         =   "Nro Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Servicio"
            Caption         =   "Servicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "FechaCierre"
            Caption         =   "Fecha Cierre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4004.788
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3750.236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2310.236
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1110
      Left            =   30
      TabIndex        =   25
      Top             =   7890
      Width           =   13710
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Salir"
         DisabledPicture =   "procesos.frx":030A
         DownPicture     =   "procesos.frx":07CE
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
         Left            =   6870
         Picture         =   "procesos.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Refrescar"
         DisabledPicture =   "procesos.frx":11A6
         DownPicture     =   "procesos.frx":1606
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
         Left            =   5310
         Picture         =   "procesos.frx":1A7B
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   405
      Left            =   8340
      TabIndex        =   1
      Top             =   -90
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox txtCabina 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   23
         Top             =   300
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Height          =   2175
         Left            =   180
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   5865
         Begin VB.TextBox txtMinCabina 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1815
            TabIndex        =   14
            Text            =   "0"
            Top             =   1110
            Width           =   1545
         End
         Begin VB.TextBox txtAviso 
            Height          =   345
            Left            =   1815
            TabIndex        =   16
            Text            =   "5"
            Top             =   1590
            Width           =   1545
         End
         Begin VB.TextBox txtClave 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1815
            PasswordChar    =   "*"
            TabIndex        =   13
            Top             =   675
            Width           =   1545
         End
         Begin VB.TextBox txtUsuario 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1815
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   225
            Width           =   1545
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Configura Cabina"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3900
            TabIndex        =   11
            Top             =   825
            Width           =   1785
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3900
            TabIndex        =   15
            Top             =   210
            Width           =   1785
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Minutos en Cabina:"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   1170
            Width           =   1365
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo aviso:"
            Height          =   195
            Left            =   135
            TabIndex        =   19
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   735
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   270
            Width           =   585
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reiniciar PC"
         Height          =   495
         Left            =   8415
         TabIndex        =   9
         Top             =   1680
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Apagar PC"
         Height          =   495
         Left            =   8415
         TabIndex        =   8
         Top             =   2220
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         Height          =   4185
         Left            =   210
         TabIndex        =   2
         Top             =   750
         Width           =   8040
         Begin VB.CommandButton Command7 
            BackColor       =   &H008080FF&
            Caption         =   "No deseo continuar usando la CABINA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   150
            MaskColor       =   &H8000000F&
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   3465
            Width           =   7755
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H008080FF&
            Caption         =   "Nuevo uso de CABINA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   120
            MaskColor       =   &H8000000F&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1695
            Width           =   7755
         End
         Begin VB.TextBox txtQuedan 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   675
            Left            =   3165
            TabIndex        =   3
            Text            =   "0"
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Minutos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   555
            Left            =   4680
            TabIndex        =   7
            Top             =   300
            Width           =   2010
         End
         Begin VB.Label lblCliente 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   105
            TabIndex        =   5
            Top             =   990
            Width           =   7770
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Te quedan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   555
            Left            =   120
            TabIndex        =   4
            Top             =   285
            Width           =   2685
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cabina Nº:"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1530
      Top             =   0
   End
   Begin VB.PictureBox picGancho 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      Caption         =   "v15102009"
      Height          =   195
      Left            =   60
      TabIndex        =   21
      Top             =   30
      Width           =   15135
   End
End
Attribute VB_Name = "Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Proceso residente en memoria
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lbProcesaReniecVSgalenhos As Boolean
Dim mo_Reniec As New ReniecGalenhos
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes  As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsHospitalizados As New Recordset
Dim lcSql As String
Dim lnCargoDesdeLoad As Integer: Dim lnNroIntentoCargarCabinas As Integer
Dim lnAnchoPantalla As Long: Dim lnLargoPantalla As Long
Dim lcClaveFecha As String
Dim lcMensajeError As String
Dim lcUsuario As String
Dim lnMinutosTranscurridos As Integer
Dim lbYaAviso As Boolean
Dim lcHorasCE As String: Dim lcHorasHosp As String
Dim lcHOrasEmergDiurno As String: Dim lcHOrasEmergNocturno As String
Const lcProblemasConReniec As String = "   <<Reniec>>"
Dim lbBuscaDNIenReniec As Boolean
Dim lbProcesaSisVSgalenhos As Boolean
'FRANK 05022014
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim oRsServiciosCuposLibre1 As New Recordset
Dim oRsServiciosCuposLibre2 As New Recordset
Dim LnTotalRegistrosGrilla1 As Integer
Const LnWidthFrame = 13575
Const LnTopFrame = 360
Const LnLeftFrame = 120
Const LnHeightFrame = 7455

'MARIO
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String

' EjemploBT ver1.0
' 1997 J.LeVasseur lvasseur@tiac.net a0@null.net
' Un ejemplo de Usar la barra de tareas en Win95/NT4
' El PictureBox picGancho sirve como gancho de los
' mensajes CallBack del API Shell_NotifyIcon. Tiene
' que ser un control con un hWnd. Todo lo interesante
' esta en el picGancho_MouseMove . Como pueden ver, un
' control MsgHook o MsgBlaster aqui sobra...
'---------------

Private Type TIPONOTIFICARICONO
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'------------------
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205


'--------------------
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    pnid As TIPONOTIFICARICONO) As Boolean
'--------------------
Private Declare Function WinExec& Lib "kernel32" _
    (ByVal lpCmdLine As String, ByVal nCmdShow As Long)
'--------------------
Dim t As TIPONOTIFICARICONO





'Minimizar ventanas abiertas---Declaración del Api keybd_event
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
                                    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Constantes
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B


'Para ocultar Barra de Tareas
Private Const SWP_HIDEWINDOW As Long = &H80&
Private Const SWP_SHOWWINDOW As Long = &H40&
'Api: busca el Handle del taskBar
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'SetWindowPos lo Oculta y lo reestablece
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long


'apagar PC
Private Type LUID
  UsedPart As Long
  IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_FORCE As Long = 4
Private Const EWX_REBOOT = 2
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

'MARIO
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
'MARIO
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
'MARIO
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Private Sub AdjustToken()
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    tkp.PrivilegeCount = 1
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub


Private Sub grdCupos1_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCupos1.Bands(0).Columns("IdServicio").Hidden = True
    grdCupos1.Bands(0).Columns("Servicio").Header.Caption = "Consultorios"
    grdCupos1.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdCupos1.Bands(0).Columns("CuposLibres").Header.Caption = "Cupos Libres"
    grdCupos1.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdCupos1.Bands(0).Columns("Turno").Activation = ssActivationActivateNoEdit
    grdCupos1.Bands(0).Columns("CuposLibres").Activation = ssActivationActivateNoEdit
'    grdCupos1.Bands(0).Columns("Servicio").Width = 3100
'    grdCupos1.Bands(0).Columns("Turno").Width = 1500
'    grdCupos1.Bands(0).Columns("CuposLibres").Width = 1500
    grdCupos1.Bands(0).Columns("CuposLibres").CellAppearance.TextAlign = ssAlignCenter
End Sub

Private Sub grdCupos2_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCupos2.Bands(0).Columns("IdServicio").Hidden = True
    grdCupos2.Bands(0).Columns("Servicio").Header.Caption = "Consultorios"
    grdCupos2.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdCupos2.Bands(0).Columns("CuposLibres").Header.Caption = "Cupos Libres"
    grdCupos2.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdCupos2.Bands(0).Columns("Turno").Activation = ssActivationActivateNoEdit
    grdCupos2.Bands(0).Columns("CuposLibres").Activation = ssActivationActivateNoEdit
'    grdCupos2.Bands(0).Columns("Servicio").Width = 3100
'    grdCupos2.Bands(0).Columns("Turno").Width = 1500
'    grdCupos2.Bands(0).Columns("CuposLibres").Width = 1500
    grdCupos2.Bands(0).Columns("CuposLibres").CellAppearance.TextAlign = ssAlignCenter
End Sub

Sub CreaTemporalCuposLibres1()
    If oRsServiciosCuposLibre1.State = 1 Then
       Set oRsServiciosCuposLibre1 = Nothing
    End If
    With oRsServiciosCuposLibre1
          .Fields.Append "IdServicio", adInteger
          .Fields.Append "Servicio", adVarChar, 255, adFldIsNullable
          .Fields.Append "Turno", adVarChar, 255, adFldIsNullable
          .Fields.Append "CuposLibres", adVarChar, 255, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdCupos1.DataSource = oRsServiciosCuposLibre1
    mo_Apariencia.ConfigurarFilasBiColores grdCupos1, SIGHEntidades.GrillaConFilasBicolor
End Sub

Sub LimpiarTemporalesCuposLibres()

    With oRsServiciosCuposLibre1
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
    
    With oRsServiciosCuposLibre2
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With

End Sub

Sub CreaTemporalCuposLibres2()
    If oRsServiciosCuposLibre2.State = 1 Then
       Set oRsServiciosCuposLibre2 = Nothing
    End If
    With oRsServiciosCuposLibre2
          .Fields.Append "IdServicio", adInteger
          .Fields.Append "Servicio", adVarChar, 255, adFldIsNullable
          .Fields.Append "Turno", adVarChar, 255, adFldIsNullable
          .Fields.Append "CuposLibres", adVarChar, 255, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdCupos2.DataSource = oRsServiciosCuposLibre2
    mo_Apariencia.ConfigurarFilasBiColores grdCupos2, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub btnAceptar_Click()
    On Error GoTo ErrCerrar
   
    If wxMuestraGrid = "CuposLibres" Then
        'Configura ventana al tamaño maximo
        lblTextoCabecera.Width = Screen.Width - 300
        FraCuposLibres.Top = 0
        FraCuposLibres.Width = Screen.Width - 300
        FraCuposLibres.Height = Screen.Height - 500   'Screen.Height - 1900
        grdCupos1.Left = 100
        grdCupos1.Width = (Screen.Width - 300) / 2 - 200
        grdCupos1.Height = Screen.Height - 600
        
        grdCupos1.Bands(0).Columns("Servicio").Width = ((Screen.Width - 300) / 2 - 200) / 2 - 150
        grdCupos1.Bands(0).Columns("Turno").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        grdCupos1.Bands(0).Columns("CuposLibres").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        
        grdCupos2.Left = 200 + (Screen.Width - 300) / 2 - 200
        grdCupos2.Width = (Screen.Width - 300) / 2 - 200
        grdCupos2.Height = Screen.Height - 600
        
        grdCupos2.Bands(0).Columns("Servicio").Width = ((Screen.Width - 300) / 2 - 200) / 2 - 150
        grdCupos2.Bands(0).Columns("Turno").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        grdCupos2.Bands(0).Columns("CuposLibres").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        
        Me.Frame4.Top = Screen.Height - 1600
        Me.Frame4.Width = Screen.Width - 300
        Me.btnAceptar.Left = (Screen.Width - 300) / 2 - 1565
        Me.btnCancelar.Left = (Screen.Width - 300) / 2 + 200
    
    
        Dim oRsTmpProgMedServicios As New Recordset
        Dim oRsTmpCitas As New Recordset
        Dim lcFecha As String
        Dim ldFechaActual As Date
        Dim lcHoraLimite As String
        Dim ldHoraActual As String
        Dim lHoraInicio As Long
        Dim lHoraFin  As Long
        Dim lHoraActual As Long
        Dim lTiempoPromedio As Long
        Dim lHoraSiguiente As Long
        Dim lHoraLimite As Long
        Dim lnTotalCupos As Integer, lnIdTurno As Long
        Dim lnTotalCuposBloqueados As Integer
        Dim lcHoraInicio As String, lcHoraFinal As String
        Dim lnTotalCitas As Integer
        Dim lcTextoTotalCupos As String
        Dim lnRegistroGrdCupos As Integer
        
        LimpiarTemporalesCuposLibres
        lnRegistroGrdCupos = 0
        
        ldFechaActual = Date
        ldHoraActual = Format$(Now, "h:mm")

'    LnTotalRegistrosGrilla1
        lblTextoCabecera.Caption = "Cupos desde: " & ldFechaActual & " " & ldHoraActual
        Set oRsTmpProgMedServicios = mo_ReglasDeProgMedica.ProgramacionMedicaServiciosSeleccionarPorFechas(ldFechaActual, ldFechaActual)
        If oRsTmpProgMedServicios.RecordCount <= 0 Then
            oRsTmpProgMedServicios.Close
            Set oRsTmpProgMedServicios = Nothing
        End If
        oRsTmpProgMedServicios.MoveFirst
        Do While Not oRsTmpProgMedServicios.EOF
            'Calcula Total Cupos
            lHoraInicio = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmpProgMedServicios.Fields!HoraInicio)
            lHoraFin = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmpProgMedServicios.Fields!HoraFin)
            lHoraActual = mo_ReglasDeProgMedica.ConvertirAMinutos(ldHoraActual)
            
            If lHoraActual <= lHoraFin Then
                lTiempoPromedio = oRsTmpProgMedServicios.Fields!TiempoPromedioAtencion
                lHoraSiguiente = lHoraInicio
                lnTotalCupos = 0
                lnTotalCuposBloqueados = 0
                lHoraLimite = 0
                Do While lHoraSiguiente < lHoraFin
                    lnTotalCupos = lnTotalCupos + 1
                    If lHoraLimite = 0 Then
                        If lHoraSiguiente <= lHoraActual And lHoraActual <= lHoraSiguiente + lTiempoPromedio Then
                            lHoraLimite = lHoraSiguiente
                            If lHoraActual = lHoraSiguiente + lTiempoPromedio Then
                                lHoraLimite = lHoraSiguiente + lTiempoPromedio
                            End If
                        End If
                    End If
                    If lHoraSiguiente + lTiempoPromedio <= lHoraActual Then
                        lnTotalCuposBloqueados = lnTotalCuposBloqueados + 1
                    End If
                    lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
'                    lcHoraInicio = mo_ReglasDeProgMedica.ConvertirAHora(lHoraInicio)
'                    lcHoraFinal = mo_ReglasDeProgMedica.ConvertirAHora(lHoraSiguiente)
'                    lHoraInicio = lHoraSiguiente
                Loop
                lcHoraLimite = mo_ReglasDeProgMedica.ConvertirAHora(lHoraLimite)
                If lHoraLimite = 0 Then lcHoraLimite = mo_ReglasDeProgMedica.ConvertirAHora(lHoraInicio)
                
                
                Set oRsTmpCitas = mo_ReglasDeProgMedica.CitasSeleccionarPorServicioTurnoFecha(ldFechaActual, oRsTmpProgMedServicios.Fields!IdServicio, oRsTmpProgMedServicios.Fields!idTurno, lcHoraLimite)
                lnTotalCitas = oRsTmpCitas.RecordCount
                
                If lnRegistroGrdCupos < LnTotalRegistrosGrilla1 Then
                    'Agregar Informacion Cupos
                    oRsServiciosCuposLibre1.AddNew
                    oRsServiciosCuposLibre1.Fields!IdServicio = oRsTmpProgMedServicios.Fields!IdServicio
                    oRsServiciosCuposLibre1.Fields!Servicio = oRsTmpProgMedServicios.Fields!Servicio
                    oRsServiciosCuposLibre1.Fields!Turno = oRsTmpProgMedServicios.Fields!HoraInicio & " - " & oRsTmpProgMedServicios.Fields!HoraFin
                    
                    lcTextoTotalCupos = "No hay"
                    If lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas > 0 Then
                        lcTextoTotalCupos = lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas
                    End If
                    oRsServiciosCuposLibre1.Fields!CuposLibres = lcTextoTotalCupos
                    oRsServiciosCuposLibre1.Update
                    lnRegistroGrdCupos = lnRegistroGrdCupos + 1 'Cuenta Registro
                Else
                    'Agregar Informacion Cupos
                    oRsServiciosCuposLibre2.AddNew
                    oRsServiciosCuposLibre2.Fields!IdServicio = oRsTmpProgMedServicios.Fields!IdServicio
                    oRsServiciosCuposLibre2.Fields!Servicio = oRsTmpProgMedServicios.Fields!Servicio
                    oRsServiciosCuposLibre2.Fields!Turno = oRsTmpProgMedServicios.Fields!HoraInicio & " - " & oRsTmpProgMedServicios.Fields!HoraFin
                    
                    lcTextoTotalCupos = "No hay"
                    If lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas > 0 Then
                        lcTextoTotalCupos = lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas
                    End If
                    oRsServiciosCuposLibre2.Fields!CuposLibres = lcTextoTotalCupos
                    oRsServiciosCuposLibre2.Update
                    lnRegistroGrdCupos = lnRegistroGrdCupos + 1 'Cuenta Registro
                End If
            End If
            oRsTmpProgMedServicios.MoveNext
        Loop
       
       Dim row As SSRow
       
       With oRsServiciosCuposLibre1
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                   Set row = Me.grdCupos1.ActiveRow
'                   row.Cells(3).Appearance.Font.Bold = True
                   If row.Cells(3) = "No hay" Then
                        row.Cells(3).Appearance.ForeColor = &HFF&
                   End If
                  .MoveNext
               Loop
            End If
        End With
        
       With oRsServiciosCuposLibre2
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                   Set row = Me.grdCupos2.ActiveRow
'                   row.Cells(3).Appearance.Font.Bold = True
                   If row.Cells(3) = "No hay" Then
                        row.Cells(3).Appearance.ForeColor = &HFF&
                   End If
                  .MoveNext
               Loop
            End If
        End With
        
        
        Exit Sub
    End If
    
    
    Dim LnGrid As Integer
    Dim lbContinua As Boolean: Dim lbMuestraEnGrid As Boolean
    Dim ml_idCuentaAtencion As Long
    Dim ldFecha1 As Date: Dim ldFecha2 As Date
    Dim ldFHCierre As Date
    Dim oRsTmp As New ADODB.Recordset
    Dim oRsTmp1 As New ADODB.Recordset
    Dim oRsTmp2 As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    
    Me.MousePointer = 11
    LnGrid = 1
    With oRsTmp1
          .Fields.Append "NroHistoria", adInteger
          .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
          .Fields.Append "NroCuenta", adInteger
          .Fields.Append "Servicio", adVarChar, 150, adFldIsNullable
          .Fields.Append "FechaCierre", adDate
          .LockType = adLockOptimistic
          .Open
    End With
    '

     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "AtencionesParaMDW"
         Set oRsTmp = .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    '
    If oRsTmp.RecordCount > 0 Then
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
            lbContinua = False: lbMuestraEnGrid = False
            ml_idCuentaAtencion = oRsTmp.Fields!idCuentaAtencion
            If oRsTmp.Fields!EsPacienteExterno = True Then
                '*****toma el mismo valor para Cerrar la cuenta que "CONSULTORIOS EXTERNOS"
                If oRsTmp.Fields!idEstado = 1 Then
                  ldFecha1 = oRsTmp.Fields!FechaCreacion
                  ldFecha2 = Now
                  If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHorasCE) Then
                     lbContinua = True
                  End If
                  lbMuestraEnGrid = True
                  ldFHCierre = DateAdd("h", Val(lcHorasCE), oRsTmp.Fields!FechaCreacion)
                End If
            Else
                Select Case oRsTmp.Fields!IdTipoServicio
                Case 1   'ce
                      If oRsTmp.Fields!idEstado = 1 Then
                        ldFecha1 = oRsTmp.Fields!FechaCreacion
                        ldFecha2 = Now
                        If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHorasCE) Then
                          lbContinua = True
                        End If
                        lbMuestraEnGrid = True
                        ldFHCierre = DateAdd("h", Val(lcHorasCE), oRsTmp.Fields!FechaCreacion)
                      End If
                Case 3   'hospitalizacion
                      If oRsTmp.Fields!idEstado = 12 Then
                        ldFecha1 = oRsTmp.Fields!FechaCreacion
                        ldFecha2 = Now
                        If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHorasHosp) Then
                           lbContinua = True
                        End If
                        lbMuestraEnGrid = True
                        ldFHCierre = DateAdd("h", Val(lcHorasHosp), oRsTmp.Fields!FechaCreacion)
                      End If
                Case Else      'emergencia
                      If oRsTmp.Fields!idEstado = 1 Or oRsTmp.Fields!idEstado = 12 Then
                        ldFecha1 = oRsTmp.Fields!FechaCreacion
                        ldFecha2 = Now
                        If oRsTmp.Fields!HoraIngreso > "06:00" And oRsTmp.Fields!HoraIngreso <= "21:59" Then
                            If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHOrasEmergDiurno) Then
                               lbContinua = True
                            End If
                        Else
                            If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHOrasEmergNocturno) Then
                               lbContinua = True
                            End If
                        End If
                        lbMuestraEnGrid = True
                        ldFHCierre = DateAdd("h", Val(lcHOrasEmergDiurno), oRsTmp.Fields!FechaCreacion)
                      End If
                End Select
            End If
            If lbMuestraEnGrid = True Then
                Select Case wxMuestraGrid
                Case "HOSPITALIZACION"
                    grdHospitalizados.Caption = "Sólo Hospitalización"
                    If oRsTmp.Fields!IdTipoServicio = 3 Then
                        oRsTmp1.AddNew
                        oRsTmp1.Fields!NroHistoria = oRsTmp.Fields!NroHistoriaClinica
                        oRsTmp1.Fields!Paciente = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                        oRsTmp1.Fields!Servicio = oRsTmp.Fields!Nombre
                        oRsTmp1.Fields!NroCuenta = oRsTmp.Fields!idCuentaAtencion
                        oRsTmp1.Fields!FechaCierre = ldFHCierre
                        oRsTmp1.Update
                    End If
                Case "EMERGENCIA"
                    grdHospitalizados.Caption = "Sólo Emergencia"
                    If (oRsTmp.Fields!IdTipoServicio = 2 Or oRsTmp.Fields!IdTipoServicio = 4) Then
                        oRsTmp1.AddNew
                        oRsTmp1.Fields!NroHistoria = oRsTmp.Fields!NroHistoriaClinica
                        oRsTmp1.Fields!Paciente = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                        oRsTmp1.Fields!Servicio = oRsTmp.Fields!Nombre
                        oRsTmp1.Fields!NroCuenta = oRsTmp.Fields!idCuentaAtencion
                        oRsTmp1.Fields!FechaCierre = ldFHCierre
                        oRsTmp1.Update
                    End If
                Case "CE"
                    grdHospitalizados.Caption = "Sólo Consulta Externa"
                    If oRsTmp.Fields!IdTipoServicio = 1 Then
                        oRsTmp1.AddNew
                        oRsTmp1.Fields!NroHistoria = oRsTmp.Fields!NroHistoriaClinica
                        oRsTmp1.Fields!Paciente = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                        oRsTmp1.Fields!Servicio = oRsTmp.Fields!Nombre
                        oRsTmp1.Fields!NroCuenta = oRsTmp.Fields!idCuentaAtencion
                        oRsTmp1.Fields!FechaCierre = ldFHCierre
                        oRsTmp1.Update
                    End If
                Case "TODOS"
                    grdHospitalizados.Caption = "CE, Emergencia y Hospitalización"
                    oRsTmp1.AddNew
                    oRsTmp1.Fields!NroHistoria = oRsTmp.Fields!NroHistoriaClinica
                    oRsTmp1.Fields!Paciente = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                    oRsTmp1.Fields!Servicio = oRsTmp.Fields!Nombre
                    oRsTmp1.Fields!NroCuenta = oRsTmp.Fields!idCuentaAtencion
                    oRsTmp1.Fields!FechaCierre = ldFHCierre
                    oRsTmp1.Update
                Case "CERRAR"
                    grdHospitalizados.Caption = "Cierra Cuentas"
                    If lbContinua = True Then
                        If oRsTmp.Fields!IdFormaPago > 1 Then
                           mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
                        End If
                        If EsSeguro(oRsTmp.Fields!IdFormaPago) = True Then
                           mo_AdminAdmision.CuentaAtencionPendientePagoSeguros ml_idCuentaAtencion, oRsTmp.Fields!IdPaciente, IIf(oRsTmp.Fields!IdTipoServicio = 1 Or oRsTmp.Fields!EsPacienteExterno = True, True, False)
                        Else
                           mo_AdminAdmision.CuentaAtencionCerradoAutomatico ml_idCuentaAtencion, oRsTmp.Fields!IdPaciente, IIf(oRsTmp.Fields!IdTipoServicio = 1 Or oRsTmp.Fields!EsPacienteExterno = 1, True, False)
                        End If
                    End If
                End Select
            End If
            oRsTmp.MoveNext
        Loop
        oRsTmp.Close
        Set oRsHospitalizados = oRsTmp1.Clone
        Set grdHospitalizados.DataSource = oRsHospitalizados
        '
        mo_ReglasArchivoClinico.ActualizaDatosConProblemas False
        '
    End If
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oConexion = Nothing
    Me.MousePointer = 1
    Exit Sub
ErrCerrar:
    If Err.Number = 3705 Then
      Select Case LnGrid
      Case 1
          oRsHospitalizados.Close
          Resume
      End Select
    Else
       MsgBox Err.Description
    End If
    Me.MousePointer = 1
End Sub



Private Sub btnCancelar_Click()
   Me.Visible = False
   txtClave.Text = "": txtUsuario.Text = ""
   MuestraBarraTarea
End Sub




Sub MaximizaPantalla()
   Me.Top = 0
   Me.Left = 0
   Me.Width = lnAnchoPantalla
   Me.Height = lnLargoPantalla

End Sub








Private Sub Form_Activate()
   MaximizaPantalla

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyTab, vbKeyControl
       KeyCode = 0
   Case 18
       KeyCode = 0
   End Select
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyTab, vbKeyControl
       KeyCode = 0
   Case 18
       KeyCode = 0
   End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrLoad
    grdCupos1.Caption = ""
    grdCupos2.Caption = ""
    '
    lcMensajeError = "CargaINI"
    CargaIni
    lcMensajeError = "timer1.interval"
    lnMinutosTranscurridos = 0
    'Timer1.Interval = 60000 * wxNumMinutosGrid   '
    
    lnCargoDesdeLoad = 0:    wxMinUltimoAvisoParaAlarma = 0:  lnNroIntentoCargarCabinas = 0
    lcMensajeError = "ocultaBarraTarea"
    OcultaBarraTarea
    lcMensajeError = "maximizaPantalla"
    lnAnchoPantalla = Screen.Width
    lnLargoPantalla = Screen.Height
     MaximizaPantalla
    lcMensajeError = " creavalorEnRegEdit..."
    
    'Frank 0602
    'Oculta Grilla y Muestra Grilla de Cupos Libres
    If wxMuestraGrid = "CuposLibres" Then
        FraProcesos.Width = 1
        FraCuposLibres.Width = LnWidthFrame
        FraCuposLibres.Top = LnTopFrame
        FraCuposLibres.Left = LnLeftFrame
        FraCuposLibres.Height = LnHeightFrame
        
        CreaTemporalCuposLibres1
        CreaTemporalCuposLibres2
    Else
        FraCuposLibres.Width = 1
        FraProcesos.Width = LnWidthFrame
        FraProcesos.Top = LnTopFrame
        FraProcesos.Left = LnLeftFrame
        FraProcesos.Height = LnHeightFrame
    End If
        
    CreaValorEnREGEDIT
    lbBuscaDNIenReniec = IIf(lcBuscaParametro.SeleccionaFilaParametro(296) = "S", True, False)
    
    'Frank 05022014
    LnTotalRegistrosGrilla1 = Val(lcBuscaParametro.SeleccionaFilaParametro(311))
    
    If lbBuscaDNIenReniec = True Then
        mo_Reniec.SeAccesaAlaWebDesdeGalenhos = False
        mo_Reniec.inicializar
    End If

    lcMensajeError = ".muestraBarraTareas"
    
    MuestraBarraTarea
    If App.PrevInstance Then
        Unload Me
        End
    End If
    

 
 Me.Visible = True
 
'---------------------------------
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
'---------------------------------
    t.szTip = "(v.3.0310a)...Cuentas que serán Anuladas  por el Sistema..." & Chr$(0) ' Es un string de "C" ( \0 )
    Shell_NotifyIcon NIM_ADD, t
    'Me.Hide
    App.TaskVisible = False
       
    lnMinutosTranscurridos = 0
    'Datos para CIERRE CUENTA
    '
    
    '
    lcMensajeError = "cargaDatosCierreCuenta"
    CargaDatosCierreCuenta
    lcMensajeError = "botonRefrescar"
    btnAceptar_Click
    Exit Sub
ErrLoad:
   MsgBox Err.Description & "  -->  " & lcMensajeError
End Sub

Sub CargaDatosCierreCuenta()
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    '
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "ParametrosSeleccionarTodos"
         Set oRsTmp = .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    '
    oRsTmp.MoveFirst
    oRsTmp.Find "idParametro=209"
    lcHorasCE = oRsTmp.Fields!ValorTexto
    oRsTmp.MoveFirst
    oRsTmp.Find "idParametro=233"
    lcHorasHosp = oRsTmp.Fields!ValorTexto
    oRsTmp.MoveFirst
    oRsTmp.Find "idParametro=234"
    lcHOrasEmergDiurno = oRsTmp.Fields!ValorTexto
    oRsTmp.MoveFirst
    oRsTmp.Find "idParametro=235"
    lcHOrasEmergNocturno = oRsTmp.Fields!ValorTexto
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub




Private Sub mnuAcerca_Click()
On Error GoTo ErrMNU
    ' Un consejo, mover un Form en estado minimizado
    ' da un GPF...
    Dim ValDev As Long
    With Me
        picGancho.Picture = Me.Icon
        Top = Screen.Height / 2 - Height / 2
        Left = Screen.Width / 2 - Width / 2
        Show
    End With

ErrMNU:
End Sub

Private Sub mnuSalir_Click(Index As Integer)
    Unload Me
End Sub

Sub HabilitaControles(lbHabilita As Boolean)
    txtMinCabina.Enabled = lbHabilita
    'txtAviso.Enabled = lbHabilita
    Command1.Enabled = lbHabilita
    Command5.Enabled = lbHabilita
End Sub



Private Sub picGancho_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, Msg As Long, ValDev As Long
    Msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:   'doble clic izquierdo
                 'ValDev = WinExec("CONTROL.EXE DESK.CPL", 1)
                 lcUsuario = ""
                 txtClave.Text = "": txtUsuario.Text = ""
                 HabilitaControles (False)
                 mnuAcerca_Click
            Case WM_LBUTTONDOWN:
                 txtClave.Text = ""
            Case WM_LBUTTONUP:
                 txtClave.Text = ""
            Case WM_RBUTTONDBLCLK:
                 txtClave.Text = ""
            Case WM_RBUTTONDOWN:        'clic derecho
            Case WM_RBUTTONUP:
                 txtClave.Text = ""
                 ' PopUp menu,2 significa Izq/Der botones en el menu, mnuAbout es BOLD
                 'Me.PopupMenu mnuBar, 2, , , mnuAcerca
            End Select
        rec = False
    End If
End Sub









Private Sub Timer1_Timer()
   FarmaciaRegeneraSaldos      'DEBB2014a
   ActualizaAdmisionCita
   If Format(Now(), "hh:mm") = "00:01" Then
      lbProcesaSisVSgalenhos = False
      lbProcesaReniecVSgalenhos = False
   End If
   '
   ImportaDatosDelSisAcreditaciones
   '
   'ComparaReniecGalenhos   'en pruebas
   '
   lnMinutosTranscurridos = lnMinutosTranscurridos + 1
   If lnMinutosTranscurridos < wxNumMinutosGrid Then
      Exit Sub
   End If
   lnMinutosTranscurridos = 0
   '
   btnAceptar_Click
   If oRsHospitalizados.State = 1 Then
        If oRsHospitalizados.RecordCount = 0 Then
           btnCancelar_Click
           Exit Sub
        End If
   End If

     MinimizaTodasLasVentanasAbiertas
     HabilitaControles (False)

     mnuAcerca_Click

     OcultaBarraTarea
     
     

End Sub

'DEBB2014a
Sub FarmaciaRegeneraSaldos()
             Dim oDOfarmAlmacen As New DOfarmAlmacen
             Dim oFarmAlmacen As New FarmAlmacen
             Dim oRsAlmacenes As New Recordset
             Dim oConexion As New Connection
             Dim lbRegeneraSaldos As Boolean
             Dim lcTexto As String, lnDiaSemana As Integer, lnFor As Integer, lcHoraActual As String
             Set oRsAlmacenes = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales<>'X' and idEstado=1")
             lcHoraActual = Format(Now(), "hh:mm")
             lnDiaSemana = Weekday(Now())
             oRsAlmacenes.Filter = "regenerarHora='" & lcHoraActual & "'"
             If oRsAlmacenes.RecordCount > 0 Then
                '****** REGENERA SALDOS *******
                Dim oRegenerarSaldo As New HerrRegeneraSaldos
                oRsAlmacenes.MoveFirst
                Do While Not oRsAlmacenes.EOF
                    lbRegeneraSaldos = False
                    '1=domingo, 2=lunes...7=sabado
                    'Ejemplo de farmAlmacen.RegeneraDias='13'      osea se REGENERA SALDOS el DOMINGO y MARTES
                    '           farmAlmacen.regeneraHora='02:15'   a las 2 mañana con 15 minutos REGENERA SALDOS
                    '           farmAlmacen.regenerarEstado=null   para que REGENERE SALDO
                    If InStr(Trim(Str(lnDiaSemana)), oRsAlmacenes.Fields!regenerarDias) > 0 And _
                                                 IsNull(oRsAlmacenes.Fields!regenerarEstado) Then
                       lbRegeneraSaldos = True
                    End If
                    If lbRegeneraSaldos = True Then
                        'actualiza ESTADO "Procesando"
                        If oConexion.State <> 1 Then
                           oConexion.CommandTimeout = 300
                           oConexion.CursorLocation = adUseClient
                           oConexion.Open SIGHEntidades.CadenaConexion
                           Set oFarmAlmacen.Conexion = oConexion
                        End If
                        oDOfarmAlmacen.IdAlmacen = oRsAlmacenes.Fields!IdAlmacen
                        If oFarmAlmacen.SeleccionarPorId(oDOfarmAlmacen) = True Then
                            oDOfarmAlmacen.regenerarEstado = ActualizaEstado("P", lnDiaSemana)   '..Procesando...
                            If oFarmAlmacen.Modificar(oDOfarmAlmacen) = True Then
                                'Regenera
                                oRegenerarSaldo.idUsuario = 0
                                oRegenerarSaldo.lcNombrePc = ""
                                oRegenerarSaldo.IdAlmacenAregenerar = oRsAlmacenes.Fields!IdAlmacen
                                oRegenerarSaldo.RegeneraDesdeUltimoMes = False
                                oRegenerarSaldo.FormularioUsadoDesdeOtroFrm = True
                                oRegenerarSaldo.Show 1
                                'actualiza ESTADO "Terminado"
                                 
                                 oDOfarmAlmacen.regenerarEstado = ActualizaEstado("T", lnDiaSemana)   '..Termino...
                                 If oFarmAlmacen.Modificar(oDOfarmAlmacen) = True Then
                                 End If
                             End If
                        End If
                        
                    End If
                    oRsAlmacenes.MoveNext
                Loop
                Set oRegenerarSaldo = Nothing
             Else
                '****** libera ESTADO=TERMINADO de días diferentes a HOY (farmAlmacen.regenerarEstado) ****
                oRsAlmacenes.Filter = ""
                If oRsAlmacenes.RecordCount > 0 Then
                   oRsAlmacenes.MoveFirst
                   Do While Not oRsAlmacenes.EOF
                       If InStr(oRsAlmacenes.Fields!regenerarEstado, "T") > 0 Or InStr(oRsAlmacenes.Fields!regenerarEstado, "P") > 0 Then
                           lcTexto = ""
                           For lnFor = 1 To 7
                               If lnDiaSemana = lnFor Then
                                  lcTexto = lcTexto & Mid(oRsAlmacenes.Fields!regenerarEstado, lnFor, 1)
                               Else
                                  lcTexto = lcTexto & " "
                               End If
                           Next
                           If Mid(oRsAlmacenes.Fields!regenerarEstado, lnDiaSemana, 1) = " " Then
                               If oConexion.State <> 1 Then
                                  oConexion.CommandTimeout = 300
                                  oConexion.CursorLocation = adUseClient
                                  oConexion.Open SIGHEntidades.CadenaConexion
                                  Set oFarmAlmacen.Conexion = oConexion
                               End If
                               oDOfarmAlmacen.IdAlmacen = oRsAlmacenes.Fields!IdAlmacen
                               If oFarmAlmacen.SeleccionarPorId(oDOfarmAlmacen) = True Then
                                    oDOfarmAlmacen.regenerarEstado = ""
                                    If oFarmAlmacen.Modificar(oDOfarmAlmacen) = True Then
                                    End If
                               End If
                           End If
                       End If
                       oRsAlmacenes.MoveNext
                   Loop
                End If
             End If
             '
             If oConexion.State = 1 Then
                oConexion.Close
             End If
             oRsAlmacenes.Close
             Set oRsAlmacenes = Nothing
             Set oDOfarmAlmacen = Nothing
             Set oFarmAlmacen = Nothing
             Set oConexion = Nothing
End Sub
'DEBB2014a
Function ActualizaEstado(lcNewEstado As String, lnDiaSemana As Integer) As String
        Dim lcTexto As String, lnFor As Integer
        lcTexto = ""
        For lnFor = 1 To 7
            If lnDiaSemana = lnFor Then
               lcTexto = lcTexto & lcNewEstado
            Else
               lcTexto = lcTexto & " "
            End If
        Next
        ActualizaEstado = lcTexto
End Function


'MARIO
Sub ActualizaAdmisionCita()
    On Error GoTo ActADCw
    Dim oConexionExterna As New Connection
    Dim oRsCitasWebCupos As New Recordset
    
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    Set oRsCitasWebCupos = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorIdEstadoCita(3, oConexionExterna)
    If oRsCitasWebCupos.RecordCount > 0 Then
        mo_ReglasDeProgMedica.ImportaCitasWebMINSA SIGHEntidades.Usuario, 683, mo_lcNombrePc
    End If
    oConexionExterna.Close
ActADCw:
    Set oConexionExterna = Nothing
    Set oRsCitasWebCupos = Nothing
End Sub


Private Sub OcultaBarraTarea()
    'Variable de retorno para el HWND
    Dim ret As Long

    'Se le pasa el nombre de clase a FindWindow
    ret = FindWindow("Shell_TrayWnd", vbNullString)
    Call SetWindowPos(ret, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Private Sub MuestraBarraTarea()
    'Variable de retorno para el HWND
    Dim ret As Long

    'Se le pasa el nombre de clase a FindWindow para obtener el handle
    ret = FindWindow("Shell_TrayWnd", vbNullString)
    Call SetWindowPos(ret, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub







Private Sub MinimizaTodasLasVentanasAbiertas()

Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)

End Sub





Sub CreaValorEnREGEDIT()
'   Dim valorDevuelto As String
'   Dim oCrypKey As New CrypKey.Util
'   valorDevuelto = ConsultarValor(&H80000002, "Software\Digital Works Corporation\SIGH", "CadenaConexionIntegrada:")
'   wxCadenaConexion = oCrypKey.DecryptString(valorDevuelto)
'   Set oCrypKey = Nothing
   'wxCadenaConexion = SIGHEntidades.CadenaConexionIntegrada
   On Error Resume Next
   wxCadenaConexion = CadenaConexionIntegrada
'   MsgBox wxCadenaConexion, vbCritical, wxCadenaConexion
End Sub

Function EsSeguro(lnIdTipoFinanciamiento As Long) As Boolean
    Dim oRsEsSeguro As New Recordset
   
    Set oRsEsSeguro = mo_ReglasComunes.TiposFinanciamientoSegunFiltro("idTipoFinanciamiento=" & lnIdTipoFinanciamiento)
    If oRsEsSeguro.RecordCount > 0 Then
       EsSeguro = oRsEsSeguro.Fields!esOficina
    End If
    oRsEsSeguro.Close
    Set oRsEsSeguro = Nothing
End Function

Function RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnNroCuenta As Long) As Double
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim lnTotal As Double
    '
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioFinanciamientosPorNroCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnNroCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   '
   lnTotal = 0
   oRecordset.Filter = "idEstadoFacturacion<>9"
   If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
          lnTotal = lnTotal + oRecordset.Fields!TotalFinanciado
          oRecordset.MoveNext
      Loop
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   RetornaConsumoPacienteServiciosConSeguroPorNroCuenta = lnTotal
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnNroCuenta As Long) As Double
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim lnTotal As Double
    '
    ms_MensajeError = ""
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFinanciamientosPorNroCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnNroCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   '
   lnTotal = 0
   oRecordset.Filter = "idEstadoMovimiento=1"
   If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
          lnTotal = lnTotal + oRecordset.Fields!TotalFinanciado
          oRecordset.MoveNext
      Loop
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta = lnTotal
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'****Barre Pacientes GAlenhos y busca por cada uno en RENIEC,
'****si Autogenerado es OK lo marca True en campo: usoWebReniec
'****sino agrege comentario en campo: OBSERVACION
Sub ComparaReniecGalenhos()
'    On Error GoTo ErrReniec
'    Dim lcHoraActual As String
'    lcHoraActual = Format(Now(), "hh:mm")
'    If lcHoraActual >= wxReniecHoraInicio And lcHoraActual <= wxReniecHoraFin And lbProcesaReniecVSgalenhos = False And lbBuscaDNIenReniec = True Then
'       lbProcesaReniecVSgalenhos = True
'       Dim oRsTmp1 As New Recordset
'       Dim oDOPaciente As New DOPaciente
'       Dim lcSql As String, lcAutogeneradoGalenHos As String, lcAutogeneradoReniec As String
'       Dim lcDNI As String
'       Dim lnNroConsultas As Long
'       lcSql = "select * from Pacientes where   idDocIdentidad=1 and len(ltrim(nroDocumento))=8 and (usoWebReniec is null) and (observacion is null)"
'       oRsTmp1.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'       lnNroConsultas = 0
'       If oRsTmp1.RecordCount > 0 Then
'          oRsTmp1.MoveFirst
'          Do While Not oRsTmp1.EOF
'             lcAutogeneradoGalenHos = oRsTmp1.Fields!Autogenerado
'             lcDNI = oRsTmp1.Fields!NroDocumento
'             mo_Reniec.SeAccesaAlaWebDesdeGalenhos = False
'             mo_Reniec.ConsultarDNIenReniec lcDNI
'             lnNroConsultas = lnNroConsultas + 1
'             If mo_Reniec.ApellidoPaterno <> "" Then
'                  oDOPaciente.ApellidoPaterno = mo_Reniec.ApellidoPaterno
'                  oDOPaciente.ApellidoMaterno = mo_Reniec.ApellidoMaterno
'                  oDOPaciente.PrimerNombre = mo_Reniec.PrimerNombre
'                  oDOPaciente.SegundoNombre = mo_Reniec.SegundoNombre
'                  oDOPaciente.FechaNacimiento = mo_Reniec.FechaNacimiento
'                  oDOPaciente.idTipoSexo = mo_Reniec.idTipoSexo
'                  lcAutogeneradoReniec = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDOPaciente)
'                  lcSql = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & Trim(oRsTmp1.Fields!PrimerNombre) & " " & IIf(IsNull(oRsTmp1.Fields!SegundoNombre), "", oRsTmp1.Fields!SegundoNombre) & "  FN: " & oRsTmp1.Fields!FechaNacimiento & "  s:" & oRsTmp1.Fields!idTipoSexo
'                  If lcAutogeneradoReniec = lcAutogeneradoGalenHos Then
'                     oRsTmp1.Fields!UsoWebReniec = True
'                  Else
'                     oRsTmp1.Fields!Observacion = Left(Trim(oRsTmp1.Fields!Observacion) & lcProblemasConReniec, 150)
'                  End If
'                  oRsTmp1.Update
'             End If
'             lcHoraActual = Format(Now(), "hh:mm")
'             If lcHoraActual > wxReniecHoraFin Then
'                lbProcesaReniecVSgalenhos = False
'                Exit Do
'             End If
'             oRsTmp1.MoveNext
'          Loop
'       End If
'       oRsTmp1.Close
'       '
'       lcSql = "update tempMovimientos set idusuario= " & lnNroConsultas & " where idMovimiento=1"
'       oRsTmp1.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'       '
'       Set oRsTmp1 = Nothing
'       Set oDOPaciente = Nothing
'       lbProcesaReniecVSgalenhos = False
'    End If
'    Exit Sub
'ErrReniec:
'    MsgBox Err.Description
'
End Sub


Sub ImportaDatosDelSisAcreditaciones()
'    On Error GoTo ErrImportarSIS
'    If lcBuscaParametro.SeleccionaFilaParametro(301) <> "S" Then
'       Exit Sub
'    End If
'    Dim lcHoraActual As String
'    lcHoraActual = Format(Now(), "hh:mm")
'    If lcHoraActual >= wxSisAcreditacioHoraInicio And lcHoraActual <= wxSisAcreditacioHoraFinal And lbProcesaSisVSgalenhos = False Then
'       lbProcesaSisVSgalenhos = True
'       '
'       Dim oRsTmp1 As New Recordset
'       Dim lcSql As String, lcFecha As String
'       Dim oConexion As New Connection
'       oConexion.CommandTimeout = 300
'       oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'       oConexion.CursorLocation = adUseClient
'       oConexion.BeginTrans
'       '
'       lcSql = "DELETE FROM SisFiliaciones"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "DROP INDEX SisFiliaciones.indDcto"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "DROP INDEX SisFiliaciones.indApellidos"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "DROP INDEX SisFiliaciones.indAfiliacion"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "alter table SisFiliaciones drop column FBajaOK"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "BULK INSERT SIGH_EXTERNA..SisFiliaciones FROM '" & lcBuscaParametro.SeleccionaFilaParametro(311) & "'" & _
'                " WITH (" & _
'                "   DATAFILETYPE = 'char'," & _
'                "   FIELDTERMINATOR = ','," & _
'                "   ROWTERMINATOR='" & Chr(10) & "'" & _
'                " )"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "ALTER TABLE SisFiliaciones add  FbajaOK datetime null"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "CREATE INDEX indDcto  ON SisFiliaciones (DocumentoNumero)"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "CREATE INDEX indApellidos  ON SisFiliaciones (Paterno,Materno,Pnombre)"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "CREATE INDEX indAfiliacion  ON SisFiliaciones (AfiliacionDisa,AfiliacionTipoFormato,AfiliacionNroFormato)"
'       If oRsTmp1.State = 1 Then oRsTmp1.Close
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       '
'       lcSql = "update parametros set ValorTexto='Afiliados SIS hasta " & Date & "/" & Time & "' where idParametro=312"
'       oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'       'actualiza fecha de baja
'       lcSql = "select * from sisFiliaciones where Fbaja<>'\N'"
'       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'       If oRsTmp1.RecordCount > 0 Then
'          oRsTmp1.MoveFirst
'          Do While Not oRsTmp1.EOF
'             lcFecha = Right(oRsTmp1.Fields!FBaja, 2) & "/" & Mid(oRsTmp1.Fields!FBaja, 5, 2) & "/" & Left(oRsTmp1.Fields!FBaja, 4)
'             If SIGHEntidades.EsFecha(lcFecha, "DD/MM/AAAA") Then
'                oRsTmp1.Fields!FBajaOK = CDate(lcFecha)
'                oRsTmp1.Update
'             End If
'             oRsTmp1.MoveNext
'          Loop
'       End If
'       oRsTmp1.Close
'       '
'       oConexion.CommitTrans
'       oConexion.Close
'       Set oRsTmp1 = Nothing
'       Set oConexion = Nothing
'
'    End If
'    Exit Sub
'ErrImportarSIS:
'
'       oConexion.RollbackTrans
'       oConexion.Close
'       Set oRsTmp1 = Nothing
'       Set oConexion = Nothing
'       lcSql = "update parametros set ValorTexto='Problemas al importar (" & Date & "/" & Format(Time, "hh:ss") & ")' where idParametro=312"
'       oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'Resume
End Sub




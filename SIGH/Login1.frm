VERSION 5.00
Begin VB.Form Login1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SGHClinicas"
   ClientHeight    =   2460
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4601.252
   ScaleMode       =   0  'User
   ScaleWidth      =   9077.404
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   5
      Top             =   1320
      Width           =   4665
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar"
         DisabledPicture =   "Login1.frx":0152
         DownPicture     =   "Login1.frx":05B2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   900
         Picture         =   "Login1.frx":0A27
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   255
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         DisabledPicture =   "Login1.frx":0FB1
         DownPicture     =   "Login1.frx":1475
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2445
         Picture         =   "Login1.frx":1961
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   4665
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1545
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2325
      End
      Begin VB.TextBox txtUsuario 
         Height          =   345
         Left            =   1545
         TabIndex        =   0
         Top             =   330
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   7
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Contraseña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   6
         Top             =   810
         Width           =   855
      End
   End
End
Attribute VB_Name = "Login1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Ingreso al Sistema (Usuario y Clave)
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mb_Autenticado As Boolean
Dim ml_IdUsuarioAutenticado As Long
Dim ms_NombreUsuarioAutenticado As String
Dim ms_UsuarioSesionSQLServer As String
Dim ms_PasswordSesionSQLServer As String
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Procesos As New SIGHProxies.Procesos
Dim mo_Teclado As New sighentidades.Teclado
Dim mb_CargaDesdeOtraOpcion As Boolean
Dim ml_NombreMaquina As String

Property Let UsuarioDeEstadoDeCuenta(lnIdEmpleado As Long)
        Me.txtUsuario.Text = mo_reglasComunes.EmpleadosDevuelveNombreUsuario(lnIdEmpleado)
End Property


Property Let CargaDesdeOtraOpcion(bValue As Boolean)
        mb_CargaDesdeOtraOpcion = bValue
End Property
Property Let Autenticado(bValue As Boolean)
        mb_Autenticado = bValue
End Property
Property Get Autenticado() As Boolean
    Autenticado = mb_Autenticado
End Property
Property Let IdUsuarioAutenticado(lValue As Long)
        ml_IdUsuarioAutenticado = lValue
End Property
Property Get IdUsuarioAutenticado() As Long
    IdUsuarioAutenticado = ml_IdUsuarioAutenticado
End Property
Property Let NombreUsuarioAutenticado(sValue As String)
        ms_NombreUsuarioAutenticado = sValue
End Property
Property Get NombreUsuarioAutenticado() As String
    NombreUsuarioAutenticado = ms_NombreUsuarioAutenticado
End Property
Property Let UsuarioSesionSQLServer(sValue As String)
        ms_UsuarioSesionSQLServer = sValue
End Property
Property Get UsuarioSesionSQLServer() As String
    UsuarioSesionSQLServer = ms_UsuarioSesionSQLServer
End Property
Property Let PasswordSesionSQLServer(sValue As String)
        ms_PasswordSesionSQLServer = sValue
End Property
Property Get PasswordSesionSQLServer() As String
    PasswordSesionSQLServer = ms_PasswordSesionSQLServer
End Property

Private Sub btnAceptar_Click()
Dim rsUsuarioAutenticado As New Recordset
Dim sCadenaConexion As String
Dim oCrypKey As New CrypKey.Util
Dim lbContinuarLogin As Boolean
On Error GoTo ErrorManager
    'Validar el usuario y clave en la tabla empleados
    
    ml_NombreMaquina = sighentidades.RetornaNombrePC
    
    
    Set rsUsuarioAutenticado = mo_AdminSeguridad.EmpleadosAutenticar(Me.txtUsuario)
    
    Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsCitasWeb As New Recordset
    mo_Procesos.SomeeActualizaDatos 2, lcMensaje, "", "", (Date - 1), (Date - 1), lbSeTerminaSistema, oRsCitasWeb
    Set oRsCitasWeb = Nothing
    
    
    
    If Not (rsUsuarioAutenticado.BOF And rsUsuarioAutenticado.EOF) Then
      If UCase(Trim(Me.txtPassword)) <> UCase(Trim(oCrypKey.DecryptString(rsUsuarioAutenticado!Clave))) Then
        MsgBox "La clave ingresada no es válida", vbInformation, Me.Caption
        Exit Sub
      End If
      Set rsUsuarioAutenticado = mo_AdminSeguridad.EmpleadosAutenticarMaquina(Me.txtUsuario, ml_NombreMaquina)
      If Not (rsUsuarioAutenticado.BOF And rsUsuarioAutenticado.EOF) Then
        lbContinuarLogin = True
        If mb_CargaDesdeOtraOpcion = False Then
        End If
        If lbContinuarLogin = True Then
            If UCase(Trim(Me.txtPassword)) <> UCase(Trim(oCrypKey.DecryptString(rsUsuarioAutenticado!Clave))) Then
                MsgBox "La clave ingresada no es válida", vbInformation, Me.Caption
            Else
                sCadenaConexion = mo_AdminSeguridad.ParametrosObtenerCadenaConexion()
                If sCadenaConexion <> "" Then
                    sighentidades.ParaAuditoria = ""
                    sighentidades.Parametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
                    sighentidades.Parametro550 = lcBuscaParametro.SeleccionaFilaParametro(550)
                    sighentidades.Parametro551 = lcBuscaParametro.SeleccionaFilaParametro(551)
                    sighentidades.Parametro556 = lcBuscaParametro.SeleccionaFilaParametro(556)
                    sighentidades.Parametro560 = lcBuscaParametro.SeleccionaFilaParametro(560)
                    sighentidades.Parametro561 = lcBuscaParametro.SeleccionaFilaParametro(561)
                    sighentidades.Parametro562 = lcBuscaParametro.SeleccionaFilaParametro(562)
                    sighentidades.Parametro568 = lcBuscaParametro.SeleccionaFilaParametro(568)
                    sighentidades.Parametro569 = lcBuscaParametro.SeleccionaFilaParametro(569)
                    sighentidades.Parametro378 = lcBuscaParametro.SeleccionaFilaParametro(378)
                    
                    sighentidades.Lx_LabVacio = "(VACIO)"
                    sighentidades.Pto = "."
                    sighentidades.Parametro301valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(301)
                    sighentidades.Parametro322valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(322)
                    sighentidades.Parametro282valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(282)
                    sighentidades.Parametro378valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(378)
                    sighentidades.Parametro387valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(387)
                    sighentidades.Parametro503valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(503)
                
                    'Guarda la cadena de conexion al registro de windows
                    sighentidades.CadenaConexion = sCadenaConexion
                    
                    sCadenaConexion = mo_AdminSeguridad.ParametrosObtenerCadenaConexionShape()
                    sighentidades.CadenaConexionShape = sCadenaConexion
                                    
                    'Guarda el usuario y nombre del ultimo empleado logeado al sistemas en el registro
                    ml_IdUsuarioAutenticado = rsUsuarioAutenticado!IdEmpleado
                    sighentidades.Usuario = ml_IdUsuarioAutenticado
                    ms_NombreUsuarioAutenticado = rsUsuarioAutenticado!ApellidoPaterno + " " + rsUsuarioAutenticado!ApellidoMaterno + " " + rsUsuarioAutenticado!Nombres
                    sighentidades.NombreUsuario = ms_NombreUsuarioAutenticado
                    
                    mb_Autenticado = True
                    mo_AdminSeguridad.LogueaUsuario 1, ml_IdUsuarioAutenticado, ml_NombreMaquina
                    ChequeaSiElInstaladorEsParaHospitalOcsYconfiguraUnaSolaVez
                    Me.Visible = False
                Else
                    MsgBox "La cadena de conexión esta vacía", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End If
      Else
        MsgBox "El usuario ya inició sesión en otra PC", vbInformation, Me.Caption
      End If
    Else
        MsgBox "El usuario ingresado no es válido.", vbInformation, Me.Caption
    End If
    
    rsUsuarioAutenticado.Close
    Set rsUsuarioAutenticado = Nothing
Exit Sub
ErrorManager:

    If InStr(Err.Description, "ConnectionOpen") > 0 Then
        MsgBox "El servidor de datos no está operativo, consulte a soporte técnico" & Chr(13) & Err.Description, vbInformation, Me.Caption
    Else
        MsgBox "No se puede conectar al servidor de datos, consulte a soporte técnico" & Chr(13) & Err.Description, vbInformation, Me.Caption
    End If
    Exit Sub
End Sub

Private Sub btnCancelar_Click()
    mb_Autenticado = False
    Me.Visible = False
    End
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub Form_Activate()
    If Len(Me.txtUsuario.Text) > 1 Then
       Me.txtPassword.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

'Sub SkinConfigura()
'  On Error GoTo ErrSkin
'  Skin1.LoadSkin App.Path & "\" & WxSkin
'  Skin1.ApplySkin Me.hwnd
'ErrSkin:
'End Sub
Private Sub Form_Load()
'    SkinConfigura
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Trim(txtUsuario.Text) <> "" And Trim(txtPassword.Text) <> "" Then
      btnAceptar_Click
    Else
      If Trim(txtUsuario.Text) = "" Then
        txtUsuario.SetFocus
      Else
        txtPassword.SetFocus
      End If
    End If
  End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Trim(txtUsuario.Text) <> "" And Trim(txtPassword.Text) <> "" Then
      btnAceptar_Click
    Else
      If Trim(txtUsuario.Text) = "" Then
        txtUsuario.SetFocus
      Else
        txtPassword.SetFocus
      End If
    End If
  End If
End Sub

Sub ChequeaSiElInstaladorEsParaHospitalOcsYconfiguraUnaSolaVez()
      Dim lcBuscaParametro As New SIGHDatos.Parametros
      On Error GoTo TermChequeo
      If lcBuscaParametro.SeleccionaFilaParametro(282) = "*" Then
         If sighentidades.EsCentroSalud <> "" Then
            Dim oDOPArametro As New DOPArametro, oParametros As New Parametros
            Dim oConexion As New Connection
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighentidades.CadenaConexion
            oConexion.BeginTrans
            Set oParametros.Conexion = oConexion
            oDOPArametro.IdParametro = 282
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               GoTo TermChequeo
            End If
            oDOPArametro.ValorTexto = IIf(sighentidades.EsCentroSalud = "S", "S", "n")
            If Not oParametros.Modificar(oDOPArametro) Then
               GoTo TermChequeo
            End If
            If sighentidades.EsCentroSalud = "S" Then
               mo_reglasComunes.RolesEliminaHospEmergDelADMINISTRADOR oConexion
            End If
            oConexion.CommitTrans
            oConexion.Close
            Set oDOPArametro = Nothing
            Set oParametros = Nothing
            Set oConexion = Nothing
          End If
      End If
      Set lcBuscaParametro = Nothing
TermChequeo:
End Sub

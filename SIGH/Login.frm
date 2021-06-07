VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIS-GalenPLUS"
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
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   7
      Top             =   1320
      Width           =   4665
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "Login.frx":0CCA
         DownPicture     =   "Login.frx":112A
         Height          =   700
         Left            =   900
         Picture         =   "Login.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "Login.frx":1A14
         DownPicture     =   "Login.frx":1ED8
         Height          =   700
         Left            =   2445
         Picture         =   "Login.frx":23C4
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
         Caption         =   "&Contraseña:"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   390
         Width           =   645
      End
   End
End
Attribute VB_Name = "Login"
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
Dim mb_Autenticado As Boolean
Dim ml_IdUsuarioAutenticado As Long
Dim ms_NombreUsuarioAutenticado As String
Dim ms_UsuarioSesionSQLServer As String
Dim ms_PasswordSesionSQLServer As String
Dim mo_Procesos As New SIGHProxies.Procesos
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New sighEntidades.Teclado
Dim mb_CargaDesdeOtraOpcion As Boolean
Dim ml_NombreMaquina As String

Property Let UsuarioDeEstadoDeCuenta(lnIdEmpleado As Long)
        Me.txtUsuario.Text = mo_ReglasComunes.EmpleadosDevuelveNombreUsuario(lnIdEmpleado)
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
Dim lbContinuarLogin As Boolean, ldHoy As Date
On Error GoTo ErrorManager
    'Validar el usuario y clave en la tabla empleados
    
    ml_NombreMaquina = sighEntidades.RetornaNombrePC
    
    
    Set rsUsuarioAutenticado = mo_AdminSeguridad.EmpleadosAutenticar(Me.txtUsuario)
    
    'SCCQ 20/11/2020 Cambio22 Inicio
    'PROCESO DE VALIDAR VERSIÓN DEL APLICATIVO
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim version As String
    version = "07062021u75hra" 'Verisión actual de los dlls del sistema
    If version <> lcBuscaParametro.SeleccionaFilaParametro(314) Then 'Parámetro 314 versión del sistema en la Base de Datos
        MsgBox "El sistema se actualizará con la última versión", vbExclamation, Me.Caption
        Dim rutaRaiz As String
        rutaRaiz = App.Path & "\actualizahra.bat"
        Shell (rutaRaiz)
        Set lcBuscaParametro = Nothing
        version = ""
        rutaRaiz = ""
        End
    End If
    Set lcBuscaParametro = Nothing
        version = ""
        rutaRaiz = ""
    'SCCQ 20/11/2020 Cambio22 Fin
    
     Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsCitasWeb As New Recordset
     lcMensaje = ""
     Set oRsCitasWeb = Nothing
'     If lbSeTerminaSistema = True Then
'        End
'     End If
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    mo_ReglasDeProgMedica.CitasBloqueadasWEBEliminarXFechas CDate("01/01/2000"), ldHoy
    
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
                    sighEntidades.ParaAuditoria = ""
                    sighEntidades.Parametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
                    sighEntidades.Parametro550 = lcBuscaParametro.SeleccionaFilaParametro(550)
                    sighEntidades.Parametro551 = lcBuscaParametro.SeleccionaFilaParametro(551)
                    sighEntidades.Parametro556 = lcBuscaParametro.SeleccionaFilaParametro(556)
                    sighEntidades.Parametro560 = lcBuscaParametro.SeleccionaFilaParametro(560)
                    sighEntidades.Parametro561 = lcBuscaParametro.SeleccionaFilaParametro(561)
                    sighEntidades.Parametro562 = lcBuscaParametro.SeleccionaFilaParametro(562)
                    sighEntidades.Parametro568 = lcBuscaParametro.SeleccionaFilaParametro(568)
                    sighEntidades.Parametro569 = lcBuscaParametro.SeleccionaFilaParametro(569)
                    sighEntidades.Parametro378 = lcBuscaParametro.SeleccionaFilaParametro(378)
                    sighEntidades.ImpresoraDefaultDeEstaPC = ImpresoraDefault
                    
                    sighEntidades.Lx_LabVacio = "(VACIO)"
                    sighEntidades.Pto = "."
                    sighEntidades.Parametro301valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(301)
                    sighEntidades.Parametro322valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(322)
                    sighEntidades.Parametro282valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(282)
                    sighEntidades.Parametro378valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(378)
                    sighEntidades.Parametro387valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(387)
                    sighEntidades.Parametro503valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(503)
                    sighEntidades.Parametro573valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(573)
                    sighEntidades.Parametro583valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(583)
                    sighEntidades.Parametro584valorInt = lcBuscaParametro.SeleccionaFilaParametroValorInt(584)
                    
                    
                    '***Parametro1valorInt =exp HIS
                    '***Parametro2valorInt =proc rep para huelga
                    '***Parametro101valorInt = Imprimir recibo
                    '***Parametro250valorInt = lic men tex
                    '***Parametro322valorInt = 9 no se usará CODIGO QR Para emitir en Ticket
                    '***Parametro500valorInt = mdw, procesar resultados laboratorio automaticos
                    '***Parametro501valorInt = mdw, actualizar citas web automáticamente
                    '***Parametro502valorInt = mdw, procesar resultados Imagenes automaticos
                    '***Parametro503valorInt = 1->es manual el importar Citas Web
                    '***Parametro559valorInt = N° maximo de años para pasar a PASIVO la HISTORIA
                    '***Parametro573valorInt=1 ->para poner SALTO DE LINEA al generar archivo JSON
                    '***Parametro573valorInt=2 ->para NO GENERAR ARCHIVOS SUNAT desde CAJA
                    '***Parametro583valorInt=1 ->Citas ->atencion ->queda HABILITADO el option "WEB"
                    '***Parametro584valorInt=1 ->Pacientes ->Agregar->el TIPO NUMERO HISTORIA es AUTOGENERADO
                
                
                    'Guarda la cadena de conexion al registro de windows
                    sighEntidades.CadenaConexion = sCadenaConexion
                    
                    sCadenaConexion = mo_AdminSeguridad.ParametrosObtenerCadenaConexionShape()
                    sighEntidades.CadenaConexionShape = sCadenaConexion
                                    
                    'Guarda el usuario y nombre del ultimo empleado logeado al sistemas en el registro
                    ml_IdUsuarioAutenticado = rsUsuarioAutenticado!IdEmpleado
                    sighEntidades.Usuario = ml_IdUsuarioAutenticado
                    ms_NombreUsuarioAutenticado = rsUsuarioAutenticado!ApellidoPaterno + " " + rsUsuarioAutenticado!ApellidoMaterno + " " + rsUsuarioAutenticado!Nombres
                    sighEntidades.NombreUsuario = ms_NombreUsuarioAutenticado
                    
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
    'End
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
         If sighEntidades.EsCentroSalud <> "" Then
            Dim oDOPArametro As New DOPArametro, oParametros As New Parametros
            Dim oConexion As New Connection
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighEntidades.CadenaConexion
            oConexion.BeginTrans
            Set oParametros.Conexion = oConexion
            oDOPArametro.IdParametro = 282
            If Not oParametros.SeleccionarPorId(oDOPArametro) Then
               GoTo TermChequeo
            End If
            oDOPArametro.ValorTexto = IIf(sighEntidades.EsCentroSalud = "S", "S", "n")
            If Not oParametros.Modificar(oDOPArametro) Then
               GoTo TermChequeo
            End If
            If sighEntidades.EsCentroSalud = "S" Then
               mo_ReglasComunes.RolesEliminaHospEmergDelADMINISTRADOR oConexion
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





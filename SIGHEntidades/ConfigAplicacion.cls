VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigAplicacion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oReg As New Registro
Const RutaDeRegistroProyecto2000 = "Software\Minsa"
Const RutaDeRegistroSIGH = "Software\Minsa\SIGH"

'**********************************
'      RUTA DE EXE PRINCIPAL
'**********************************
Property Get RutaExePrincipal() As String
    RutaExePrincipal = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "RutaExePrincipal:")
End Property
Property Let RutaExePrincipal(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "RutaExePrincipal:", sValor
End Property

'**********************************
'BASE DE DATOS
'**********************************
Property Get CadenaConexion() As String
    CadenaConexion = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexion:")
End Property
Property Let CadenaConexion(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "CadenaConexion:", sValor
End Property
Property Get BaseDatosUserId() As String
    BaseDatosUserId = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "BaseDatosUserId:")
End Property
Property Let BaseDatosUserId(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "BaseDatosUserId:", sValor
End Property
Property Get BaseDatosClave() As String
    BaseDatosClave = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "BaseDatosClave:")
End Property
Property Let BaseDatosClave(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "BaseDatosClave:", sValor
End Property
Property Get BaseDatosServidor() As String
    BaseDatosServidor = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "BaseDatosServidor:")
End Property
Property Let BaseDatosServidor(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "BaseDatosServidor:", sValor
End Property

'**********************************
'SISTEMA
'**********************************
Property Get NombreSistema() As String
    NombreSistema = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Nombre:")
End Property
Property Let NombreSistema(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Nombre:", sValor
End Property

Property Get RutaReportes() As String
    RutaReportes = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "RutaReportes:")
End Property
Property Let RutaReportes(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "RutaReportes:", sValor
End Property

'**********************************
' ESTABLECIMIENTO
'**********************************
Property Get IdEstablecimiento() As String
    IdEstablecimiento = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "IdEstablecimiento:")
End Property
Property Let IdEstablecimiento(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "IdEstablecimiento:", sValor
End Property
Property Get DISA() As String
    DISA = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "DISA:")
End Property
Property Let DISA(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "DISA:", sValor
End Property
'**********************************
'      USUARIO
'**********************************
Property Get Usuario() As String
    Usuario = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Usuario:")
End Property
Property Let Usuario(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Usuario:", sValor
End Property
Property Get NombreUsuario() As String
    NombreUsuario = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "NombreUsuario:")
End Property
Property Let NombreUsuario(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "NombreUsuario:", sValor
End Property
Property Get Periodo() As String
    Periodo = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Periodo:")
End Property
Property Let Periodo(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Periodo:", sValor
End Property

Private Sub Class_Terminate()
    Set oReg = Nothing
End Sub



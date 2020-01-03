VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFuaAtencionSMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para la estructura de la tabla SisFuaAtencionSMI
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_id As Long
Dim ml_idCuentaAtencion As Long
Dim ms_IntervencionesPreventivas As String
Dim ms_Valor As String
Dim ms_CabDniUsuarioRegistra As String
Dim ms_CabFechaFuaPrimeraVez As String
Dim ms_CabEstado As String
Dim ms_CabNroEnvioAlSIS As String
Dim ml_CabCodigoPuntoDigitacion As Long
Dim ms_CabCodigoUDR As String
Dim ms_FuaDisa As String
Dim ms_FuaLote As String
Dim ms_FuaNumero As String
Dim ms_CabOrigenDelRegistro As String
Dim ms_CabVersionAplicativo As String
Dim ml_CabIdentificacionPaquete As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let id(iValue As Long)
   ml_id = iValue
End Property
Property Get id() As Long
   id = ml_id
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IntervencionesPreventivas(sValue As String)
   ms_IntervencionesPreventivas = sValue
End Property
Property Get IntervencionesPreventivas() As String
   IntervencionesPreventivas = ms_IntervencionesPreventivas
End Property
Property Let Valor(sValue As String)
   ms_Valor = sValue
End Property
Property Get Valor() As String
   Valor = ms_Valor
End Property
Property Let CabDniUsuarioRegistra(sValue As String)
   ms_CabDniUsuarioRegistra = sValue
End Property
Property Get CabDniUsuarioRegistra() As String
   CabDniUsuarioRegistra = ms_CabDniUsuarioRegistra
End Property
Property Let CabFechaFuaPrimeraVez(sValue As String)
   ms_CabFechaFuaPrimeraVez = sValue
End Property
Property Get CabFechaFuaPrimeraVez() As String
   CabFechaFuaPrimeraVez = ms_CabFechaFuaPrimeraVez
End Property
Property Let CabEstado(sValue As String)
   ms_CabEstado = sValue
End Property
Property Get CabEstado() As String
   CabEstado = ms_CabEstado
End Property
Property Let CabNroEnvioAlSIS(sValue As String)
   ms_CabNroEnvioAlSIS = sValue
End Property
Property Get CabNroEnvioAlSIS() As String
   CabNroEnvioAlSIS = ms_CabNroEnvioAlSIS
End Property
Property Let CabCodigoPuntoDigitacion(iValue As Long)
   ml_CabCodigoPuntoDigitacion = iValue
End Property
Property Get CabCodigoPuntoDigitacion() As Long
   CabCodigoPuntoDigitacion = ml_CabCodigoPuntoDigitacion
End Property
Property Let CabCodigoUDR(sValue As String)
   ms_CabCodigoUDR = sValue
End Property
Property Get CabCodigoUDR() As String
   CabCodigoUDR = ms_CabCodigoUDR
End Property
Property Let FuaDisa(sValue As String)
   ms_FuaDisa = sValue
End Property
Property Get FuaDisa() As String
   FuaDisa = ms_FuaDisa
End Property
Property Let FuaLote(sValue As String)
   ms_FuaLote = sValue
End Property
Property Get FuaLote() As String
   FuaLote = ms_FuaLote
End Property
Property Let FuaNumero(sValue As String)
   ms_FuaNumero = sValue
End Property
Property Get FuaNumero() As String
   FuaNumero = ms_FuaNumero
End Property
Property Let CabOrigenDelRegistro(sValue As String)
   ms_CabOrigenDelRegistro = sValue
End Property
Property Get CabOrigenDelRegistro() As String
   CabOrigenDelRegistro = ms_CabOrigenDelRegistro
End Property
Property Let CabVersionAplicativo(sValue As String)
   ms_CabVersionAplicativo = sValue
End Property
Property Get CabVersionAplicativo() As String
   CabVersionAplicativo = ms_CabVersionAplicativo
End Property
Property Let CabIdentificacionPaquete(iValue As Long)
   ml_CabIdentificacionPaquete = iValue
End Property
Property Get CabIdentificacionPaquete() As Long
   CabIdentificacionPaquete = ml_CabIdentificacionPaquete
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFuaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla SisFuaUsuario
'        Programado por: Barrantes D
'        Fecha: Julio 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idUsuario As Long
Dim ms_DNI As String
Dim ms_TipoDoc As String
Dim ms_ApellidoPat As String
Dim ms_ApellidoMat As String
Dim ms_PrimerNombre As String
Dim ms_SegundoNombre As String
Dim ml_NroEnvio As Long
Dim ms_Periodo As String
Dim ms_Mes As String
Dim ms_CodigoEstablecimiento As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let DNI(sValue As String)
   ms_DNI = sValue
End Property
Property Get DNI() As String
   DNI = ms_DNI
End Property
Property Let TipoDoc(sValue As String)
   ms_TipoDoc = sValue
End Property
Property Get TipoDoc() As String
   TipoDoc = ms_TipoDoc
End Property
Property Let ApellidoPat(sValue As String)
   ms_ApellidoPat = sValue
End Property
Property Get ApellidoPat() As String
   ApellidoPat = ms_ApellidoPat
End Property
Property Let ApellidoMat(sValue As String)
   ms_ApellidoMat = sValue
End Property
Property Get ApellidoMat() As String
   ApellidoMat = ms_ApellidoMat
End Property
Property Let PrimerNombre(sValue As String)
   ms_PrimerNombre = sValue
End Property
Property Get PrimerNombre() As String
   PrimerNombre = ms_PrimerNombre
End Property
Property Let SegundoNombre(sValue As String)
   ms_SegundoNombre = sValue
End Property
Property Get SegundoNombre() As String
   SegundoNombre = ms_SegundoNombre
End Property
Property Let NroEnvio(iValue As Long)
   ml_NroEnvio = iValue
End Property
Property Get NroEnvio() As Long
   NroEnvio = ml_NroEnvio
End Property
Property Let Periodo(sValue As String)
   ms_Periodo = sValue
End Property
Property Get Periodo() As String
   Periodo = ms_Periodo
End Property
Property Let Mes(sValue As String)
   ms_Mes = sValue
End Property
Property Get Mes() As String
   Mes = ms_Mes
End Property
Property Let CodigoEstablecimiento(sValue As String)
   ms_CodigoEstablecimiento = sValue
End Property
Property Get CodigoEstablecimiento() As String
   CodigoEstablecimiento = ms_CodigoEstablecimiento
End Property


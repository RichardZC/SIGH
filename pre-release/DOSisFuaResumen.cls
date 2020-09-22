VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFuaResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla SisFuaResumen
'        Programado por: Barrantes D
'        Fecha: Julio 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idResumen As Long
Dim ms_Anio As String
Dim ms_Mes As String
Dim ms_NroEnvio As String
Dim ms_NomPaquete As String
Dim ms_VersionGTI As String
Dim ml_CantFilATE As Long
Dim ml_CantFilSMI As Long
Dim ml_CantFilDIA As Long
Dim ml_CantFilMED As Long
Dim ml_CantFilINS As Long
Dim ml_CantFilPRO As Long
Dim ml_CantFilUSU As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idResumen(iValue As Long)
   ml_idResumen = iValue
End Property
Property Get idResumen() As Long
   idResumen = ml_idResumen
End Property
Property Let Anio(sValue As String)
   ms_Anio = sValue
End Property
Property Get Anio() As String
   Anio = ms_Anio
End Property
Property Let Mes(sValue As String)
   ms_Mes = sValue
End Property
Property Get Mes() As String
   Mes = ms_Mes
End Property
Property Let NroEnvio(sValue As String)
   ms_NroEnvio = sValue
End Property
Property Get NroEnvio() As String
   NroEnvio = ms_NroEnvio
End Property
Property Let NomPaquete(sValue As String)
   ms_NomPaquete = sValue
End Property
Property Get NomPaquete() As String
   NomPaquete = ms_NomPaquete
End Property
Property Let VersionGTI(sValue As String)
   ms_VersionGTI = sValue
End Property
Property Get VersionGTI() As String
   VersionGTI = ms_VersionGTI
End Property
Property Let CantFilATE(iValue As Long)
   ml_CantFilATE = iValue
End Property
Property Get CantFilATE() As Long
   CantFilATE = ml_CantFilATE
End Property
Property Let CantFilSMI(iValue As Long)
   ml_CantFilSMI = iValue
End Property
Property Get CantFilSMI() As Long
   CantFilSMI = ml_CantFilSMI
End Property
Property Let CantFilDIA(iValue As Long)
   ml_CantFilDIA = iValue
End Property
Property Get CantFilDIA() As Long
   CantFilDIA = ml_CantFilDIA
End Property
Property Let CantFilMED(iValue As Long)
   ml_CantFilMED = iValue
End Property
Property Get CantFilMED() As Long
   CantFilMED = ml_CantFilMED
End Property
Property Let CantFilINS(iValue As Long)
   ml_CantFilINS = iValue
End Property
Property Get CantFilINS() As Long
   CantFilINS = ml_CantFilINS
End Property
Property Let CantFilPRO(iValue As Long)
   ml_CantFilPRO = iValue
End Property
Property Get CantFilPRO() As Long
   CantFilPRO = ml_CantFilPRO
End Property
Property Let CantFilUSU(iValue As Long)
   ml_CantFilUSU = iValue
End Property
Property Get CantFilUSU() As Long
   CantFilUSU = ml_CantFilUSU
End Property



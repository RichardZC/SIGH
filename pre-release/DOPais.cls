VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Paises
'        Programado por: Barrantes D
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_nombre As String
Dim ml_IdPais As Long
Dim ms_codigoPostal As String
Property Let codigoPostal(sValue As String)
   ms_codigoPostal = sValue
End Property
Property Get codigoPostal() As String
   codigoPostal = ms_codigoPostal
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let nombre(sValue As String)
   ms_nombre = sValue
End Property
Property Get nombre() As String
   nombre = ms_nombre
End Property
Property Let IdPais(lValue As Long)
   ml_IdPais = lValue
End Property
Property Get IdPais() As Long
   IdPais = ml_IdPais
End Property



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DONacionalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Nombre As String
Dim ml_IdPais As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Descripcion(sValue As String)
   ms_Nombre = sValue
End Property
Property Get Descripcion() As String
   nombre = ms_Nombre
End Property
Property Let IdPais(lValue As Long)
   ml_IdPais = lValue
End Property
Property Get IdPais() As Long
   IdPais = ml_IdPais
End Property


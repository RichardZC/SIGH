VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DODistrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Distritos
'        Programado por: Castro W
'        Fecha: Agosto 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdProvincia As Long
Dim ms_Nombre As String
Dim ml_IdDistrito As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdProvincia(lValue As Long)
   ml_IdProvincia = lValue
End Property
Property Get IdProvincia() As Long
   IdProvincia = ml_IdProvincia
End Property
Property Let Nombre(sValue As String)
   ms_Nombre = sValue
End Property
Property Get Nombre() As String
   Nombre = ms_Nombre
End Property
Property Let IdDistrito(lValue As Long)
   ml_IdDistrito = lValue
End Property
Property Get IdDistrito() As Long
   IdDistrito = ml_IdDistrito
End Property



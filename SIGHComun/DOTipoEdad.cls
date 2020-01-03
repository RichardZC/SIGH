VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTipoEdad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TiposEdad
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Descripcion As String
Dim ms_Codigo As String
Dim ml_IdTipoEdad As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let IdTipoEdad(lValue As Long)
   ml_IdTipoEdad = lValue
End Property
Property Get IdTipoEdad() As Long
   IdTipoEdad = ml_IdTipoEdad
End Property



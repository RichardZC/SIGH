VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCatalogoServicioSubGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactCatalogoServiciosSubGrupo
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Codigo As String
Dim ml_IdServicioGrupo As Long
Dim ms_Descripcion As String
Dim ml_IdServicioSubGrupo As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let IdServicioGrupo(lValue As Long)
   ml_IdServicioGrupo = lValue
End Property
Property Get IdServicioGrupo() As Long
   IdServicioGrupo = ml_IdServicioGrupo
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdServicioSubGrupo(lValue As Long)
   ml_IdServicioSubGrupo = lValue
End Property
Property Get IdServicioSubGrupo() As Long
   IdServicioSubGrupo = ml_IdServicioSubGrupo
End Property



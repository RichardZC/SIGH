VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFuaEstadosTrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla SisFuaEstadosTrama
'        Programado por: Barrantes D
'        Fecha: Julio 2013
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_id As Long
Dim ms_tabla As String
Dim ms_campo As String
Dim mb_estado As Boolean

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
Property Let tabla(sValue As String)
   ms_tabla = sValue
End Property
Property Get tabla() As String
   tabla = ms_tabla
End Property
Property Let campo(sValue As String)
   ms_campo = sValue
End Property
Property Get campo() As String
   campo = ms_campo
End Property
Property Let estado(bValue As Boolean)
   mb_estado = bValue
End Property
Property Get estado() As Boolean
   estado = mb_estado
End Property


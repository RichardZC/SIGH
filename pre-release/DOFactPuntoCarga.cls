VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFactPuntoCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactPuntosCarga
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_TipoPunto As String
Dim ms_Descripcion As String
Dim ml_idPuntoCarga As Long
Dim ml_IdUPS As Long
Dim ml_IdServicio As Long

Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property

Property Let IdUPS(lValue As Long)
   ml_IdUPS = lValue
End Property
Property Get IdUPS() As Long
   IdUPS = ml_IdUPS
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let TipoPunto(sValue As String)
   ms_TipoPunto = sValue
End Property
Property Get TipoPunto() As String
   TipoPunto = ms_TipoPunto
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let idPuntoCarga(lValue As Long)
   ml_idPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property



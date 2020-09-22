VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEnfermeria_TratamientoDosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Enfermeria_tratamiento
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdVisita As Long
Dim ml_IdDiaVisita As Long
Dim ml_IdReceta As Long
Dim ml_idItem As Long
Dim ml_Dosis As Long
Dim ml_DatoProrenata As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdVisita(iValue As Long)
   ml_IdVisita = iValue
End Property
Property Get IdVisita() As Long
   IdVisita = ml_IdVisita
End Property
Property Let IdDiaVisita(iValue As Long)
   ml_IdDiaVisita = iValue
End Property
Property Get IdDiaVisita() As Long
   IdDiaVisita = ml_IdDiaVisita
End Property
Property Let IdReceta(iValue As Long)
   ml_IdReceta = iValue
End Property
Property Get IdReceta() As Long
   IdReceta = ml_IdReceta
End Property
Property Let idItem(iValue As Long)
   ml_idItem = iValue
End Property
Property Get idItem() As Long
   idItem = ml_idItem
End Property
Property Let Dosis(iValue As Long)
   ml_Dosis = iValue
End Property
Property Get Dosis() As Long
   Dosis = ml_Dosis
End Property
Property Let DatoProrenata(iValue As Long)
   ml_DatoProrenata = iValue
End Property
Property Get DatoProrenata() As Long
   DatoProrenata = ml_DatoProrenata
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOImagSalaPuntoCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ImagSalaPuntoCarga
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdSala As Long
Dim ml_IdPuntoCarga As Long
Dim mb_EsActivo As Boolean
Dim mda_FechsCrea As Date
Dim mda_FechaEdita As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdSala(iValue As Long)
   ml_IdSala = iValue
End Property
Property Get IdSala() As Long
   IdSala = ml_IdSala
End Property
Property Let IdPuntoCarga(iValue As Long)
   ml_IdPuntoCarga = iValue
End Property
Property Get IdPuntoCarga() As Long
   IdPuntoCarga = ml_IdPuntoCarga
End Property
Property Let EsActivo(bValue As Boolean)
   mb_EsActivo = bValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = mb_EsActivo
End Property
Property Let FechsCrea(daValue As Date)
   mda_FechsCrea = daValue
End Property
Property Get FechsCrea() As Date
   FechsCrea = mda_FechsCrea
End Property
Property Let FechaEdita(daValue As Date)
   mda_FechaEdita = daValue
End Property
Property Get FechaEdita() As Date
   FechaEdita = mda_FechaEdita
End Property

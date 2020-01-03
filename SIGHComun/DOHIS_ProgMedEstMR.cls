VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_ProgMedEstMR"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_ProgMedEstMR
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdHisProgMedEstMR As Long
Dim ml_idMedico As Long
Dim ml_IdServicio As Long
Dim ml_IdEstablecimiento As Long
Dim mda_FechaProgramada As Date
Dim ml_idTurno As Long
Dim ml_IdUsuarioAuditoria As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdHisProgMedEstMR(iValue As Long)
   ml_IdHisProgMedEstMR = iValue
End Property
Property Get IdHisProgMedEstMR() As Long
   IdHisProgMedEstMR = ml_IdHisProgMedEstMR
End Property
Property Let idMedico(iValue As Long)
   ml_idMedico = iValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let IdServicio(iValue As Long)
   ml_IdServicio = iValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let IdEstablecimiento(iValue As Long)
   ml_IdEstablecimiento = iValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let FechaProgramada(daValue As Date)
   mda_FechaProgramada = daValue
End Property
Property Get FechaProgramada() As Date
   FechaProgramada = mda_FechaProgramada
End Property
Property Let idTurno(iValue As Long)
   ml_idTurno = iValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property



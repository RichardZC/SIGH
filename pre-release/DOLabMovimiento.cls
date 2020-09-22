VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOLabMovimiento"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabMovimiento
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ms_movTipo As String
Dim ml_idTipoConcepto As Long
Dim ml_idPuntoCarga As Long
Dim mda_fecha As Date
Dim ml_idUsuario As Long
Dim ml_IdLabEstado As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property

Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property

Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property

Property Let movTipo(sValue As String)
   ms_movTipo = sValue
End Property

Property Get movTipo() As String
   movTipo = ms_movTipo
End Property

Property Let idTipoConcepto(iValue As Long)
   ml_idTipoConcepto = iValue
End Property

Property Get idTipoConcepto() As Long
   idTipoConcepto = ml_idTipoConcepto
End Property

Property Let idPuntoCarga(iValue As Long)
   ml_idPuntoCarga = iValue
End Property

Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property

Property Let fecha(daValue As Date)
   mda_fecha = daValue
End Property

Property Get fecha() As Date
   fecha = mda_fecha
End Property

Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property

Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let IdLabEstado(iValue As Long)
   ml_IdLabEstado = iValue
End Property

Property Get IdLabEstado() As Long
   IdLabEstado = ml_IdLabEstado
End Property


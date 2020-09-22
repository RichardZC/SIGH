VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOfarmAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmInventarioCabecera
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_idAlmacen As Long
Dim ms_Descripcion As String
Dim ms_idTipoLocales As String
Dim ms_idTipoSuministro As String
Dim ml_idEstado As Long
Dim ml_codigoSISMED As String
Dim ml_regenerarDias As String
Dim ml_regenerarHora As String
Dim ml_regenerarEstado As String
Dim ml_esUnidosis As Long
Property Get esUnidosis() As Long
   esUnidosis = ml_esUnidosis
End Property
Property Let esUnidosis(iValue As Long)
   ml_esUnidosis = iValue
End Property

Property Let regenerarEstado(iValue As String)
   ml_regenerarEstado = iValue
End Property
Property Get regenerarEstado() As String
   regenerarEstado = ml_regenerarEstado
End Property

Property Let regenerarHora(iValue As String)
   ml_regenerarHora = iValue
End Property
Property Get regenerarHora() As String
   regenerarHora = ml_regenerarHora
End Property


Property Let regenerarDias(iValue As String)
   ml_regenerarDias = iValue
End Property
Property Get regenerarDias() As String
   regenerarDias = ml_regenerarDias
End Property


Property Let codigoSISMED(iValue As String)
   ml_codigoSISMED = iValue
End Property
Property Get codigoSISMED() As String
   codigoSISMED = ml_codigoSISMED
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let idAlmacen(iValue As Long)
   ml_idAlmacen = iValue
End Property
Property Get idAlmacen() As Long
   idAlmacen = ml_idAlmacen
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let idTipoLocales(sValue As String)
   ms_idTipoLocales = sValue
End Property
Property Get idTipoLocales() As String
   idTipoLocales = ms_idTipoLocales
End Property
Property Let idTipoSuministro(sValue As String)
   ms_idTipoSuministro = sValue
End Property
Property Get idTipoSuministro() As String
   idTipoSuministro = ms_idTipoSuministro
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property

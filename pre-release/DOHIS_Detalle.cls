VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHIS_Detalle"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HIS_Detalle
'        Programado por: Cachay F
'        Fecha: Octubre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisDetalle As Long
Dim ml_IdHisCabecera As Long
Dim ml_IdTipoAtencion As Long
Dim ml_DiaAtencion As Long
Dim ml_IdHisPaciente As Long
Dim ms_CodigoActividad As String
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdDistrito As Long
Dim ml_IdTipoEdad As Long
Dim ml_Edad As Long
Dim ml_Talla As Long
Dim ms_Peso As String
Dim ml_IdEstadoaEstablec As Long
Dim ml_IdEstadoaServicio As Long
Dim ml_NroRegistroLote As Integer
Dim ml_NroRegistroHoja As Integer

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdHisDetalle(iValue As Long)
   ml_IdHisDetalle = iValue
End Property
Property Get IdHisDetalle() As Long
   IdHisDetalle = ml_IdHisDetalle
End Property
Property Let IdHisCabecera(iValue As Long)
   ml_IdHisCabecera = iValue
End Property
Property Get IdHisCabecera() As Long
   IdHisCabecera = ml_IdHisCabecera
End Property
Property Let IdTipoAtencion(iValue As Long)
   ml_IdTipoAtencion = iValue
End Property
Property Get IdTipoAtencion() As Long
   IdTipoAtencion = ml_IdTipoAtencion
End Property
Property Let DiaAtencion(iValue As Long)
   ml_DiaAtencion = iValue
End Property
Property Get DiaAtencion() As Long
   DiaAtencion = ml_DiaAtencion
End Property
Property Let IdHisPaciente(iValue As Long)
   ml_IdHisPaciente = iValue
End Property
Property Get IdHisPaciente() As Long
   IdHisPaciente = ml_IdHisPaciente
End Property
Property Let CodigoActividad(sValue As String)
   ms_CodigoActividad = sValue
End Property
Property Get CodigoActividad() As String
   CodigoActividad = ms_CodigoActividad
End Property
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let IdDistrito(iValue As Long)
   ml_IdDistrito = iValue
End Property
Property Get IdDistrito() As Long
   IdDistrito = ml_IdDistrito
End Property
Property Let IdTipoEdad(iValue As Long)
   ml_IdTipoEdad = iValue
End Property
Property Get IdTipoEdad() As Long
   IdTipoEdad = ml_IdTipoEdad
End Property
Property Let Edad(iValue As Long)
   ml_Edad = iValue
End Property
Property Get Edad() As Long
   Edad = ml_Edad
End Property
Property Let Talla(iValue As Long)
   ml_Talla = iValue
End Property
Property Get Talla() As Long
   Talla = ml_Talla
End Property
Property Let Peso(sValue As String)
   ms_Peso = sValue
End Property
Property Get Peso() As String
   Peso = ms_Peso
End Property
Property Let IdEstadoaEstablec(iValue As Long)
   ml_IdEstadoaEstablec = iValue
End Property
Property Get IdEstadoaEstablec() As Long
   IdEstadoaEstablec = ml_IdEstadoaEstablec
End Property
Property Let IdEstadoaServicio(iValue As Long)
   ml_IdEstadoaServicio = iValue
End Property
Property Get IdEstadoaServicio() As Long
   IdEstadoaServicio = ml_IdEstadoaServicio
End Property
Property Let NroRegistroLote(iValue As Integer)
   ml_NroRegistroLote = iValue
End Property
Property Get NroRegistroLote() As Integer
   NroRegistroLote = ml_NroRegistroLote
End Property
Property Let NroRegistroHoja(iValue As Integer)
   ml_NroRegistroHoja = iValue
End Property
Property Get NroRegistroHoja() As Integer
   NroRegistroHoja = ml_NroRegistroHoja
End Property


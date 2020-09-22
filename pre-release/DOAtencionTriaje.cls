VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionTriaje
'        Programado por: Barrantes D
'        Fecha: Febrero 2011
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdAtencion As Long
Dim ms_Presion As String
Dim ms_Temperatura As String
Dim ms_Peso As String
Dim ms_Talla As String
Dim mda_FechaTriaje As Date
Dim ml_IdUsuarioCreo As Long
Dim mda_FechaModifico As Date
Dim ml_IdUsuarioModifico As Long


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let Presion(sValue As String)
   ms_Presion = sValue
End Property
Property Get Presion() As String
   Presion = ms_Presion
End Property
Property Let Temperatura(sValue As String)
   ms_Temperatura = sValue
End Property
Property Get Temperatura() As String
   Temperatura = ms_Temperatura
End Property
Property Let Peso(sValue As String)
   ms_Peso = sValue
End Property
Property Get Peso() As String
   Peso = ms_Peso
End Property
Property Let Talla(sValue As String)
   ms_Talla = sValue
End Property
Property Get Talla() As String
   Talla = ms_Talla
End Property
Property Let FechaTriaje(daValue As Date)
   mda_FechaTriaje = daValue
End Property
Property Get FechaTriaje() As Date
   FechaTriaje = mda_FechaTriaje
End Property
Property Let IdUsuarioCreo(iValue As Long)
   ml_IdUsuarioCreo = iValue
End Property
Property Get IdUsuarioCreo() As Long
   IdUsuarioCreo = ml_IdUsuarioCreo
End Property
Property Let FechaModifico(daValue As Date)
   mda_FechaModifico = daValue
End Property
Property Get FechaModifico() As Date
   FechaModifico = mda_FechaModifico
End Property
Property Let IdUsuarioModifico(iValue As Long)
   ml_IdUsuarioModifico = iValue
End Property
Property Get IdUsuarioModifico() As Long
   IdUsuarioModifico = ml_IdUsuarioModifico
End Property



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEstanciaHospitalaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesEstanciaHospitalaria
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_Auditoria As Long
Dim md_DiasEstancia As Double
Dim ml_IdAtencion As Long
Dim ml_IdFacturacionServicio As Long
Dim ml_IdMedicoOrdena As Long
Dim ml_IdCama As Long
Dim ml_IdServicio As Long
Dim ms_HoraDesocupacion As String
Dim mda_FechaDesocupacion As Date
Dim ms_HoraOcupacion As String
Dim mda_FechaOcupacion As Date
Dim ml_Secuencia As Long
Dim ml_IdEstanciaHospitalaria As Long
Dim ml_LlegoAlServicio As Long
Dim ml_IdProducto As Long
Dim ml_IdMedicoOrdenaOrigen As Long

Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property


Property Let LlegoAlServicio(lValue As Long)
   ml_LlegoAlServicio = lValue
End Property
Property Get LlegoAlServicio() As Long
   LlegoAlServicio = ml_LlegoAlServicio
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let DiasEstancia(dValue As Double)
   md_DiasEstancia = dValue
End Property
Property Get DiasEstancia() As Double
   DiasEstancia = md_DiasEstancia
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdFacturacionServicio(lValue As Long)
   ml_IdFacturacionServicio = lValue
End Property
Property Get IdFacturacionServicio() As Long
   IdFacturacionServicio = ml_IdFacturacionServicio
End Property
Property Let IdMedicoOrdena(lValue As Long)
   ml_IdMedicoOrdena = lValue
End Property
Property Get IdMedicoOrdena() As Long
   IdMedicoOrdena = ml_IdMedicoOrdena
End Property
Property Let IdCama(lValue As Long)
   ml_IdCama = lValue
End Property
Property Get IdCama() As Long
   IdCama = ml_IdCama
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let HoraDesocupacion(sValue As String)
   ms_HoraDesocupacion = sValue
End Property
Property Get HoraDesocupacion() As String
   HoraDesocupacion = ms_HoraDesocupacion
End Property
Property Let FechaDesocupacion(daValue As Date)
   mda_FechaDesocupacion = daValue
End Property
Property Get FechaDesocupacion() As Date
   FechaDesocupacion = mda_FechaDesocupacion
End Property
Property Let HoraOcupacion(sValue As String)
   ms_HoraOcupacion = sValue
End Property
Property Get HoraOcupacion() As String
   HoraOcupacion = ms_HoraOcupacion
End Property
Property Let FechaOcupacion(daValue As Date)
   mda_FechaOcupacion = daValue
End Property
Property Get FechaOcupacion() As Date
   FechaOcupacion = mda_FechaOcupacion
End Property
Property Let Secuencia(lValue As Long)
   ml_Secuencia = lValue
End Property
Property Get Secuencia() As Long
   Secuencia = ml_Secuencia
End Property
Property Let IdEstanciaHospitalaria(lValue As Long)
   ml_IdEstanciaHospitalaria = lValue
End Property
Property Get IdEstanciaHospitalaria() As Long
   IdEstanciaHospitalaria = ml_IdEstanciaHospitalaria
End Property
'actualizado 20142009
Property Let IdMedicoOrdenaOrigen(lValue As Long)
   ml_IdMedicoOrdenaOrigen = lValue
End Property
Property Get IdMedicoOrdenaOrigen() As Long
   IdMedicoOrdenaOrigen = ml_IdMedicoOrdenaOrigen
End Property



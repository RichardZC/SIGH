VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Servicios
'        Programado por: Castro W
'        Fecha: Marzo 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_idProducto As Long
Dim ms_SVG As String
Dim ms_Codigo As String
Dim ml_IdServicio As Long
Dim ms_nombre As String
Dim ml_IdEspecialidad As Long
Dim ml_IdTipoServicio As Long
Dim ml_soloTipoSexo As Long
Dim ml_maximaEdad As Long
Dim ms_codigoServicioSEM As String
Dim ms_ubicacionSEM As String
Dim ms_codigoServicioHIS As String
Dim ms_CostoCeroCE As String
Dim ml_minimaEdad As Long
Dim ml_idEstado As Long
Dim mb_Triaje As Boolean
Dim mb_EsObservacionEmergencia As Boolean
Dim mb_UsaModuloNinoSano As Boolean
Dim mb_UsaModuloMaterno As Boolean
Dim mb_UsaGalenHos As Boolean
Dim ml_TipoEdad As Long
Dim mb_UsaFUA As Boolean
Dim ms_codigoServicioSuSalud As String
Dim ms_codigoServicioFUA As String
Dim ml_FuaTipoAnexo2015 As Long
Dim ml_MaxCuposCitasAdelantadas As Long
Dim ml_MaxCuposAdicionales As Long
Dim ml_MaxCuposCitasHoySIS As Long
Dim ml_MaxCuposCitasAdelandatasSIS As Long
Dim ml_usaNacimiento As Long
Dim ms_emergenciaCorrelativo As String
Dim ml_NoUsaMensajeTexto As Long

Property Let NoUsaMensajeTexto(lValue As Long)
   ml_NoUsaMensajeTexto = lValue
End Property
Property Get NoUsaMensajeTexto() As Long
   NoUsaMensajeTexto = ml_NoUsaMensajeTexto
End Property


Property Let emergenciaCorrelativo(sValue As String)
   ms_emergenciaCorrelativo = sValue
End Property
Property Get emergenciaCorrelativo() As String
   emergenciaCorrelativo = ms_emergenciaCorrelativo
End Property


Property Let usaNacimiento(lValue As Long)
   ml_usaNacimiento = lValue
End Property
Property Get usaNacimiento() As Long
   usaNacimiento = ml_usaNacimiento
End Property

Property Let MaxCuposCitasAdelandatasSIS(lValue As Long)
   ml_MaxCuposCitasAdelandatasSIS = lValue
End Property
Property Get MaxCuposCitasAdelandatasSIS() As Long
   MaxCuposCitasAdelandatasSIS = ml_MaxCuposCitasAdelandatasSIS
End Property

Property Let MaxCuposCitasHoySIS(lValue As Long)
   ml_MaxCuposCitasHoySIS = lValue
End Property
Property Get MaxCuposCitasHoySIS() As Long
   MaxCuposCitasHoySIS = ml_MaxCuposCitasHoySIS
End Property


Property Let MaxCuposCitasAdelantadas(lValue As Long)
   ml_MaxCuposCitasAdelantadas = lValue
End Property
Property Get MaxCuposCitasAdelantadas() As Long
   MaxCuposCitasAdelantadas = ml_MaxCuposCitasAdelantadas
End Property

Property Let MaxCuposAdicionales(lValue As Long)
   ml_MaxCuposAdicionales = lValue
End Property
Property Get MaxCuposAdicionales() As Long
   MaxCuposAdicionales = ml_MaxCuposAdicionales
End Property

Property Let FuaTipoAnexo2015(lValue As Long)
   ml_FuaTipoAnexo2015 = lValue
End Property
Property Get FuaTipoAnexo2015() As Long
   FuaTipoAnexo2015 = ml_FuaTipoAnexo2015
End Property
Property Let codigoServicioFUA(sValue As String)
   ms_codigoServicioFUA = sValue
End Property
Property Get codigoServicioFUA() As String
   codigoServicioFUA = ms_codigoServicioFUA
End Property
Property Let codigoServicioSuSalud(sValue As String)
   ms_codigoServicioSuSalud = sValue
End Property
Property Get codigoServicioSuSalud() As String
   codigoServicioSuSalud = ms_codigoServicioSuSalud
End Property
Property Let UsaFUA(bValue As Boolean)
   mb_UsaFUA = bValue
End Property
Property Get UsaFUA() As Boolean
   UsaFUA = mb_UsaFUA
End Property

Property Let TipoEdad(lValue As Long)
   ml_TipoEdad = lValue
End Property
Property Get TipoEdad() As Long
   TipoEdad = ml_TipoEdad
End Property

Property Let UsaGalenHos(bValue As Boolean)
   mb_UsaGalenHos = bValue
End Property
Property Get UsaGalenHos() As Boolean
   UsaGalenHos = mb_UsaGalenHos
End Property

Property Let UsaModuloMaterno(bValue As Boolean)
   mb_UsaModuloMaterno = bValue
End Property
Property Get UsaModuloMaterno() As Boolean
   UsaModuloMaterno = mb_UsaModuloMaterno
End Property


Property Let UsaModuloNinoSano(bValue As Boolean)
   mb_UsaModuloNinoSano = bValue
End Property
Property Get UsaModuloNinoSano() As Boolean
   UsaModuloNinoSano = mb_UsaModuloNinoSano
End Property


Property Let EsObservacionEmergencia(bValue As Boolean)
   mb_EsObservacionEmergencia = bValue
End Property
Property Get EsObservacionEmergencia() As Boolean
   EsObservacionEmergencia = mb_EsObservacionEmergencia
End Property



Property Let Triaje(bValue As Boolean)
   mb_Triaje = bValue
End Property
Property Get Triaje() As Boolean
   Triaje = mb_Triaje
End Property

Property Let idEstado(lValue As Long)
   ml_idEstado = lValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property

Property Let minimaEdad(lValue As Long)
   ml_minimaEdad = lValue
End Property
Property Get minimaEdad() As Long
   minimaEdad = ml_minimaEdad
End Property

Property Let CostoCeroCE(sValue As String)
   ms_CostoCeroCE = sValue
End Property
Property Get CostoCeroCE() As String
   CostoCeroCE = ms_CostoCeroCE
End Property
Property Let codigoServicioHIS(sValue As String)
   ms_codigoServicioHIS = sValue
End Property
Property Get codigoServicioHIS() As String
   codigoServicioHIS = ms_codigoServicioHIS
End Property
Property Let ubicacionSEM(sValue As String)
   ms_ubicacionSEM = sValue
End Property
Property Get ubicacionSEM() As String
   ubicacionSEM = ms_ubicacionSEM
End Property
Property Let codigoServicioSEM(sValue As String)
   ms_codigoServicioSEM = sValue
End Property
Property Get codigoServicioSEM() As String
   codigoServicioSEM = ms_codigoServicioSEM
End Property
Property Let maximaEdad(lValue As Long)
   ml_maximaEdad = lValue
End Property
Property Get maximaEdad() As Long
   maximaEdad = ml_maximaEdad
End Property
Property Let soloTipoSexo(lValue As Long)
   ml_soloTipoSexo = lValue
End Property
Property Get soloTipoSexo() As Long
   soloTipoSexo = ml_soloTipoSexo
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let nombre(sValue As String)
   ms_nombre = sValue
End Property
Property Get nombre() As String
   nombre = ms_nombre
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let idProducto(lValue As Long)
   ml_idProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property
Property Let SVG(sValue As String)
   ms_SVG = sValue
End Property
Property Get SVG() As String
   SVG = ms_SVG
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property


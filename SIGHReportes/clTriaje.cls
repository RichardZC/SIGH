VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para el mantenimiento de Triaje o Signos Vitales
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim mi_Opcion As sghOpciones
Dim ml_idAtencion As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim oFormulario As New Triaje
Dim ml_triajeOrigen As sightriajeorigen
Dim ml_idCuentaAtencion As Long
Dim mb_EsAtencionCRED As Boolean
Dim mb_GuardoTriaje As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property

Property Let TriajeOrigen(lValue As sightriajeorigen)
   ml_triajeOrigen = lValue
End Property

Property Get TriajeOrigen() As sightriajeorigen
   TriajeOrigen = ml_triajeOrigen
End Property

Property Let GuardoTriaje(bValue As Boolean)
   mb_GuardoTriaje = bValue
End Property
Property Get GuardoTriaje() As Boolean
   GuardoTriaje = mb_GuardoTriaje
End Property

Sub MostrarFormulario()
    oFormulario.Opcion = mi_Opcion
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.idAtencion = ml_idAtencion
    oFormulario.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oFormulario.lcNombrePc = mo_lcNombrePc
    oFormulario.TriajeOrigen = IIf(ml_triajeOrigen = 0, sightriajeorigen.Triaje, ml_triajeOrigen)
    oFormulario.IdCuentaAtencion = ml_idCuentaAtencion
    oFormulario.EsAtencionCRED = mb_EsAtencionCRED
    oFormulario.GuardoTriaje = False
    oFormulario.Show 1
    mb_GuardoTriaje = oFormulario.GuardoTriaje
    Set oFormulario = Nothing
End Sub

Sub MostrarFormularioDesdeAtenciones(lTriajeOrigen As sightriajeorigen, lIdCuentaAtencion As Long, _
                Optional bEsAtencionCRED As Boolean = False)
    ml_triajeOrigen = lTriajeOrigen
    mo_lnIdTablaLISTBARITEMS = 1303
    ml_idCuentaAtencion = lIdCuentaAtencion
    mb_EsAtencionCRED = bEsAtencionCRED
    MostrarFormulario
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaMedicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Médico
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oFormulario As New MedicosBusqueda
Dim ml_IdRegistroSeleccionado As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdEspecialidad  As Long
Dim lcHoraProgramada As String
Dim ldFechaProgramada As Date
Dim lnIdTipoServicio  As Long
Dim lcNombreMedico  As String
Dim ml_NoMuestraInactivos As Boolean
Property Let NoMuestraInactivos(lValue As Boolean)
    ml_NoMuestraInactivos = lValue
End Property

Property Let NombreMedico(lValue As String)
    lcNombreMedico = lValue
End Property

Property Let IdEspecialidad(lValue As Long)
    ml_IdEspecialidad = lValue
End Property

Property Let HoraProgramada(lValue As String)
    lcHoraProgramada = lValue
End Property
Property Let FechaProgramada(lValue As Date)
    ldFechaProgramada = lValue
End Property

Property Let idTipoServicio(lValue As Long)
    lnIdTipoServicio = lValue
End Property



Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property

Sub MostrarFormulario()
    oFormulario.IdEspecialidad = ml_IdEspecialidad
    oFormulario.FechaProgramada = ldFechaProgramada
    oFormulario.HoraProgramada = lcHoraProgramada
    oFormulario.idTipoServicio = lnIdTipoServicio
    oFormulario.NombreMedico = lcNombreMedico
    oFormulario.NoMuestraInactivos = ml_NoMuestraInactivos
    oFormulario.Show 1
    ml_IdRegistroSeleccionado = oFormulario.IdRegistroSeleccionado
    mi_BotonPresionado = oFormulario.BotonPresionado
    Set oFormulario = Nothing
End Sub

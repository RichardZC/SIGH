VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clBuscaReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Receta
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_RecetaBusqueda As New SIGHNegocios.BuscaReceta
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdRecetaSeleccionada As Long
Dim ml_IdPuntoCarga As sghPuntosCargaBasicos
Dim ml_idCuentaAtencion As Long

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property
Property Let IdPuntoCarga(lValue As sghPuntosCargaBasicos)
    ml_IdPuntoCarga = lValue
End Property


Sub MostrarFormulario()
    mo_RecetaBusqueda.IdPuntoCarga = ml_IdPuntoCarga
    mo_RecetaBusqueda.idCuentaAtencion = ml_idCuentaAtencion
    mo_RecetaBusqueda.Show 1
    ml_IdRecetaSeleccionada = mo_RecetaBusqueda.idReceta
    mi_BotonPresionado = mo_RecetaBusqueda.BotonPresionado
    ml_IdPuntoCarga = mo_RecetaBusqueda.IdPuntoCarga
    Set mo_RecetaBusqueda = Nothing
End Sub

Property Let IdRecetaSeleccionada(lValue As Long)
    ml_IdRecetaSeleccionada = lValue
End Property
Property Get IdRecetaSeleccionada() As Long
    IdRecetaSeleccionada = ml_IdRecetaSeleccionada
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Get IdPuntoCarga() As Long
    IdPuntoCarga = ml_IdPuntoCarga
End Property


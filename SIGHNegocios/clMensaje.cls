VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Mensajes en los formularios
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oBusqueda As New SIGHNegocios.Mensajes
Dim mi_BotonPresionado As sghBotonDetallePresionado
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Sub MostrarFormulario(Mensaje As String, Titulo As String, Optional LongitudLetra As Integer, _
                      Optional EsNegrita As Boolean, Optional ColorLetra As sghColores, _
                      Optional UsaBotonesAceptarCancelar As Boolean)
    If LongitudLetra > 0 Then
       oBusqueda.LongitudLetra = LongitudLetra
    End If
    If EsNegrita = True Then
       oBusqueda.EsNegrita = EsNegrita
    End If
    If ColorLetra > 0 Then
       oBusqueda.ColorLetra = ColorLetra
    End If
    oBusqueda.Mensaje = Mensaje
    oBusqueda.Titulo = Titulo
    oBusqueda.UsaBotonesAceptarCancelar = UsaBotonesAceptarCancelar
    oBusqueda.Show 1
    mi_BotonPresionado = oBusqueda.BotonPresionado
    Set oBusqueda = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Archivo
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ArchivoBusqueda As New ArchivoBusqueda
Dim mo_ArchivoBusqueda1 As New Camara1
Dim lb_MuestraImagen As Boolean
Dim lc_PathDefault As String
Dim lc_ArchivoElegido As String
Dim lc_TipoArchivo As String     'DEBB2014a

'DEBB2014a
Property Let TipoArchivo(lValue As String)
    lc_TipoArchivo = lValue
End Property
Property Get ArchivoElegido() As String
    ArchivoElegido = lc_ArchivoElegido
End Property
Property Let PathDefault(lValue As String)
    lc_PathDefault = lValue
End Property

Property Let MuestraImagen(lValue As Boolean)
    lb_MuestraImagen = lValue
End Property


Sub MostrarFormulario()
    mo_ArchivoBusqueda.MuestraImagen = lb_MuestraImagen
    mo_ArchivoBusqueda.PathDefault = lc_PathDefault
    mo_ArchivoBusqueda.TipoArchivo = lc_TipoArchivo   'DEBB2014a
    mo_ArchivoBusqueda.Show 1
    lc_ArchivoElegido = mo_ArchivoBusqueda.ArchivoElegido
    Set mo_ArchivoBusqueda = Nothing
End Sub

Sub MuestraCamara()
    mo_ArchivoBusqueda1.Show 1
End Sub


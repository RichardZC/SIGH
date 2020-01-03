VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsumoWebUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim ms_MensajeError As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


Public Function PeticionHttpEsCorrecta(oHttReq As XMLHTTPRequest) As Boolean
    PeticionHttpEsCorrecta = False
    If oHttReq.Status = sghEstadoPeticionHttp.RespuestaOk Then
        PeticionHttpEsCorrecta = True
    End If
End Function

Public Function ParseMensageError(oHttReq As XMLHTTPRequest)
    Select Case oHttReq.Status
        Case sghEstadoPeticionHttp.RespuestaOk:
            ms_MensajeError = ""
            
        Case sghEstadoPeticionHttp.PagNoEncontrada:
            ms_MensajeError = "Paginá no encontrada en el servidor que contiene el servicio solicitado," & _
                            Chr(13) & " verifique que la dirección del servicio solicitado este " & _
                            " especificada correctamente"
                            
        Case sghEstadoPeticionHttp.ErrorInternoServidor:
            ms_MensajeError = "El Servidor que contiene el servicio solicitado a devuelto un error :" & _
                                oHttReq.statusText
            
        Case sghEstadoPeticionHttp.ServidorNoEncontrado:
            ms_MensajeError = "Servidor que contiene el servicio, no ha sido encontrado, " & _
                            Chr(13) & " verifique que la dirección del servicio solicitado " & _
                            "este especificada correctamente"
                            
        Case sghEstadoPeticionHttp.TiempoEsperaAgotado:
            ms_MensajeError = "Tiempo de Espera agotado para la solicitud, Esto se puede deber a: " & _
                                Chr(13) & "Dirección que contiene el servicio solicitado no es la correcta, " & _
                                Chr(13) & "a lentitud en su conexón a internet o la pagina que contiene el servicio"
                                
        Case sghEstadoPeticionHttp.SinConexion:
            ms_MensajeError = "No se tiene conexion a internet: " & _
                                Chr(13) & "- Verifique que su cable este conectado y operativo" & _
                                Chr(13) & "- Si usa una conexión inalambrica asegurese que esta conectado a su punto de acceso" & _
                                Chr(13) & "- Si se usa un servidor proxy para la salida a internet, asegurese " & _
                                "que su Computadora este configurada correctamente"
                 
                 
        Case Else:
            ms_MensajeError = "Error desconocido :" & oHttReq.statusText
    End Select
    ParseMensageError = ms_MensajeError
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisConsumoWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para  WEB CON EL SIS
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado

Private Const SIS_BuscarEESSxCodigo = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<Buscar_EESSxCodigo xmlns=""http://www.sis.gob.pe/"">" & _
              "<IdEess>string</IdEess>" & _
              "<Usuario>string</Usuario>" & _
              "<Clave>string</Clave>" & _
            "</Buscar_EESSxCodigo>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"

Private Const SIS_FFamaceuticaxCodigo = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<Buscar_FFamaceuticaxCodigo xmlns=""http://www.sis.gob.pe/"">" & _
              "<Codigo>string</Codigo>" & _
              "<Usuario>string</Usuario>" & _
              "<Clave>string</Clave>" & _
            "</Buscar_FFamaceuticaxCodigo>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"
        
Private Const SIS_Buscar_InsumosxCodigo = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<Buscar_InsumosxCodigo xmlns=""http://www.sis.gob.pe/"">" & _
              "<Codigo>string</Codigo>" & _
              "<Usuario>string</Usuario>" & _
              "<Clave>string</Clave>" & _
            "</Buscar_InsumosxCodigo>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"

Private Const SIS_Buscar_MedicamentosxCodigo = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<Buscar_MedicamentosxCodigo xmlns=""http://www.sis.gob.pe/"">" & _
              "<Codigo>string</Codigo>" & _
              "<Usuario>string</Usuario>" & _
              "<Clave>string</Clave>" & _
            "</Buscar_MedicamentosxCodigo>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"

Private Const SIS_Buscar_PerSaludxNroDoc = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<Buscar_PerSaludxNroDoc xmlns=""http://www.sis.gob.pe/"">" & _
              "<Codigo>string</Codigo>" & _
              "<Usuario>string</Usuario>" & _
              "<Clave>string</Clave>" & _
            "</Buscar_PerSaludxNroDoc>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"
Private Const SIS_VerificaFUADuplicado = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<VerificaFUADuplicado xmlns=""http://www.sis.gob.pe/"">" & _
              "<Disa>string</Disa>" & _
              "<TipoFormato>string</TipoFormato>" & _
              "<Numero>string</Numero>" & _
            "</VerificaFUADuplicado>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"
Private Const SIS_Buscar_AfiliadoxNroDoc = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
            "<soap:Header>" & _
              "<CredencialWS xmlns=""http://www.sis.gob.pe/"">" & _
                "<Usuario>string</Usuario>" & _
                "<Clave>string</Clave>" & _
              "</CredencialWS>" & _
            "</soap:Header>" & _
            "<soap:Body>" & _
              "<BuscarAseguradosDocIdent xmlns=""http://www.sis.gob.pe/"">" & _
                "<IdEess>string</IdEess>" & _
                "<TipoDoc>string</TipoDoc>" & _
                "<Documento>string</Documento>" & _
              "</BuscarAseguradosDocIdent>" & _
            "</soap:Body>" & _
        "</soap:Envelope>"
'FRANK 03082015
Private Const SIS_Buscar_AfiliadoxNroAfiliado = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
            "<soap:Body>" & _
              "<BuscarAsegurados xmlns=""http://www.sis.gob.pe/"">" & _
                "<Disa>string</Disa>" & _
                "<TipoFormato>string</TipoFormato>" & _
                "<Contrato>string</Contrato>" & _
                "<Correlativo>string</Correlativo>" & _
                "<CodTabla>string</CodTabla>" & _
              "</BuscarAsegurados>" & _
            "</soap:Body>" & _
        "</soap:Envelope>"

Private Const wUsuario = "siges"
Private Const wClave = "pr2cnze10as"

' Definicion de Ruta de Servicio Web
Private Const RutaWeb = "http://app.sis.gob.pe/edi/maestros.asmx"
Private Const RutaWebAdi = "http://app.sis.gob.pe/edi/RecepcionTrama.asmx"

' Aqui se guardara la estructura de datos que se recibe del servicio web
Dim XmlStructura As String

Dim oConexion As New Connection
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


Sub EstableceConexion()
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
End Sub

'debb-21/10/2015
Function ConsultarServicioBuscarEESSxCodigo(ms_Codigo As String, ByRef rsTmp As Recordset) As Boolean
    If lcBuscaParametro.SeleccionaFilaParametro(322) <> "S" Then
       Exit Function
    End If
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    Dim lpos As Integer
    Dim lcDatosWeb As String
    Dim res() As String
    Dim wlcad As String
    Dim indice As Integer
    Dim lbSalida As Boolean
    
    Dim oM_eess As New m_eess
    Dim ODom_eess As New Dom_eess
    
 
    Set rsTmp = m_eessEleccionarPorCodigoRenaes(ms_Codigo)
    If rsTmp.RecordCount > 0 Then
         lbSalida = True
    Else
        Set parser = New DOMDocument
        parser.loadXML SIS_BuscarEESSxCodigo
    
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_EESSxCodigo/IdEess").Text = ms_Codigo
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_EESSxCodigo/Usuario").Text = lcBuscaParametro.SeleccionaFilaParametro(365)    'wUsuario
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_EESSxCodigo/Clave").Text = lcBuscaParametro.SeleccionaFilaParametro(366)   ' wClave
        
        Set oHttReq = New XMLHTTPRequest
        oHttReq.Open "POST", RutaWeb, False
        oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/Buscar_EESSxCodigo "
        On Error Resume Next
        oHttReq.send parser.XML
        If Err.Number > 0 Then
             MsgBox Err.Description
             Set parser = Nothing
             Set oHttReq = Nothing
             Set oConexion = Nothing
             Exit Function
        End If
        Set parser = New DOMDocument
        parser.loadXML oHttReq.responseText
        lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_EESSxCodigoResponse/Buscar_EESSxCodigoResult").XML
        
        If Err.Number > 0 Then
'            MsgBox Err.Description 'Actualizado 07102014
            If Err.Number = 91 Then
              MsgBox "No hay conexiòn con el SIS, Verificar su conexciòn, Se buscarà en la base local", vbInformation, "PROBLEMAS CON LA WEB"
            Else
                MsgBox Err.Description
            End If
        Else
            res = XmlaArreglo(lcDatosWeb)
        End If
        If res(0) = "0" Then
            lbSalida = False
        Else
            EstableceConexion
            Set oM_eess.Conexion = oConexion
            
            'oDom_eess.IdUsuarioAuditoria=
            ODom_eess.pre_Afilia = res(2)
            ODom_eess.pre_Aisped = res(11)
            ODom_eess.pre_CodEjeAdm = res(8)
            ODom_eess.pre_CodigoRENAES = res(12)
            ODom_eess.pre_esmn = "N"
            ODom_eess.pre_idCategoriaEESS = res(4)
            ODom_eess.pre_IdDisa = res(5)
            ODom_eess.pre_IdEESS = res(0)
            ODom_eess.pre_IdEstado = res(13)
            ODom_eess.pre_IdOdsis = res(6)
            ODom_eess.pre_idUbigeo = res(7)
            ODom_eess.pre_nombre = (1)
            ODom_eess.pre_UCI = res(3)
            ODom_eess.pre_Umbral = res(10)
            ODom_eess.pre_Vrae = res(9)
            lbSalida = oM_eess.Insertar(ODom_eess)
        End If
        Set oHttReq = Nothing
        Set parser = Nothing
        Set oConexion = Nothing
    End If
    ConsultarServicioBuscarEESSxCodigo = lbSalida
End Function

'debb-21/10/2015
Function ConsultarServicioFFamaceuticaxCodigo(ms_Codigo As String, ByRef rsTmp As Recordset) As Boolean
    If lcBuscaParametro.SeleccionaFilaParametro(322) <> "S" Then
       Exit Function
    End If
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    Dim lpos As Integer
    Dim lcDatosWeb As String
    Dim res() As String
    Dim wlcad As String
    Dim indice As Integer

    Dim oA_presentaciones As New a_presentaciones
    Dim oDoa_presentaciones As New Doa_presentaciones
    Dim lbSalida As Boolean

 
        Set rsTmp = a_presentacionPorAbreviatura(ms_Codigo)
        If rsTmp.RecordCount > 0 Then
             lbSalida = True
        Else 'Modificado para busqueda web*************************************************

            Set parser = New DOMDocument
            parser.loadXML SIS_FFamaceuticaxCodigo
        
            parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_FFamaceuticaxCodigo/Codigo").Text = ms_Codigo
            parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_FFamaceuticaxCodigo/Usuario").Text = lcBuscaParametro.SeleccionaFilaParametro(365)   'wUsuario
            parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_FFamaceuticaxCodigo/Clave").Text = lcBuscaParametro.SeleccionaFilaParametro(366) 'wClave
        
        
            Set oHttReq = New XMLHTTPRequest
            oHttReq.Open "POST", RutaWeb, False
            oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
            oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/Buscar_FFamaceuticaxCodigo "
            On Error Resume Next
            oHttReq.send parser.XML
            If Err.Number > 0 Then
                 MsgBox Err.Description
                 Set parser = Nothing
                 Set oHttReq = Nothing
                 Set oConexion = Nothing
                 Exit Function
            End If
            Set parser = New DOMDocument
            parser.loadXML oHttReq.responseText
            lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_FFamaceuticaxCodigoResponse/Buscar_FFamaceuticaxCodigoResult").XML
        
            If Err.Number > 0 Then
'                MsgBox Err.Description 'Actualizado 07102014
                If Err.Number = 91 Then
                  MsgBox "No hay conexiòn con el SIS, Verificar su conexciòn, Se buscarà en la base local", vbInformation, "PROBLEMAS CON LA WEB"
                Else
                    MsgBox Err.Description
                End If
            Else
                res = XmlaArreglo(lcDatosWeb)
            End If
            If res(0) = "0" Then
                lbSalida = False
            Else
        
                EstableceConexion
                Set oA_presentaciones.Conexion = oConexion
                'oDoa_presentaciones.IdUsuarioAuditoria =
                oDoa_presentaciones.tpre_Abreviatura = res(0)
                oDoa_presentaciones.tpre_Descripcion = res(1)
                oDoa_presentaciones.tpre_IdEstado = res(5)
                oDoa_presentaciones.tpre_IdPresentacion = AutogeneraSAM
                oDoa_presentaciones.tpre_TopeHosp = res(4)
                oDoa_presentaciones.tpre_TopeMinimo = res(2)
                oDoa_presentaciones.tpre_TopeNoHosp = res(3)
        
                lbSalida = oA_presentaciones.Insertar(oDoa_presentaciones)
        
            End If
            
            Set oHttReq = Nothing
            Set parser = Nothing
            Set oConexion = Nothing
        End If
        ConsultarServicioFFamaceuticaxCodigo = lbSalida
End Function

'debb-21/10/2015
Function ConsultarServicioInsumosxCodigo(ms_Codigo As String, ByRef rsTmp As Recordset) As Boolean
    If lcBuscaParametro.SeleccionaFilaParametro(322) <> "S" Then
       Exit Function
    End If
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    Dim lpos As Integer
    Dim lcDatosWeb As String
    Dim res() As String
    Dim wlcad As String
    Dim indice As Integer
    Dim ano As Integer
    
    Dim oM_insumos As New m_insumos
    Dim oDom_insumos As New Dom_insumos
    Dim lbSalida As Boolean
    Dim wfechaBaja As String

 
    Set rsTmp = m_insumosSeleccionarPorId(ms_Codigo)
    If rsTmp.RecordCount > 0 Then
         lbSalida = True
    Else 'Modificado para busqueda web*************************************************
        Set parser = New DOMDocument
        parser.loadXML SIS_Buscar_InsumosxCodigo
        
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_InsumosxCodigo/Codigo").Text = ms_Codigo
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_InsumosxCodigo/Usuario").Text = lcBuscaParametro.SeleccionaFilaParametro(365) 'wUsuario
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_InsumosxCodigo/Clave").Text = lcBuscaParametro.SeleccionaFilaParametro(366)   'wClave
        
        Set oHttReq = New XMLHTTPRequest
        oHttReq.Open "POST", RutaWeb, False
        oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/Buscar_InsumosxCodigo "
        On Error Resume Next
        oHttReq.send parser.XML
        If Err.Number > 0 Then
             MsgBox Err.Description
             Set parser = Nothing
             Set oHttReq = Nothing
             Set oConexion = Nothing
             Exit Function
        End If
        Set parser = New DOMDocument
        parser.loadXML oHttReq.responseText
        lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_InsumosxCodigoResponse/Buscar_InsumosxCodigoResult").XML
        
        If Err.Number > 0 Then
            MsgBox Err.Description
              MsgBox "Puede ser problema de CABLE," & Chr(13) & "No hay servicio de INTERNET," & Chr(13) & _
                     "hay problemas en la página SIS," & Chr(13) & "Usuario o clave incorrecto", vbInformation, "PROBLEMAS CON LA WEB"
        Else
            res = XmlaArreglo(lcDatosWeb)
        End If
        If res(0) = 0 Then
            lbSalida = False
        Else
            EstableceConexion
            Set oM_insumos.Conexion = oConexion
            
            'odom_insumos.IdUsuarioAuditoria=
            oDom_insumos.ins_CodIns = res(0)
            'odom_insumos.ins_Concen=
            oDom_insumos.ins_Costo = res(3)
            'odom_insumos.ins_DocBaja =
            ano = CInt(Left(res(5), 4))
            If ano > 2079 Then ano = 2078
            wfechaBaja = "28/" & Right(res(5), 2) & "/" & Trim(Str(ano))
            oDom_insumos.ins_FecBaja = wfechaBaja
            oDom_insumos.ins_FormaFarmaceutica = res(2)
            oDom_insumos.ins_IdEstado = res(6)
            oDom_insumos.ins_Nombre = res(1)
            'odom_insumos.ins_Observacion =
            'odom_insumos.ins_Petitorio=
            'odom_insumos.ins_Presen =
            
            lbSalida = oM_insumos.Insertar(oDom_insumos)
            
        End If
        Set oConexion = Nothing
        Set oHttReq = Nothing
        Set parser = Nothing
    End If
    ConsultarServicioInsumosxCodigo = lbSalida
End Function

'debb-21/10/2015
Function ConsultarServicioMedicamentosxCodigo(ms_Codigo As String, ByRef rsTmp As Recordset) As Boolean
    If lcBuscaParametro.SeleccionaFilaParametro(322) <> "S" Then
       Exit Function
    End If
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    Dim lpos As Integer
    Dim lcDatosWeb As String
    Dim res() As String
    Dim wlcad As String
    Dim indice As Integer
    Dim ano As Integer
    
    Dim oM_medicamentos As New m_medicamentos
    Dim oDom_medicamentos As New Dom_medicamentos
    Dim lbSalida As Boolean
    Dim wfechaBaja As String
 
    Set rsTmp = m_medicamentosSeleccionarPorId(ms_Codigo)
    If rsTmp.RecordCount > 0 Then
         lbSalida = True
    Else 'Modificado para busqueda web*************************************************

        Set parser = New DOMDocument
        parser.loadXML SIS_Buscar_MedicamentosxCodigo
        
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_MedicamentosxCodigo/Codigo").Text = ms_Codigo
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_MedicamentosxCodigo/Usuario").Text = lcBuscaParametro.SeleccionaFilaParametro(365)  ' wUsuario
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_MedicamentosxCodigo/Clave").Text = lcBuscaParametro.SeleccionaFilaParametro(366)     'wClave
        
        Set oHttReq = New XMLHTTPRequest
        oHttReq.Open "POST", RutaWeb, False
        oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/Buscar_MedicamentosxCodigo"
        On Error Resume Next
        oHttReq.send parser.XML
        If Err.Number > 0 Then
             MsgBox Err.Description
             Set parser = Nothing
             Set oHttReq = Nothing
             Set oConexion = Nothing
             Exit Function
        End If
        
        Set parser = New DOMDocument
        parser.loadXML oHttReq.responseText
        lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_MedicamentosxCodigoResponse/Buscar_MedicamentosxCodigoResult").XML
        
        If Err.Number > 0 Then
            MsgBox Err.Description
              MsgBox "Puede ser problema de CABLE," & Chr(13) & "No hay servicio de INTERNET," & Chr(13) & _
                     "hay problemas en la página SIS," & Chr(13) & "usuario o clave incorrecto", vbInformation, "PROBLEMAS CON LA WEB"
        Else
            res = XmlaArreglo(lcDatosWeb)
        End If
        If res(0) = 0 Then
            lbSalida = False
        Else
            EstableceConexion
            Set oM_medicamentos.Conexion = oConexion
            'odom_medicamentos.IdUsuarioAuditoria=
            oDom_medicamentos.med_CodMed = res(0)
            'odom_medicamentos.med_Concen=
            oDom_medicamentos.med_Costo = res(4)
            ano = CInt(Left(res(7), 4))
            If ano > 2079 Then ano = 2078
            wfechaBaja = "28/" & Right(res(7), 2) & "/" & ano
            oDom_medicamentos.med_FecBaja = wfechaBaja
            'odom_medicamentos.med_FFDigemid=
            oDom_medicamentos.med_FormaFarmaceutica = res(3)
            oDom_medicamentos.med_IdEstado = res(9)
            oDom_medicamentos.med_Nombre = res(1)
            oDom_medicamentos.med_Petitorio = res(5)
            oDom_medicamentos.med_Petitorio2005 = res(5)
            oDom_medicamentos.med_Petitorio2010 = res(5)
            oDom_medicamentos.med_Presen = res(2)
            
            lbSalida = oM_medicamentos.Insertar(oDom_medicamentos)
        End If
        Set oConexion = Nothing
        Set oHttReq = Nothing
        Set parser = Nothing
    End If
    ConsultarServicioMedicamentosxCodigo = lbSalida
End Function

Function XmlaArreglo(wlcad As String) As String()
    Dim lpos As Integer
    Dim res() As String
    Dim indice As Integer
    
    indice = 0
    Do While Len(wlcad) > 0
        If Mid(wlcad, 1, 1) = "<" Then
                lpos = InStr(wlcad, ">")
                wlcad = Right(wlcad, Len(wlcad) - lpos)
        Else
            lpos = InStr(wlcad, "<")
            ReDim Preserve res(indice + 1)
            res(indice) = Left(wlcad, lpos - 1)
            indice = indice + 1
            wlcad = Right(wlcad, Len(wlcad) - lpos + 1)
        End If
    Loop
    If indice = 0 Then
        ReDim res(1)
        res(0) = 0
    End If
    XmlaArreglo = res
End Function

'debb-21/10/2015
Function ConsultarServicioPerSaludxNroDoc(ms_Codigo As String, ByRef rsTmp As Recordset) As Boolean
    If lcBuscaParametro.SeleccionaFilaParametro(322) <> "S" Then
       Exit Function
    End If
    
    Dim woResA_resatencion As New a_resatencion
    Dim woDoa_resAtencion As New Doa_resatencion
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    Dim lpos As Integer
    Dim lcDatosWeb As String
    Dim res() As String
    Dim wlcad As String
    Dim indice As Integer
    Dim lbSalida As Boolean
    
 
    Set rsTmp = a_resatencionSeleccionarPorId(ms_Codigo)
    If rsTmp.RecordCount > 0 Then
         lbSalida = True
    Else 'Modificado para busqueda web*************************************************

        Set parser = New DOMDocument
        parser.loadXML SIS_Buscar_PerSaludxNroDoc
        
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_PerSaludxNroDoc/Codigo").Text = ms_Codigo
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_PerSaludxNroDoc/Usuario").Text = lcBuscaParametro.SeleccionaFilaParametro(365)   'wUsuario
        parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_PerSaludxNroDoc/Clave").Text = lcBuscaParametro.SeleccionaFilaParametro(366)     'wClave
        
        Set oHttReq = New XMLHTTPRequest
        oHttReq.Open "POST", RutaWeb, False
        oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/Buscar_PerSaludxNroDoc"
        On Error Resume Next
        oHttReq.send parser.XML
        If Err.Number > 0 Then
             MsgBox Err.Description
             Set parser = Nothing
             Set oHttReq = Nothing
             Set oConexion = Nothing
             Exit Function
        End If
        Set parser = New DOMDocument
        parser.loadXML oHttReq.responseText
        lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/Buscar_PerSaludxNroDocResponse/Buscar_PerSaludxNroDocResult").XML
        
        If Err.Number > 0 Then
            MsgBox Err.Description
              MsgBox "Puede ser problema de CABLE," & Chr(13) & "No hay servicio de INTERNET," & Chr(13) & _
                     "hay problemas en la página SIS," & Chr(13) & "usuario o clave incorrecto", vbInformation, "PROBLEMAS CON LA WEB"
        Else
            res = XmlaArreglo(lcDatosWeb)
        End If
        If res(0) = 0 Then
            lbSalida = False
        Else
            EstableceConexion
            Set woResA_resatencion.Conexion = oConexion
            
            'woDoa_resAtencion.IdUsuarioAuditoria =
            woDoa_resAtencion.pers_ApeMaterno = res(3)
            woDoa_resAtencion.pers_ApePaterno = res(2)
            woDoa_resAtencion.pers_Colegiatura = res(5)
            'woDoa_resAtencion.pers_IdEspecialidad = 1
            woDoa_resAtencion.pers_IdEstado = res(6)
            woDoa_resAtencion.pers_IdResAtencion = res(0)
            woDoa_resAtencion.pers_IdTipoDocumento = res(1)
            woDoa_resAtencion.pers_IdTipoPersonalSalud = 1
            'woDoa_resAtencion.pers_NroEspecialidad = 0
            'woDoa_resAtencion.pers_OtrNombre = ""
            woDoa_resAtencion.pers_PriNombre = res(4)
    
            If woResA_resatencion.Insertar(woDoa_resAtencion) Then
                lbSalida = True
            Else
                MsgBox ms_MensajeError & Chr(13) & woResA_resatencion.MensajeError
                lbSalida = False
            End If
        End If
        Set oHttReq = Nothing
        Set parser = Nothing
        Set oConexion = Nothing
    End If
    ConsultarServicioPerSaludxNroDoc = lbSalida
End Function

Private Function AutogeneraSAM() As String
    Dim num As Integer
    Dim oRst As New Recordset
    Dim oCommand As New ADODB.Command
   ' EstableceConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "a_presentacionesSeleccionarMaximoCodigo"
        Set oRst = .Execute
'        Set oRst.ActiveConnection = Nothing
    End With

    num = CInt(oRst.Fields(0)) + 1
    AutogeneraSAM = Trim(Str(num))
End Function

Function a_presentacionPorAbreviatura(lcAbre As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "a_presentacionesSeleccionarPorAbreviatura"
        Set oParameter = .CreateParameter("@tpre_Abreviatura", adVarChar, adParamInput, 15, lcAbre): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set a_presentacionPorAbreviatura = oRecordset
   Set oCommand = Nothing
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
Function m_eessEleccionarPorCodigoRenaes(lcCodRenaes As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "m_eessSeleccionarPorcodigoRenaes"
        Set oParameter = .CreateParameter("@codigoRENAES", adVarChar, adParamInput, 50, lcCodRenaes): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set m_eessEleccionarPorCodigoRenaes = oRecordset
   Set oCommand = Nothing
   If oConexionSIS Is Nothing Then
      Set oConexion = Nothing
   End If
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

    '**** Programa: se arreglo la cantidad de caracteres  Right("00000" + lcCodigoMed, 5))
    '**** Programado por:Eder Yamill Palomino Espinoza
    '**** Fecha: 06102014
Function m_insumosSeleccionarPorId(lcCodigoMed As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "m_insumosSeleccionarPorId"
        Set oParameter = .CreateParameter("@ins_CodIns", adVarChar, adParamInput, 5, Right("00000" + lcCodigoMed, 5)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set m_insumosSeleccionarPorId = oRecordset
   Set oCommand = Nothing
   If oConexionSIS Is Nothing Then
       Set oConexion = Nothing
   End If
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function m_medicamentosSeleccionarPorId(lcCodigoMed As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "m_medicamentosSeleccionarPorId"
        Set oParameter = .CreateParameter("@med_CodMed", adVarChar, adParamInput, 13, lcCodigoMed): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set m_medicamentosSeleccionarPorId = oRecordset
   Set oCommand = Nothing
   If oConexionSIS Is Nothing Then
      Set oConexion = Nothing
   End If
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

'a_resatencionSeleccionarPorId

Function a_resatencionSeleccionarPorId(lcDniMedico As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "a_resatencionSeleccionarPorId"
        Set oParameter = .CreateParameter("@pers_IdResAtencion", adVarChar, adParamInput, 9, Left(lcDniMedico, 9)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@pers_IdTipoDocumento", adVarChar, adParamInput, 1, 1): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set a_resatencionSeleccionarPorId = oRecordset
   Set oCommand = Nothing
   If oConexionSIS Is Nothing Then
      Set oConexion = Nothing
   End If
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


Function m_eessSelecionarXcodigoRenaes(lcCodRenaes As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "m_eessSelecionarXcodigoRenaes"
        Set oParameter = .CreateParameter("@codigoRENAES", adVarChar, adParamInput, 50, lcCodRenaes): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set m_eessSelecionarXcodigoRenaes = oRecordset
   Set oCommand = Nothing
   If oConexionSIS Is Nothing Then
      Set oConexion = Nothing
   End If
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

'debb-20/10/2015
Function WebServiceSISBuscarAfiliado(lcDNI As String, lcDIsa As String, lcLote As String, lcNumero As String, _
                                             lcCorrelativo As String, lcTabla As String, _
                                             wxParametro323 As String) As Recordset
    On Error GoTo IR_ERROR
    Dim clnt
    Dim Datos() As String
    Dim Datos2() As String
    Dim oRsTmpAfiliadoSIS As New Recordset
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    Dim lpos As Integer
    Dim lcDatosWeb As String
    Dim wlcad As String
    Dim lcDniConsultado As String
    Dim lcTipoDocConsultado As String
    With oRsTmpAfiliadoSIS
          .Fields.Append "ApPaterno", adVarChar, 40, adFldIsNullable
          .Fields.Append "ApMaterno", adVarChar, 40, adFldIsNullable
          .Fields.Append "PNombre", adVarChar, 70, adFldIsNullable
          .Fields.Append "SNombre", adVarChar, 70, adFldIsNullable
          .Fields.Append "Fnacimiento", adDate, 8, adFldIsNullable
          .Fields.Append "cAfiliacion", adVarChar, 50, adFldIsNullable
          .Fields.Append "estadoSis", adVarChar, 1, adFldIsNullable
          .Fields.Append "fBajaOK", adDate, 8, adFldIsNullable
          .Fields.Append "DNI", adVarChar, 10, adFldIsNullable
          .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
          .Fields.Append "distritoDomicilio", adVarChar, 6, adFldIsNullable
          .Fields.Append "cDisa", adVarChar, 3, adFldIsNullable
          .Fields.Append "cFormato", adVarChar, 2, adFldIsNullable
          .Fields.Append "cNumero", adVarChar, 10, adFldIsNullable
          .Fields.Append "AfiliacionNroIntegrante", adVarChar, 2, adFldIsNullable
          .Fields.Append "codigo", adVarChar, 2, adFldIsNullable
          .Fields.Append "idSiaSis", adInteger
          .Fields.Append "MotivoBaja", adVarChar, 70, adFldIsNullable
          .Fields.Append "CodigoEstablAdscripcion", adVarChar, 10, adFldIsNullable
          .Fields.Append "AfiliacionFecha", adDate, 8, adFldIsNullable
          .Fields.Append "DocumentoTipo", adInteger                                         'debb-27/09/2019
          .LockType = adLockOptimistic
          .Open
    End With
    Set WebServiceSISBuscarAfiliado = oRsTmpAfiliadoSIS
    
    lcDniConsultado = "": lcTipoDocConsultado = ""
    Set parser = New DOMDocument
    If lcDNI <> "" Then
        parser.loadXML SIS_Buscar_AfiliadoxNroDoc
        parser.selectSingleNode("/soap:Envelope/soap:Header/CredencialWS/Usuario").Text = lcBuscaParametro.SeleccionaFilaParametro(363)  ' "sisws"
        parser.selectSingleNode("/soap:Envelope/soap:Header/CredencialWS/Clave").Text = lcBuscaParametro.SeleccionaFilaParametro(364)   '"g525nke54su"
        parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAseguradosDocIdent/IdEess").Text = lcBuscaParametro.SeleccionaFilaParametro(280)
        parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAseguradosDocIdent/TipoDoc").Text = "1"
        parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAseguradosDocIdent/Documento").Text = lcDNI
        Set oHttReq = New XMLHTTPRequest
        oHttReq.Open "POST", "http://app.sis.gob.pe/edi/sisws.asmx", False
        oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/BuscarAseguradosDocIdent"
        On Error Resume Next
        oHttReq.send parser.XML
        Set parser = New DOMDocument
        parser.loadXML oHttReq.responseText
        lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAseguradosDocIdentResponse/BuscarAseguradosDocIdentResult").XML
        If Err.Number > 0 Then
'            MsgBox "Puede ser problema de CABLE," & Chr(13) & "No hay servicio de INTERNET," & Chr(13) & _
'                   "hay problemas en la página SIS," & Chr(13) & "usuario o clave incorrecto", vbInformation, "PROBLEMAS CON LA WEB"
'             Set parser = Nothing
'             Set oHttReq = Nothing
'             Set oConexion = Nothing
'             Exit Function
            lcDIsa = ""
            lcLote = "2"
            lcNumero = lcDNI
            lcCorrelativo = ""
            lcTabla = ""
        Else
            Datos = XmlTabuladoraArreglo(lcDatosWeb)
            ObtenerTablaDisaLoteNumeroAfiliacion lcDIsa, lcLote, lcNumero, lcTabla, Datos(13), Datos(15)
            lcCorrelativo = ""
            lcTipoDocConsultado = Datos(0)
            If lcTipoDocConsultado = "1" Then
               lcDniConsultado = Datos(1)
            End If
        End If
    End If
    'Buscar por Disa, Lote, Numero
    parser.loadXML SIS_Buscar_AfiliadoxNroAfiliado
    parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAsegurados/Disa").Text = lcDIsa
    parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAsegurados/TipoFormato").Text = lcLote
    parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAsegurados/Contrato").Text = lcNumero
    parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAsegurados/Correlativo").Text = lcCorrelativo
    parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAsegurados/CodTabla").Text = lcTabla
    Set oHttReq = New XMLHTTPRequest
    oHttReq.Open "POST", "http://app.sis.gob.pe/edi/RecepcionTrama.asmx", False
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "SOAPAction", "http://www.sis.gob.pe/BuscarAsegurados"
    On Error Resume Next
    oHttReq.send parser.XML
    Set parser = New DOMDocument
    parser.loadXML oHttReq.responseText
    lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/BuscarAseguradosResponse/BuscarAseguradosResult").XML
    If Err.Number > 0 Then
        MsgBox "Puede ser problema de CABLE," & Chr(13) & "No hay servicio de INTERNET," & Chr(13) & _
               "hay problemas en la página SIS," & Chr(13) & "usuario o clave incorrecto", vbInformation, "PROBLEMAS CON LA WEB"
         Set parser = Nothing
         Set oHttReq = Nothing
         Set oConexion = Nothing
         Exit Function
    Else
        Datos2 = XmlTabuladoraArreglo(lcDatosWeb)
        If lcDatosWeb <> "" Then
            If Datos2(9) <> "" Then
                
                
                oRsTmpAfiliadoSIS.AddNew
                oRsTmpAfiliadoSIS.Fields!apPaterno = mo_Teclado.DevuelveTextoSINtildes(UCase(Datos2(9)))
                oRsTmpAfiliadoSIS.Fields!apMaterno = mo_Teclado.DevuelveTextoSINtildes(UCase(Datos2(10)))
                oRsTmpAfiliadoSIS.Fields!Pnombre = mo_Teclado.DevuelveTextoSINtildes(UCase(Datos2(11)))
                oRsTmpAfiliadoSIS.Fields!sNombre = mo_Teclado.DevuelveTextoSINtildes(mo_Teclado.CapitalizarNombres(Datos2(12)))
                oRsTmpAfiliadoSIS.Fields!Sexo = Datos2(13)
                If Datos2(14) <> "" Then
                   oRsTmpAfiliadoSIS.Fields!Fnacimiento = CDate(Datos2(14))
                End If
                oRsTmpAfiliadoSIS.Fields!DistritoDomicilio = Datos2(15)
                oRsTmpAfiliadoSIS.Fields!estadoSis = Datos2(16)
                If Datos2(17) <> "" Then
                    If IsDate(Datos2(17)) Then
                        oRsTmpAfiliadoSIS.Fields!fbajaOK = Format(Datos2(17), sighentidades.FormatoFechaCorta)
                    Else
                        oRsTmpAfiliadoSIS.Fields!fbajaOK = "" 'CDate(Datos(17))
                    End If
                End If
                If lcDNI <> "" Then
                    oRsTmpAfiliadoSIS.Fields!DNI = lcDNI
                Else
                    If Datos2(6) = "1" Then
                       oRsTmpAfiliadoSIS.Fields!DNI = Datos2(18)
                    End If
                End If
                oRsTmpAfiliadoSIS.Fields!cDisa = Datos2(2)
                oRsTmpAfiliadoSIS.Fields!cFormato = Datos2(3)
                oRsTmpAfiliadoSIS.Fields!cNumero = Datos2(4)
                oRsTmpAfiliadoSIS.Fields!idSiaSis = Val(Datos2(0))
                oRsTmpAfiliadoSIS.Fields!AfiliacionNroIntegrante = Datos2(5)
                oRsTmpAfiliadoSIS.Fields!Codigo = Datos2(1)
                oRsTmpAfiliadoSIS.Fields!CodigoEstablAdscripcion = Datos2(7)
                oRsTmpAfiliadoSIS.Fields!AfiliacionFecha = Datos2(8)
                oRsTmpAfiliadoSIS.Fields!cAfiliacion = Datos2(2) & " " & Datos2(3) & " " & Datos2(4)
                oRsTmpAfiliadoSIS.Update
            End If
        End If
    End If
    Set WebServiceSISBuscarAfiliado = oRsTmpAfiliadoSIS

IR_ERROR:
    If Err.Number <> 0 Then
        Resume Next
        'MsgBox Err.Description
    End If
    Exit Function
    Resume
End Function


Public Sub ObtenerTablaDisaLoteNumeroAfiliacion(ByRef lcDIsa As String, ByRef lcLote As String, ByRef lcNumero As String, _
                                                ByRef lcTabla As String, lcComponente As String, lcNumAfiliacion As String)
    Dim lnGuionDisa As Integer
    Dim lnGuionLote As Integer
    Dim orsTemp As New Recordset
                                           
    lnGuionDisa = InStr(lcNumAfiliacion, "-")
    lcDIsa = "200" 'Disa por defecto para consulta
    lcLote = Mid(lcNumAfiliacion, 1, lnGuionDisa - 1)
    lcNumero = Mid(lcNumAfiliacion, lnGuionDisa + 1, Len(lcNumAfiliacion) - lnGuionDisa)
    lnGuionLote = InStr(lcNumero, "-")
    If lnGuionLote <> -1 And lnGuionLote <> 0 Then
        lcDIsa = lcLote
        lcLote = Mid(lcNumero, 1, lcNumero - 1)
        lcNumero = Mid(lcNumero, lnGuionLote + 1, Len(lcNumero) - lnGuionLote)
    End If
    
    'Obtener Tabla
    Set orsTemp = ObtenerLoteTablaSiaSIS(lcComponente, lcLote)
    If orsTemp.RecordCount > 0 Then
         orsTemp.MoveFirst
         Do While Not orsTemp.EOF
             lcTabla = orsTemp.Fields!lot_idtablasiasis
             orsTemp.MoveNext
         Loop
    End If
    orsTemp.Close
    Set orsTemp = Nothing
End Sub

Function XmlTabuladoraArreglo(wlcad As String) As String()
    Dim lpos As Integer
    Dim res() As String
    If Len(wlcad) > 0 Then
        lpos = InStr(wlcad, ">")
        wlcad = Right(wlcad, Len(wlcad) - lpos)
        lpos = InStr(wlcad, "<")
        If lpos > 0 Then
            wlcad = Left(wlcad, lpos - 1)
        End If
        res = Split(wlcad, "|")
    End If
    XmlTabuladoraArreglo = res
End Function

Function ObtenerLoteTablaSiaSIS(lcIdComponente As String, lcLote As String) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ObtenerLoteTablaSiaSIS"
        Set oParameter = .CreateParameter("@IdComponente", adVarChar, adParamInput, 1, lcIdComponente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 2, lcLote): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set ObtenerLoteTablaSiaSIS = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set lcBuscaParametro = Nothing
End Function



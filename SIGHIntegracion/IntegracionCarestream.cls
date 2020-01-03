VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntegracionCarestream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: clase para integracion de pacientes con sistemas del proveedor Carestream
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_paciente As doPaciente
'====================================================================================
Dim mo_Medico As DOMedico
Dim mo_Empleado As DOEmpleado
Dim mo_Especialidades As Collection
'====================================================================================

Dim mo_DOPais As New DOPais
Dim mo_DoDistrito As New DODistrito
Dim mo_IntegrationUtil As New IntegrationUtil

Dim ml_IntegrationType As Integer
Const DIRECTORY_SEPARATOR = "/"
Const TAG_HEADER = "SENDEVENT"
Const FORMAT_FECHA_NACIMIENTO = "yyyymmdd"
Const FORMAT_FECHA_GENERAL = "yyyymmdd"
Const FORMAT_FECHA_UPDATED = "dd/mm/yyyy"
Const FORMAT_FECHA_HORA_GENERAL = "yyyymmddhhmmss"



Private Sub Class_Initialize()
    Call setIntegrationType
End Sub

Property Set Paciente(iValue As doPaciente)
   Set mo_paciente = iValue
End Property

Property Get Paciente() As doPaciente
   Paciente = mo_paciente
End Property

Property Set Distrito(iValue As DODistrito)
   Set mo_DoDistrito = iValue
End Property

Property Get Distrito() As DODistrito
   Set Distrito = mo_DoDistrito
End Property

Property Set Pais(iValue As DOPais)
   Set mo_DOPais = iValue
End Property

Property Get Pais() As DOPais
   Set Pais = mo_DOPais
End Property

'====================================================================================
Property Set Medico(iValue As DOMedico)
   Set mo_Medico = iValue
End Property

Property Get Medico() As DOMedico
   Set Medico = mo_Medico
End Property

Property Set Empleado(iValue As DOEmpleado)
   Set mo_Empleado = iValue
End Property

Property Get Empleado() As DOEmpleado
   Set Empleado = mo_Empleado
End Property

Property Set Especialidades(iValue As Collection)
   Set mo_Especialidades = iValue
End Property

Property Get Especialidades() As Collection
   Set Especialidades = mo_Especialidades
End Property
'====================================================================================


Private Function setIntegrationType() As Integer
    '1 Tipo de integracion usando xml
    '2 Tipo de Integracion
    ml_IntegrationType = 1
End Function

'Ruta donde se guardaran los archivos para paciente
Private Function getPathFileRepositorySourcePatient() As String
    Dim sFileName As String
    sFileName = App.Path & DIRECTORY_SEPARATOR & "ArchivosXml"
    getPathFileRepositorySourcePatient = sFileName
End Function

'Ruta donde se guardaran los archivos para medicos
Private Function getPathFileRepositorySourceDoctor() As String
    Dim sFileName As String
    sFileName = App.Path & DIRECTORY_SEPARATOR & "ArchivosXml"
    getPathFileRepositorySourceDoctor = sFileName
End Function

'devuelve el id de sexo de paciente para Carestream
Private Function getGenderForCarestream(lidGender As Long) As Integer
    Dim l_idGender As Long
    Select Case lidGender
        Case sghsexo.Femenino:
            l_idGender = enumGenderCarestream.Femenino
        Case sghsexo.Masculino
            l_idGender = enumGenderCarestream.Masculino
        Case Else
    
    End Select
    getGenderForCarestream = l_idGender
End Function

'devuelve el id de sexo de paciente para ser almacenado en galenhos
Private Function getGenderForGalenhos(lidGender As Long) As Integer
    Dim l_idGender As Long
    Select Case lidGender
        Case enumGenderCarestream.Femenino:
            l_idGender = sghsexo.Femenino
        Case enumGenderCarestream.Masculino
            l_idGender = sghsexo.Masculino
        Case Else
    
    End Select
    getGenderForGalenhos = l_idGender
End Function

'inicializa objetos en blanco
Private Function verifyDataObject()
    If mo_DOPais Is Nothing Then
        Set mo_DOPais = New DOPais
    End If
    If mo_DoDistrito Is Nothing Then
        Set mo_DoDistrito = New DODistrito
    End If
End Function

'==============================================================================
'MEDICOS
'==============================================================================

Public Function NuevoMedico() As DOMedico
    Select Case ml_IntegrationType
            
        Case Else:
            Call GenerateXMLNewDoctor
    End Select
End Function

Public Function EditarMedico() As DOMedico
    Select Case ml_IntegrationType
            
        Case Else:
            Call GenerateXMLNewDoctor
    End Select
End Function

Public Function UnificarMedico(oDoMedicoAntiguo As DOMedico) As DOMedico
    Select Case ml_IntegrationType
            
        Case Else:
            Call GenerateXMLUnificarDoctor(oDoMedicoAntiguo)
    End Select
End Function

'nombre el archivo xml para ser trasferido al sistema del proveedor
Private Function generatorFileNameDoctor(oDoMedico As DOMedico) As String
    generatorFileNameDoctor = "Doctor" & oDoMedico.Colegiatura & ".xml"
End Function

Private Function GenerateXMLNewDoctor() As DOMedico
    'documento xml
    Dim oDocXml As MSXML.DOMDocument
    'primer node del archivo
    Dim oElementHeader As MSXML.IXMLDOMElement
    
    Call verifyDataObject
    
    Dim oElementParent As MSXML.IXMLDOMElement

    Set oDocXml = New MSXML.DOMDocument
    'oDocXml.appendChild oDocXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

    Set oElementHeader = oDocXml.createElement(TAG_HEADER)
    
    Call generateXmlDoctor(oDocXml, oElementHeader)
    
    Call mo_IntegrationUtil.CreateElementXml("MESSAGE_TYPE", "MASTER_FILE_NOTIFICATION", oDocXml, oElementHeader)
    
    
    oDocXml.appendChild oElementHeader
    
    
    
    If mo_IntegrationUtil.ExistsPathFileRepositorySource(getPathFileRepositorySourceDoctor()) = True Then
        oDocXml.Save getPathFileRepositorySourceDoctor() & DIRECTORY_SEPARATOR & generatorFileNameDoctor(mo_Medico)
        MsgBox "Archivo generado en : " & getPathFileRepositorySourceDoctor() & DIRECTORY_SEPARATOR & generatorFileNameDoctor(mo_Medico)
    End If
    '// Destruir objetos
    Set oElementHeader = Nothing
    Set oDocXml = Nothing
    
'    Set GenerateXMLNewPatient = mo_paciente
End Function

'=============================================================================================================
'unificar para pacientes
'=============================================================================================================
'genera el xml para unir los datos de dos registros que pertenezca a un mismo paciente
Private Function GenerateXMLUnificarDoctor(oDoMedicoAntiguo As DOMedico) As DOMedico
    'documento xml
    Dim oDocXml As MSXML.DOMDocument
    'primer node del archivo
    Dim oElementHeader As MSXML.IXMLDOMElement
    
    Call verifyDataObject
    
    Dim oElementParent As MSXML.IXMLDOMElement

    Set oDocXml = New MSXML.DOMDocument
    'oDocXml.appendChild oDocXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

    Set oElementHeader = oDocXml.createElement(TAG_HEADER)
    
    Call generateXmlPatient(oDocXml, oElementHeader)
    
    'declaracion de nuevo nodo xml
    Dim oElementMerged As MSXML.IXMLDOMElement

    Set oElementMerged = mo_IntegrationUtil.CreateElementXml("MERGED", "", oDocXml, oElementHeader)
    
'    Call mo_IntegrationUtil.CreateElementXml("PATIENT_ID_MERGED", oDoMedicoAntiguo.NroDocumento, oDocXml, oElementMerged)
'    Call mo_IntegrationUtil.CreateElementXml("SSN_MERGED", oDoMedicoAntiguo.NroDocumento, oDocXml, oElementMerged)
'    Call mo_IntegrationUtil.CreateElementXml("ISSUER_MERGED", "CSH", oDocXml, oElementMerged)
    
    Call mo_IntegrationUtil.CreateElementXml("MESSAGE_TYPE", "MASTER_FILE_NOTIFICATION", oDocXml, oElementHeader)
    
    oDocXml.appendChild oElementHeader
    
    If mo_IntegrationUtil.ExistsPathFileRepositorySource(getPathFileRepositorySourceDoctor()) = True Then
        oDocXml.Save getPathFileRepositorySourceDoctor() & DIRECTORY_SEPARATOR & generatorFileNameDoctor(mo_Medico)
        MsgBox "Archivo generado en : " & getPathFileRepositorySourceDoctor() & DIRECTORY_SEPARATOR & generatorFileNameDoctor(mo_Medico)
    End If
    '// Destruir objetos
    Set oElementHeader = Nothing
    Set oDocXml = Nothing
End Function

'Genera xml para los datos del paciente
Private Function generateXmlDoctor(ByRef oDocXml As MSXML.DOMDocument, _
                oElementHeader As MSXML.IXMLDOMElement)

    Dim oElementDoctorHeader As MSXML.IXMLDOMElement
    Dim oElementDoctorData As MSXML.IXMLDOMElement
    Dim oElementPersonData As MSXML.IXMLDOMElement
    Dim oElementAddressData As MSXML.IXMLDOMElement
    Dim oElementOtherData As MSXML.IXMLDOMElement
    Dim oElementPVHeader As MSXML.IXMLDOMElement
    Dim oElementVisitData As MSXML.IXMLDOMElement
    
    Dim sEspecialidad As String
    
    sEspecialidad = ""
    
    Call verifyDataObject
  
    Set oElementDoctorHeader = mo_IntegrationUtil.CreateElementXml("MFN", "", oDocXml, oElementHeader)
    
    'Datos Paciente
    mo_IntegrationUtil.CreateElementXml "COD_MEDICO", mo_Medico.Colegiatura, oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "NOMBRE_MEDICO", mo_Empleado.ApellidoPaterno & " " & mo_Empleado.ApellidoMaterno _
                                            & "^" & mo_Empleado.Nombres, oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "ESPECIALIDAD", getEspecialidad(), oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "FECHA_NACIMIENTO", Format(mo_Empleado.FechaNacimiento, FORMAT_FECHA_GENERAL), oDocXml, oElementDoctorHeader
'    mo_IntegrationUtil.CreateElementXml "GENDER_KEY", getGenderForCarestream(mo_Empleado.idTipoSexo), oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "GENDER_KEY", "3", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "COD_MEDICO_ALTERNATIVO", "", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "STREET_ADDRESS", "SIN DIRECCION", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "CITY", "SIN COMUNA", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "COUNTRY", "PERU", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "PHONE_NUMBER", "", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "EMAIL", "", oDocXml, oElementDoctorHeader
    mo_IntegrationUtil.CreateElementXml "START_DATETIME", Format(Now, FORMAT_FECHA_HORA_GENERAL), oDocXml, oElementDoctorHeader
    
End Function

Private Function getEspecialidad() As String
    Dim sEspecilidad As String
    
    If Not (mo_Especialidades Is Nothing) Then
        If mo_Especialidades.Count > 0 Then
            Dim oDoEspecialidad As New DOMedicoEspecialidad
            Set oDoEspecialidad = mo_Especialidades(1)
            sEspecilidad = "MEDICINA GENERAL"
        End If
    End If
    
    getEspecialidad = sEspecilidad
End Function


'==============================================================================
'PACIENTE
'==============================================================================

Public Function NuevoPaciente() As doPaciente
    Select Case ml_IntegrationType
            
        Case Else:
            Call GenerateXMLNewPatient
    End Select
End Function

Public Function EditarPaciente() As doPaciente
    Select Case ml_IntegrationType
            
        Case Else:
            Call GenerateXMLUpdatePatient
    End Select
End Function

Public Function UnificarPaciente(oDoPacienteAntiguo As doPaciente) As doPaciente
    Select Case ml_IntegrationType
            
        Case Else:
            Call GenerateXMLUnificarPatient(oDoPacienteAntiguo)
    End Select
End Function

'nombre el archivo xml para ser trasferido al sistema del proveedor
Private Function generatorFileName(oDoPaciente As doPaciente) As String
    generatorFileName = "Patient" & oDoPaciente.NroHistoriaClinica & ".xml"
End Function

'Private Function getPathFileRepository() As String
'    getPathFileRepository = App.Path & "ArchivosParaCarestream"
'End Function

'=============================================================================================================
'xml para pacientes
'=============================================================================================================
'genera el xml para registar un nuevo paciente
Private Function GenerateXMLNewPatient() As doPaciente
    'documento xml
    Dim oDocXml As MSXML.DOMDocument
    'primer node del archivo
    Dim oElementHeader As MSXML.IXMLDOMElement
    
    Call verifyDataObject
    
    Dim oElementParent As MSXML.IXMLDOMElement

    Set oDocXml = New MSXML.DOMDocument
    'oDocXml.appendChild oDocXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

    Set oElementHeader = oDocXml.createElement(TAG_HEADER)
    
    Call generateXmlPatient(oDocXml, oElementHeader)

    
    Call mo_IntegrationUtil.CreateElementXml("MESSAGE_TYPE", "PATIENT_NEW", oDocXml, oElementHeader)
    Call mo_IntegrationUtil.CreateElementXml("MESSAGETYPEDETAILS", "ADMIN_VISIT_NOTIFICATION", oDocXml, oElementHeader)
    
    oDocXml.appendChild oElementHeader
    
    If mo_IntegrationUtil.ExistsPathFileRepositorySource(getPathFileRepositorySourcePatient()) = True Then
        oDocXml.Save getPathFileRepositorySourcePatient() & DIRECTORY_SEPARATOR & generatorFileName(mo_paciente)
        MsgBox "Archivo generado en : " & getPathFileRepositorySourcePatient() & DIRECTORY_SEPARATOR & generatorFileName(mo_paciente)
    End If
    '// Destruir objetos
    Set oElementHeader = Nothing
    Set oDocXml = Nothing
    
'    Set GenerateXMLNewPatient = mo_paciente
End Function


'genera el xml para registar un nuevo paciente
Private Function GenerateXMLUpdatePatient() As doPaciente
    'documento xml
    Dim oDocXml As MSXML.DOMDocument
    'primer node del archivo
    Dim oElementHeader As MSXML.IXMLDOMElement
    
    Call verifyDataObject
    
    Dim oElementParent As MSXML.IXMLDOMElement

    Set oDocXml = New MSXML.DOMDocument
    'oDocXml.appendChild oDocXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

    Set oElementHeader = oDocXml.createElement(TAG_HEADER)
    
    Call generateXmlPatient(oDocXml, oElementHeader)

    
    Call mo_IntegrationUtil.CreateElementXml("MESSAGE_TYPE", "PATIENT_UPDATED", oDocXml, oElementHeader)
    Call mo_IntegrationUtil.CreateElementXml("MESSAGETYPEDETAILS", "ADMIN_VISIT_NOTIFICATION", oDocXml, oElementHeader)
    
    oDocXml.appendChild oElementHeader
    
    If mo_IntegrationUtil.ExistsPathFileRepositorySource(getPathFileRepositorySourcePatient()) = True Then
        oDocXml.Save getPathFileRepositorySourcePatient() & DIRECTORY_SEPARATOR & generatorFileName(mo_paciente)
        MsgBox "Archivo generado en : " & getPathFileRepositorySourcePatient() & DIRECTORY_SEPARATOR & generatorFileName(mo_paciente)
    End If
    '// Destruir objetos
    Set oElementHeader = Nothing
    Set oDocXml = Nothing
    
'    Set GenerateXMLNewPatient = mo_paciente
End Function
'=============================================================================================================
'unificar para pacientes
'=============================================================================================================
'genera el xml para unir los datos de dos registros que pertenezca a un mismo paciente
Private Function GenerateXMLUnificarPatient(oDoPacienteAntiguo As doPaciente) As doPaciente
    'documento xml
    Dim oDocXml As MSXML.DOMDocument
    'primer node del archivo
    Dim oElementHeader As MSXML.IXMLDOMElement
    
    Call verifyDataObject
    
    Dim oElementParent As MSXML.IXMLDOMElement

    Set oDocXml = New MSXML.DOMDocument
    'oDocXml.appendChild oDocXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

    Set oElementHeader = oDocXml.createElement(TAG_HEADER)
    
    Call generateXmlPatient(oDocXml, oElementHeader)
    
    'declaracion de nuevo nodo xml
    Dim oElementMerged As MSXML.IXMLDOMElement

    Set oElementMerged = mo_IntegrationUtil.CreateElementXml("MERGED", "", oDocXml, oElementHeader)
    
    Call mo_IntegrationUtil.CreateElementXml("PATIENT_ID_MERGED", oDoPacienteAntiguo.NroDocumento, oDocXml, oElementMerged)
    Call mo_IntegrationUtil.CreateElementXml("SSN_MERGED", oDoPacienteAntiguo.NroDocumento, oDocXml, oElementMerged)
    Call mo_IntegrationUtil.CreateElementXml("ISSUER_MERGED", "CSH", oDocXml, oElementMerged)
    
    Call mo_IntegrationUtil.CreateElementXml("MESSAGE_TYPE", "PATIENT_MERGED", oDocXml, oElementHeader)
    Call mo_IntegrationUtil.CreateElementXml("MESSAGETYPEDETAILS", "ADT40", oDocXml, oElementHeader)
    
    oDocXml.appendChild oElementHeader
    
    If mo_IntegrationUtil.ExistsPathFileRepositorySource(getPathFileRepositorySourcePatient()) = True Then
        oDocXml.Save getPathFileRepositorySourcePatient() & DIRECTORY_SEPARATOR & generatorFileName(mo_paciente)
        MsgBox "Archivo generado en : " & getPathFileRepositorySourcePatient() & DIRECTORY_SEPARATOR & generatorFileName(mo_paciente)
    End If
    '// Destruir objetos
    Set oElementHeader = Nothing
    Set oDocXml = Nothing
End Function

'Genera xml para los datos del paciente
Private Function generateXmlPatient(ByRef oDocXml As MSXML.DOMDocument, _
                oElementHeader As MSXML.IXMLDOMElement)

    Dim oElementPatientHeader As MSXML.IXMLDOMElement
    Dim oElementPatientData As MSXML.IXMLDOMElement
    Dim oElementPersonData As MSXML.IXMLDOMElement
    Dim oElementAddressData As MSXML.IXMLDOMElement
    Dim oElementOtherData As MSXML.IXMLDOMElement
    Dim oElementPVHeader As MSXML.IXMLDOMElement
    Dim oElementVisitData As MSXML.IXMLDOMElement
    
    Call verifyDataObject
  
    Set oElementPatientHeader = mo_IntegrationUtil.CreateElementXml("PATIENT", "", oDocXml, oElementHeader)
    
    'Datos Paciente
    Set oElementPatientData = mo_IntegrationUtil.CreateElementXml("PATIENT", "", oDocXml, oElementPatientHeader)
    mo_IntegrationUtil.CreateElementXml "PATIENT_PERSON_KEY", mo_paciente.NroDocumento, oDocXml, oElementPatientData
    mo_IntegrationUtil.CreateElementXml "DEATH_INDICATOR", "N", oDocXml, oElementPatientData
    mo_IntegrationUtil.CreateElementXml "CAT_NAME", mo_paciente.PrimerNombre, oDocXml, oElementPatientData
    
    'Datos personales
    Set oElementPersonData = mo_IntegrationUtil.CreateElementXml("PERSON", "", oDocXml, oElementPatientHeader)
    mo_IntegrationUtil.CreateElementXml "PERSON_KEY", mo_paciente.NroDocumento, oDocXml, oElementPersonData
    mo_IntegrationUtil.CreateElementXml "LAST_NAME", mo_paciente.ApellidoPaterno & " " & mo_paciente.ApellidoMaterno, oDocXml, oElementPersonData
    mo_IntegrationUtil.CreateElementXml "FIRST_NAME", mo_paciente.PrimerNombre & " " & mo_paciente.SegundoNombre & " " & mo_paciente.TercerNombre, oDocXml, oElementPersonData
    mo_IntegrationUtil.CreateElementXml "BIRTH_DATE", Format(mo_paciente.FechaNacimiento, FORMAT_FECHA_NACIMIENTO), oDocXml, oElementPersonData

    mo_IntegrationUtil.CreateElementXml "MULTIINDEXREF", "", oDocXml, oElementPersonData
    mo_IntegrationUtil.CreateElementXml "GENDER_KEY", getGenderForCarestream(mo_paciente.idTipoSexo), oDocXml, oElementPersonData
    mo_IntegrationUtil.CreateElementXml "LAST_UPDATED", Format(Date, FORMAT_FECHA_GENERAL), oDocXml, oElementPersonData
    mo_IntegrationUtil.CreateElementXml "SSN", mo_paciente.NroDocumento, oDocXml, oElementPersonData
    
    'Datos personales
    Set oElementAddressData = mo_IntegrationUtil.CreateElementXml("ADDRESS", "", oDocXml, oElementPatientHeader)
    mo_IntegrationUtil.CreateElementXml "ADDRESS_KEY", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "PERSON_KEY", mo_paciente.NroDocumento, oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "ADDRESS_USAGE_KEY", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "STREET_ADDRESS", mo_paciente.DireccionDomicilio, oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "OTHER_DESIGNATION", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "CITY", mo_DoDistrito.Nombre, oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "COUNTRY", mo_DOPais.Nombre, oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "SEQUENCE_ID", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "PHONE_NUMBER", mo_paciente.Telefono, oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "PRIMARY", "Y", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "LAST_UPDATED", Format(Date, FORMAT_FECHA_GENERAL), oDocXml, oElementAddressData
    
    Set oElementOtherData = mo_IntegrationUtil.CreateElementXml("PATIENT_ID_LIST", "", oDocXml, oElementPatientHeader)
    mo_IntegrationUtil.CreateElementXml "PATIENT_ID_LIST_KEY", "", oDocXml, oElementOtherData
    mo_IntegrationUtil.CreateElementXml "PATIENT_PERSON_KEY", "", oDocXml, oElementOtherData
    mo_IntegrationUtil.CreateElementXml "PATIENT_ID", mo_paciente.NroDocumento, oDocXml, oElementOtherData
    mo_IntegrationUtil.CreateElementXml "PRIMARY", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "SEQUENCE_ID", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "ISSUER_OF_PID_KEY", "", oDocXml, oElementAddressData
    mo_IntegrationUtil.CreateElementXml "LAST_UPDATED", "", oDocXml, oElementOtherData
    
    Set oElementPVHeader = mo_IntegrationUtil.CreateElementXml("PV1", "", oDocXml, oElementHeader)
    Set oElementVisitData = mo_IntegrationUtil.CreateElementXml("VISIT", "", oDocXml, oElementPVHeader)
    
    mo_IntegrationUtil.CreateElementXml "VISIT_KEY", "", oDocXml, oElementVisitData
    mo_IntegrationUtil.CreateElementXml "PATIENT_PERSON_KEY", "", oDocXml, oElementVisitData
    mo_IntegrationUtil.CreateElementXml "PATIENT_CLASS_KEY", enumPatientClassKeyCarestream.Ambulatorio, oDocXml, oElementVisitData
    mo_IntegrationUtil.CreateElementXml "VISIT_NUMBER", "", oDocXml, oElementVisitData
    mo_IntegrationUtil.CreateElementXml "CREATE_BY_PERSON_KEY", "", oDocXml, oElementVisitData
    mo_IntegrationUtil.CreateElementXml "CREATE_ON_DATE", "", oDocXml, oElementVisitData
    mo_IntegrationUtil.CreateElementXml "IS_MASTER", "", oDocXml, oElementVisitData
End Function

'=============================================================================================================
'=============================================================================================================


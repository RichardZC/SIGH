VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de API
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_SET_VALUE = &H2
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_NOTIFY = &H10
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API
' Registry API prototypes
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Nombre de la PC
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' PAra deshabilitar el menú y otros
Private Const MF_BYPOSITION = &H400&
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Property Get HKEY_CLASSES_ROOT() As Long
    HKEY_CLASSES_ROOT = &H80000000
End Property
Property Get HKEY_CURRENT_USER() As Long
    HKEY_CURRENT_USER = &H80000001
End Property
Property Get HKEY_LOCAL_MACHINE() As Long
    HKEY_LOCAL_MACHINE = &H80000002
End Property
Property Get HKEY_USERS() As Long
    HKEY_USERS = &H80000003
End Property
Property Get HKEY_PERFORMANCE_DATA() As Long
    HKEY_PERFORMANCE_DATA = &H80000004
End Property

Public Sub EliminarClave(ByVal hKey As Long, ByVal strPath As String)
Dim lRegResult As Long

    lRegResult = RegDeleteKey(hKey, strPath)

End Sub
Public Sub GuardarClave(hKey As Long, strPath As String)
Dim keyhand&
Dim r As Integer

    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
    
End Sub

Public Sub CrearClave(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)
If lRegResult <> ERROR_SUCCESS Then
' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Sub
Public Function ObtenerCadena(hKey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim lValueType As Long

    lResult = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                ObtenerCadena = Left$(strBuf, intZeroPos - 1)
            Else
                ObtenerCadena = strBuf
            End If
        End If
    End If
    
End Function

Public Sub GuardarCadena(hKey As Long, strPath As String, strValue As String, strData As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
r = RegCloseKey(keyhand)
End Sub


Function ObtenerDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

    r = RegOpenKey(hKey, strPath, keyhand)
    lDataBufSize = 4
        
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            ObtenerDWord = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)

End Function

Function GuardarDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
Dim lResult As Long
Dim keyhand As Long
Dim r As Long
    
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)

End Function

Public Function EliminarValor(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
Dim r As Integer

    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)

End Function


Public Function RetornaNombrePC() As String
  Dim dwLen As Long
  Dim strString As String
  dwLen = MAX_COMPUTERNAME_LENGTH + 1
  strString = String(dwLen, "X")
  GetComputerName strString, dwLen
  strString = Left(strString, dwLen)
  RetornaNombrePC = strString
End Function

Public Function ArchivoExiste(lcPathYnombreArchivo As String) As Boolean
    ArchivoExiste = IIf(PathFileExists(lcPathYnombreArchivo) > 0, True, False)
End Function

Public Sub OcultaBotonXdelFormulario(lc_hwnd As Long)
    Dim Hwnd_Menu As Long
    ' Obtiene el Hwnd del menú para usar con el Api DeleteMenu
    Hwnd_Menu = GetSystemMenu(lc_hwnd, False)
    ' botón Cerrar
    Call DeleteMenu(Hwnd_Menu, 6, MF_BYPOSITION)
End Sub

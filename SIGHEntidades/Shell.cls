VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shell"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de tipo DOS
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1

Private Type STARTUPINFO
         cb As Long
         lpReserved As String
         lpDesktop As String
         lpTitle As String
         dwX As Long
         dwY As Long
         dwXSize As Long
         dwYSize As Long
         dwXCountChars As Long
         dwYCountChars As Long
         dwFillAttribute As Long
         dwFlags As Long
         wShowWindow As Integer
         cbReserved2 As Integer
         lpReserved2 As Long
         hStdInput As Long
         hStdOutput As Long
         hStdError As Long
      End Type

Private Type PROCESS_INFORMATION
        hProcess As Long
         hThread As Long
         dwProcessID As Long
         dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const ZSO_MINORCANCEL = &H2

Public Sub ejecutarComando(cmdline$)
Dim proc As PROCESS_INFORMATION
Dim ret As Long
         
         Dim start As STARTUPINFO
         ' Initialize the STARTUPINFO structure:
         start.cb = Len(start)
         ' Start the shelled application:
         ret = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
         ' Wait for the shelled application to finish:
         ret = WaitForSingleObject(proc.hProcess, INFINITE)
        DoEvents
         ret = CloseHandle(proc.hProcess)
         
End Sub

Public Sub CargarRutaWeb(lcRutaWeb As String, lnHwnd As Long)
Dim X
X = ShellExecute(lnHwnd, "Open", lcRutaWeb, &O0, &O0, SW_NORMAL)
End Sub
         
Function LogError(sLog As String)
    
    Dim I As Integer
    I = FreeFile
    Open "c:\galenhos.log" For Append As #I
    Print #I, sLog, Now
    Close #I

End Function



Sub DescomprimeArchivoZIP(lcClaveZip As String, lcNombreArchivoZip As String, lcRutaDondeSeDescomprime As String, lbMuestraMensajeDescomprimioOK As Boolean)
       Dim oUnzip As New duzactxctrl
       With oUnzip
            .ActionDZ = 0
            .BackgroundProcessFlag = True
            .ConvertLFtoCRLFFlag = False
            .DecryptCode = lcClaveZip
            .DecryptFlag = True
            .Destination = ""
            .DiagnosticFlag = False
            .ErrorCode = 0
            .Filespec = ""
            .FreshenFlag = False
            .MajorStatusFlag = False
            .MessageCallbackFlag = False
            .MinorStatusFlag = False
            .NoDirectoryItemsFlag = True
            .NoDirectoryNamesFlag = True
            '*Para que chanque los archivos si existe
            .OverwriteFlag = True
            .QuietFlag = True
            .MessageCallbackFlag = True
            .RecurseFlag = True
            .ReturnCount = 0
            .ReturnString = ""
            .TestFlag = False
            .UnZIPIndex = -1
            .UnZipSubOptions = 0
            .UpdateFlag = False
            .zi_attr = 0
            .zi_cMethod = 0
            .zi_cPathType = 0
            .zi_crc_32 = 0
            .zi_cSize = 0
            .zi_DateTime = ""
            .zi_FileName = ""
            .zi_index = 0
            .zi_oSize = 0
            .ZipFile = ""
            '* added for rev 3.00
            .RenameCallbackFlag = False
            .ExtProgTitle = ""
            .UnZIPStringSize = 0
            .UnZIPStringOffset = 0
            .MemoryBlockSize = 0
            .StartingOffset = 0
             
            .MajorStatusFlag = True
            .MinorStatusFlag = True
            .BackgroundProcessFlag = True
            '*----Nombre del paquete ZIP con toda la ruta completa
            .ZipFile = lcNombreArchivoZip
            .Filespec = "*.*"
            '*----Donde se va ha desempaquetar
            .Destination = lcRutaDondeSeDescomprime
            .ActionDZ = 8
       End With
       If oUnzip.ErrorCode = 0 Then
          If lbMuestraMensajeDescomprimioOK = True Then
              MsgBox "Descomprimió en forma Correcta" & Chr(13) & Chr(13) & "2)Renombre archivo PR????.dbf como PR.DBF (Ruta SIGH)" & Chr(13) & "3)Debe existir nuePrec.dbf (Ruta SIGH)" & Chr(13) & "4)Pulse clic en el botón 'Aceptar (F2)'", vbInformation, "ZIP"
          End If
       Else
          MsgBox "Error al descomprimir (verifique el archivo ZIP)", vbInformation, "Deszipear"
       End If
       Exit Sub
ErrActPRe:
    MsgBox Err.Description
End Sub

Sub ComprimeArchivoZip(lcExpTablaZIP As String, lcClaveZip As String, lcArchivosAexportar As String)
    Dim oZipear As New dzactxctrl
    
    With oZipear
         .AddCommentFlag = 0
         .AfterDateFlag = 0
         .BackgroundProcessFlag = True
         .Comment = ""
         .CompressionFactor = 5
         .ConvertLFtoCRLFFlag = 0
         .Date = ""
         .DeleteOriginalFlag = 0
         .DiagnosticFlag = 0
         .DontCompressTheseSuffixesFlag = 0
         .DosifyFlag = 0
         .EncryptCode = lcClaveZip
         .EncryptFlag = True
         .ErrorCode = 0
         .ExcludeFollowing = ""
         .ExcludeFollowingFlag = 0
         .ExtProgTitle = ""
         .FixFlag = False
         .FixHarderFlag = 0
         .GrowExistingFlag = 0
         .IncludeFollowing = ""
         .IncludeOnlyFollowingFlag = 0
         .IncludeVolumeFlag = 0
         .ItemList = lcArchivosAexportar
         .MajorStatusFlag = 1
         .MessageCallbackFlag = True
         .MinorStatusFlag = 1
         .MultiVolumeControl = 0
         .NoDirectoryEntriesFlag = False
         .NoDirectoryNamesFlag = True
         .OldAsLatestFlag = 0
         .PathForTempFlag = 0
         .QuietFlag = 0
         .RecurseFlag = 0
         .RenameCallbackFlag = 0
         .StoreSuffixes = ""
         .TempPath = ""
         .ZipFile = lcExpTablaZIP
         .ZIPString = ""
         .ZipSubOptions = 0
         .ZipSubOptions = .ZipSubOptions + ZSO_MINORCANCEL
         .ActionDZ = 4
         If .ErrorCode <> 0 Then
            MsgBox "Error al comprimir", vbInformation, "Zipear"
         End If
   End With

End Sub

'Sub CreandoRutaEnDisco(lcRutaTemporal As String)
'    Dim lcSql As String
'    On Error Resume Next
'    lcSql = "c:"
'    ejecutarComando lcSql
'    lcSql = "cd \"
'    ejecutarComando lcSql
'    lcSql = "md " & lcRutaTemporal
'    ejecutarComando lcSql
'    lcSql = "cd " & lcRutaTemporal
'    ejecutarComando lcSql
'    lcSql = "del *.*"
'
'End Sub

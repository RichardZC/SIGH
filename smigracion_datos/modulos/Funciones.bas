Attribute VB_Name = "Funciones"
'Declaraciones para el Numero de Serie del Disco
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
                                                                                                      lpVolumeSerialNumber As Long, LpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)


Sub PAbreBD()
Dim wxbd As String
       On Error GoTo ErrorPabreBD
     
       Select Case wxTipoBD
       Case "1"
            wxbd = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & wxRutaBD & "dbfs\ETTRESA.mdb;Mode=ReadWrite"
       End Select
       wxConexion.CommandTimeout = 150
       wxConexion.Open wxbd
       
       Exit Sub
ErrorPabreBD:
     MsgBox "No se pudo Abrir la BD, Error (" & Err.Number & ") " & Err.Description, vbCritical, "Error"
     End
End Sub




Public Sub PutImageInField(wwf As ADODB.Field, wwFile As String)
    Dim wwb() As Byte
    Dim wwff  As Long
    Dim wwn   As Long

    On Error GoTo ErrHandler
    wwff = FreeFile
    Open wwFile For Binary Access Read As wwff
    wwn = LOF(wwff)
    If wwn Then
       ReDim wwb(1 To wwn) As Byte
       Get wwff, , wwb()
    End If
    Close wwff
    wwf.Value = wwb()
    Exit Sub

ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Sub

Public Function GetImageFromField(wwf As ADODB.Field) As StdPicture

    Dim wwb()  As Byte
    Dim wwff   As Long
    Dim wwFile As String

    On Error GoTo ErrHandler
    Call GetRandomFileName(wwFile)
    wwff = FreeFile
    Open wwFile For Binary Access Write As wwff
    wwb() = wwf.Value
    Put wwff, , wwb()
    Close wwff
    Erase wwb
    Set GetImageFromField = LoadPicture(wwFile)
    Kill wwFile
    Exit Function

ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Function

'Guardar el contenido del Picture en el campo de la base
Public Sub GuardarBinary(wwcampoBinary As ADODB.Field, wwunPicture As PictureBox)
    
    Dim wwDataFile As Integer
    Dim wwChunk() As Byte
    Const wwconChunkSize As Integer = 16384

    Dim wwi As Integer
    Dim wwFragment As Integer, wwFl As Long, wwChunks As Integer
    'NOTA:
    ' El recordset debe estar preparado para Editar o Añadir
    'Guardar el contenido del picture en un fichero temporal
    SavePicture wwunPicture.Picture, "pictemp"
    'Leer el fichero y guardarlo en el campo
    wwDataFile = FreeFile
    Open "pictemp" For Binary Access Read As wwDataFile
    wwFl = LOF(wwDataFile) ' Longitud de los datos en el archivo
    If wwFl = 0 Then Close wwDataFile: Exit Sub
    wwChunks = wwFl \ wwconChunkSize
    wwFragment = wwFl Mod wwconChunkSize
    ReDim wwChunk(wwFragment)
    Get wwDataFile, , wwChunk()
    wwcampoBinary.AppendChunk wwChunk()
    ReDim Chunk(wwconChunkSize)
    For i = 1 To wwChunks
        Get wwDataFile, , wwChunk()
        wwcampoBinary.AppendChunk wwChunk()
    Next i
    Close wwDataFile
    'Ya no necesitamos el fichero, así que borrarlo
    On Local Error Resume Next
    If Len(Dir$("pictemp")) Then
        Kill "pictemp"
    End If
    Err = 0
End Sub

Private Sub GetRandomFileName(ByRef wwFile As String)
    Randomize Timer
    wwFile = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\") & Format(Rnd() * 1000000, "00000000") & ".tmp"
End Sub

Sub PAgregaTGENERAL()
      Dim wrs_uhc As New ADODB.Recordset
      wrs_uhc.Open "select * from Tgeneral", wxConexion, adOpenKeyset, adLockOptimistic
      If wrs_uhc.RecordCount = 0 Then
          frm_DatosG.Show 1
      End If
      wrs_uhc.Close
End Sub


'Function FrepararCompactarBD()
' Dim wcompactar As New DBEngine
' Dim wbd_backup As String
'    'bd
'    wbd_backup = App.Path & "\bk" & Trim(Str(Day(Date))) & Trim(Str(Month(Date))) & ".mdb"
'    wcompactar.RepairDatabase (wxRutaBD & "dbfs\shc.mdb")
'    If Dir(wbd_backup) <> "" Then
'       Kill wbd_backup
'    End If
'    wcompactar.CompactDatabase wxRutaBD & "dbfs\shc.mdb", wbd_backup
'    Kill wxRutaBD & "dbfs\shc.mdb"
'    FileCopy wbd_backup, wxRutaBD & "dbfs\shc.mdb"
'
'    'temporal
'    wbd_backup = App.Path & "\001.mdb"
'    wcompactar.RepairDatabase (App.Path & "\bdTmp\tmp.mdb")
'    If Dir(wbd_backup) <> "" Then
'       Kill wbd_backup
'    End If
'    wcompactar.CompactDatabase App.Path & "\bdTmp\tmp.mdb", wbd_backup
'    Kill App.Path & "\bdTmp\tmp.mdb"
'    FileCopy wbd_backup, App.Path & "\bdTmp\tmp.mdb"
'    Kill wbd_backup
'
'    'Vademecum
'    wbd_backup = App.Path & "\001.mdb"
'    wcompactar.RepairDatabase (App.Path & "\bdrefer\bdrefer.mdb")
'    If Dir(wbd_backup) <> "" Then
'       Kill wbd_backup
'    End If
'    wcompactar.CompactDatabase App.Path & "\bdrefer\bdrefer.mdb", wbd_backup
'    Kill App.Path & "\bdrefer\bdrefer.mdb"
'    FileCopy wbd_backup, App.Path & "\bdrefer\bdrefer.mdb"
'    Kill wbd_backup
'
'End Function
'

Sub terminar()
    wxConexion.Close
'    wxRefConexion.Close
'    wxConexionTmp.Close
    End
End Sub

'Calcula la Edad
Function FEdadActual(wFechaNac As String) As Integer
    If wFechaNac <> "__/__/____" Then
         wuno = 0
         If Month(Date) < Month(CDate(wFechaNac)) Then
            wuno = 1
         Else
            If Month(Date) = Month(CDate(wFechaNac)) Then
               If Not (Day(Date) >= Day(CDate(wFechaNac))) Then
                  wuno = 1
               End If
            End If
         End If
         FEdadActual = Year(Date) - Year(CDate(wFechaNac)) - wuno
     Else
         FEdadActual = 0
     End If
End Function



Function ConvNumLetra(wnumero As Double) As String
    Dim wuni(1 To 99) As String, wcen(1 To 9) As String
    
    wuni(1) = "UN"
    wuni(2) = "DOS"
    wuni(3) = "TRES"
    wuni(4) = "CUATRO"
    wuni(5) = "CINCO"
    wuni(6) = "SEIS"
    wuni(7) = "SIETE"
    wuni(8) = "OCHO"
    wuni(9) = "NUEVE"
    wuni(10) = "DIEZ"
    wuni(11) = "ONCE"
    wuni(12) = "DOCE"
    wuni(13) = "TRECE"
    wuni(14) = "CATORCE"
    wuni(15) = "QUINCE"
    wuni(16) = "DIECISEIS"
    wuni(17) = "DIECISIETE"
    wuni(18) = "DIECIOCHO"
    wuni(19) = "DIECINUEVE"
    wuni(20) = "VEINTE"
    wuni(21) = "VEINTIUN"
    wuni(22) = "VEINTIDOS"
    wuni(23) = "VEINTITRES"
    wuni(24) = "VEINTICUATRO"
    wuni(25) = "VEINTICINCO"
    wuni(26) = "VEINTISEIS"
    wuni(27) = "VEINTISIETE"
    wuni(28) = "VEINTIOCHO"
    wuni(29) = "VEINTINUEVE"
    wuni(30) = "TREINTA"
    wuni(31) = "TREINTIUN"
    wuni(32) = "TREINTIDOS"
    wuni(33) = "TREINTITRES"
    wuni(34) = "TREINTICUATRO"
    wuni(35) = "TREINTICINCO"
    wuni(36) = "TREINTISEIS"
    wuni(37) = "TREINTISIETE"
    wuni(38) = "TREINTIOCHO"
    wuni(39) = "TREINTINUEVE"
    wuni(40) = "CUARENTA"
    wuni(41) = "CUARENTIUN"
    wuni(42) = "CUARENTIDOS"
    wuni(43) = "CUARENTITRES"
    wuni(44) = "CUARENTICUATRO"
    wuni(45) = "CUARENTICINCO"
    wuni(46) = "CUARENTISEIS"
    wuni(47) = "CUARENTISIETE"
    wuni(48) = "CUARENTIOCHO"
    wuni(49) = "CUARENTINUEVE"
    wuni(50) = "CINCUENTA"
    wuni(51) = "CINCUENTIUN"
    wuni(52) = "CINCUENTIDOS"
    wuni(53) = "CINCUENTITRES"
    wuni(54) = "CINCUENTICUATRO"
    wuni(55) = "CINCUENTICINCO"
    wuni(56) = "CINCUENTISEIS"
    wuni(57) = "CINCUENTISIETE"
    wuni(58) = "CINCUENTIOCHO"
    wuni(59) = "CINCUENTINUEVE"
    wuni(60) = "SESENTA"
    wuni(61) = "SESENTIUN"
    wuni(62) = "SESENTIDOS"
    wuni(63) = "SESENTITRES"
    wuni(64) = "SESENTICUATRO"
    wuni(65) = "SESENTICINCO"
    wuni(66) = "SESENTISEIS"
    wuni(67) = "SESENTISIETE"
    wuni(68) = "SESENTIOCHO"
    wuni(69) = "SESENTINUEVE"
    wuni(70) = "SETENTA"
    wuni(71) = "SETENTIUN"
    wuni(72) = "SETENTIDOS"
    wuni(73) = "SETENTITRES"
    wuni(74) = "SETENTICUATRO"
    wuni(75) = "SETENTICINCO"
    wuni(76) = "SETENTISEIS"
    wuni(77) = "SETENTISIETE"
    wuni(78) = "SETENTIOCHO"
    wuni(79) = "SETENTINUEVE"
    wuni(80) = "OCHENTA"
    wuni(81) = "OCHENTIUN"
    wuni(82) = "OCHENTIDOS"
    wuni(83) = "OCHENTITRES"
    wuni(84) = "OCHENTICUATRO"
    wuni(85) = "OCHENTICINCO"
    wuni(86) = "OCHENTISEIS"
    wuni(87) = "OCHENTISIETE"
    wuni(88) = "OCHENTIOCHO"
    wuni(89) = "OCHENTINUEVE"
    wuni(90) = "NOVENTA"
    wuni(91) = "NOVENTIUN"
    wuni(92) = "NOVENTIDOS"
    wuni(93) = "NOVENTITRES"
    wuni(94) = "NOVENTICUATRO"
    wuni(95) = "NOVENTICINCO"
    wuni(96) = "NOVENTISEIS"
    wuni(97) = "NOVENTISIETE"
    wuni(98) = "NOVENTIOCHO"
    wuni(99) = "NOVENTINUEVE"
    
    wcen(1) = "CIENTO"
    wcen(2) = "DOSCIENTOS"
    wcen(3) = "TRESCIENTOS"
    wcen(4) = "CUATROCIENTOS"
    wcen(5) = "QUINIENTOS"
    wcen(6) = "SEISCIENTOS"
    wcen(7) = "SETECIENTOS"
    wcen(8) = "OCHOCIENTOS"
    wcen(9) = "NOVECIENTOS"

    'wente = Trim(Str(Int(wnumero), 20))
    wente = Trim(Str(Int(wnumero)))
    wdeci = Trim(Str(Round((wnumero - Int(wnumero)) * 100, 2)))
    wlen = Len(wente)
    wletras = ""
    wcon = 0
    ndig = wlen
    Do While ndig > 0
        wcon = wlen - ndig + 1
        dig = Val(Mid(wente, wcon, 1))
        dig1 = IIf(wcon < wlen, Val(Mid(wente, wcon + 1, 1)), 0)
        wcad = ""
        Select Case ndig
        Case 11, 5, 10, 4
            If dig > 0 Then
                If ndig = 11 Or ndig = 5 Then
                    wcad = " " & wuni(dig * 10 + dig1)
                    ndig = ndig - 1
                Else
                    wcad = " " & wuni(dig)
                End If
            End If
            wcad = wcad & " MIL"
        Case 8, 7
            If dig > 0 Then
                If ndig = 8 Then
                    wcad = " " & wuni(dig * 10 + dig1)
                    ndig = ndig - 1
                Else
                    wcad = " " & wuni(dig)
                End If
            End If
            wcad = wcad & " MILLONES"
        Case 1, 2
            If dig > 0 Then
                wcad = wuni(dig * 10 + dig1)
                If ndig = 2 Then
                    wcad = " " & wuni(dig * 10 + dig1)
                    ndig = ndig - 1
                Else
                    wcad = " " & wuni(dig)
                End If
            End If
            wcad = wcad
        Case 9, 6, 3
            If dig > 0 Then
                If dig = 1 And dig1 = 0 And Val(Mid(wente, wcon + 2, 1)) = 0 Then
                    wcad = "CIEN"
                Else
                    wcad = " " & wcen(dig)
                End If
            End If
        End Select
        wletras = wletras & wcad
        ndig = ndig - 1
    Loop
    wletras = wletras & " Y " & Right("00" + wdeci, 2) & "/100  NUEVOS SOLES"
    ConvNumLetra = wletras
End Function


Function BuscaSiAccesa(wprg As String) As Boolean
     Dim wrs_Prg As New ADODB.Recordset
     wok = False
     wprg = UCase(wprg)
     With wrs_Prg
         .Open "select * from Tprogramas where prgdebb='" & wprg & "'", wxConexion, adOpenKeyset, adLockOptimistic
         If .RecordCount > 0 Then
            wprog = .Fields("codpro").Value
            .Close
            .Open "select * from taccesos where codpro='" & wprog & "' and usuario='" & wxUsuaSist & "'", wxConexion, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                  wok = True
            End If
         End If
         .Close
     End With
     If wok Then
        BuscaSiAccesa = True
     Else
        MsgBox "No tiene Acceso", vbCritical, wxSistema
        BuscaSiAccesa = False
     End If
End Function



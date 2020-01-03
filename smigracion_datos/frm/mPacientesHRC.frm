VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mPacientesHRC 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   300
         TabIndex        =   9
         Top             =   7380
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CheckBox chkEliminaPacientes 
         Caption         =   "Elimina todos los Pacientes"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   6930
         Width           =   4695
      End
      Begin VB.TextBox txtFechaIni 
         Height          =   345
         Left            =   1860
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   6300
         Width           =   1185
      End
      Begin VB.TextBox txtFechaFin 
         Height          =   345
         Left            =   3540
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton cmdProcesa 
         Caption         =   "Proceso de Migración de Pacientes"
         Height          =   345
         Left            =   6180
         TabIndex        =   3
         Top             =   6330
         Width           =   3195
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consideraciones:"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   6075
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9585
         Begin VB.ListBox List1 
            Height          =   5130
            Left            =   150
            TabIndex        =   2
            Top             =   420
            Width           =   9315
         End
      End
      Begin VB.Label Label1 
         Caption         =   "F.Ingreso Hospital:"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   6330
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   3300
         TabIndex        =   6
         Top             =   6360
         Width           =   120
      End
   End
End
Attribute VB_Name = "mPacientesHRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Pacientes Lolcli
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcSql As String

Private Sub cmdProcesa_Click()
    On Error GoTo err_proceso
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexionExcel As New ADODB.Connection
       Dim oRsExcel1 As New ADODB.Recordset
       Dim oRsTmp1 As New Recordset
       Dim lnRegAct As Long, lntotReg As Long, lnTipoSexo As Long
       Dim lcFec_nac_pac As String, lcSegundoNombre As String
       Dim lnIdPaciente As Long, lcAutogenerado As String
       oConexionExcel.CommandTimeout = 150
       oConexionExcel.CursorLocation = adUseServer
       oConexionExcel.Open "dsn=Pacientes"
       lcSql = "SELECT dni, fec_nac_pac, sexo_pac, ap_pac, am_pac, nom_pac, " & _
               " hc_paciente , Segundo_nombre" & _
               " From DBA_paciente"
       oRsExcel1.Open lcSql, oConexionExcel, adOpenKeyset, adLockOptimistic
       lntotReg = oRsExcel1.RecordCount
       If Me.chkEliminaPacientes.Value = 1 Then
            lcSql = "delete * from HistoriasClinicas"
            oRsTmp1.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
            lcSql = "delete * from Pacientes"
            oRsTmp1.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       End If
       If lntotReg > 0 Then
          Me.MousePointer = 11
          ProgressBar1.Min = 0
          ProgressBar1.Max = lntotReg
          lnRegAct = 0
          oRsExcel1.MoveFirst
          Do While Not oRsExcel1.EOF
             lnRegAct = lnRegAct + 1: ProgressBar1.Value = lnRegAct
             If (Not IsNull(oRsExcel1.Fields!ap_pac)) And (Not IsNull(oRsExcel1.Fields!am_pac)) And (Not IsNull(oRsExcel1.Fields!nom_pac)) Then
                lcSql = "select * from Pacientes where NroHistoriaClinica=" & oRsExcel1.Fields!hc_paciente
                oRsTmp1.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                If IsNull(oRsExcel1.Fields!fec_nac_pac) Then
                   lcFec_nac_pac = "01/01/1990"
                Else
                   lcFec_nac_pac = oRsExcel1.Fields!fec_nac_pac
                End If
                lnTipoSexo = IIf(UCase(oRsExcel1.Fields!sexo_pac) = "F", 2, 1)
                lcSegundoNombre = IIf(IsNull(oRsExcel1.Fields!Segundo_nombre), "", Left(oRsExcel1.Fields!Segundo_nombre, 20))
                lcAutogenerado = PacienteCrearNroAutogenerado(lcFec_nac_pac, Left(oRsExcel1.Fields!ap_pac, 20), Left(oRsExcel1.Fields!am_pac, 20), Left(oRsExcel1.Fields!nom_pac, 20), lcSegundoNombre, lnTipoSexo)
                If oRsTmp1.RecordCount = 0 Then
                   oRsTmp1.AddNew
                   oRsTmp1.Fields!NroHistoriaClinica = Val(oRsExcel1.Fields!hc_paciente)
                   oRsTmp1.Fields!IdTipoNumeracion = 2
                   oRsTmp1.Fields!Autogenerado = lcAutogenerado
                   oRsTmp1.Fields!ApellidoPaterno = Left(oRsExcel1.Fields!ap_pac, 20)
                   oRsTmp1.Fields!ApellidoMaterno = Left(oRsExcel1.Fields!am_pac, 20)
                   oRsTmp1.Fields!PrimerNombre = Left(oRsExcel1.Fields!nom_pac, 20)
                   If Len(lcSegundoNombre) > 0 Then
                      oRsTmp1.Fields!SegundoNombre = lcSegundoNombre
                   End If
                   oRsTmp1.Fields!idTipoSexo = lnTipoSexo
                   oRsTmp1.Fields!FechaNacimiento = CDate(lcFec_nac_pac)
                   oRsTmp1.Update
                   lnIdPaciente = oRsTmp1.Fields!IdPaciente
                   oRsTmp1.Close
                   lcSql = "insert into HistoriasClinicas (" & _
                                   " IdTipoNumeracionAnterior,NroHistoriaClinicaAnterior," & _
                                   " IdTipoNumeracion,FechaCreacion,FechaPasoAPasivo," & _
                                   " IdTipoHistoria,IdEstadoHistoria,IdPaciente," & _
                                   " NroHistoriaClinica) values (" & _
                                   "0,0," & _
                                   "2,'01/01/2010',null," & _
                                   "1,1," & lnIdPaciente & "," & _
                                   oRsExcel1.Fields!hc_paciente & ")"
                   oRsTmp1.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                Else
                   oRsTmp1.Fields!ApellidoPaterno = Left(oRsExcel1.Fields!ap_pac, 20)
                   oRsTmp1.Fields!ApellidoMaterno = Left(oRsExcel1.Fields!am_pac, 20)
                   oRsTmp1.Fields!PrimerNombre = Left(oRsExcel1.Fields!nom_pac, 20)
                   If Len(lcSegundoNombre) > 0 Then
                      oRsTmp1.Fields!SegundoNombre = lcSegundoNombre
                   End If
                   oRsTmp1.Fields!idTipoSexo = lnTipoSexo
                   oRsTmp1.Fields!FechaNacimiento = CDate(lcFec_nac_pac)
                   If Len(oRsExcel1.Fields!DNI) > 0 Then
                      oRsTmp1.Fields!NroDocumento = ""
                      oRsTmp1.Fields!IdDocIdentidad = Null
                   Else
                      oRsTmp1.Fields!NroDocumento = oRsExcel1.Fields!DNI
                      oRsTmp1.Fields!IdDocIdentidad = 1
                   End If
                   oRsTmp1.Update
                   oRsTmp1.Close
                End If
             End If
             oRsExcel1.MoveNext
          Loop
       End If
       oRsExcel1.Close
    End If
    Me.MousePointer = 1
    Unload Me
    Exit Sub
err_proceso:
    MsgBox Err.Description
    Me.MousePointer = 1
    'Resume
End Sub

Private Sub Form_Load()
    List1.AddItem "1-Crear ODBC llamado 'Pacientes':"
    List1.AddItem "  * microsof excel Driver (*.xls)"
    List1.AddItem "  * el archivo que debe apuntar es c:\barrantes\pacientes.xls"
    List1.AddItem ""
    List1.AddItem "2-Solo se considera Pacientes con datos:"
    List1.AddItem "  apellido Paterno, apellido Materno, Primer Nombre, Sexo "
    List1.AddItem "  sino existen alguno de estos datos no se considera en la Migración"
    txtFechaIni.Text = Date
    txtFechaFin.Text = Date

End Sub


Function PacienteCrearNroAutogenerado(lcFechaNacimiento As String, lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String, lcSegundoNombre As String, lnIdTipoSexo As Long)
Dim P1 As String    'Primer digito del apellido paterno
Dim P4 As String    'Cuarto Digito del apellido paterno
Dim M1 As String    'Primer digito del apellido materno
Dim M4 As String    'Cuarto digito del apellido materno
Dim N11 As String   'Primer digito del primer nombre
Dim N41 As String   'Cuarto digito del primer materno
Dim N12 As String   'Primer digito del Ultimo materno
Dim N42 As String   'Cuarto digito del Ultimo materno
Dim D As String     'Digito de verificacion
Dim DD As String
Dim MM As String
Dim AAA As String
Dim sTemp  As String

        DD = Left(lcFechaNacimiento, 2)
        MM = Mid(lcFechaNacimiento, 4, 2)
        AAA = Mid(lcFechaNacimiento, 8, 3)
        DevuelvePrimeryCuartoCaracter lcApellidoPaterno, P1, P4
        DevuelvePrimeryCuartoCaracter lcApellidoMaterno, M1, M4
        DevuelvePrimeryCuartoCaracter lcPrimerNombre, N11, N41
        DevuelvePrimeryCuartoCaracter lcSegundoNombre, N12, N42
        sTemp = AAA + MM + DD & lnIdTipoSexo & P1 + P4 + M1 + M4 + N11 + N41 + N12 + N42
        PacienteCrearNroAutogenerado = sTemp & Modulo10(sTemp)
        
End Function


Sub DevuelvePrimeryCuartoCaracter(sPalabra As String, C1 As String, C2 As String)
Dim sTemp As String
        If sPalabra <> "" Then
            sTemp = ObtenerUltimaPalabra(EliminarConjunciones(sPalabra))
            C1 = Left(sTemp, 1)
            C2 = DevuelveCuartoCaracter(sTemp)
        Else
            C1 = "X"
            C2 = "X"
        End If
End Sub
Function DevuelveCuartoCaracter(sPalabra) As String
    If Len(sPalabra) <= 4 Then
        DevuelveCuartoCaracter = Right(sPalabra, 1)
    Else
        DevuelveCuartoCaracter = Mid(sPalabra, 4, 1)
    End If
End Function

Function ObtenerUltimaPalabra(sTexto As String) As String
Dim p As String
Dim iUltBlanco As Integer
Dim sTemp As String


    sTemp = Trim(sTexto)

    p = InStr(sTemp, " ")
    iUltBlanco = 0
    Do While p > 0
        iUltBlanco = p
        p = InStr(p + 1, sTemp, " ")
    Loop
    If iUltBlanco > 0 Then
        ObtenerUltimaPalabra = Mid(sTemp, iUltBlanco + 1)
    Else
        ObtenerUltimaPalabra = sTemp
    End If
End Function

Function EliminarConjunciones(sPalabra As String)
Dim sTemp As String

        sTemp = ReemplazarCadena(sPalabra, " DE ", " ")
        sTemp = ReemplazarCadena(sTemp, " DEL ", " ")
        sTemp = ReemplazarCadena(sTemp, " EL ", " ")
        sTemp = ReemplazarCadena(sTemp, " LA ", " ")
        sTemp = ReemplazarCadena(sTemp, " LOS ", " ")
        sTemp = ReemplazarCadena(sTemp, " LAS ", " ")

        EliminarConjunciones = sTemp

End Function
Function Modulo10(sValor As String) As Integer
Dim sTemp As String
Dim I As Integer
Dim k As Integer
Dim iTotal As Integer

    sTemp = ""
    
    For I = 1 To Len(sValor)
        If IsNumeric(Mid(sValor, I, 1)) Then
            sTemp = sTemp + Mid(sValor, I, 1)
        Else
            sTemp = sTemp + DevuelveValorEnNumeros(Mid(sValor, I, 1))
        End If
    Next I

    'Acumula total de digitos
    iTotal = 0
    For I = 1 To Len(sTemp)
        If I Mod 2 <> 0 Then
            k = CInt(Mid(sTemp, I, 1)) * 2
            iTotal = iTotal + (k - (k Mod 10)) / 10 + (k Mod 10)
        Else
            iTotal = iTotal + CInt(Mid(sTemp, I, 1))
        End If
    Next I

    If (iTotal Mod 10) = 0 Then
        Modulo10 = 0
    Else
        Modulo10 = 10 - (iTotal Mod 10)
    End If



End Function


Function DevuelveValorEnNumeros(sCaracter As String) As String

    Select Case sCaracter
    Case "A" To "N"
        DevuelveValorEnNumeros = Asc(sCaracter) - 55
    Case "Ñ"
        DevuelveValorEnNumeros = 24
    Case "O" To "Z"
        DevuelveValorEnNumeros = Asc(sCaracter) - 54
    End Select

End Function



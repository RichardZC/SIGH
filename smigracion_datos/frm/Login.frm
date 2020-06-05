VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "28092015u73"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2955
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones"
      Height          =   1575
      Left            =   30
      TabIndex        =   10
      Top             =   2730
      Width           =   2895
      Begin VB.TextBox Text1 
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Login.frx":000C
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.TextBox txtClave2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1350
      Width           =   1605
   End
   Begin VB.TextBox txtConexion 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1260
      TabIndex        =   0
      Text            =   $"Login.frx":00A3
      Top             =   150
      Width           =   1605
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   1470
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   525
      Left            =   30
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtClave 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   930
      Width           =   1605
   End
   Begin VB.TextBox txtUsuario 
      Height          =   345
      Left            =   1260
      TabIndex        =   1
      Top             =   540
      Width           =   1605
   End
   Begin VB.Label Label4 
      Caption         =   "Clave2"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Conexion"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   210
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Clave1"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   930
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   915
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Pide Usuario y Clave, además de CLAVE ESPECIAL
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Private Sub cmdAceptar_Click()
    Dim oRsTmp As New ADODB.Recordset
    Dim oCrypKey As New CrypKey.Util
    'PAbreBDhbtMono "dsn=GALENHOS"
    PAbreBDhbtMono SIGHEntidades.CadenaConexion
    oRsTmp.Open "select  * from Empleados where usuario='" & txtUsuario.Text & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       If UCase(txtClave.Text) = UCase(oCrypKey.DecryptString(oRsTmp.Fields!Clave)) Then
          'If Val(Left(Me.txtClave2.Text, 2)) = Month(Date) And Val(Right(Me.txtClave2.Text, 2)) = Day(Date) Then
          If SIGHEntidades.VerificaClaveMesDia(Me.txtClave2.Text) Then
             'Segunda Clave=Mes+dia....incluye ceros..ej: 09/06/2011...clave2=0609
             wxVersionBDactualizada = Me.Caption
             ActualizaCorrelativoAutomatico
             MDIfrmControl.Show
          Else
             MsgBox "Clave incorrecta", vbInformation, Me.Caption
             txtUsuario.Text = "": txtClave.Text = "": Me.txtClave2.Text = ""
          End If
       Else
          MsgBox "Clave incorrecta", vbInformation, Me.Caption
          txtUsuario.Text = "": txtClave.Text = "": Me.txtClave2.Text = ""
          'End
       End If
    Else
        MsgBox "Usuario ó Claves Incorrectas", vbInformation, Me.Caption
        txtUsuario.Text = "": txtClave.Text = "": Me.txtClave2.Text = ""
       'End
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    ActualizaCorrelativoAutomatico
End Sub

Sub ActualizaCorrelativoAutomatico()
    Dim wrs_GalenHos2 As New Recordset
    Dim lnNroHistoriaClinica As Long
    On Error Resume Next 'Actualizado 07102014
'    wrs_GalenHos2.Open "select top 1 nroHistoriaClinica from Pacientes where idTipoorder by nroHistoriaClinica desc", wxConexionRed, adOpenKeyset, adLockOptimistic
    wrs_GalenHos2.Open "select top 1 nroHistoriaClinica from Pacientes order by nroHistoriaClinica desc", wxConexionRed, adOpenKeyset, adLockOptimistic
    lnNroHistoriaClinica = wrs_GalenHos2.Fields!NroHistoriaClinica
    wrs_GalenHos2.Close
    wrs_GalenHos2.Open "update generadorNroHistoriaClinica set nroHistoriaClinica=" & lnNroHistoriaClinica & " where idNumerador=17", wxConexionRed, adOpenKeyset, adLockOptimistic
End Sub

Private Sub cmdCancelar_Click()
    End
End Sub

Private Sub Form_Load()
    If wxSistema = "CR" Then
        Dim oConfReg As New ConfigRegional
        oConfReg.FormatoFechaCorta = "dd/MM/yyyy"
        oConfReg.SeparadorDecimal = "."
        oConfReg.SeparadorDeMiles = ","
        oConfReg.SeparadorDecimalDeMonedas = "."
        oConfReg.SeparadorDeMilesDeMonedas = ","
        oConfReg.FormatoDeHoras = "hh:mm:ss tt"
        MsgBox "cambió la CONFIGURACION REGIONAL como estaba"
        End
    End If
    CargaVersionSQL
    If wxVersionSQL = sghVersionBD.sighSql2000 Then
       Frame1.Caption = "Consideraciones en BD SQL2000"
    Else
       Frame1.Caption = "Consideraciones en BD SQL2008"
    End If
    'Me.Caption = Year(Date)
End Sub

Sub CargaVersionSQL()
    On Error Resume Next
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    wxVersionSQL = sghVersionBD.sighSql2000
    If InStr(lcBuscaParametro.RetornaVersionServidorSQLserver, "SQL Server  2000") = 0 Then
       wxVersionSQL = sghVersionBD.sighSql2008
    End If
    Set lcBuscaParametro = Nothing
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdAceptar.SetFocus
    End If
End Sub









Private Sub txtClave2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtUsuario.SetFocus
    End If
End Sub

Private Sub txtConexion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtUsuario.SetFocus
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtClave.SetFocus
    End If
End Sub



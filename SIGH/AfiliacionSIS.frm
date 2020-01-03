VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AfiliacionSIS 
   Caption         =   "Filiación SIS en SisGalenPlus"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   6390
   Icon            =   "AfiliacionSIS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmOpciones 
      Height          =   1335
      Left            =   0
      TabIndex        =   22
      Top             =   3420
      Width           =   6330
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AfiliacionSIS.frx":000C
         DownPicture     =   "AfiliacionSIS.frx":046C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1553
         Picture         =   "AfiliacionSIS.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   1485
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AfiliacionSIS.frx":0D56
         DownPicture     =   "AfiliacionSIS.frx":121A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   3353
         Picture         =   "AfiliacionSIS.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   150
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del SIS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      TabIndex        =   16
      Top             =   2235
      Width           =   6330
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2685
         MaxLength       =   10
         TabIndex        =   2
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4455
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "99"
         Top             =   210
         Width           =   1785
      End
      Begin VB.TextBox txtFormato 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2085
         MaxLength       =   2
         TabIndex        =   1
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox txtIdSiaSis 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         TabIndex        =   17
         Text            =   "99999999"
         Top             =   230
         Width           =   1680
      End
      Begin VB.TextBox txtDisa 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         MaxLength       =   3
         TabIndex        =   0
         Top             =   630
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "IdSiaSis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   21
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "N° Afiliación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   20
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3855
         TabIndex        =   19
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.Frame fraDatosPaciente 
      Caption         =   "Datos del Paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6330
      Begin VB.TextBox txtNroDocumento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4455
         TabIndex        =   24
         Top             =   1425
         Width           =   1770
      End
      Begin VB.TextBox txtPrimerNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   8
         Top             =   618
         Width           =   1695
      End
      Begin VB.TextBox txtApellidoPaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   7
         Top             =   230
         Width           =   1680
      End
      Begin VB.TextBox txtSegundoNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4455
         MaxLength       =   40
         TabIndex        =   6
         Top             =   630
         Width           =   1785
      End
      Begin VB.TextBox txtApellidoMaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4455
         MaxLength       =   40
         TabIndex        =   5
         Top             =   210
         Width           =   1800
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   330
         Left            =   4470
         TabIndex        =   9
         Top             =   1020
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDocumentoTipo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1530
         TabIndex        =   28
         Top             =   1485
         Width           =   180
      End
      Begin VB.Label lblSexo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1530
         TabIndex        =   27
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Doc&umento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   26
         Top             =   1455
         Width           =   960
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Nº Dcto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3735
         TabIndex        =   25
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Nacimien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3570
         TabIndex        =   15
         Top             =   1035
         Width           =   870
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Segundo Nom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3300
         TabIndex        =   14
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   1065
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apell. &Materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3270
         TabIndex        =   12
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Primer Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   671
         Width           =   1215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido &Paterno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   285
         Width           =   1335
      End
   End
End
Attribute VB_Name = "AfiliacionSIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mo_Teclado As New sighEntidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lc_ApellidoPaterno As String
Dim lc_ApellidoMaterno As String
Dim lc_PrimerNombre As String
Dim lc_SegundoNombre As String
Dim ln_idTipoSexo As Long
Dim lc_FechaNacimiento As String
Dim lc_DocumentoTipo As String
Dim lc_DocumentoTipo1 As String
Dim lc_DocumentoNro As String
Property Let DocumentoNro(lValue As String)
      lc_DocumentoNro = lValue
      txtNroDocumento.Text = Left(lValue, 10)
End Property

Property Let DocumentoTipo(lValue As String)
      lc_DocumentoTipo = lValue
      Me.lblDocumentoTipo.Caption = lValue
End Property
Property Let DocumentoTipo1(lValue As String)
      lc_DocumentoTipo1 = lValue
End Property
Property Let FechaNacimiento(lValue As String)
      lc_FechaNacimiento = lValue
      Me.txtFechaNacimiento.Text = lValue
End Property

Property Let idTipoSexo(lValue As Long)
      ln_idTipoSexo = lValue
      If ln_idTipoSexo = 1 Then
        Me.lblSexo.Caption = "Masculino"
      ElseIf ln_idTipoSexo = 2 Then
         Me.lblSexo.Caption = "Femenino"
      Else
         Me.lblSexo.Caption = ""
      End If
End Property

Property Let SegundoNombre(lValue As String)
      lc_SegundoNombre = lValue
      Me.txtSegundoNombre.Text = lValue
End Property
Property Let PrimerNombre(lValue As String)
      lc_PrimerNombre = lValue
      Me.txtPrimerNombre.Text = lValue
End Property
Property Let ApellidoMaterno(lValue As String)
      lc_ApellidoMaterno = lValue
      Me.txtApellidoMaterno.Text = lValue
End Property
Property Let ApellidoPaterno(lValue As String)
      lc_ApellidoPaterno = lValue
      Me.txtApellidoPaterno = lValue
End Property

Private Sub btnAceptar_Click()
      If ValidaDatosObligatorios Then
         AgregaAfiliadosSIS lcBuscaParametro.SeleccionaFilaParametro(313), lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
      End If
End Sub

Private Sub btnCancelar_Click()
        Unload Me
End Sub

Function ValidaDatosObligatorios() As Boolean
      ValidaDatosObligatorios = False
      If txtDisa.Text = "" Or txtFormato.Text = "" Or txtNumero.Text = "" Then
         MsgBox "Falta datos del N° FILIACION", vbInformation, ""
         Exit Function
      End If
      If txtApellidoPaterno.Text = "" Then
         MsgBox "Falta datos de    APELLIDO PATERNO, hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If txtApellidoMaterno.Text = "" Then
         MsgBox "Falta datos de    APELLIDO MATERNO, hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If Me.txtPrimerNombre.Text = "" Then
         MsgBox "Falta datos de    PRIMER NOMBRE, hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If lblSexo.Caption = "" Then
         MsgBox "Falta datos de  SEXO  , hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If txtFechaNacimiento.Text = "__/__/____" Then
         MsgBox "Falta datos de    FECHA NACIMIENTO, hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If Me.lblDocumentoTipo.Caption = "" Then
         MsgBox "Falta datos de    TIPO DOCUMENTO, hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If Me.txtNroDocumento.Text = "" Then
         MsgBox "Falta datos de    NUMERO DE DOCUMENTO, hacerlo en ventana anterior", vbInformation, ""
         Exit Function
      End If
      If ChequeaSiExisteDocumentoYtipo = True Then
         Exit Function
      End If
      
      ValidaDatosObligatorios = True
End Function

Function ChequeaSiExisteDocumentoYtipo() As Boolean
     ChequeaSiExisteDocumentoYtipo = False
     Dim oRsAfiliadosSIS As New Recordset
     Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
     Dim oConexionExterna As New Connection
     oConexionExterna.CommandTimeout = 900
     oConexionExterna.CursorLocation = adUseClient
     oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisFiliacionesXdocumento(txtNroDocumento, lc_DocumentoTipo1, oConexionExterna)
     If oRsAfiliadosSIS.RecordCount > 0 Then
        MsgBox "Ese TIPO y NUMERO DE DOCUMENTO ya existen para: " & Chr(13) & _
               oRsAfiliadosSIS!paterno & " " & oRsAfiliadosSIS!materno & " " & oRsAfiliadosSIS!Pnombre, vbInformation, ""
        ChequeaSiExisteDocumentoYtipo = True
     End If
     oRsAfiliadosSIS.Close
     oConexionExterna.Close
     Set oRsAfiliadosSIS = Nothing
     Set mo_ReglasSISgalenhos = Nothing
     Set oConexionExterna = Nothing
End Function





Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
End Sub



Private Sub txtDisa_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtDisa
End Sub



Private Sub txtFormato_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtFormato
End Sub

Private Sub txtIdSiaSis_KeyDown(KeyCode As Integer, Shift As Integer)
          mo_Teclado.RealizarNavegacion KeyCode, txtIdSiaSis
End Sub



Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtNumero
End Sub
Sub AgregaAfiliadosSIS(lcParametro313 As String, lcConexionExterna As String)
        
        Me.MousePointer = 11
        Dim oConexionExterna As New Connection
        Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
        Dim oRsTmp As New Recordset
        Dim ms_MensajeError As String
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcConexionExterna
        
        ms_MensajeError = "select top 1 idSiaSis from sisFiliaciones where codigo='99' order by idSiaSis desc"
        oRsTmp.Open ms_MensajeError, oConexionExterna, adOpenKeyset, adLockOptimistic
        If oRsTmp.RecordCount > 0 Then
           txtIdSiaSis.Text = oRsTmp!idSiaSis + 1
           txtCodigo.Text = "99"
        End If
        oRsTmp.Close
        
        ms_MensajeError = ""
        Dim lcRango As String, lnFila As Long
        Dim lcIdSiaSis As Long, lcCodigo As String, lcCdisa As String, lcCformato As String, lcCnumero As String
        Dim lcAfiliacionNroIntegrante As String, lcTipoDocumento As String, lcCodigoEstablAdscripcion As String
        Dim lcAfiliacionFecha As String, lcApPaterno As String, lcApMaterno As String, lcPnombre As String
        Dim lcSnombre As String, lcSexo As String, lcFnacimiento As String, lcDistritoDomicilio As String
        Dim lcEstadoSis As String, lcFbajaok As String, lcDNI As String, lcMotivoBaja As String
        Dim ldAfiliacionFecha As Date, ldFNacimiento As Date, ldFbajaOk As Date
                    
                    lcIdSiaSis = txtIdSiaSis.Text
                    lcCodigo = txtCodigo.Text
                    lcCdisa = txtDisa.Text
                    lcCformato = Me.txtFormato.Text
                    lcCnumero = txtNumero.Text
                    lcAfiliacionNroIntegrante = "x"
                    lcTipoDocumento = lc_DocumentoTipo1
                    lcCodigoEstablAdscripcion = "x"
                    lcAfiliacionFecha = "2019-01-01"
                    ldAfiliacionFecha = 0
                    If IsDate(lcAfiliacionFecha) Then
                       ldAfiliacionFecha = CDate(lcAfiliacionFecha)
                    End If
                    lcApPaterno = Me.txtApellidoPaterno.Text
                    lcApMaterno = Me.txtApellidoMaterno.Text
                    lcPnombre = Me.txtPrimerNombre.Text
                    lcSnombre = Me.txtSegundoNombre.Text
                    lcSexo = IIf(ln_idTipoSexo = 1, "1", "0")
                    lcFnacimiento = Right(Me.txtFechaNacimiento.Text, 4) & "-" & Mid(Me.txtFechaNacimiento.Text, 4, 2) & "-" & Left(Me.txtFechaNacimiento.Text, 2)
                    ldFNacimiento = 0
                    If IsDate(lcFnacimiento) Then
                       ldFNacimiento = CDate(lcFnacimiento)
                    End If
                    lcDistritoDomicilio = "x"
                    lcEstadoSis = "0"
                    lcFbajaok = ""
                    ldFbajaOk = 0
                    If IsDate(lcFbajaok) Then
                       ldFbajaOk = CDate(lcFbajaok)
                    End If
                    lcDNI = Me.txtNroDocumento.Text
                    lcMotivoBaja = ""
                    ms_MensajeError = mo_ReglasSISgalenhos.SisFiliacionesBuscaYactualizaDatosXafiliado(oConexionExterna, _
                                                    Val(lcIdSiaSis), _
                                                    lcCodigo, _
                                                    lcCdisa, _
                                                    lcCformato, _
                                                    lcCnumero, _
                                                    lcAfiliacionNroIntegrante, _
                                                    lcTipoDocumento, _
                                                    lcCodigoEstablAdscripcion, _
                                                    ldAfiliacionFecha, _
                                                    lcApPaterno, _
                                                    lcApMaterno, _
                                                    lcPnombre, _
                                                    lcSnombre, _
                                                    lcSexo, _
                                                    ldFNacimiento, _
                                                    lcDistritoDomicilio, _
                                                    lcEstadoSis, _
                                                    ldFbajaOk, _
                                                    lcDNI, _
                                                    lcMotivoBaja)
         
      
      MsgBox "Generó en forma correcta", vbInformation, ""
      Unload Me
Error_AgregaAfiliadosSIS:
      If Err.Number <> 0 Then
         MsgBox Err.Description
      End If
      oConexionExterna.Close
      Set oConexionExterna = Nothing
      Set mo_ReglasSISgalenhos = Nothing
      Me.MousePointer = 1
      Exit Sub
      Resume
End Sub



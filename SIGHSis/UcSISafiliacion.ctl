VERSION 5.00
Begin VB.UserControl UcSISafiliacion 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   LockControls    =   -1  'True
   ScaleHeight     =   585
   ScaleWidth      =   2565
   Begin VB.TextBox txtNumero 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   2
      Top             =   220
      Width           =   1725
   End
   Begin VB.TextBox txtLote 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   220
      Width           =   375
   End
   Begin VB.TextBox txtDisa 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   220
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "------- N° de Afiliación (SIS)  -------"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2505
   End
End
Attribute VB_Name = "UcSISafiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para el Número de Afiliación del Paciente
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
Dim oRsAfiliadosSIS As New Recordset
Dim ldFechaActual As Date
Dim lcSql As String
Const lnHeigh As Integer = 555
Const lnWidth As Integer = 2505
'Public Event OnClick(oRecordSet As Recordset)
Public Event OnLostFocus(lcDisa As String, lcLote As String, lcNumero As String)
Dim ml_DNI As String
Dim ml_Apaterno As String
Dim ml_Amaterno As String
Dim ml_Pnombre As String
Dim ml_Onombre As String
Dim ml_DistritoDomicilio As Long
Dim ml_Sexo As Long
Dim ml_FNacimiento As Date
Dim ms_BusquedaDNI As String
Dim ms_BusquedaApaterno As String
Dim ms_BusquedaAmaterno As String
Dim ms_BusquedaPnombre As String
Dim ms_BusquedaSnombre As String
Dim wxParametroJAMO As String

Property Let BusquedaDNI(sValue As String)
   ms_BusquedaDNI = sValue
End Property
Property Let BusquedaApaterno(sValue As String)
   ms_BusquedaApaterno = sValue
End Property
Property Let BusquedaAmaterno(sValue As String)
   ms_BusquedaAmaterno = sValue
End Property
Property Let BusquedaPnombre(sValue As String)
   ms_BusquedaPnombre = sValue
End Property
Property Let BusquedaSnombre(sValue As String)
   ms_BusquedaSnombre = sValue
End Property

Property Get Fnacimiento() As Long
    Fnacimiento = ml_FNacimiento
End Property


Property Get Sexo() As Long
    Sexo = ml_Sexo
End Property

Property Get DistritoDomicilio() As Long
    DistritoDomicilio = ml_DistritoDomicilio
End Property

Property Get DNI() As String
    DNI = ml_DNI
End Property

Property Get Apaterno() As String
    Apaterno = ml_Apaterno
End Property

Property Get Amaterno() As String
    Amaterno = ml_Amaterno
End Property

Property Get Pnombre() As String
    Pnombre = ml_Pnombre
End Property

Property Get Onombre() As String
    Onombre = ml_Onombre
End Property






Function ValidaSiEsAfiliadoActualDelSIS() As Boolean
    ValidaSiEsAfiliadoActualDelSIS = False
    On Error GoTo ErrValAfil
    If (IsNull(oRsAfiliadosSIS.Fields!fbajaOK) Or (ldFechaActual <= oRsAfiliadosSIS.Fields!fbajaOK)) And Val(oRsAfiliadosSIS.Fields!estadoSis) = 0 Then
       ValidaSiEsAfiliadoActualDelSIS = True
    End If
    If ValidaSiEsAfiliadoActualDelSIS = False Then
        lcSql = "La afiliación de este paciente tiene problemas: " & Chr(13) & Chr(13) & _
                "Motivo de Baja: " & IIf(IsNull(oRsAfiliadosSIS.Fields!MotivoBaja), "", oRsAfiliadosSIS.Fields!MotivoBaja) & Chr(13) & _
                "Fecha Baja: " & oRsAfiliadosSIS.Fields!fbajaOK & Chr(13) & _
                "Estado: " & IIf(oRsAfiliadosSIS.Fields!estadoSis = 0, "Activo", "Inactivo")
        MsgBox lcSql, vbExclamation, "SIS"
    End If
ErrValAfil:
End Function





Private Sub txtDisa_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDisa
End Sub


Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtLote
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNumero
End Sub

Private Sub txtNumero_LostFocus()
    On Error GoTo ErrFocus
    UserControl.Height = lnHeigh
    UserControl.Width = lnWidth
    If (txtDisa.Text <> "" And txtLote.Text <> "" And txtNumero.Text <> "") Then
       RaiseEvent OnLostFocus(txtDisa.Text, txtLote.Text, txtNumero.Text)
    Else
       SendKeys "{tab}"
    End If
    Exit Sub
ErrFocus:
   If Err.Number = 3705 Then
      oRsAfiliadosSIS.Close
      Resume
   End If
End Sub




Sub FiltraPacientesSIS(lcWhereOrder As String)
       Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcWhereOrder, wxParametroJAMO)
End Sub

Public Sub Inicializar()
    ldFechaActual = CDate(Format(lcBuscaParametro.RetornaFechaHoraServidorSQL, SIGHEntidades.DevuelveFechaSoloFormato_DMY))
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
End Sub

Public Sub Limpiar()
   UserControl.Height = lnHeigh
   UserControl.Width = lnWidth
   txtDisa.Text = ""
   txtLote.Text = ""
   txtNumero.Text = ""
End Sub

Public Sub InabilitaControles(lbTrueFalse As Boolean)
    mo_Formulario.HabilitarDeshabilitar txtDisa, lbTrueFalse
    mo_Formulario.HabilitarDeshabilitar txtLote, lbTrueFalse
    mo_Formulario.HabilitarDeshabilitar txtNumero, lbTrueFalse
    Limpiar
End Sub

Public Function VerificaAcreditacionSIS(lcDNI As String) As Boolean
   FiltraPacientesSIS " where documentoNumero='" & lcDNI & "'"
   VerificaAcreditacionSIS = ValidaSiEsAfiliadoActualDelSIS
End Function

Public Sub DevuelveValoresDeFiliacion(ByRef lcDisa As String, ByRef lcLote As String, ByRef lcNumero As String)
    lcDisa = txtDisa.Text
    lcLote = txtLote.Text
    lcNumero = txtNumero.Text
End Sub

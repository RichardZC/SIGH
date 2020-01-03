VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Begin VB.UserControl UcSISafiliacion 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   LockControls    =   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   3915
   Begin VB.TextBox txtCorrelativo 
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
      Left            =   3420
      MaxLength       =   3
      TabIndex        =   4
      ToolTipText     =   "Correlativo"
      Top             =   220
      Width           =   435
   End
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
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "Número"
      Top             =   220
      Width           =   1155
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
      Left            =   1860
      TabIndex        =   2
      ToolTipText     =   "Lote"
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
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Disa"
      Top             =   220
      Width           =   435
   End
   Begin PVCOMBOLibCtl.PVComboBox cmbTipoFormatoSIS 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Tabla"
      Top             =   240
      Width           =   1455
      _Version        =   524288
      _cx             =   2566
      _cy             =   582
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Locked          =   0   'False
      Style           =   0
      Sorted          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowPictures    =   0   'False
      ColumnHeaders   =   -1  'True
      PrimaryColumn   =   1
      VisibleItems    =   10
      ColumnHeaderHeight=   20
      ListMember      =   ""
      ColumnHeaderForeColor=   0
      ColumnHeaderBackColor=   13160660
      SelectedForeColor=   16777215
      SelectedBackColor=   6956042
      AlternateBackColor=   16777215
      ItemLabelStyle  =   1
      ItemLabelType   =   0
      ItemLabelWidth  =   40
      ItemLabelForeColor=   0
      ItemLabelBackColor=   13160660
      ColumnHeaderStyle=   1
      VerticalGridLines=   -1  'True
      HorizontalGridLines=   -1  'True
      ColumnResize    =   0   'False
      ItemLabelResize =   0   'False
      AllowDBAutoConfig=   0   'False
      GridLineColor   =   13421772
      List            =   ""
      NullString      =   "[NULL]"
      DropShadow      =   -1  'True
      Text            =   ""
      SortOnColumnHeaderClick=   0   'False
      DropEffect      =   1
      ColumnCount     =   3
      Column0.Heading =   "Id"
      Column0.Width   =   20
      Column0.Alignment=   0
      Column0.Hidden  =   -1  'True
      Column0.Name    =   "lot_IdTablaSiasis"
      Column0.Format  =   ""
      Column0.Bound   =   -1  'True
      Column0.Locked  =   0   'False
      Column0.HeaderAlignment=   0
      Column1.Heading =   "Formato"
      Column1.Width   =   100
      Column1.Alignment=   0
      Column1.Hidden  =   0   'False
      Column1.Name    =   "tfrm_Descripcion"
      Column1.Format  =   ""
      Column1.Bound   =   -1  'True
      Column1.Locked  =   0   'False
      Column1.HeaderAlignment=   0
      Column2.Heading =   "Componente"
      Column2.Width   =   100
      Column2.Alignment=   0
      Column2.Hidden  =   0   'False
      Column2.Name    =   "com_Descripcion"
      Column2.Format  =   ""
      Column2.Bound   =   -1  'True
      Column2.Locked  =   0   'False
      Column2.HeaderAlignment=   0
      SortKey1.Column =   -1
      SortKey1.Ascending=   -1  'True
      SortKey1.CaseInsensitive=   -1  'True
      SortKey2.Column =   -1
      SortKey2.Ascending=   -1  'True
      SortKey2.CaseInsensitive=   -1  'True
      SortKey3.Column =   -1
      SortKey3.Ascending=   -1  'True
      SortKey3.CaseInsensitive=   -1  'True
      BoundColumn     =   ""
      Border          =   -1  'True
      VertAlign       =   1
      Format          =   ""
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "------------------- N° de Afiliación (SIS)  ---------------------"
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
      TabIndex        =   5
      Top             =   0
      Width           =   4065
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
'        Programa: Control para lista de Historia Clinica
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim oRsAfiliadosSIS As New Recordset
Dim ldFechaActual As Date
Dim lcSql As String
Const lnHeigh As Integer = 555
Const lnWidth As Integer = 2505
Const lnHeigh_Combo As Integer = 400
Const lcTipoFormatoDefault As String = "7"     'Afiliacion AUS
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

Property Get FNacimiento() As Long
    FNacimiento = ml_FNacimiento
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
    If (IsNull(oRsAfiliadosSIS.Fields!fBajaOK) Or (ldFechaActual <= oRsAfiliadosSIS.Fields!fBajaOK)) And Val(oRsAfiliadosSIS.Fields!estadoSis) = 0 Then
       ValidaSiEsAfiliadoActualDelSIS = True
    End If
    If ValidaSiEsAfiliadoActualDelSIS = False Then
        lcSql = "La afiliación de este paciente tiene problemas: " & Chr(13) & Chr(13) & _
                "Motivo de Baja: " & IIf(IsNull(oRsAfiliadosSIS.Fields!MotivoBaja), "", oRsAfiliadosSIS.Fields!MotivoBaja) & Chr(13) & _
                "Fecha Baja: " & oRsAfiliadosSIS.Fields!fBajaOK & Chr(13) & _
                "Estado: " & IIf(oRsAfiliadosSIS.Fields!estadoSis = 0, "Activo", "Inactivo")
        MsgBox lcSql, vbExclamation, "SIS"
    End If
ErrValAfil:
End Function

Public Sub TipoFormatoSISvisible(lbEsVisible As Boolean)
End Sub

Private Sub cmbTipoFormatoSIS_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoFormatoSIS
End Sub


Private Sub txtCorrelativo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCorrelativo
End Sub

Private Sub txtDisa_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDisa
End Sub

Private Sub txtDisa_LostFocus()
    If txtDisa.Text <> "" Then
       TipoFormatoSISvisible True
       On Error Resume Next
       txtLote.SetFocus
    End If
End Sub

Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtLote
End Sub

Private Sub txtLote_LostFocus()
    If txtLote.Text <> "" Then
       TipoFormatoSISvisible True
       On Error Resume Next
       txtNumero.SetFocus
    End If
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNumero
End Sub

Private Sub txtNumero_LostFocus()
    On Error GoTo ErrFocus
    If txtNumero.Text <> "" Then
        TipoFormatoSISvisible True
'        UserControl.Height = lnHeigh
'        UserControl.Width = lnWidth
        If (txtDisa.Text <> "" And txtLote.Text <> "" And txtNumero.Text <> "") Then
           RaiseEvent OnLostFocus(txtDisa.Text, txtLote.Text, txtNumero.Text)
        Else
           SendKeys "{tab}"
        End If
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
    ldFechaActual = CDate(Format(lcBuscaParametro.RetornaFechaHoraServidorSQL, sighentidades.DevuelveFechaSoloFormato_DMY))
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    Set cmbTipoFormatoSIS.ListSource = mo_ReglasSISgalenhos.TipoFormatoSIS()
'    cmbTipoFormatoSIS.Visible = False
    txtLote.Text = 2
    cmbTipoFormatoSIS_UbicaPosicion lcTipoFormatoDefault
End Sub
Public Sub Limpiar()
'   UserControl.Height = 1100 'lnHeigh
'   UserControl.Width = lnWidth
   txtDisa.Text = ""
   txtLote.Text = ""
   txtNumero.Text = ""
   'Actualizado yamill Palomino 20102014
   txtCorrelativo = ""
   TipoFormatoSISvisible False
   txtLote.Text = 2
   cmbTipoFormatoSIS_UbicaPosicion lcTipoFormatoDefault
End Sub


Sub cmbTipoFormatoSIS_UbicaPosicion(lcLot_idTablaSiaSis As String)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbTipoFormatoSIS.ListCount - 1)
        cmbTipoFormatoSIS.ListIndex = lnFor
        If cmbTipoFormatoSIS.SubItem(cmbTipoFormatoSIS.ListIndex, 0) = lcLot_idTablaSiaSis Then
           Exit For
        End If
    Next
End Sub

Public Sub InabilitaControles(lbTrueFalse As Boolean)
    If lbTrueFalse = False Then
        cmbTipoFormatoSIS.Enabled = False 'A. Yañez 13112014
        'mo_Formulario.HabilitarDeshabilitar cmbTipoFormatoSIS, True 'A. Yañez 13112014
    Else
        cmbTipoFormatoSIS.Enabled = True 'A. Yañez 13112014
    End If
    mo_Formulario.HabilitarDeshabilitar txtDisa, lbTrueFalse
    mo_Formulario.HabilitarDeshabilitar txtLote, lbTrueFalse
    mo_Formulario.HabilitarDeshabilitar txtNumero, lbTrueFalse
    mo_Formulario.HabilitarDeshabilitar txtCorrelativo, lbTrueFalse
    Limpiar
End Sub

Public Function VerificaAcreditacionSIS(lcDNI As String) As Boolean
   VerificaAcreditacionSIS = True
   If lcBuscaParametro.SeleccionaFilaParametro(326) = "S" Then
        FiltraPacientesSIS " where documentoNumero='" & lcDNI & "'"
        VerificaAcreditacionSIS = ValidaSiEsAfiliadoActualDelSIS
   End If
End Function

Public Sub DevuelveValoresDeFiliacion(ByRef lcDisa As String, ByRef lcLote As String, ByRef lcNumero As String, ByRef lcTipoFormato As String, ByRef lcNumCorrelativo As String)
    Dim oCampos() As String
    lcDisa = txtDisa.Text
    lcLote = txtLote.Text
    lcNumero = txtNumero.Text
    lcNumCorrelativo = txtCorrelativo.Text
    '
    If cmbTipoFormatoSIS.ListIndex < 0 Then
       lcTipoFormato = lcTipoFormatoDefault
    Else
       oCampos = Split(cmbTipoFormatoSIS.List(cmbTipoFormatoSIS.ListIndex), "|")
       lcTipoFormato = oCampos(0)
    End If
    '
End Sub

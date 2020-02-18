VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Begin VB.UserControl ucSISfuaCodPrestacion 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   315
   ScaleWidth      =   5385
   Begin PVCOMBOLibCtl.PVComboBox txtCodigoPrestacion 
      Height          =   330
      Left            =   1740
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _Version        =   524288
      _cx             =   1508
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
         Size            =   9
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
      ItemLabelWidth  =   20
      ItemLabelForeColor=   0
      ItemLabelBackColor=   13160660
      ColumnHeaderStyle=   0
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
      ColumnCount     =   2
      Column0.Heading =   "Descripción"
      Column0.Width   =   200
      Column0.Alignment=   0
      Column0.Hidden  =   0   'False
      Column0.Name    =   "ser_Descripcion"
      Column0.Format  =   ""
      Column0.Bound   =   -1  'True
      Column0.Locked  =   0   'False
      Column0.HeaderAlignment=   0
      Column1.Heading =   "Id"
      Column1.Width   =   40
      Column1.Alignment=   0
      Column1.Hidden  =   0   'False
      Column1.Name    =   "ser_IdServicio"
      Column1.Format  =   ""
      Column1.Bound   =   -1  'True
      Column1.Locked  =   0   'False
      Column1.HeaderAlignment=   0
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Cod.Prestación (SIS)"
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
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   1695
   End
   Begin VB.Label lblPrestacion 
      AutoSize        =   -1  'True
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2610
      TabIndex        =   1
      Top             =   30
      Width           =   135
   End
End
Attribute VB_Name = "ucSISfuaCodPrestacion"
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
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim oRsPrestaciones As New Recordset
Dim mo_CodigoPrestacion As String
Dim mo_Prestacion As String
Public Event SePresionoTeclaEspecial(KeyCode As Integer)

Property Let CodigoPrestacion(lValue As String)
    mo_CodigoPrestacion = lValue
    txtCodigoPrestacion.Text = lValue
    txtCodigoPrestacion_LostFocus
End Property
Property Get Prestacion() As String
    Prestacion = mo_Prestacion
End Property

Property Let Prestacion(lValue As String)
    mo_Prestacion = lValue
    lblPrestacion.Caption = lValue
End Property
Property Get CodigoPrestacion() As String
    CodigoPrestacion = mo_CodigoPrestacion
End Property


Public Sub Inicializar()
End Sub


Sub ReglasDeConsistenciasAntesDeCargarFormulario(ml_IdTipoServicio As sghTipoServicio, lcTipoSexo As String, ml_edad_En_YYYYMMDD As String)
        Dim lcFiltro As String
        lcFiltro = ""
        If ml_IdTipoServicio > 0 Then
            If ml_IdTipoServicio <> sghHospitalizacion Then
               lcFiltro = " ser_Hosp='N' "
            Else
               lcFiltro = " ser_Hosp='S' "
            End If
            
        End If
        If lcTipoSexo <> "" Then
            If lcTipoSexo = "M" Then
               lcFiltro = lcFiltro & " and (rc01_idSexo='2' or rc01_idSexo='1')"
            Else
               lcFiltro = lcFiltro & " and (rc01_idSexo='2' or rc01_idSexo='0')"
            End If
        End If
        If ml_edad_En_YYYYMMDD <> "" Then
           lcFiltro = lcFiltro & " and ('" & ml_edad_En_YYYYMMDD & "'>=rc01_edadMin  and '" & ml_edad_En_YYYYMMDD & "'<= rc01_edadMax)"
        End If
        If Left(lcFiltro, 4) = " and" Then
           lcFiltro = Mid(lcFiltro, 5, 200)
        End If
        Set oRsPrestaciones = mo_ReglasSISgalenhos.SisServiciosSeleccionarPorFiltro(lcFiltro)
        Set txtCodigoPrestacion.ListSource = oRsPrestaciones
End Sub

Private Sub txtCodigoPrestacion_Click()
    
    BuscaDescripcionPrestacion
End Sub

Private Sub txtCodigoPrestacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoPrestacion
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtCodigoPrestacion_LostFocus()
    BuscaDescripcionPrestacion
End Sub

Sub BuscaDescripcionPrestacion()
    lblPrestacion.Caption = ""
    mo_CodigoPrestacion = ""
    If oRsPrestaciones.State <> 1 Then
       Set oRsPrestaciones = mo_ReglasSISgalenhos.SisServiciosSeleccionarPorFiltro("")
    End If
    If txtCodigoPrestacion.Text <> "" And oRsPrestaciones.RecordCount > 0 Then
       oRsPrestaciones.MoveFirst
       oRsPrestaciones.Find "ser_IdServicio='" & txtCodigoPrestacion.Text & "'"
       If Not oRsPrestaciones.EOF Then
            lblPrestacion.Caption = oRsPrestaciones.Fields!ser_descripcion
            mo_CodigoPrestacion = txtCodigoPrestacion.Text
            mo_Prestacion = lblPrestacion.Caption
       End If
    End If
End Sub

Public Sub AsignaDescripcionSegunCodigoPrestacion()
    lblPrestacion.Caption = mo_ReglasSISgalenhos.SisServiciosDevuelveDescripcion(txtCodigoPrestacion.Text)
End Sub

Public Sub HabilitaCodigoPrestacion(lbHabilitar As Boolean)
       If lbHabilitar = False Then
            txtCodigoPrestacion.Locked = True
            txtCodigoPrestacion.BackColor = &HF9EADF
            txtCodigoPrestacion.ForeColor = &H808080
       Else
            txtCodigoPrestacion.Locked = False
            txtCodigoPrestacion.BackColor = &HFFFFFF
            txtCodigoPrestacion.ForeColor = &H0&
       End If
End Sub

Public Sub FocusEnCodigoPrestacion()
    On Error Resume Next
    txtCodigoPrestacion.SetFocus
End Sub

VERSION 5.00
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "OWC11.DLL"
Begin VB.Form VisitEnferGraficos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grafico"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "VisitEnferGraficos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   10440
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "VisitEnferGraficos.frx":0CCA
         DownPicture     =   "VisitEnferGraficos.frx":118E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4680
         Picture         =   "VisitEnferGraficos.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
   End
   Begin OWC11.ChartSpace CSGrafico 
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "VisitEnferGraficos.frx":1B66
      TabIndex        =   2
      Top             =   0
      Width           =   10530
   End
End
Attribute VB_Name = "VisitEnferGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Gráficos para Enfermería
'        Programado por: Franklin C
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
'------------------------------------------------------------------------------------
''       Autor: Franklin Cachay Velasquez
'        Fecha: 04/04/2014 12:26:57 a.m.
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_idUsuario As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdVisita As Integer
Dim ml_IdVariable As Integer
Dim mc_TituloGrafico As String
Dim mi_Opcion As sghOpciones
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_lcNombrePc As String
Dim mb_EsNuevaVisita As Boolean
Dim mc_TextoVariable As String
'------------------------------------------------------------------------------------
'                               GRAFICOS
'------------------------------------------------------------------------------------
Dim xValues As Variant, yValues As Variant
Dim owcChart As OWC11.ChChart
Dim owcSeries As OWC11.ChSeries
Dim lnNroPuntosGraficos As Integer
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdVisita(lValue As Integer)
   ml_IdVisita = lValue
End Property
Property Get IdVisita() As Integer
   IdVisita = ml_IdVisita
End Property
Property Let IdVariable(lValue As Integer)
   ml_IdVariable = lValue
End Property
Property Get IdVariable() As Integer
   IdVariable = ml_IdVariable
End Property
Property Let TituloGrafico(lValue As String)
   mc_TituloGrafico = lValue
End Property
Property Get TituloGrafico() As String
   TituloGrafico = mc_TituloGrafico
End Property
Property Let EsNuevaVisita(lValue As Boolean)
   mb_EsNuevaVisita = lValue
End Property
Property Get EsNuevaVisita() As Boolean
   EsNuevaVisita = mb_EsNuevaVisita
End Property
Property Let TextoVariable(lValue As String)
   mc_TextoVariable = lValue
End Property
Property Get TextoVariable() As String
   TextoVariable = mc_TextoVariable
End Property

Sub CargarDatosAlFormulario()
    Me.Caption = "GRAFICO: " & mc_TituloGrafico & " por visita"
    CargaGraficoChartSpace
End Sub

Sub CargaGraficoChartSpace()
    Dim lnFor As Integer
    Dim lnUltimoPunto As Integer
    Dim oConexion As New Connection
    Dim orsTemp As New ADODB.Recordset
    Dim lnRangoMaximo As Long
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion

    xValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
    yValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)

    If ml_IdVisita = 0 Then Exit Sub
    Set orsTemp = mo_AdminAdmision.enfermeria_ConsultaValoresVariableGrafico(oConexion, ml_idCuentaAtencion, ml_IdVariable, ml_IdVisita)
    
    If mb_EsNuevaVisita = True Then
        lnUltimoPunto = orsTemp.RecordCount
    Else
        lnUltimoPunto = orsTemp.RecordCount - 1
    End If
    
    ReDim xValues(lnUltimoPunto)
    lnRangoMaximo = 110
    If orsTemp.RecordCount > 0 Then
        orsTemp.MoveLast
        For lnFor = (orsTemp.RecordCount - 1) To 0 Step -1
           xValues(lnFor) = orsTemp.Fields!IdVisita
           yValues(lnFor) = orsTemp.Fields!VariableDato
           If Val(orsTemp.Fields!VariableDato) > lnRangoMaximo Then
                lnRangoMaximo = Val(orsTemp.Fields!VariableDato) + 50
           End If
           orsTemp.MovePrevious
        Next
    End If
'    If mb_EsNuevaVisita = True Then
        xValues(lnUltimoPunto) = ml_IdVisita
        yValues(lnUltimoPunto) = IIf(mc_TextoVariable = "", 0, Val(mc_TextoVariable))
'    End If
    '
    CSGrafico.Clear
    CSGrafico.DisplayToolbar = False
    Set owcChart = CSGrafico.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = mc_TituloGrafico '+ " vs Visita"
    owcChart.Title.Font.Name = "Tahoma"
    owcChart.Title.Font.Size = 13
    owcChart.Title.Font.Color = vbBlue
    owcChart.Axes(chAxisPositionBottom).Font.Name = "Tahoma"
    owcChart.Axes(chAxisPositionBottom).Font.Size = 9
    owcChart.Axes(chAxisPositionBottom).Font.Color = vbBlue
    owcChart.Axes(chAxisPositionBottom).Scaling.Minimum = 0
    owcChart.Axes(chAxisPositionLeft).Font.Name = "Tahoma"
    owcChart.Axes(chAxisPositionLeft).Font.Size = "9"
    owcChart.Axes(chAxisPositionLeft).Font.Color = vbBlue
    owcChart.Axes(chAxisPositionLeft).Scaling.Minimum = 0
    owcChart.Axes(chAxisPositionLeft).Scaling.Maximum = lnRangoMaximo '110
    owcChart.Axes(1).Font.Size = 9
'    owcChart.Axes(1).HasTitle = 1
'    owcChart.Axes(1).Font.Name = "Arial Narrow"
'    owcChart.Axes(1).Font.Size = 8
'    owcChart.Axes(1).Font.Color = vbBlue
'    owcChart.Axes(1).Title.Caption = ""
'    owcChart.Axes(1).Title.Font.Name = "Arial Narrow"
'    owcChart.Axes(1).Title.Font.Size = 8
'    owcChart.Axes(1).Title.Font.Color = vbBlue
    '
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = ""
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValues
        .Type = chChartTypeLineMarkers
        .Line.Color = vbRed
        .Line.Weight = 4
        .Marker.Style = chMarkerStyleCircle
        .Line.DashStyle = chLineSolid
        .DataLabelsCollection.Add
    End With
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub Form_Load()
       CargarDatosAlFormulario
End Sub

Private Sub btnCancelar_Click()
'   Unload Me
    Me.Visible = False
End Sub

Public Sub MostrarFormulario()
    Me.Show 1
End Sub

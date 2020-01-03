VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIfrmControl 
   BackColor       =   &H8000000F&
   Caption         =   "Utilidades externas para SIS-GalenPlus"
   ClientHeight    =   6510
   ClientLeft      =   -75
   ClientTop       =   -900
   ClientWidth     =   8880
   Icon            =   "principa.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "principa.frx":0442
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1191
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   510
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu menArchivos 
      Caption         =   "Archivos"
      Begin VB.Menu mnuGasto 
         Caption         =   "Usuarios y opciones que tienen asignados"
      End
      Begin VB.Menu mnuPacLolCli 
         Caption         =   "Mantenimiento de Pacientes LolCli"
      End
      Begin VB.Menu menSep1 
         Caption         =   "-"
      End
      Begin VB.Menu menSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu menProcesos 
      Caption         =   "Procesos"
      Begin VB.Menu mnuDP 
         Caption         =   "Carga EXCEL con CPT"
      End
      Begin VB.Menu migrarHBT 
         Caption         =   "Migra HBT"
      End
      Begin VB.Menu mnuJamoPac 
         Caption         =   "Migra Pacientes HRA, JAMO, SICUANI, Nazareno, San Juan"
      End
      Begin VB.Menu mnuOt11 
         Caption         =   "Compara Ubigeos GalenHos vs LolCli"
      End
      Begin VB.Menu mnuInvent 
         Caption         =   "Insertar Inventario segun ICI, IDI"
      End
      Begin VB.Menu mnuPtoCarga 
         Caption         =   "Varios Procesos"
      End
      Begin VB.Menu mnuAcTSaldo 
         Caption         =   "Actualiza Saldos y Fecha Vencimiento"
      End
      Begin VB.Menu mnuCuboSismed 
         Caption         =   "Reportes CUBO desde el SISMEDV2"
      End
   End
   Begin VB.Menu manAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu menAyu 
         Caption         =   "Contenido"
      End
      Begin VB.Menu menSep4 
         Caption         =   "-"
      End
      Begin VB.Menu menAcerca 
         Caption         =   "Acerca del Sistema "
      End
   End
End
Attribute VB_Name = "MDIfrmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programa principal
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Private Sub MDIForm_Load()
   'PAbreBD
'   st_Gen.Panels(1).Text = Right(wxBDLolcli, 70)
'   st_Gen.Panels(2).Text = Right(wxBDGalenhos, 70)
   'st_Gen.Panels(3).Text = RetornaFechaHoraServidorSQL()
End Sub

Function RetornaFechaServidorSQL() As String
    Dim oRsTmp As New ADODB.Recordset
    oRsTmp.Open "select  getdate() as FechaHoraSQL", wxConexionRed, adOpenKeyset, adLockOptimistic
    RetornaFechaServidorSQL = Format(oRsTmp.Fields!FechaHoraSQL, "DD/MM/YYYY")
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function
Function RetornaHoraServidorSQL() As String
    Dim oRsTmp As New ADODB.Recordset
    oRsTmp.Open "select  getdate() as FechaHoraSQL", wxConexionRed, adOpenKeyset, adLockOptimistic
    RetornaHoraServidorSQL = Format(oRsTmp.Fields!FechaHoraSQL, "HH:MM")
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function
Function RetornaFechaHoraServidorSQL() As String
    Dim oRsTmp As New ADODB.Recordset
    oRsTmp.Open "select  getdate() as FechaHoraSQL", wxConexionRed, adOpenKeyset, adLockOptimistic
    RetornaFechaHoraServidorSQL = Format(oRsTmp.Fields!FechaHoraSQL, "DD/MM/YYYY HH:MM")
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Private Sub MDIForm_Unload(Cancel As Integer)
   CierraConexiones
End Sub

Sub CierraConexiones()
   On Error Resume Next
   wxConexion.Close
   wxConexionRed.Close
   End
End Sub

Private Sub menSalir_Click()
   CierraConexiones
End Sub



Private Sub migrarHBT_Click()
   migraHBT.Show 1
End Sub

Private Sub mnuAcTSaldo_Click()
 '  HerrActualizaSaldo.Show 1
End Sub

Private Sub mnuCuboSismed_Click()
    ' rReportesSismedv2.Show 1
End Sub

Private Sub mnuDP_Click()
   
  mMantTabla.Show 1
End Sub

Private Sub mnuGasto_Click()
   mTablitas.Show 1
End Sub

Private Sub mnuInvent_Click()
   'InventarioInicial.Show 1
End Sub

Private Sub mnuJamoPac_Click()
   mPacientes.Show 1

End Sub

Private Sub mnuOt11_Click()
   'mUbigeo.Show 1
End Sub

Private Sub mnuPacLolCli_Click()
  frmActualizaTablas.Show 1
End Sub

Private Sub mnuPtoCarga_Click()
    mVariosProcesos.Show 1
End Sub

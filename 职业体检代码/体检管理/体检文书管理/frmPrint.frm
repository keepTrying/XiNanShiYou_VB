VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9555
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 cRepPrint 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      lastProp        =   600
      _cx             =   16748
      _cy             =   14208
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width > 120 Then
        cRepPrint.Width = Me.Width - 120
    End If
    cRepPrint.Left = 0
    cRepPrint.Top = 0
    If Me.Height > 500 Then
        cRepPrint.Height = Me.Height - 500
    End If
End Sub


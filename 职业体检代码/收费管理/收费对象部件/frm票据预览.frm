VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frm票据预览 
   Caption         =   "Form2"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9795
   StartUpPosition =   2  '屏幕中心
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8415
      lastProp        =   500
      _cx             =   14843
      _cy             =   12726
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton Ccmd退出 
      Caption         =   "返  回"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frm票据预览"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

'Private Sub Form_Resize()
'Debug.Print Width, Height

'End Sub
Private Sub Ccmd退出_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    CrystalReport1.WindowLeft = 0
'    CrystalReport1.WindowTop = 0
'    CrystalReport1.WindowWidth = Me.ScaleWidth
'    CrystalReport1.WindowHeight = Me.ScaleHeight
End Sub

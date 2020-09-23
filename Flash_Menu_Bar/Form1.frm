VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10b.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8775
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Swf1 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      _cx             =   23733
      _cy             =   15478
      FlashVars       =   ""
      Movie           =   "md"
      Src             =   "md"
      WMode           =   "Transparent"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "true"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Swf1.Movie = App.Path & "\MyMenuBar.Swf"
    Me.Top = 500
    Me.Left = 500
End Sub

Private Sub Swf1_FSCommand(ByVal command As String, ByVal args As String)
    Select Case command
        Case "Exit"
            If MsgBox("Do you want to exit?", vbQuestion + vbYesNo, "Exit") = vbYes Then
                Unload Me
            End If
        Case Else
            MsgBox "You clicked on " & UCase(command), vbInformation, "Clicked"
    End Select
    
End Sub



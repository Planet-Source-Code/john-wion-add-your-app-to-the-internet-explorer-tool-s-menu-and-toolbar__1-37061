VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Add your App to MSIE's Tools Menu and an Icon on the Toolbar (MSIE 5.x or higher)"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   4725
   ScaleWidth      =   9315
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddMSIE 
         Caption         =   "&Add to MSIE Tools Menu && Toolbar"
      End
      Begin VB.Menu mnuDeleteMSIE 
         Caption         =   "&Delete from MSIE Tools Menu && Toolbar"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To be honest I stolen this code from Chuck D

Option Explicit

Private Sub Form_Load()
' (OPTIONAL)
' Sets the Hook To SubClass the Menu (Label1.Caption)
' For a normal Menu remove Menustat.bas, Label1, and in Private Sub Form_Load()
' delete the line origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
' In Private Sub Form_Unload(Cancel As Integer) delete the line SetWindowLong hwnd, GWL_WNDPROC, origWndProc
' In Private Sub Form_Resize() delete the line Label1.Move 100, ScaleHeight - Label1.Height, ScaleWidth
'
origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
'
' Checks the Registry for the install path of MSIE
' then checks to see if the file exists in that path
' if the file exists in that path, MSIE's version info is checked
' to see if the version info = 5.x or higher. If it is, the menu allows you to add or delete
' your app in MSIE's Tools Menu & add an Icon on the Toolbar
' IF the file doesn't exists in that path, or MSIE's version < 5
' then the menu doesn't allow you to add or delete your app
' in MSIE's Tools Menu or add an Icon on the Toolbar
'
' Note: If Your App is already in MSIE's Tools Menu & an Icon is on the Toolbar
' the Menu structure will reflect that (Checks the Registry on Form_Load)
' A new instance of MSIE is required for changes to be seen (Add or Delete)
'
DetectIE
'
' SAFE to delete the following 2 MsgBoxes
'MsgBox "Make sure you compile into an exe, then run the exe (Running in design mode will reference IETOOLS.vbp instead of SampleApp.exe and MSIE will not find an *.exe to run !!!)"
MsgBox "Note: If the user has customized the toolbar, the button will not appear on the toolbar automatically. The toolbar button will be added to the choices in the Customize Toolbar dialog box and will appear if the toolbar is reset"
'
' Your Code...

End Sub

Private Sub Form_Resize()
' Move Label1
Label1.Move 100, ScaleHeight - Label1.Height, ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
' UnHooks the Menu SubClassing
SetWindowLong hwnd, GWL_WNDPROC, origWndProc
Shell "START http://www.wionj.net/"
End Sub

Private Sub mnuAddMSIE_Click()
' Adds Your App to MSIE's Tools Menu and add an Icon on the Toolbar
mnuAddIE
End Sub

Private Sub mnuDeleteMSIE_Click()
' Deletes Your App from MSIE's Tools Menu and the Icon on the Toolbar
mnuDeleteIE
End Sub

Private Sub mnuExit_Click()
' Unloads yor App
Unload Me
End Sub

 
   

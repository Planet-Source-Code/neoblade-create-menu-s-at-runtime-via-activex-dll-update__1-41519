VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dynamic Menus via Dll"
   ClientHeight    =   1950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   4440
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picBitmap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   4680
      Picture         =   "frmMain.frx":04CE
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblDecscription 
      Caption         =   "The Menu Dll now support PopUp menu's, and icons on the menu items. Right Click to see the Menu specified PopUp."
      Height          =   675
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   3150
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   45
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'NOTE: A menu must be added to the form using traditional methods
'any entries in this menu will be deleted, it just needs the object initially

'Make sure withevents is declared to get the mouse events
Private WithEvents MenuBar As clsNewMenu
Attribute MenuBar.VB_VarHelpID = -1


Private Sub Form_Load()

Set MenuBar = New clsNewMenu


With MenuBar
.Create Me.hWnd
.AddMenu "MENU 0"
.AddMenu "MENU 1"
.AddMenu "MENU 2"

.SubMenu(0).AddMenu "MENU 0 - SUBMENU 0"
'''To add an icon to the menu you need to set the picture to the image
picIcon.Picture = picIcon.Image
.SubMenu(0).AddPicture picIcon.Picture

.SubMenu(1).AddMenu "MENU 1 - SUBMENU 0"
.SubMenu(1).AddMenu "MENU 1 - SUBMENU 1"
.SubMenu(1).SubMenu(1).AddMenu "MENU 1 - SUB-SUBMENU 0"
.SubMenu(1).SubMenu(1).AddPicture picBitmap.Picture
.SubMenu(1).AddMenu "MENU 1 - SUBMENU 2"


.SubMenu(2).AddMenu "MENU 2 - SUBMENU 0"
.SubMenu(2).SubMenu(0).AddMenu "MENU 2 - SUB-SUBMENU 0"
.SubMenu(2).SubMenu(0).SubMenu(0).AddMenu "MENU 2 - SUB-SUB-SUBMENU 0"
.SubMenu(2).AddMenu "Seperator", True
.SubMenu(2).AddMenu "MENU 2 - SUBMENU 2"
.SubMenu(2).SubMenu(2).AddMenu "MENU 2 - SUB-SUBMENU 0"

lblCount.Caption = .SubMenu(1).Caption & " contains " & .SubMenu(1).Count & " menu's"

End With
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then MenuBar.PopUp MenuBar.SubMenu(1)
End Sub

Private Sub MenuBar_MenuClicked(MenuId As Long, Caption As String)
MsgBox "Caption: " & Caption & vbCrLf & "MenuId: " & MenuId
End Sub

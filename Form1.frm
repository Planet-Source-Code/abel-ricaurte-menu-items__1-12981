VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0090
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0120
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":01B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0240
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0360
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":03F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0488
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Begin VB.Menu mnuClip 
            Caption         =   "Clip On / Off"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private Sub Form_Load()
  Dim mHandle As Long, lRet As Long, sHandle As Long, sHandle2 As Long

  mHandle = GetMenu(hwnd)
  sHandle = GetSubMenu(mHandle, 0)  ' Menu #0

  lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, ImageList.ListImages(1).Picture, ImageList.ListImages(1).Picture)
  lRet = SetMenuItemBitmaps(sHandle, 1, MF_BYPOSITION, ImageList.ListImages(2).Picture, ImageList.ListImages(2).Picture)
  lRet = SetMenuItemBitmaps(sHandle, 2, MF_BYPOSITION, ImageList.ListImages(3).Picture, ImageList.ListImages(3).Picture)
  lRet = SetMenuItemBitmaps(sHandle, 4, MF_BYPOSITION, ImageList.ListImages(4).Picture, ImageList.ListImages(4).Picture)

  mHandle = GetMenu(hwnd)
  sHandle = GetSubMenu(mHandle, 1)  ' Menu #1

  lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, ImageList.ListImages(5).Picture, ImageList.ListImages(5).Picture)
  lRet = SetMenuItemBitmaps(sHandle, 1, MF_BYPOSITION, ImageList.ListImages(6).Picture, ImageList.ListImages(6).Picture)
  lRet = SetMenuItemBitmaps(sHandle, 2, MF_BYPOSITION, ImageList.ListImages(7).Picture, ImageList.ListImages(7).Picture)

  sHandle = GetSubMenu(mHandle, 1)  ' Menu #1
  sHandle2 = GetSubMenu(sHandle, 4) ' SubMenu Position

  lRet = SetMenuItemBitmaps(sHandle2, 0, MF_BYPOSITION, ImageList.ListImages(8).Picture, ImageList.ListImages(9).Picture)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' mnuCut Bold Effect
  If Button = vbRightButton Then PopupMenu mnuEdit, 0, , , mnuCut
End Sub

Private Sub mnuClip_Click()
  mnuClip.Checked = Not mnuClip.Checked ' Hide/Show Clip item
End Sub

Private Sub mnuQuit_Click()
  Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Menu Example"
   ClientHeight    =   2415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdAdd2 
         Caption         =   "Add to Sample Menu 2"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtNew 
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   2
         Text            =   "Ne&w Menu Item 1"
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton cmdAdd1 
         Caption         =   "Add to Sample Menu 1"
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: adding an ampersand (&&) to the menu item creates a shortcut key!"
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type the name of the new menu item in the box below:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Menu mnuSampleMenu1 
      Caption         =   "&Sample Menu 1"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSampleMenu2 
      Caption         =   "S&ample Menu 2"
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Background Color"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Type Declarations
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'Private API Declarations
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

'Private Constants
Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_SEPARATOR = &H800
Private Const MFT_STRING = &H0
Private Const MFS_ENABLED = &H0
Private Const MFS_CHECKED = &H8

'Form variables
Dim Message As Variant 'for displaying message boxes

Private Sub cmdAdd1_Click()
    Call AddNewMenuItem(0) 'add a new menu items to the first sample menu (zero-based array)
End Sub

Private Sub cmdAdd2_Click()
    Call AddNewMenuItem(1) 'add a new menu items to the second sample menu (zero-based array)
End Sub

Private Sub Form_Load()
    itemID = 1000 'initialize to one thousand in order to avoid conflicts with existing menu item IDs
    oldWindowProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WindowProc) 'set up a new window procedure for this form and save a pointer to the original one as 'oldWindowProc'
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim retval As Long  'holds the return value

    Message = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit")
    If Message = vbNo Then 'if user clicked 'No'
        Cancel = 1 'don't unload this form
    Else
        retval = SetWindowLong(Me.hWnd, GWL_WNDPROC, oldWindowProc) 'restore this window's original procedure before it unloads
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me 'Unload this form
End Sub

Private Sub mnuRestore_Click()
    Me.BackColor = &H8000000F
End Sub

Private Sub AddNewMenuItem(SubMenu As Integer)
    Dim newItem As String 'holds the item that the user wants to add to the menu
    Dim hMenu As Long 'handle to this form's menu
    Dim hSubMenu 'handle to one of the sub menus
    Dim iNum As Long 'the number of items currently on the menu
    Dim menuInfo As MENUITEMINFO 'holds info about the current menu
    
    newItem = txtNew.Text
    If Len(newItem) < 1 Then 'if the user didn't enter anything
        Message = MsgBox("Please enter the name of the new menu item.", vbExclamation + vbOKOnly, "Error")
        Exit Sub
    End If
    
    hMenu = GetMenu(Me.hWnd) 'retreive a handle to this form's menu
    hSubMenu = GetSubMenu(hMenu, SubMenu) 'retreive a handle to the submenu (0-based array)
    iNum = GetMenuItemCount(hSubMenu) 'determine how many items are currently on this menu
    
    If iNum = 1 Then 'if there is currently only one menu item in the menu (Exit) then add a separator bar to the menu
        With menuInfo 'add a separator bar to this menu
            .cbSize = Len(menuInfo) 'set the length of the menu structure
            .fMask = MIIM_ID Or MIIM_TYPE 'declare which parts of the menu structure to use
            .fType = MFT_SEPARATOR 'define the type of menu item (separator)
            .wID = itemID 'set the structure ID
        End With
        Call InsertMenuItem(hSubMenu, iNum - 1, 1, menuInfo) 'add the separator bar to the menu
        itemID = itemID + 1 'increment for the next new menu item
    End If

    With menuInfo 'add the new menu item to the top of this menu
        .cbSize = Len(menuInfo) 'set the length of the menu structure
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE 'declare which parts of the menu structure to use
        .fType = MFT_STRING 'define the type of menu item (text)
        .fState = MFS_ENABLED 'this item should be enabled on the menu
        .dwTypeData = newItem 'the text of the new menu item
        .cch = Len(.dwTypeData)
        .wID = itemID 'set the structure ID (this ID is used to add functionality to this menu item)
    End With
    Call InsertMenuItem(hSubMenu, 0, 1, menuInfo)
    itemID = itemID + 1 'increment for the next new menu item
    
    txtNew.Text = "" 'clear the text box
End Sub

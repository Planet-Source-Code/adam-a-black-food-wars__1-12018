VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W-MART"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmStore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&BUY"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Add to shopping basket"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&EXIT"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstWeapons 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   2566
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Weapon"
         Text            =   "Weapon"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   2999
      EndProperty
   End
   Begin MSComctlLib.ListView lstBought 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   2566
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Weapon"
         Text            =   "Weapon"
         Object.Width           =   6438
      EndProperty
   End
   Begin VB.Label lblYourItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Items"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   750
   End
   Begin VB.Label lblWeapons 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weapons"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Price(1 To 4) As Long

Private Sub cmdAdd_Click()
    Dim i As Byte
    If YW(lstWeapons.SelectedItem.Index) = True Then Exit Sub
    If lstWeapons.SelectedItem.ListSubItems(1).text > Credit Then
        MsgBox "You cannot afford to pay for this"
        Else
        Credit = Credit - lstWeapons.SelectedItem.ListSubItems(1).text
        YW(lstWeapons.SelectedItem.Index) = True
        frmMain.lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
        lstBought.ListItems.Clear
        For i = 1 To UBound(Weapon)
        If YW(i) = True Then
            lstBought.ListItems.Add , , Weapon(i)
        End If
    Next
    End If
    If YW(1) = True And YW(2) = True And YW(3) = True And YW(4) = True Then cmdAdd.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload frmStore
End Sub

Private Sub Form_Load()
    Dim i As Byte
    Price(1) = 60000
    Price(2) = 100000
    Price(3) = 1500000
    Price(4) = 3500000
    If YW(1) = True And YW(2) = True And YW(3) = True And YW(4) = True Then cmdAdd.Enabled = False
    For i = 1 To UBound(Weapon)
        lstWeapons.ListItems.Add i, , Weapon(i)
        lstWeapons.ListItems(i).ListSubItems.Add , , Format(Price(i), "###,###,###")
    Next
    For i = 1 To 4
        If YW(i) = True Then
            lstBought.ListItems.Add , , Weapon(i)
        End If
    Next
End Sub

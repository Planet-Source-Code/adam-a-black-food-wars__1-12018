VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Food Wars"
   ClientHeight    =   6705
   ClientLeft      =   615
   ClientTop       =   615
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History"
      Height          =   375
      Left            =   2820
      TabIndex        =   9
      ToolTipText     =   "Prices History"
      Top             =   5715
      Width           =   1095
   End
   Begin VB.CommandButton cmdDoctor 
      Caption         =   "Doctor"
      Height          =   375
      Left            =   2820
      TabIndex        =   7
      Top             =   4875
      Width           =   1095
   End
   Begin VB.CommandButton cmdStore 
      Caption         =   "Store"
      Height          =   375
      Left            =   2820
      TabIndex        =   8
      Top             =   5295
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   3960
      ScaleHeight     =   1755
      ScaleWidth      =   3075
      TabIndex        =   22
      Top             =   360
      Width           =   3135
      Begin VB.CommandButton cmdBILO 
         Caption         =   "&BI-LO"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdFranklins 
         Caption         =   "&Franklins"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdIGA 
         Caption         =   "&IGA"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSafeway 
         Caption         =   "&Safeway"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdMaxi 
         Caption         =   "&Maxi"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdColes 
         Caption         =   "&Coles"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   375
      Left            =   2820
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculator"
      Height          =   375
      Left            =   2820
      TabIndex        =   6
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdFinances 
      Caption         =   "Finances"
      Height          =   375
      Left            =   2820
      TabIndex        =   5
      Top             =   4035
      Width           =   1095
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "<< Sell"
      Height          =   375
      Left            =   2820
      TabIndex        =   4
      Top             =   3615
      Width           =   1095
   End
   Begin VB.CommandButton cmdSteal 
      Caption         =   "Steal >>"
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      Top             =   3200
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy >>"
      Height          =   375
      Left            =   2820
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   9
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":071A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstFoods 
      Height          =   4095
      Left            =   75
      TabIndex        =   0
      Top             =   2520
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   7223
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ColHdrIcons     =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Food"
         Text            =   "Food"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView lstItems 
      Height          =   4095
      Left            =   3960
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7223
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Food"
         Text            =   "Food"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Qty"
         Text            =   "Qty"
         Object.Width           =   1288
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   3315
      TabIndex        =   17
      Top             =   120
      Width           =   3375
      Begin FoodWars.ProgBar pbHealth 
         Height          =   255
         Left            =   1080
         Top             =   1575
         Width           =   2055
         _extentx        =   3625
         _extenty        =   450
         backcolor       =   16384
         barcolor        =   16711680
         value           =   100
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Day:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Items:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Top             =   480
         Width           =   2190
      End
      Begin VB.Label lblCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   960
         TabIndex        =   24
         Top             =   120
         Width           =   2190
      End
      Begin VB.Label dspHealth 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Health:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label dspCash 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   675
      End
      Begin VB.Label dspDebit 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.Label lblPlace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You are now in "
      Height          =   210
      Left            =   3960
      TabIndex        =   28
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblYourItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Items"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3960
      TabIndex        =   19
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label lblAvailable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supermarket Items:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   18
      Top             =   2280
      Width           =   1515
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFinances 
         Caption         =   "&Finances"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewScores 
         Caption         =   "&High Scores"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "&History"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSounds 
         Caption         =   "&Enable Sounds"
         Begin VB.Menu mnuTruck 
            Caption         =   "&Truck sound"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOther 
            Caption         =   "&Other sounds"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelpMain 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------'
'Welcome to the code for Food Wars. some of this code is quite complex and there'
'is some parts of code you may not understand but some of it is commented so I  '
'hope you can learn from some of the code here and enjoy the clone of Dopewars. '
'The High score list module wasn't created by myself and I actually obtained it '
'from planet-source-code but every other piece of code is.                      '
'-------------------------------------------------------------------------------'
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const BM_SETSTYLE = &HF4
Private Const BS_SOLID = 0
Public Day As Integer

Private Sub cmdBilo_Click()
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.15)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in BILO"
    lblDay.Caption = "Day: " & Day & " of 35"
    cmdBILO.Enabled = False
    EnableOthers cmdColes, cmdSafeway, cmdIGA, cmdFranklins, cmdMaxi
    Call AddPrices
End Sub

Private Sub cmdBuy_Click()
    If iSpace = 0 Then Exit Sub
    If lstFoods.SelectedItem.ListSubItems(1).text > Credit Then
        MsgBox "You can't afford it, borrow some money if you really want it!", vbExclamation
        Else
        frmBuy.Show vbModal
    End If
End Sub

Private Sub cmdCalc_Click()
    On Error Resume Next
    Shell "Calc", vbNormalFocus
End Sub

Private Sub cmdColes_Click()
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.15)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Coles"
    lblDay.Caption = "Day: " & Day & " of 35"
    cmdColes.Enabled = False
    Call EnableOthers(cmdFranklins, cmdIGA, cmdMaxi, cmdSafeway, cmdBILO)
    Call AddPrices
End Sub

Private Sub cmdDoctor_Click()
    If Health = 100 Then MsgBox "You don't need a doctor at the moment": Exit Sub
    If Credit < 10000 Then MsgBox "You can't afford a doctor": Exit Sub
    PlaySound SDir & "Doctor.wav", 0, 3
    frmDoctor.Show vbModal
End Sub

Private Sub cmdFinances_Click()
    frmFinances.Show vbModal
End Sub

Private Sub cmdFranklins_Click()
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.15)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Franklins"
    lblDay.Caption = "Day: " & Day & " of 35"
    cmdFranklins.Enabled = False
    EnableOthers cmdIGA, cmdMaxi, cmdColes, cmdSafeway, cmdBILO
    Call AddPrices
End Sub

Private Sub cmdHistory_Click()
    frmHistory.Start
    frmHistory.Show vbModal
End Sub

Private Sub cmdIGA_click()
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.15)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in IGA"
    lblDay.Caption = "Day: " & Day & " of 35"
    cmdIGA.Enabled = False
    EnableOthers cmdFranklins, cmdColes, cmdMaxi, cmdBILO, cmdSafeway
    Call AddPrices
End Sub

Private Sub cmdMaxi_Click()
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.15)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Maxi"
    lblDay.Caption = "Day: " & Day & " of 35"
    cmdMaxi.Enabled = False
    EnableOthers cmdBILO, cmdColes, cmdFranklins, cmdIGA, cmdSafeway
    Call AddPrices
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdNewGame_Click()
    mnuNew_Click
    cmdBuy.SetFocus
End Sub

Private Sub cmdSafeway_Click()
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.15)
    lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Safeway"
    lblDay.Caption = "Day: " & Day & " of 35"
    cmdSafeway.Enabled = False
    EnableOthers cmdBILO, cmdColes, cmdFranklins, cmdIGA, cmdMaxi
    Call AddPrices
End Sub

Private Sub cmdSell_Click()
    If Used = 0 Then
        MsgBox "You have nothing to sell", vbInformation
        Else
        frmSell.Show vbModal
    End If
End Sub

Private Sub cmdSteal_Click()
    If iSpace > 0 Then frmSteal.Show vbModal
End Sub

Private Sub cmdStore_Click()
    frmStore.Show vbModal
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then End
    If Right(App.Path, 1) = "\" Then
        SDir = App.Path & "Sounds\"
        Else
        SDir = App.Path & "\Sounds\"
    End If
    Truck = SDir & "Truck.wav"
    CButton cmdColes
    CButton cmdMaxi
    CButton cmdSafeway
    CButton cmdIGA
    CButton cmdFranklins
    CButton cmdBILO
    Call Init
End Sub

Public Sub Init()
    Dim Rand As Byte
    Randomize
    mnuNew.Enabled = False
    cmdNewGame.Enabled = False
    Sound
    EnableControls
    ResetCaptions
    YW(1) = False
    YW(2) = False
    YW(3) = False
    YW(4) = False
    Call HSL.SetListCount(20)
    Call HSL.FileName(HSL.DefaultFileName)
    FinalScore = 0
    Credit = 1000
    Debit = 3000
    Day = 1
    iSpace = 200
    Used = 0
    TotalSpace = 200
    Caught = 0
    Health = 100
    Stole = False
    Attacked = False
    Sold = False
    cmdNewGame.Enabled = False
    cmdColes.Enabled = False
    EnableOthers cmdFranklins, cmdBILO, cmdIGA, cmdSafeway, cmdMaxi
    lblDay.Caption = "Day: " & Day & " of 35"
    lblPlace.Caption = "You are now in Coles"
    lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
    lblItems.Caption = "Items: 0 of 200"
    pbHealth.Value = Health
    lstFoods.ListItems.Clear
    lstItems.ListItems.Clear
    Erase Avg
    Erase Quantity
    Call AddFoods
    Call AddPrices
    Call AddWeapons
End Sub

Private Sub lblDay_Change()
    Dim iRnd As Byte
    Dim Ans As Byte
    Dim j As Byte, i As Byte
    Randomize
    If Day >= 15 Then
        cmdNewGame.Enabled = True
        mnuNew.Enabled = True
    End If
    If Day > 1 Then
        PlaySound Truck, 0, 3
        iRnd = Int(Rnd * 70)
        Select Case iRnd
            Case 5, 32, 18
            'Checks to see if you own the 200 space truck and you can afford it.
            If Credit > 200000 And TotalSpace = 200 Then
                Ans = MsgBox("There is a new truck available that can hold a total of 400 items, do you want to buy this for 200,000?", vbYesNo + vbQuestion)
                If Ans = vbYes Then
                    Credit = Credit - 200000
                    TotalSpace = 400
                    iSpace = iSpace + 200
                    lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
                    lblItems.Caption = "Items: " & Used & " of " & TotalSpace
                End If
            End If
            
            Case 13
            If iSpace >= 10 And Quantity(13) = 0 Then
                MsgBox "You found 10 packets of chocolates on a dead dude in the subway!", vbInformation
                lstItems.ListItems.Clear
                Quantity(13) = 10
                Avg(13) = 0
                j = 1
                Used = Used + 10
                iSpace = iSpace - 10
                lblItems.Caption = "Items: " & Used & " of " & TotalSpace
                For i = 1 To 17
                    If Quantity(i) > 0 Then
                        frmMain.lstItems.ListItems.Add j, , frmMain.lstFoods.ListItems(i)
                        frmMain.lstItems.ListItems(j).ListSubItems.Add , , Round(Avg(i), 0)
                        frmMain.lstItems.ListItems(j).ListSubItems.Add , , Quantity(i)
                        j = j + 1
                    End If
                Next
            End If
            
            Case 30
            If Credit >= 10 Then
                MsgBox "Two guys attacked you cause they wanted some money, you gave them " & Format(Round(Credit / 3, 2), "###,###,###") & " and ran.", vbExclamation
                Credit = Credit - (Credit / 3)
                lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
            End If
            
            Case 41
            If Credit > 500000 And TotalSpace = 400 Then
                Ans = MsgBox("There is a new truck available that can hold a total of 600 items, do you want to buy this for 500,000?", vbYesNo + vbQuestion)
                If Ans = vbYes Then
                    Credit = Credit - 500000
                    TotalSpace = 600
                    iSpace = iSpace + 200
                    lblItems.Caption = "Items: " & Used & " of " & TotalSpace
                    lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
                End If
            End If
            
            Case 53
            If Credit >= 10 Then
                MsgBox "The cops insist that you stole a bike, even though you didn't. You have to pay a fine.", vbExclamation
                Credit = Credit - (Credit * 0.1) 'or credit/10 of course
                lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
            End If
            
            Case 9, 68, 24, 27, 15, 47, 42, 54, 58
            If Sold = True Then
                frmManager.Show vbModal
            End If
            
            Case 3, 26, 40, 55, 34, 38, 12, 45, 59, 1, 48
            If Attacked = True Or Stole = True Then
                PlaySound SDir & "Police.wav", 0, 3
                frmPolice.Show vbModal
            End If
        End Select
    End If
End Sub

Private Sub lblDebit_Change()
    If Debit > 3500000 And Debit < 4000000 Then MsgBox "The bank want their money, they don't trust you with that huge debit. Pay up or else they will take you to court", vbExclamation
    If Debit > 4000000 Then
        If Credit > Debit Then
            Dim Temp As Long
            MsgBox "The bank wants their money, you have enough money to pay off this loan. You have to pay it now", vbInformation
            Do
                frmFinances.Show vbModal
            Loop Until Debit = 0
            lblDebit.Caption = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
            lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
        End If
        If Credit < Debit Then
            Dim Ret As Byte
            MsgBox "You don't have enough money to pay of this loan. Next time try to keep track of your debit", vbCritical
            Ret = MsgBox("Do you want to play again?", vbYesNo + vbQuestion)
            If Ret = vbYes Then
                frmMain.mnuNew_Click
                Else
                End
            End If
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmSplash.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuFinances_Click()
    frmFinances.Show
End Sub

Private Sub mnuHelp_Click()
    MsgBox "I didn't bother adding help since this game shouldn't be hard to understand. If there is a bug or you cant understand how to use a feature on this game email adz8@softhome.net"
End Sub

Private Sub mnuHistory_Click()
    frmHistory.Start
    frmHistory.Show vbModal
End Sub

Public Sub mnuNew_Click()
    Call Init
End Sub

Private Function EnableOthers(Cmd1 As CommandButton, Cmd2 As CommandButton, Cmd3 As CommandButton, Cmd4 As CommandButton, Cmd5 As CommandButton)
    Cmd1.Enabled = True
    Cmd2.Enabled = True
    Cmd3.Enabled = True
    Cmd4.Enabled = True
    Cmd5.Enabled = True
End Function

Private Function CButton(Button As CommandButton) As Long
    SendMessage Button.hwnd, BM_SETSTYLE, BS_SOLID, 1
End Function

Private Sub mnuOther_Click()
    If mnuOther.Checked = True Then
        mnuOther.Checked = False
        Else
        mnuOther.Checked = True
    End If
    SaveSetting "Food Wars", "Options", "Sounds", mnuOther.Checked
    Sound
End Sub

Private Sub mnuTruck_Click()
    If mnuTruck.Checked = True Then
        mnuTruck.Checked = False
        Else
        mnuTruck.Checked = True
    End If
    SaveSetting "Food Wars", "Options", "Truck", mnuTruck.Checked
    Sound
End Sub

Private Sub mnuViewScores_Click()
    Call HSL.FillScoreList(frmScores.lstScores)
    frmScores.Show vbModal
End Sub

Private Function CheckScore() As Boolean
    Dim i As Integer
    Dim Pass As Boolean
    Dim Cnt As Byte
    Dim Success As Boolean
    Success = False
    Cnt = 0
    CheckScore = False
    If Day = 34 Then CheckScore = True: ChangeCaptions
    If Day <> 35 Then CheckScore = True: Exit Function
    FinalScore = Credit
    For i = 1 To 17
        If Quantity(i) > 0 Then
            Cnt = Cnt + 1
        End If
    Next
    If Cnt > 0 Then MsgBox "You have to get rid of all your food before you can finish off"
    If Cnt = 0 Then
        DisableControls
        Call HSL.FillScoreList(frmScores.lstScores)
        Success = HSL.AddHighScore(frmScores.lstScores, User, Int(FinalScore))
        CheckScore = True
        If Success = True Then
            frmScores.Show vbModal
            Exit Function
            Else
            i = MsgBox("Your time has run out. You can play a new game or quit?")
        End If
    End If
End Function

Private Sub ChangeCaptions()
    cmdColes.Caption = "FINISH"
    cmdMaxi.Caption = "FINISH"
    cmdSafeway.Caption = "FINISH"
    cmdIGA.Caption = "FINISH"
    cmdBILO.Caption = "FINISH"
    cmdFranklins.Caption = "FINISH"
End Sub

Private Sub ResetCaptions()
    cmdColes.Caption = "&Coles"
    cmdMaxi.Caption = "&Maxi"
    cmdSafeway.Caption = "&Safeway"
    cmdIGA.Caption = "&IGA"
    cmdBILO.Caption = "&BI-LO"
    cmdFranklins.Caption = "&Franklins"
End Sub

Private Sub DisableControls()
    'This sub will be called when you approach the end of the days
    'first it disables all the COMMAND BUTTONS and then enabled the few needed
    'so the user can still view high scores, start a new game and view prices from
    'other days (this is more efficent than disabling only the command buttons
    'needed)
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Controls.Count - 1
        If TypeOf Controls(i) Is CommandButton Then
            Controls(i).Enabled = False
        End If
    Next
    cmdHistory.Enabled = True 'So they can view prices
    cmdNewGame.Enabled = True 'So they can start a new game
    mnuFinances.Enabled = False 'so they cant mess around with their previous game
End Sub

Private Sub EnableControls()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Controls.Count - 1
        If TypeOf Controls(i) Is CommandButton Then
            Controls(i).Enabled = True
        End If
    Next
    mnuFinances.Enabled = True
End Sub

Private Sub Sound()
    mnuTruck.Checked = GetSetting("Food Wars", "Options", "Truck", True)
    mnuOther.Checked = GetSetting("Food Wars", "Options", "Sounds", True)
    If mnuOther.Checked = False Then
        'SDir is needed for the program to know where the sounds are and resetting
        'SDir to nothing is easier than making IF statements for every time a
        'sound is going to be played.
        SDir = ""
        Else
        If Right(App.Path, 1) = "\" Then
            SDir = App.Path & "Sounds\"
            Else
            SDir = App.Path & "\Sounds\"
        End If
    End If
    If mnuTruck.Checked = False Then
        Truck = vbNullString
        Else
        If Right(App.Path, 1) = "\" Then
            Truck = App.Path & "Sounds\"
            Else
            Truck = App.Path & "\Sounds\"
        End If
        Truck = Truck & "Truck.wav"
    End If
End Sub

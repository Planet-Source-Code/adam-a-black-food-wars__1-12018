VERSION 5.00
Begin VB.Form frmSteal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Steal"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmSteal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.HScrollBar scr 
      Height          =   135
      Left            =   120
      Max             =   1000
      Min             =   1
      TabIndex        =   1
      Top             =   1320
      Value           =   1
      Width           =   855
   End
   Begin VB.CommandButton cmdSteal 
      Caption         =   "&Steal"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item(s)"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   1005
      Width           =   465
   End
   Begin VB.Label lblInform 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSteal.frx":0442
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSteal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmSteal
End Sub

Private Sub cmdSteal_Click()
    On Error Resume Next
    Randomize
    Dim Ans As Byte
    Dim Rand As Integer
    Dim i As Integer
    Dim j As Byte
    Dim EndMsg As String
    j = 1
    'If you don't have enough space for the amount of units you want to buy don't go ahead
    If txtQty.text > iSpace Then Exit Sub
    Rand = Int(5 * Rnd) + 1
    Stole = True
    'If Rand=2 then you get away with stealing. (1 in 5 chance)
    If Not Rand = 2 Then
        'If you have stolen 3 times since the start of the new game
        If Caught = 2 Then
            Ans = MsgBox("You have stolen too many times and will be arrested. Do you want to play again?", vbYesNo + vbQuestion)
            Unload frmSteal
            If Ans = vbYes Then
                frmMain.mnuNew_Click
                Exit Sub
            End If
            End
            Exit Sub
        End If
        If Credit < 1 Then
            EndMsg = "You have been caught, 0 has been deducted from your account"
            Else
            EndMsg = Credit * 0.9 & " has been deducted from your account"
        End If
        MsgBox EndMsg, vbExclamation
        'Deduct 90% of money from account as consequence
        Credit = Credit - (Credit * 9 / 10)
        Caught = Caught + 1
        frmMain.lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
        Unload frmSteal
        Exit Sub
    End If
    frmMain.lstItems.ListItems.Clear
    'The exact same code is commented in the frmBuy form.
    If Quantity(frmMain.lstFoods.SelectedItem.Index) > 0 Then
        Quantity(frmMain.lstFoods.SelectedItem.Index) = Quantity(frmMain.lstFoods.SelectedItem.Index) + txtQty.text
        Else
        Quantity(frmMain.lstFoods.SelectedItem.Index) = txtQty.text
    End If
    If Avg(frmMain.lstFoods.SelectedItem.Index) > 0 Or Not Quantity(frmMain.lstFoods.SelectedItem.Index) = 0 Then
        Avg(frmMain.lstFoods.SelectedItem.Index) = ((Quantity(frmMain.lstFoods.SelectedItem.Index) - txtQty.text) * (Avg(frmMain.lstFoods.SelectedItem.Index)) + txtQty.text * 0) / (txtQty.text + (Quantity(frmMain.lstFoods.SelectedItem.Index) - txtQty.text))
        Else
        Avg(frmMain.lstFoods.SelectedItem.Index) = 0
    End If
    For i = 1 To 17
        If Quantity(i) > 0 Then
            frmMain.lstItems.ListItems.Add j, , frmMain.lstFoods.ListItems(i)
            frmMain.lstItems.ListItems(j).ListSubItems.Add , , Round(Avg(i), 0)
            frmMain.lstItems.ListItems(j).ListSubItems.Add , , Quantity(i)
            j = j + 1
        End If
    Next
    iSpace = iSpace - txtQty.text
    Used = Used + txtQty.text
    frmMain.lblItems = "Items: " & Used & " of " & TotalSpace
    PlaySound SDir & "cashreg.wav", 0, 3
    Unload frmSteal
End Sub

Private Sub Form_Load()
    scr.Max = iSpace
    scr.Value = iSpace
    txtQty.text = iSpace
    txtQty.SelLength = Len(txtQty)
End Sub

Private Sub scr_Change()
    txtQty.text = scr.Value
End Sub


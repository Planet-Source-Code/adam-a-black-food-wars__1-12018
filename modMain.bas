Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public User As String 'Name when they login to Food Wars
Public History(1 To 35, 1 To 17)
Public Foods(1 To 17) As String 'Foods available to buy
Public Prices(1 To 17) As Long 'Price of foods
Public Quantity(1 To 17) As Integer 'Quantity of items purchased
Public Weapon(1 To 4) As String 'Available weapons
Public YW(1 To 4) As Boolean 'The weapons you have
Public Credit As Double 'your money
Public Debit As Long 'no need for double cant go past 4,000,000
Public Avg(1 To 17) As Double 'Avg of bought prices
Public iSpace As Integer 'iSpace remaining
Public Used As Integer 'Used truck spaces
Public TotalSpace As Integer 'Total truck space
Public Caught As Byte 'how many times caught
Public Stole As Boolean 'If items stolen before
Public SDir As String 'Directory with sound files
Public Truck As String 'Directory with Truck wav
Public Sold As Boolean 'Whether you have traded
Public Attacked As Boolean 'Whether you have attacked anyone with a weapon
Public Health As Integer
Public HSL As New HighScoreList
Public FinalScore As Double 'Final score.

Public Sub AddWeapons()
    Weapon(1) = "9mm"
    Weapon(2) = "Magnum"
    Weapon(3) = "Machine gun"
    Weapon(4) = "Rocket launcher"
End Sub

Public Sub AddFoods()
    Dim i As Byte 'Counter
    Foods(1) = "Apricots"
    Foods(2) = "Biscuits"
    Foods(3) = "Bread"
    Foods(4) = "Butter"
    Foods(5) = "Cake Mix"
    Foods(6) = "Cream"
    Foods(7) = "Frozen Burgers"
    Foods(8) = "Frozen Chips"
    Foods(9) = "Frozen Pizza"
    Foods(10) = "Icecream"
    Foods(11) = "Jelly"
    Foods(12) = "Milo"
    Foods(13) = "Chocolates"
    Foods(14) = "BBQ Chicken"
    Foods(15) = "Drinks"
    Foods(16) = "Chips (12 pack)"
    Foods(17) = "Rolls (16 pack)"
    For i = 1 To 17
        frmMain.lstFoods.ListItems.Add , , Foods(i)
    Next
End Sub

Public Sub AddPrices()
    Dim Temp(1 To 17) As Integer 'Store actual prices before randomized
    Dim iRnd As Byte
    Dim j As Integer
    Dim i As Byte 'Counter
    Randomize
    Temp(1) = 100
    Temp(2) = 40
    Temp(3) = 70
    Temp(4) = 430
    Temp(5) = 120
    Temp(6) = 300
    Temp(7) = 600
    Temp(8) = 800
    Temp(9) = 6000
    Temp(10) = 500
    Temp(11) = 90
    Temp(12) = 410
    Temp(13) = 1000
    Temp(14) = 22000
    Temp(15) = 900
    Temp(16) = 130
    Temp(17) = 70
    For i = 1 To 17
        Do
            Prices(i) = Int((Temp(i) / 0.7) * Rnd) + 1
        Loop Until Prices(i) > Temp(i) - Temp(i) * 0.3
    Next
    iRnd = Int(6 * Rnd) + 1
    If iRnd = 2 Then
        j = Int(Rnd * 17) + 1
        Prices(j) = Prices(j) * 5
        MsgBox "The price of " & Foods(j) & " has raised!", vbInformation
    End If
    If iRnd = 4 Then
        j = Int(Rnd * 17) + 1
        Prices(j) = Prices(j) / 5
        MsgBox "Today's special is on " & Foods(j), vbInformation
    End If
    For i = 1 To 17
        frmMain.lstFoods.ListItems(i).ListSubItems.Clear
    Next
    For i = 1 To 17
        History(frmMain.Day, i) = Prices(i)
        frmMain.lstFoods.ListItems(i).ListSubItems.Add , , Prices(i)
        If Prices(i) > Temp(i) Then
            frmMain.lstFoods.ListItems(i).SmallIcon = frmMain.imgList.ListImages(1).Index
        End If
        If Prices(i) < Temp(i) Then
            frmMain.lstFoods.ListItems(i).SmallIcon = frmMain.imgList.ListImages(2).Index
        End If
        If Prices(i) = Temp(i) Then
            frmMain.lstFoods.ListItems(i).SmallIcon = frmMain.imgList.ListImages(3).Index
        End If
    Next
End Sub

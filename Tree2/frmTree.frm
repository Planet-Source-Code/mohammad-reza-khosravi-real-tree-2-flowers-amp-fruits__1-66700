VERSION 5.00
Begin VB.Form frmTree 
   AutoRedraw      =   -1  'True
   Caption         =   "Tree"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "frmTree.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6705
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrFruits 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1680
      Top             =   2430
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Index           =   4
      Left            =   5760
      Picture         =   "frmTree.frx":0442
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   12060
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5820
      Index           =   3
      Left            =   4170
      Picture         =   "frmTree.frx":5B4F
      ScaleHeight     =   5760
      ScaleWidth      =   8010
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3330
      Visible         =   0   'False
      Width           =   8070
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8160
      Index           =   2
      Left            =   2370
      Picture         =   "frmTree.frx":AB51
      ScaleHeight     =   8100
      ScaleWidth      =   11520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3450
      Visible         =   0   'False
      Width           =   11580
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6300
      Index           =   1
      Left            =   4920
      Picture         =   "frmTree.frx":DA48
      ScaleHeight     =   6240
      ScaleWidth      =   9000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6180
      Index           =   0
      Left            =   3210
      Picture         =   "frmTree.frx":12D53
      ScaleHeight     =   6120
      ScaleWidth      =   9000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.Image imgFruit2 
      Height          =   645
      Left            =   2130
      Picture         =   "frmTree.frx":17AB4
      Top             =   1620
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgFruit1 
      Height          =   645
      Left            =   870
      Picture         =   "frmTree.frx":181E7
      Top             =   1650
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgFlower 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   1590
      Picture         =   "frmTree.frx":1892B
      Stretch         =   -1  'True
      Top             =   1020
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pi As Single = 3.141592

Private Sub Form_Activate()
    ' when you move between Forms , this procedure became active, but I want to run it once, so I use goRender variable
    If goRender Then goRender = False: Call renderMe
End Sub


Private Sub renderMe()
    timerState (False) ' stop last actions of timer (if any).
    refreshMe

    Me.AutoRedraw = Not fastPaint ' disabling of this property make faster drawing, but your tree is not stable and when you move another Form on the Tree Form, the tree may be clear.
    goOut = False ' this variable is for manually stop
    myTree.currentStep = 0    ' this is the first step of function so it start from 0
    nextBranch Me.Width / 2, Me.Height - Me.Height / 5, 180
    
    If Not goOut And myTree.flowerNeed > 0 And myTree.fruitMaxChange > 0 Then timerState (True) ' activate Timer for changing flowers to fruits
    
    frmMain.cmdOK.Visible = True
    frmMain.cmdStop.Visible = False
    frmMain.ProgressBar.Value = 0
End Sub


' this procedure only clean Form and shows a new picture as background and reset some values
Private Sub refreshMe()
    Static newBackground As Byte
    
    myTree.currentProgress = 0  ' reset progress bar value in main form
    myTree.recisionValueHi = 0
    myTree.recisionValueLow = 0
    Call cleanFlowers
    Me.AutoRedraw = True
    Me.Cls
    If randomBackground Then newBackground = Int(Rnd(1) * 5)
    Me.PaintPicture picBack(newBackground).Picture, 0, 0, Me.Width, Me.Height, 0, 0, picBack(newBackground).Width, picBack(newBackground).Height
    DoEvents
End Sub

' this procedure remove flowers of last run, Me.Cls can't remove them because they are objects not paint.
Private Sub cleanFlowers()
    Dim j As Integer
        
    For j = 1 To imgFlower.UBound
        Unload imgFlower(j)
    Next
End Sub


' this is the main function that several times refer to itself
Private Function nextBranch(ByVal startX As Single, ByVal startY As Single, ByVal startAngle As Single) As Boolean
    On Error GoTo mustOut 'this is for overflow error controlling when size of a branche became very high in random states
    Dim j As Byte
    Dim branchSize As Single, branchWidth As Single
    Dim endX As Single, endY As Single, angleGrow As Single, newAngle As Single
    Dim branchCount As Single
    Dim branchColor As Long
    
    If myTree.currentStep >= myTree.totalSteps Or goOut Then Exit Function
    
    If myTree.currentStep = 3 Then ' for showing progressbar, and also "Cut Operation" in update 3, because of affecting by several items (BB,BPS,BPS Scale...), it's not too accurate at this time. any suggestion ?
        myTree.currentProgress = myTree.currentProgress + 100 / myTree.branchPerStep ^ myTree.currentStep
        frmMain.ProgressBar.Value = IIf(myTree.currentProgress < 100, myTree.currentProgress, 100)
        If myTree.cutOperation And myTree.currentProgress >= myTree.cutOperationValue Then myTree.totalSteps = 0: Exit Function
    End If
    
    If myTree.brokenBranches > 0 And myTree.currentStep > 2 Then If Rnd(1) * 100 < myTree.brokenBranches + (myTree.brokenBranches * ((myTree.brokenBranchesScale - 50) / 50) * ((myTree.currentStep * 2 - myTree.totalSteps) / myTree.totalSteps)) * 0.7 * (1 - (myTree.brokenBranchesScale - 50) / 100) Then Exit Function ' this is for making broken branches
    
    If myTree.recisionLevelHi > 0 And myTree.recisionValueHi = 0 And myTree.currentStep = myTree.recisionLevelHi - 1 Then myTree.recisionValueHi = startY ' this line and following are for recision level for higher branches
    If myTree.recisionValueHi > 0 And startY < myTree.recisionValueHi Then Exit Function    ' It cut high branches.
    
    If myTree.recisionLevelLow > 0 And myTree.recisionValueLow = 0 And myTree.currentStep = myTree.recisionLevelLow - 1 Then myTree.recisionValueLow = startY ' this line and following are for recision level for lower branches
    If myTree.recisionValueLow > 0 And startY > myTree.recisionValueLow And myTree.currentStep > myTree.recisionLevelLow * 0.8 Then startAngle = startAngle Mod 180 + 90 'Exit Function   ' I prefer to change branch direction to up, but you can cut it.
    
    DoEvents
    
    myTree.currentStep = myTree.currentStep + 1
    
    
    ' different width for branches from root to leaves.
    branchWidth = (myTree.widthSize / 50) * (15 + myTree.totalSteps / 4) / (myTree.currentStep ^ ((myTree.widthScale ^ 1.2 - 50) / 75 + 0.2))
    If branchWidth < 1 Then branchWidth = 1

    
    ' length of branch.
    branchSize = (myTree.totalSteps - myTree.currentStep * myTree.leafLevel / 50) * IIf(myTree.leafLevel >= 50, myTree.leafLevel / 50, (myTree.leafLevel + 50) / 100)
    branchSize = IIf(branchSize > 0, branchSize, 1) * (Me.Height / myTree.totalSteps ^ 1.9) * IIf(myTree.fixSize, 1, Rnd(1) * 1.7 + 0.1) * myTree.maxSize / 80
    
    ' more control for height of tree's trunk.
    If myTree.currentStep < 3 Then branchSize = branchSize + Me.Height / (30 * myTree.currentStep) + branchSize * (myTree.trunkHeight / 10 - 4) / (myTree.currentStep + 1.5) * myTree.totalSteps / 15
    
    'calculating end points. [ * Pi / 180] is for changing degrees to radians.
    endX = Sin(startAngle * Pi / 180) * branchSize + startX
    endY = Cos(startAngle * Pi / 180) * branchSize + startY
    
    ' Color of branch
    branchColor = RGB(100, 255 * myTree.currentStep / myTree.totalSteps, 35)
    
    ' this paint tree branch.
    If myTree.currentStep >= myTree.startingBranch Then
        Me.DrawWidth = branchWidth
        Me.Line (startX, startY)-(endX, endY), branchColor
    End If
    
    'this is for making flowers and is the main enhancement from last version.
    If myTree.flowerNeed > 0.1 Then makeFlower endX, endY
    
    
    ' this calculate degree for next branch. beginners can try angleGrow=30
    angleGrow = myTree.sizeOfAngel * IIf(myTree.fixAngel, 1, Rnd(1) * 4 - 2)
    
    ' following line is only for widening effect, in ver 2.
    angleGrow = angleGrow + Sin(angleGrow * Pi / 180) * ((myTree.wideLevel - 500) * 0.3) * myTree.totalSteps / myTree.currentStep


    ' in ver 2, I added this item for different number of branches per step from root to leaves. it uses B.P.S Scale Slider
    branchCount = myTree.branchPerStep + myTree.branchPerStep * 2 * ((myTree.branchPerStepScale - 50) / 50) * ((myTree.currentStep) / myTree.totalSteps)
    If branchCount < 1 Then branchCount = 1
    
    ' this is the place that function run itself again.
    For j = 1 To Int(branchCount)
        newAngle = startAngle * (1 - (myTree.windDirection - 50) / 700) - angleGrow / 2 + angleGrow * (j - branchCount / 2) * IIf(myTree.fixAngel, 1, 1.1 - Rnd(1) * 0.25) ' Mod 360  'for Wind effect I added [* (1 - (myTree.windDirection - 50) / 700)] .
        nextBranch endX, endY, newAngle
    Next
    
    'nextBranch = True
mustOut:
    myTree.currentStep = myTree.currentStep - 1
End Function


'**************************************************************************
'Following lines added in version 2 (except: Form_QueryUnload).



' this procedure makes one Flower and put it on Tree
' because of some reasons, I used another way for making Flowers (not painting), maybe a little difficult for beginners. (one reason : I want to change flowers to fruits after a few seconds)
' I make all of flowers from one base flower at run time [imgFlower(0)], this is like when you make TextBoxes or Menus in run time.
Private Sub makeFlower(ByVal endX As Integer, ByVal endY As Integer)
    Dim i As Integer
    
    If myTree.currentStep > myTree.totalSteps / 2 Then ' flowers are only on higher branches
        If Int(Rnd(1) * myTree.currentStep * 30 ^ ((10 + myTree.branchPerStep) / 10) / myTree.flowerNeed) = 1 Then ' how many flowers ?
            i = imgFlower.UBound + 1 ' get highest number of index (last flower) and add one
            If i > 32760 Then myTree.flowerNeed = 0.1 'almost my Maximum Flowers, this change in variable "flowerNeed" is for preventing from running this procedure again. see "nextBranch" function (I can use another new variable for this.)
            Load imgFlower(i)        ' make new flower from base flower
            With imgFlower(i)        ' following lines set properties of new flower
                .Picture = imgFlower(0).Picture ' I used only one type of flower, you can use multi shapes and use a random statement.
                .Width = 25 + Int(Rnd(1) * ((myTree.flowerSize / 5) ^ 1.6) * 3)
                .Height = .Width
                .Left = endX
                .Top = endY
                .Tag = "Flower" ' I used this property to know this object is a flower or a fruit
                .Visible = True
            End With
        End If
    End If
End Sub


'Timer for changing flowers
Private Sub tmrFruits_Timer()
    Dim myNumber As Integer
    Dim j As Byte
    
    If tmrFruits.Interval > 1000 Then tmrFruits.Interval = 10 + (100 - myTree.fruitSpeed) * 5
    If imgFlower.UBound = 0 Or imgFlower.UBound >= 32760 Then timerState (False): Exit Sub ' this is for preventing from errors when be in some random state.
    
    If myTree.fruitCounter >= imgFlower.UBound / 100 * myTree.fruitMaxChange Then
        timerState (False)
    Else
        For j = 0 To myTree.fruitSpeed / 12
            myNumber = Int(Rnd(1) * imgFlower.UBound) + 1 'select a random flower (or fruit) for change
            changeFlower myNumber
            DoEvents
        Next
    End If
End Sub


' this procedure only change a flower to green and then red fruit, but I put an option for a fun part to removing red fruit
Private Sub changeFlower(ByVal myNumber As Integer, Optional ByVal canRemove As Boolean = False)
    On Error GoTo errHandler ' preventing of one run time error : when timer wants to change flowers, you unload flowers (by closing Form)
    
    Select Case imgFlower(myNumber).Tag
        Case "Flower"
            myTree.fruitCounter = myTree.fruitCounter + 1 ' if you put this line in the Case "Fruit" , results became based on Red Fruits. try It !
            imgFlower(myNumber).Tag = "Fruit" ' this means that this is a Green fruit
            imgFlower(myNumber).Picture = imgFruit1.Picture
    
        Case "Fruit"
            imgFlower(myNumber).Tag = "FinalFruit" ' this means that this is a Red final fruit
            imgFlower(myNumber).Picture = imgFruit2.Picture
            ' I prefer that size of Final fruit became a few bigger . if you don't like, you can remove all 2 following lines
            imgFlower(myNumber).Width = imgFlower(myNumber).Width * 1.7
            imgFlower(myNumber).Height = imgFlower(myNumber).Height * 1.7
    
        Case "FinalFruit"
            ' following statement is only for fun part, but I remark it because when it remove a red fruit, it also destroy branches. try it !
'           If canRemove Then imgFlower(myNumber).Visible = False
            
    End Select
   
    ' you can add several cases and make several new differences  in your flowers and fruits (such as new size and color) by controlling "Tag" property

errHandler:    'just go out
End Sub


' this is for reseting, activating and deactivating Fruit Timer
Private Sub timerState(ByVal myAction As Boolean)
    If myAction Then ' when True
        tmrFruits.Interval = 3000
        tmrFruits.Enabled = True
        frmMain.cmdTimerStop.Visible = True
        If playSound Then sndPlaySound App.Path & "\fruit.wav", 1
    Else
        tmrFruits.Enabled = False
        frmMain.cmdTimerStop.Visible = False
        myTree.fruitCounter = 0
    End If
End Sub


'this is only for fun. when you click on a Flower it chang to a green fruit, then red fruit and then remove ( see changeFlower procedure)
Private Sub imgFlower_Click(Index As Integer)
        changeFlower Index, True
End Sub

'this procedure became active when user want to close Form, maybe timer be active or tree is not complete, I must manually deactivate these actions. ( some changes in frmMain)
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    timerState (False)
    goOut = True
End Sub

' for saving picture by pressing "F2"
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then savePicture
End Sub


'*****************************************************************************
' Saving picture on hard disk in Jpeg format (added in update 2)
' this part uses cJpeg.cls
Public Sub savePicture()
    On Error GoTo errHandler
    Dim myPath As String
    Dim myJpeg As New cJpeg
    
    Me.Caption = "Saving..."
    myPath = App.Path & "\SavedPics\"
    If Dir(myPath, vbDirectory) = "" Then MkDir myPath
    
    myJpeg.Quality = 85
    myJpeg.Comment = "Real Tree 2"
    myJpeg.SampleHDC frmTree.hDC, Me.ScaleWidth / Screen.TwipsPerPixelX, Me.ScaleHeight / Screen.TwipsPerPixelY
    myJpeg.SaveFile myPath & "Pic_" & Format(Now, "yyyy-mm-dd___hh-mm-ss") & ".jpg"
    Set myJpeg = Nothing
    
    Me.Caption = "Tree"
Exit Sub
errHandler:
    MsgBox Err.Description, vbExclamation, "Error Saving Picture"
End Sub



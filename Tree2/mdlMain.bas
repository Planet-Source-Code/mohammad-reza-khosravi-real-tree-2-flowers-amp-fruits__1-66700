Attribute VB_Name = "mdlMain"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Type treeInfo
    totalSteps As Byte
    currentStep As Byte
    branchPerStep As Single
    startingBranch As Byte
    ' fixSize As Byte           'I split this variable to 2 follwing variables in  ver 2
    fixSize As Boolean
    maxSize As Single
    recisionLevelHi As Byte
    recisionValueHi As Single
    recisionLevelLow As Byte
    recisionValueLow As Single
    ' fixAngel As Integer       'I split this variable to 2 follwing variables in  ver 2
    fixAngel As Boolean
    sizeOfAngel As Single
    brokenBranches As Single
    brokenBranchesScale As Single
    branchPerStepScale As Single
    leafLevel As Single
    trunkHeight As Single
    widthSize As Single
    widthScale As Single
    windDirection As Single
    wideLevel As Single
    flowerNeed As Single
    flowerSize As Single
    fruitSpeed As Single
    fruitMaxChange As Byte
    fruitCounter As Integer
    currentProgress As Single
    cutOperation As Boolean
    cutOperationValue As Integer
End Type

Public myTree As treeInfo
Public onTop As Boolean
Public fastPaint As Boolean
Public playSound As Boolean
Public randomBackground As Boolean
Public goRender As Boolean
Public goOut As Boolean


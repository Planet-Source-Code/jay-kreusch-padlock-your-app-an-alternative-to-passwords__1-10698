VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create PadLock Key"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'Very simple Proc to create new keys while in
    'Program
    On Error GoTo errtrap:
    
    Dim jLoop As Byte
    Dim jTemp As String
    
    'This loop will allow us to prompt after each creation
    Do Until jLoop = vbNo
        'Make sure a disk is in the drive
        jTemp = Dir("a:\")
        'Create the PadLock Key
        CreateKey
        'Prompt for another disk to be created
        jLoop = MsgBox("Key Created. Do you wish to create another PadLock Key?", vbYesNo)
    Loop
        
Exit Sub
errtrap:
Select Case Err.Number
    Case 52
        Exit Sub
    Case Else
        Stop
End Select
End Sub

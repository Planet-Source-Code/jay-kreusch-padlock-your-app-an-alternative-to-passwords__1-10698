Attribute VB_Name = "basMain"
Option Explicit

'This is a hardcoded password key for the encryption
Private Const MY_PASSWORD As String = "hfdsjhljkas"

'This is used to obtain the serial number of the diskette
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'This function obtains the disk's serial number
Private Function GetSerial() As Long
    On Error GoTo errtrap
    
    Dim jResponse As Long
    Dim jSerial As Long
            
    'call the API to get the serial number
    jResponse = GetVolumeInformation("a:\", vbNullString, 0, jSerial, 0, 0, vbNullString, 0)

    GetSerial = jSerial
    Exit Function

errtrap:
    Select Case Err.Number
        Case 52 'Cancelled prompt for disk
            GetSerial = 0
            Exit Function
        Case Else
            Stop
    End Select
End Function

'This function creates the key file on the disk
'This file is basically the serial number of the diskette
'itself, written to the disk in encrypted form
'Copying the disk using DiskCopy or copying the file to
'another disk will not create a valid PadLock Key.
'Creating an image might work, but I haven't tested
'this.
Public Sub CreateKey()
    Dim sEncrypted As String
    Dim jTemp As Long
    
    'Convert the serial to a string so we can encrypt
    jTemp = CStr(GetSerial)
    
    'Run the encryption stuff courtesy of Barry Dunne
    EncryptionCSPConnect
    sEncrypted = EncryptData(jTemp, MY_PASSWORD)
    EncryptionCSPDisconnect
    
    'Write the return value to the diskette itself
    Open "a:\Key.dat" For Random As #1
        Put #1, , sEncrypted
    Close #1
End Sub

Public Function ReadKey() As Boolean
    Dim sEncrypted As String
    Dim jTemp As Long
    
    'grab the encrypted data from the floppy
    Open "a:\Key.dat" For Random As #1
        Get #1, , sEncrypted
    Close #1
    
    'If it is a null string, it didn't exist
    If sEncrypted = vbNullString Then
        ReadKey = False
        Exit Function
    End If
    
    'decrypt the data
    EncryptionCSPConnect
    jTemp = CLng(DecryptData(sEncrypted, MY_PASSWORD))
    EncryptionCSPDisconnect

    'make sure it is equal to the disks serial number
    If jTemp = GetSerial Then
        ReadKey = True
    Else
        ReadKey = False
    End If
End Function


Sub Main()
    On Error GoTo errtrap
    
    Dim jTemp As String
    
    If GetSetting(App.Title, "Settings", "FirstUse", 0) = 0 Then
        'first use of program
        jTemp = InputBox("This is the first use of this program. Please enter the initial password")
        'This is the first-time use password. Use it or change it.
        If UCase(jTemp) <> UCase("Crack") Then
            MsgBox "That is not the right password!"
            End
        Else
            'If the password is correct, create a new PadLock Key
            MsgBox "The first PadLock Key will now be created! Please insert a disk int drive A"
            'This line of code will cause a prompt if there is no disk in the drive
            'If the user clicks the cancel button on the prompt, error 52 will be returned
            jTemp = Dir("a:\")
            'Run the createKey Proc
            CreateKey
            'Indicate in the registry that the program has been run
            SaveSetting App.Title, "Settings", "FirstUse", 1
            'Show the main form
            frmMain.Show
        End If
    Else
        'Make sure disk is in drive
        jTemp = Dir("a:\")
        'See if the key is valid
        If ReadKey = True Then
            'If it is, then show the program
            frmMain.Show
        Else
            'Otherwise, tell them they used a wrong key and kick them out
            MsgBox "Your Key Diskette is invalid, this program will now end"
            End
        End If
    End If

Exit Sub
errtrap:
Select Case Err.Number
    Case 52 'disk cancelled
        End
    Case Else
        Stop
End Select
End Sub


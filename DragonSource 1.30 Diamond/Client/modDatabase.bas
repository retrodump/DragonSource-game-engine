Attribute VB_Name = "modDatabase"

' Created with DragonSource Diamond. Copyright � 2012 DragonSource


Option Explicit
Public Const MAX_PATH As Long = 260
Private Const ERROR_NO_MORE_FILES = 18
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public SOffsetX As Integer
Public SOffsetY As Integer

Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Function ExistVar(File As String, Header As String, Var As String) As Boolean

    ExistVar = (GetVar(File, Header, Var) <> "")

End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    If Trim$(Value) = "0" Or Trim$(Value) = vbNullString Then
        If ExistVar(File, Header, Var) Then
            Call DelVar(File, Header, Var)
        End If
    Else
        Call WritePrivateProfileString(Header, Var, Value, File)
    End If
End Sub

Public Sub DelVar(sFileName As String, sSection As String, sKey As String)

   If Len(Trim$(sKey)) <> 0 Then
      WritePrivateProfileString sSection, sKey, _
         vbNullString, sFileName
   Else
      WritePrivateProfileString _
         sSection, sKey, vbNullString, sFileName
   End If
End Sub

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Function FileExist(ByVal FileName As String) As Boolean
    If Dir(App.Path & "\" & FileName) = vbNullString Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Function ScreenFileExist(ByVal FileName As String) As Boolean
    If Dir(App.Path & "\Screenshots\" & FileName) = "" Then
        ScreenFileExist = False
    Else
        ScreenFileExist = True
    End If
End Function

Sub AddLog(ByVal Text As String)
Dim FileName As String
Dim f As Long

    'If Trim$(Command) = "-debug" Then
    If GoDebug = YES Then
        If frmDebug.Visible = False Then
            frmDebug.Visible = True
        End If
        
        FileName = App.Path & "\debug.txt"
    
        If Not FileExist("debug.txt") Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub SaveLocalMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
                            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    If FileExist("maps\map" & MapNum & ".dat") = False Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Map(MapNum)
    Close #f

End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
End Function

Function ListMusic(ByVal sStartDir As String)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
    
    On Error Resume Next
    
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    frmMapProperties.lstMusic.Clear
    
    frmMapProperties.lstMusic.AddItem "None", 0
    
    sStartDir = sStartDir & "*.*"
    
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                    If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                        sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                        If Right$(sTemp, 4) = ".mid" Then
                            frmMapProperties.lstMusic.AddItem sTemp
                        End If
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

Function ListSounds(ByVal sStartDir As String, ByVal Form As Long)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
    
    On Error Resume Next
    
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    If Form = 1 Then
        frmSound.lstSound.Clear
    ElseIf Form = 2 Then
        frmNotice.lstSound.Clear
    ElseIf Form = 3 Then
        frmEmoticonEditor.lstSound.Clear
    End If
    
    sStartDir = sStartDir & "*.*"
    
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                    sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                    If Right$(sTemp, 4) = ".wav" Then
                        If Form = 1 Then
                            frmSound.lstSound.AddItem sTemp
                        ElseIf Form = 2 Then
                            frmNotice.lstSound.AddItem sTemp
                        ElseIf Form = 3 Then
                            frmEmoticonEditor.lstSound.AddItem sTemp
                        End If
                    End If
                End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PB.Left = PB.Left + x - SOffsetX
        PB.Top = PB.Top + y - SOffsetY
    End If
End Sub

Function AllGUIFilesExist() As Boolean
Dim FileName As String
Dim Error As Byte

    AllGUIFilesExist = False
    
    FileName = "GUI\"
    Error = NO
    
    Do Until Error = YES
        If Not FileExist(FileName & "content.gif") Then Error = YES
        If Not FileExist(FileName & "contentlist.gif") Then Error = YES
        If Not FileExist(FileName & "contentsmall.gif") Then Error = YES
        If Not FileExist(FileName & "contentstatus.gif") Then Error = YES
        If Not FileExist(FileName & "game.gif") Then Error = YES
        If Not FileExist(FileName & "loading.gif") Then Error = YES
        If Not FileExist(FileName & "mainmenu.gif") Then Error = YES
        If Not FileExist(FileName & "medium.gif") Then Error = YES
        If Not FileExist(FileName & "mediumlist.gif") Then Error = YES
        If Not FileExist(FileName & "shop.gif") Then Error = YES
        Exit Do
    Loop
    
    If Not Error Then
        AllGUIFilesExist = True
    End If

End Function

Sub LoadAllGUIPictures()

    frmChars.Picture = LoadPicture(App.Path & "\GUI\mediumlist.gif")
    'Unload frmChars
    frmCredits.Picture = LoadPicture(App.Path & "\GUI\mediumlist.gif")
    'Unload frmCredits
    frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\medium.gif")
    'Unload frmDeleteAccount
    frmFixItem.Picture = LoadPicture(App.Path & "\GUI\content.gif")
    'Unload frmFixItem
    frmGuild.Picture = LoadPicture(App.Path & "\GUI\content.gif")
    'Unload frmGuild
    frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\medium.gif")
    'Unload frmIpconfig
    frmLogin.Picture = LoadPicture(App.Path & "\GUI\medium.gif")
    'Unload frmLogin
    frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\mainmenu.gif")
    'Unload frmMainMenu
    frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\medium.gif")
    'Unload frmNewAccount
    frmNewChar.Picture = LoadPicture(App.Path & "\GUI\medium.gif")
    'Unload frmNewChar
    frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\content.gif")
    'Unload frmPlayerChat
    frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\content.gif")
    'Unload frmPlayerTrade
    frmTalk.Picture = LoadPicture(App.Path & "\GUI\content.gif")
    Unload frmTalk
    frmTrade.Picture = LoadPicture(App.Path & "\GUI\shop.gif")
    'Unload frmTrade

End Sub

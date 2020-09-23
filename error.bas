Attribute VB_Name = "error"
'error.bas
'By: érrør
'-I strongly encourage everyone who downloads this bas to read the
' READ ME! file.
'-error.bas, beta 2, was released December 9, 2000 at 11:45 P.M.
'-error.bas was made to be used with Visual Basic 5.0 and Visual Basic 6.0,
' and possibly any higher versions if yet out.
'-error.bas was designed to work with AOL 6.0 and AIM 4.0 and higher, some
' subs and functions may work with lower versions.
'-As with all beta versions of anything, there are going to be bugs. There are
' subs and functions in this bas that do not work, if you can fix them, please
' let me know how you did it by emailing me.
'-Look for updates of error.bas on www.planetsourcecode.com
'-Please email me your ideas for subs, functions, or anything at all. If you do
' choose to email me, I ask that you please title the subject of your mail
' appropriately to ensure that your mail gets read. I will not read mail with
' no subjects or unappropriate subject titles.
'-Email: [errorandskoal@lycos.com]
Public Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal DWreserved As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long  'formerly integer
  '  Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type
Private Const PING_TIMEOUT = 200
Public Const LB_GETCOUNT = &H18B
Public Const LB_SETCURSEL = &H186
Public Const WM_LBUTTONDBLCLK = &H203
Public Const SW_SHOWNORMAL = 1
Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3
Public Const WM_CLOSE = &H10
Public Const WM_CHAR = &H102
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_SETTEXT = &HC
Public Const BM_SETCHECK = &HF1
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONUP = &H202
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const ENTER_KEY = 13


Global Const GW_CHILD = 5

Public Type POINTAPI
      X As Long
      Y As Long
End Type
Public Sub Keyword(Keyword As String)
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMEssageByString(EditWin&, WM_SETTEXT, 0&, Keyword$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub RunToolbar(iconnumber&, letter$)
Dim aolframe As Long, menu As Long, aoltoolbar1 As Long
Dim aoltoolbar2 As Long, AOLIcon As Long, Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2 = FindWindowEx(aoltoolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(aoltoolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To iconnumber
AOLIcon = FindWindowEx(aoltoolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter$ = Asc(letter)
Call PostMessage(menu, WM_CHAR, letter$, 0&)
End Sub
Public Sub Chat_EnterPrivateChatRoom(RoomName As String)
Keyword "aol://2719:2-2-" & RoomName$
End Sub
Public Sub Chat_EnterMemberChatRoom(RoomName As String)
Keyword "aol://2719:61-2-" & RoomName$
End Sub
Public Sub ClickMailMenu_ReadMail_NewMail()
aolframe& = FindWindow("AOL Frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(aoltoolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_Chat_ChangeCaption(NewCaption As String)
Dim aimchatwnd As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Call SendMEssageByString(aimchatwnd, WM_SETTEXT, 0&, NewCaption$)
End Sub
Sub Aim_Chat_Ignore_Person(Who As String)
Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, Buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer, BuddyTree As Long

    ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatRoom& <> 0 Then
        Do
            BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMEssageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            Buffer$ = String$(NameLen, 0)
            Moo2 = SendMEssageByString(BuddyTree&, LB_GETTEXT, MooLoo, Buffer$)
            TabPos = InStr(Buffer$, Chr$(9))
            NameText$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            If name$ = Who$ Then GoTo Igorn
        Next MooLoo
    End If
Igorn:
    Dim ChatWindz As Long, IM As Long, IgnoreBut As Long, Klick As Long

    ChatWindz& = FindWindow("AIM_ChatWnd", vbNullString)
    IM& = FindWindowEx(ChatWindz&, 0, "_Oscar_IconBtn", vbNullString)
    IgnoreBut& = FindWindowEx(ChatWindz&, IM&, "_Oscar_IconBtn", vbNullString)
    Klick& = SendMessage(IgnoreBut&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(IgnoreBut&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub ClickMailMenu_Write()
Call RunToolbar("0", "W")
End Sub
Public Sub ClickMailMenu_AddressBook()
Call RunToolbar("0", "A")
End Sub
Public Sub ClickMailMenu_MailCenter()
Call RunToolbar("0", "M")
End Sub
Public Sub ClickMailMenu_RecentlyDeletedMail()
Call RunToolbar("0", "D")
End Sub
Public Sub ClickMailMenu_FilingCabinet()
Call RunToolbar("0", "F")
End Sub
Public Sub ClickMailMenu_MailWaitingToBeSent()
Call RunToolbar("0", "b")
End Sub
Public Sub ClickMailMenu_AutomaticAol()
Call RunToolbar("0", "u")
End Sub
Public Sub ClickMailMenu_MailSignatures()
Call RunToolbar("0", "S")
End Sub

Public Function Mail_ChangeWriteMailCaption(NewCaption As String)
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
richcntl& = FindWindowEx(aolchild&, 0&, "RICHCNTL", vbNullString)
Call SendMEssageByString(aolchild&, WM_SETTEXT, 0&, NewCaption$)
End Function
Public Sub ClickMailMenu_MailControls()
Call RunToolbar("0", "C")
End Sub
Public Function SignOnScreen_EnterPassword(Text As String)
Dim aoledit As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Sign On")
aoledit& = FindWindowEx(aolchild&, 0&, "_AOL_Edit", vbNullString)
Call SendMEssageByString(aoledit&, WM_SETTEXT, 0&, Text$)
End Function

Public Function Chat_ShowChat()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
Call showwindow(aolchild&, SW_HIDE)
End Function

Public Sub ClickMailMenu_MailPreferences()
Call RunToolbar("0", "P")
End Sub
Public Sub ClickMailMenu_GreetingsAndMailExtras()
Call RunToolbar("0", "G")
End Sub
Public Sub ClickMailMenu_Newsletters()
Call RunToolbar("0", "N")
End Sub
Public Sub ClickPeopleMenu_SendInstantMessage()
Call RunToolbar("1", "I")
End Sub
Public Sub ClickPeopleMenu_ChatPeopleConnection()
Call RunToolbar("1", "C")
End Sub
Public Sub ClickPeopleMenu_ChatNow()
Call RunToolbar("1", "N")
End Sub
Public Sub ClickPeopleMenu_FindAChat()
Call RunToolbar("1", "F")
End Sub
Public Sub ClickPeopleMenu_StartYourOwnChat()
Call RunToolbar("1", "S")
End Sub
Public Sub ClickPeopleMenu_LiveEvents()
Call RunToolbar("1", "e")
End Sub
Public Sub ClickPeopleMenu_Buddylist()
Call RunToolbar("1", "B")
End Sub
Public Sub ClickPeopleMenu_GetDirectoryListing()
Call RunToolbar("1", "G")
End Sub
Public Sub ClickPeopleMenu_LocateMemberOnline()
Call RunToolbar("1", "L")
End Sub
Public Sub ClickPeopleMenu_SendMessageToPager()
Call RunToolbar("1", "M")
End Sub
Public Sub ClickPeopleMenu_SignOnAFriend()
Call RunToolbar("1", "o")
End Sub
Public Sub ClickPeopleMenu_AolHometown()
Call RunToolbar("1", "H")
End Sub
Public Sub ClickPeopleMenu_GroupsAtAol()
Call RunToolbar("1", "A")
End Sub
Public Function BuddyList_ClickChat()
'clicks the chat button on your buddylist
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function Mail_Write()
aolframe& = FindWindow("AOL Frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(aoltoolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function BuddyList_ClickLocate()
'clicks locate button on blist
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function


Public Function Mail_Send()
'clicks the Send Button on Write Mail
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 17&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function Mail_ClickSendLater()
'clicks the send later button on write mail
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 18&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function


Public Function Mail_ClickAddressBook()
'clicks the addy book button on write mail
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 19&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function Mail_ClickGreetings()
'clicks greetings icon on write mail
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 20&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function Mail_ClickSignOnFriend()
'clicks the Sign On Friend Button on write mail
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 10&
    aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Sub ClickPeopleMenu_Invitations()
Call RunToolbar("1", "v")
End Sub
Public Sub ClickPeopleMenu_PeopleDirectory()
Call RunToolbar("1", "D")
End Sub
Public Sub ClickPeopleMenu_Personals()
Call RunToolbar("1", "P")
End Sub
Public Sub ClickPeopleMenu_WhitePages()
Call RunToolbar("1", "W")
End Sub
Public Sub ClickPeopleMenu_YellowPages()
Call RunToolbar("1", "Y")
End Sub
Public Sub ClickAolServicesMenu_ShopAtAol()
Call RunToolbar("2", "S")
End Sub
Public Sub ClickAolServicesMenu_Internet_InternetConnection()
Call RunToolbar2("2", "I", "I")
End Sub
Public Sub ClickAolServicesMenu_Internet_GoToTheWeb()
Call RunToolbar2("2", "I", "G")
End Sub
Public Sub ClickAolServicesMenu_Internet_SearchTheWeb()
Call RunToolbar2("2", "I", "S")
End Sub
Public Sub ClickAolServicesMenu_Internet_NewsGroups()
Call RunToolbar2("2", "I", "N")
End Sub
Public Sub ClickAolServicesMenu_Internet_FTP()
Call RunToolbar2("2", "I", "F")
End Sub
Public Sub ClickAolServicesMenu_AddToMyCalendar()
Call RunToolbar("2", "A")
End Sub
Public Sub ClickAolServicesMenu_AolHelp()
Call RunToolbar("2", "H")
End Sub
Public Sub ClickAolServicesMenu_Calendar()
Call RunToolbar("2", "C")
End Sub
Public Sub ClickAolServicesMenu_CarBuying()
Call RunToolbar("2", "B")
End Sub
Public Sub ClickAolServicesMenu_DownloadCenter()
Call RunToolbar("2", "D")
End Sub
Public Sub ClickAolServicesMenu_GovernmentGuide()
Call RunToolbar("2", "u")
End Sub
Public Sub ClickAolServicesMenu_HomeworkHelp()
Call RunToolbar("2", "k")
End Sub
Public Sub ClickAolServicesMenu_MapsAndDirections()
Call RunToolbar("2", "M")
End Sub
Public Sub ClickAolServicesMenu_MedicalReferences()
Call RunToolbar("2", "R")
End Sub
Public Sub ClickAolServicesMenu_MemberRewards()
Call RunToolbar("2", "e")
End Sub
Public Sub ClickAolServicesMenu_MovieShowTimes()
Call RunToolbar("2", "w")
End Sub
Public Sub ClickAolServicesMenu_OnlineGreetings()
Call RunToolbar("2", "G")
End Sub
Public Sub ClickAolServicesMenu_Personals()
Call RunToolbar("2", "P")
End Sub
Public Sub ClickAolServicesMenu_RecipeFinder()
Call RunToolbar("2", "F")
End Sub
Public Sub ClickAolServicesMenu_SportsScores()
Call RunToolbar("2", "o")
End Sub
Public Sub ClickAolServicesMenu_StockPortfolios()
Call RunToolbar("2", "l")
End Sub
Public Sub ClickAolServicesMenu_StockQuotes()
Call RunToolbar("2", "Q")
End Sub
Public Sub ClickAolServicesMenu_TravelReservations()
Call RunToolbar("2", "v")
End Sub
Public Sub ClickAolServicesMenu_TVListings()
Call RunToolbar("2", "T")
End Sub
Public Sub ClickSettingsMenu_AolAnywhere()
Call RunToolbar("3", "A")
End Sub
Public Sub ClickSettingsMenu_Preferences()
Call RunToolbar("3", "P")
End Sub
Public Sub ClickSettingsMenu_ParentalControls()
Call RunToolbar("3", "C")
End Sub
Public Sub ClickSettingsMenu_MyDirectoryListing()
Call RunToolbar("3", "M")
End Sub
Public Sub ClickSettingsMenu_ScreenNames()
Call RunToolbar("3", "S")
End Sub
Public Sub ClickSettingsMenu_Passwords()
Call RunToolbar("3", "a")
End Sub
Public Sub ClickSettingsMenu_BillingCenter()
Call RunToolbar("3", "B")
End Sub
Public Sub ClickSettingsMenu_AolQuickCheckout()
Call RunToolbar("3", "Q")
End Sub
Public Sub ClickFavoritesMenu_FavoritePlaces()
Call RunToolbar("4", "F")
End Sub
Public Sub ClickFavoritesMenu_AddTopWindowToFavorites()
Call RunToolbar("4", "A")
End Sub
Public Sub ClickFavoritesMenu_GoToKeyword()
Call RunToolbar("4", "G")
End Sub
Public Function Get_Caption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Public Function Get_User() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    GetUser$ = ""
End Function
Public Sub Chat_Send(WhatToSay As String)
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMEssageByString(AORich2, WM_SETTEXT, 0&, WhatToSay$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub Pause(duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= duration
        DoEvents
    Loop
End Sub
Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Public Function Find_Room() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, aolstatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    aolstatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And aolstatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            aolstatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And aolstatic& <> 0& Then
                Find_Room& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    Find_Room& = child&
End Function
Public Function ClickSendIM()
'when your having a conversation with someone, this
'clicks that send button
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Sub InstantMessage_Send(Person As String, Message As String)
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMEssageByString(Rich&, WM_SETTEXT, 0&, Message$)
    ClickSendIM
    Do
    DoEvents
    OK& = FindWindow("#32770", "America Online")
    IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or IM& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Function Minimize_Aol()
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call showwindow(aolframe&, SW_MINIMIZE)
End Function
Public Sub Form_Center(TheForm As Form)
With TheForm
.Left = (Screen.Width - .Width) / 2
.Top = (Screen.Height - .Height) / 2
End With
End Sub

Public Function Maximize_Aol()
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call showwindow(aolframe&, SW_MAXIMIZE)
End Function
Public Function Change_AolCaption(NewCaption As String)
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call SendMEssageByString(aolframe&, WM_SETTEXT, 0&, NewCaption$)
End Function
Public Sub GoToWeb(Address As String)
Keyword Address$
End Sub
Sub Change_WelcomeCaption(NewCaption As String)
Dim aolframe As Long, mdiclient As Long, aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Call SendMEssageByString(aolchild, WM_SETTEXT, 0&, NewCaption)
End Sub

Sub Show_Welcome()
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Call showwindow(aolchild, SW_SHOW)
End Sub
Sub Hide_Welcome()
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Call showwindow(aolchild, SW_HIDE)
End Sub
Public Sub Windows_Reboot()
Call ExitWindowsEx(EWX_REBOOT, 0)
End Sub
Sub AIM_ChangeCaption(NewCaption As String)
Dim oscarbuddylistwin As Long
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
Call SendMEssageByString(oscarbuddylistwin, WM_SETTEXT, 0&, NewCaption$)
End Sub
Sub Windows_Show_ProgramTray()
Dim shelltraywnd As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
Call showwindow(shelltraywnd, SW_SHOW)
End Sub
Sub Windwos_Freeze()
    Call ExitWindowsEx(EWX_FORCE, 0)
End Sub
Public Sub Aim_InstantMessage_MassIm(List As ListBox, Message As String)
Dim Scrll As Integer, Num As Integer, Str As String
Num% = 0
For Scrll% = 0 To List.ListCount - 1
    Str$ = List.List(Scrll%)
    If LCase(Get_User) = LCase(Str$) Then
    Else
        If Num% >= 5 Then
            Pause (20)
            Num% = 0
        End If
        IM_Send Str$, Message$
        Pause (0.2)
    End If
    Num% = Num% + 1
    DoEvents
Next
End Sub
Public Sub Form_MakeTransparent(TransForm As Form)
Dim ErrorTest As Double
    On Error Resume Next
    Dim Regn As Long
    Dim TmpRegn As Long
    Dim TmpControl As Control
    Dim LinePoints(4) As POINTAPI
    TransForm.ScaleMode = 3
    If TransForm.BorderStyle <> 0 Then MsgBox "Change the borderstyle to 0!", vbCritical, "ACK!": End
    Regn = CreateRectRgn(0, 0, 0, 0)
 For Each TmpControl In TransForm
        If TypeOf TmpControl Is Line Then
            If Abs((TmpControl.Y1 - TmpControl.Y2) / (TmpControl.X1 - TmpControl.X2)) > 1 Then
                LinePoints(0).X = TmpControl.X1 - 1
                LinePoints(0).Y = TmpControl.Y1
                LinePoints(1).X = TmpControl.X2 - 1
                LinePoints(1).Y = TmpControl.Y2
                LinePoints(2).X = TmpControl.X2 + 1
                LinePoints(2).Y = TmpControl.Y2
                LinePoints(3).X = TmpControl.X1 + 1
                LinePoints(3).Y = TmpControl.Y1
            Else

                LinePoints(0).X = TmpControl.X1
                LinePoints(0).Y = TmpControl.Y1 - 1
                LinePoints(1).X = TmpControl.X2
                LinePoints(1).Y = TmpControl.Y2 - 1
                LinePoints(2).X = TmpControl.X2
                LinePoints(2).Y = TmpControl.Y2 + 1
                LinePoints(3).X = TmpControl.X1
                LinePoints(3).Y = TmpControl.Y1 + 1
            End If
    TmpRegn = CreatePolygonRgn(LinePoints(0), 4, 1)
        ElseIf TypeOf TmpControl Is Shape Then
            If TmpControl.Shape = 0 Then
                TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
            ElseIf TmpControl.Shape = 1 Then
                If TmpControl.Width < TmpControl.Height Then
                    TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width)
                Else
                    TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height)
                End If
            ElseIf TmpControl.Shape = 2 Then
                TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
            ElseIf TmpControl.Shape = 3 Then
                If TmpControl.Width < TmpControl.Height Then
                    TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 0.5)
                Else
                    TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
                End If
            ElseIf TmpControl.Shape = 4 Then
    
                If TmpControl.Width > TmpControl.Height Then
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
                Else
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Width / 4, TmpControl.Width / 4)
                End If
            ElseIf TmpControl.Shape = 5 Then
   
                If TmpControl.Width > TmpControl.Height Then
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
                Else
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 1, TmpControl.Width / 4, TmpControl.Width / 4)
                End If
            End If
         If TmpControl.BackStyle = 0 Then
      CombineRgn Regn, Regn, TmpRegn, RGN_XOR
                If TmpControl.Shape = 0 Then
             TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + TmpControl.Height - 1)
                ElseIf TmpControl.Shape = 1 Then
                    If TmpControl.Width < TmpControl.Height Then
                        TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 1)
                    Else
                        TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 1, TmpControl.Top + TmpControl.Height - 1)
                    End If
                ElseIf TmpControl.Shape = 2 Then
                    TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
                ElseIf TmpControl.Shape = 3 Then
                    If TmpControl.Width < TmpControl.Height Then
                        TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 0.5)
                    Else
                        TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
                    End If
                ElseIf TmpControl.Shape = 4 Then
                    If TmpControl.Width > TmpControl.Height Then
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
                    Else
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Width / 4, TmpControl.Width / 4)
                    End If
                ElseIf TmpControl.Shape = 5 Then
                    If TmpControl.Width > TmpControl.Height Then
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
                    Else
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width, TmpControl.Width / 4, TmpControl.Width / 4)
                    End If
                End If
            End If
        Else
      TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
        End If
    ErrorTest = 0
            ErrorTest = TmpControl.Width
            If ErrorTest <> 0 Or TypeOf TmpControl Is Line Then
                CombineRgn Regn, Regn, TmpRegn, RGN_XOR
            End If
    Next TmpControl
    SetWindowRgn TransForm.hWnd, Regn, True
End Sub
Sub Aim_InstantMessage_Send2(Person As String, SayWhat As String)
Dim aimimessage As Long
Dim oscarpersistantcombo As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim editx As Long
Dim TextSet As Long
Dim TextEditBox As Long
Dim SendText As Long
Dim oscariconbtn As Long
Dim send As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
TextSet& = FindWindowEx(aimimessage, 0&, "_oscar_persistantcombo", vbNullString)
editx = FindWindowEx(oscarpersistantcombo, 0&, "edit", vbNullString)
TextEditBox& = SendMEssageByString(TextSet&, WM_SETTEXT, 0, Person$)
Pause 0.1
aimimessage = FindWindow("aim_imessage", vbNullString)
wndateclass = FindWindowEx(aimimessage, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimimessage, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
SendText& = SendMEssageByString(ateclass, WM_SETTEXT, 0, SayWhat$)
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
send& = SendMessage(oscariconbtn, WM_LBUTTONDOWN, 0, 0&)
send& = SendMessage(oscariconbtn, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Aim_InstantMessage_HideSend()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_InstantMessage_ChangeInstantMessageCaption(NewCaption As String)
Dim aimimessage As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
Call SendMEssageByString(aimimessage, WM_SETTEXT, 0&, NewCaption$)
End Sub
Sub Aim_InstantMessage_Close()
Dim aimimessage As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
Call SendMessageLong(aimimessage, WM_CLOSE, 0&, 0&)
End Sub
Sub Aim_InstantMessage_ClickAddBuddy()
Dim aimimessage As Long
Dim oscariconbtn As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_InstantMessage_ClickBlock()
Dim aimimessage As Long
Dim oscariconbtn As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_InstantMessage_ClickInfo()
Dim aimimessage As Long, oscariconbtn As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub Aim_InstantMessage_HideAddBuddy()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_InstantMessage_HideBlock()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_InstantMessage_HideInfo()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_InstantMessage_HideTalk()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_InstantMessage_HideWarn()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_InstantMessage_Minimize()
Dim aimimessage As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
Call showwindow(aimimessage, SW_MINIMIZE)
End Sub
Sub Aim_InstantMessage_Maximize()
'Example
'Call IM_Max
Dim aimimessage As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
Call showwindow(aimimessage, SW_MAXIMIZE)
End Sub
Sub Aim_InstantMessage_OpenAnInstantMessage()
Dim oscarbuddylistwin As Long
Dim oscartabgroup As Long
Dim btn As Long
Dim send As Long
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
Call showwindow(oscarbuddylistwin, SW_SHOW)
Pause 0.1
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup = FindWindowEx(oscarbuddylistwin, 0&, "_oscar_tabgroup", vbNullString)
btn& = FindWindowEx(oscartabgroup, 0&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(btn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(btn&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub Aim_InstantMessage_Send(Person As String, SayWhat As String)
Dim aimimessage As Long
Dim oscarpersistantcombo As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim editx As Long
Dim TextSet As Long
Dim TextEditBox As Long
Dim SendText As Long
Dim oscariconbtn As Long
Dim send As Long
Call IM_Open
Pause 0.1
aimimessage = FindWindow("aim_imessage", vbNullString)
TextSet& = FindWindowEx(aimimessage, 0&, "_oscar_persistantcombo", vbNullString)
editx = FindWindowEx(oscarpersistantcombo, 0&, "edit", vbNullString)
TextEditBox& = SendMEssageByString(TextSet&, WM_SETTEXT, 0, Person$)
Pause 0.1
aimimessage = FindWindow("aim_imessage", vbNullString)
wndateclass = FindWindowEx(aimimessage, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimimessage, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
SendText& = SendMEssageByString(ateclass, WM_SETTEXT, 0, SayWhat$)
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
send& = SendMessage(oscariconbtn, WM_LBUTTONDOWN, 0, 0&)
send& = SendMessage(oscariconbtn, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Aim_InstantMessage_ShowAddBuddy()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_InstantMessage_ShowBlock()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_InstantMessage_ShowInfo()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_InstantMessage_ShowSend()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_InstantMessage_ShowTalk()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_InstantMessage_ShowWarn()
Dim aimimessage As Long
Dim oscariconbtn As Long
Dim X As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
X& = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_InstantMessage_ClickTalk()
Dim aimimessage As Long
Dim oscariconbtn As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_InstantMessage_ClickWarn()
Dim aimimessage As Long
Dim oscariconbtn As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Function Ping(ByVal hostnameOrIpaddress As String, Optional timeOutmSec As Long = PING_TIMEOUT) As Boolean
Dim echoValues As ICMP_ECHO_REPLY
Dim pos As Integer
Dim Count As Integer
Dim returnIp As Collection
    On Error GoTo e_Trap
    If Trim(hostnameOrIpaddress) = "" Then
        Ping = False
        Exit Function
    End If
    If SocketsInitialize() Then
        If InStr(1, hostnameOrIpaddress, ".", vbTextCompare) <> 0 Then
            If IsNumeric(Mid(hostnameOrIpaddress, 1, InStr(1, hostnameOrIpaddress, ".") - 1)) = False Then
                Set returnIp = ResolveIpaddress(hostnameOrIpaddress)
                If returnIp.Count = 0 Then
                    Ping = False
                    Exit Function
                Else
                    hostnameOrIpaddress = returnIp.Item(1)
                End If
            End If
        End If
        Call PingAddress((hostnameOrIpaddress), echoValues, timeOutmSec)
        
        If Left$(echoValues.Data, 1) <> Chr$(0) Then
           pos = InStr(echoValues.Data, Chr$(0))
           echoValues.Data = Left$(echoValues.Data, pos - 1)
        Else
              echoValues.Data = ""
        End If
             
        SocketsCleanup
        
        If echoValues.Status <> 0 Then
            Ping = False
        Else
            Ping = True
        End If
    End If
    Exit Function
e_Trap:
    Ping = False
End Function
Private Function PingAddress(szAddress As String, ECHO As ICMP_ECHO_REPLY, Optional TimeOut As Long = PING_TIMEOUT) As Long
   Dim hPort As Long
   Dim dwAddress As Long
   Dim sDataToSend As String
   Dim iOpt As Long
   sDataToSend = "Echo This"
   dwAddress = AddressStringToLong(szAddress)
   hPort = IcmpCreateFile()
   If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), TimeOut) Then
         PingAddress = ECHO.RoundTripTime
   Else: PingAddress = ECHO.Status * -1
   End If
   Call IcmpCloseHandle(hPort)
End Function
Sub Win_Hide_ProgramTray()
Dim shelltraywnd As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
Call showwindow(shelltraywnd, SW_HIDE)
End Sub
Sub Windows_Show_Tray()
Dim shelltraywnd As Long
Dim traynotifywnd As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
Call showwindow(traynotifywnd, SW_SHOW)
End Sub
Sub Windows_Hide_Tray()
Dim shelltraywnd As Long
Dim traynotifywnd As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
Call showwindow(traynotifywnd, SW_HIDE)
End Sub
Public Sub Windows_Shutdown()
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub
Sub Windows_Hide_Clock()
Dim shelltraywnd As Long
Dim traynotifywnd As Long
Dim trayclockwclass As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
trayclockwclass = FindWindowEx(traynotifywnd, 0&, "trayclockwclass", vbNullString)
Call showwindow(trayclockwclass, SW_HIDE)
End Sub
Sub Windows_Show_Clock()
Dim shelltraywnd As Long
Dim traynotifywnd As Long
Dim trayclockwclass As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
traynotifywnd = FindWindowEx(shelltraywnd, 0&, "traynotifywnd", vbNullString)
trayclockwclass = FindWindowEx(traynotifywnd, 0&, "trayclockwclass", vbNullString)
Call showwindow(trayclockwclass, SW_SHOW)
End Sub
Public Function Hide_AOL()
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call showwindow(aolframe&, SW_HIDE)
End Function
Public Function Show_AOL()
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call showwindow(aolframe&, SW_SHOWNORMAL)
End Function
Public Function File_Exists(FileName As String) As Boolean
    If Len(FileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(FileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Function Form_OnTop(Form As Form)
'Example:
'Call error.FormOnTop Me
Call SetWindowPos(Form.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Function
Public Function Form_NotOnTop(Form As Form)
'Example:
'Call error.FormNotOnTop Me
Call SetWindowPos(Form.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Function
Public Function Close_Aol()
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call SendMessage(aolframe&, WM_CLOSE, 0&, 0&)
End Function
Public Function IntsantMessage_Close()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Send Instant Message")
Call SendMessage(aolchild&, WM_CLOSE, 0&, 0&)
End Function
Public Function File_GetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function
Public Function Chat_Maximize()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
'finds frame
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
Call showwindow(aolchild&, SW_MAXIMIZE)
End Function
Public Function InstantMessage_Minimize()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Send Instant Message")
Call showwindow(aolchild&, SW_MINIMIZE)
End Function
Public Function InstantMessage_Maximize()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Send Instant Message")
Call showwindow(aolchild&, SW_MAXIMIZE)
End Function
Public Function Chat_Minimize()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
Call showwindow(aolchild&, SW_MINIMIZE)
End Function
Sub File_Copy(FileName$, DestinationFile$)
If Not File_IfileExists(FileName$) Then Exit Sub
FileCopy FileName$, DestinationFile$
End Sub
Public Function Close_BuddyList()
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Welcome, Itz Da  s ourc e!")
aolchild& = FindWindowEx(mdiclient&, aolchild&, "AOL Child", vbNullString)
Call SendMessage(aolchild&, WM_CLOSE, 0&, 0&)
End Function
Sub Aim_Chat_Clear()
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Call SendMEssageByString(ateclass, WM_SETTEXT, 0&, "")
End Sub
Sub Aim_Chat_Hide_PeopleList()
'Example
'Call Chat_Hide_Buddies
Dim aimchatwnd As Long
Dim oscartree As Long
Dim X As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscartree = FindWindowEx(aimchatwnd, 0&, "_oscar_tree", vbNullString)
X& = showwindow(oscartree, SW_HIDE)
End Sub
Sub Aim_Chat_Show_SendButton()
Dim aimchatwnd As Long
Dim oscariconbtn As Long
Dim X As String
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
X = showwindow(oscariconbtn, SW_SHOW)
End Sub
Sub Aim_Chat_Show_Meter()
Dim aimchatwnd As Long
Dim oscarratemeter As Long
Dim X As String
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscarratemeter = FindWindowEx(aimchatwnd, 0&, "_oscar_ratemeter", vbNullString)
X = showwindow(oscarratemeter, SW_SHOW)
End Sub
Sub Aim_Chat_Minimize()
Dim aimchatwnd As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Call showwindow(aimchatwnd, SW_MINIMIZE)
End Sub
Sub Aim_Chat_ClickLess()
Dim aimchatwnd As Long
Dim Button As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Button = FindWindowEx(aimchatwnd, 0&, "button", vbNullString)
Button = FindWindowEx(aimchatwnd, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub Aim_Chat_ClickIgnore()
Dim aimchatwnd As Long
Dim oscariconbtn As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub Aim_Chat_ClickIM()
Dim aimchatwnd As Long
Dim oscariconbtn As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_Chat_ClickInfo()
Dim aimchatwnd As Long
Dim oscariconbtn As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_Chat_Maximize()
Dim aimchatwnd As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Call showwindow(aimchatwnd, SW_MAXIMIZE)
End Sub
Sub Aim_Chat_ClickMore()
Dim aimchatwnd As Long
Dim Button As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Button = FindWindowEx(aimchatwnd, 0&, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub Aim_Chat_Hide_SendButton()

Dim aimchatwnd As Long
Dim oscariconbtn As Long
Dim X As String
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
X = showwindow(oscariconbtn, SW_HIDE)
End Sub
Sub Aim_Chat_Hide_Meter()
'Example
'Call Chat_Hide_Meter
Dim aimchatwnd As Long
Dim oscarratemeter As Long
Dim X As String
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscarratemeter = FindWindowEx(aimchatwnd, 0&, "_oscar_ratemeter", vbNullString)
X = showwindow(oscarratemeter, SW_HIDE)
End Sub
Sub Aim_Chat_ClickGetInfo()
Dim aimchatwnd As Long
Dim oscariconbtn As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub Aim_Chat_AddRoomToList(LISTNAME As ListBox)
    Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, Buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer, BuddyTree As Long

    ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatRoom& <> 0 Then
        Do
            BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMEssageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            Buffer$ = String$(NameLen, 0)
            Moo2 = SendMEssageByString(BuddyTree&, LB_GETTEXT, MooLoo, Buffer$)
            TabPos = InStr(Buffer$, Chr$(9))
            NameText$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            For mooz = 0 To LISTNAME.ListCount - 1
                If name$ = LISTNAME.List(mooz) Then
                    Well% = 123
                    GoTo Endz
                End If
            Next mooz
            If Well% <> 123 Then
                LISTNAME.AddItem name$
            Else
            End If
Endz:
        Next MooLoo
    End If
End Sub
Sub Aim_Chat_Hide()
Dim aimchatwnd As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Call showwindow(aimchatwnd, SW_HIDE)
End Sub

Sub Aim_Chat_Sounds_Off()
'Example
'Call Chat_Soundz_Off
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)

    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_Chat_Sounds_On()
'Example
'Call Chat_Soundz_On
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)

    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Aim_SiteNavigation(Address As String)
Dim oscarbuddylistwin As Long
Dim editx As Long
Dim send As Long
Dim oscariconbtn As Long
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
Call showwindow(oscarbuddylistwin, SW_SHOW)
Pause 0.1
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
editx = FindWindowEx(oscarbuddylistwin, 0&, "edit", vbNullString)
Call SendMEssageByString(editx&, WM_SETTEXT, 0&, Address$)
Pause 0.1
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn = FindWindowEx(oscarbuddylistwin, 0&, "_oscar_iconbtn", vbNullString)
send& = SendMessage(oscariconbtn, WM_LBUTTONDOWN, 0, 0&)
send& = SendMessage(oscariconbtn, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Aim_Chat_Show()
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim X As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
X& = showwindow(wndateclass, SW_SHOW)
End Sub

Sub Aim_Chat_Show_PeopleList()
Dim aimchatwnd As Long
Dim oscartree As Long
Dim X As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscartree = FindWindowEx(aimchatwnd, 0&, "_oscar_tree", vbNullString)
X& = showwindow(oscartree, SW_SHOW)
End Sub
Sub Aim_Chat_Send(WhatToSend As String)
'Example
'Call Chat_Send(Text1)
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim oscariconbtn As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Call SendMEssageByString(ateclass&, WM_SETTEXT, 0&, WhatToSend$)
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub Play_MIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub Stop_MIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Sub Stop_CD()
     Dim lRet As Long
     lRet = mciSendString("stop cd wait", 0&, 0, 0)
     DoEvents
     lRet = mciSendString("close cd", 0&, 0, 0)
End Sub
Sub Play_CD(TRack$)
     Dim lRet As Long
     Dim nCurrentTrack As Integer
     lRet = mciSendString("open cdaudio alias cd wait", 0&, 0, 0)
     lRet = mciSendString("set cd time format tmsf", 0&, 0, 0)
     lRet = mciSendString("play cd", 0&, 0, 0)
     nCurrentTrack = TRack
     lRet = mciSendString("play cd from" & Str(nCurrentTrack), 0&, 0, 0)
     End Sub

Public Sub File_Set_Normal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub File_Set_ReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub File_Set_Hidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Sub Delete_File(TheFile As String)
Kill TheFile$
End Sub
End Sub
Sub Create_Folder(FolderToCreate As String)
MkDir FolderToCreate$
End Sub
Sub Delete_Folder(FolderToDelete As String)
RmDir FolderToDelete$
End Sub
Sub Disable_Ctrl_Alt_Delete()
Call SystemParametersInfo(97, True, 0&, 0)
End Sub
Sub Enable_Ctrl_Alt_Delete()
Call SystemParametersInfo(97, False, 0&, 0)
End Sub
Public Function Close_ConnectionLog()
'This will close the connection log when you are trying to sign on.
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Connection Log")
Call SendMessage(aolchild&, WM_CLOSE, 0&, 0&)
End Function
Function LastChatLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Public Function AutoClose_ConnectionLog() As Long
'This will close the connection log when you are trying to sign on.
'Note:
'For this Function to work, you must put the following in a timer:
'If error.AutoCloseConnectionLog <> 0& then
'call error.CloseConnectionLog
'else
'end if
Dim Counter As Long
Dim AOLView As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
AOLView& = FindWindowEx(aolchild&, 0&, "_AOL_View", vbNullString)
Do While (Counter& <> 100&) And (AOLView& = 0&): DoEvents
    aolchild& = FindWindowEx(mdiclient&, aolchild&, "AOL Child", vbNullString)
    AOLView& = FindWindowEx(aolchild&, 0&, "_AOL_View", vbNullString)
    If AOLView& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    AutoCloseConnectionLog& = aolchild&
    Exit Function
End If
End Function
Public Function Mail_ChangeCopyToCaption(NewCaption As String)
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
Call SendMEssageByString(aolstatic&, WM_SETTEXT, 0&, NewCaption$)
End Function
Public Function Mail_ChangeSubjectCaption(NewCaption As String)
Dim i As Long
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
Next i&
Call SendMEssageByString(aolstatic&, WM_SETTEXT, 0&, NewCaption$)
End Function
Public Function Buudylist_ClickAwayNotice()
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(aolchild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function Mail_ChangeSendToCaption(NewCaption As String)
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Write Mail")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
Call SendMEssageByString(aolstatic&, WM_SETTEXT, 0&, NewCaption$)
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Sub Kill_File(FileToDelete As String)
Kill FileToDelete$
End Sub
Public Function Get_FromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub Play_Wav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
Public Function Chat_CountInRoom() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    Chat_CountInRoom& = Count&
End Function
Public Sub Form_Drag(TheForm As Form)
'Example:
'Call error.FormDrag Me
    Call ReleaseCapture
    Call SendMessage(TheForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Function SignOnScreen_ChangeEnterPasswordCaption(Text As String)
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Sign On")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
Call SendMEssageByString(aolstatic&, WM_SETTEXT, 0&, Text$)
End Function
Public Function SignOnscreen_ChangeSelectLocationCaption(Text As String)
Dim i As Long
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Sign On")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 3&
    aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
Next i&
Call SendMEssageByString(aolstatic&, WM_SETTEXT, 0&, Text$)
End Function
Public Function SignOnScreen_ChangeScreenNameCaption(Text As String)
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", "Sign On")
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
Call SendMEssageByString(aolstatic&, WM_SETTEXT, 0&, Text$)
End Function
Public Sub Form_Exit_Down(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub

Public Sub Form_Exit_Left(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) - 300))
    Loop Until TheForm.Left < -TheForm.Width
End Sub
Public Sub Form_Exit_Diagonally(TheForm As Form)
'This makes a form exit right diaganol
'ex. call formexitrightdiaganol(form1)
'Pr0g's Bas
Do
Pause 0.04
TheForm.Move TheForm.Left + 120, TheForm.Top + 120
Loop Until TheForm.Top > Screen.Height
If TheForm.Top > Screen.Height Then
End
End If
End Sub

Public Sub Form_Exit_Right(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) + 300))
    Loop Until TheForm.Left > Screen.Width
End Sub

Public Sub Form_Exit_Up(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub
Public Sub RunToolbar_2(iconnumber&, letter$, letter2$)
Dim aolframe As Long, menu As Long, aoltoolbar1 As Long
Dim aoltoolbar2 As Long, AOLIcon As Long, Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2 = FindWindowEx(aoltoolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(aoltoolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To iconnumber
AOLIcon = FindWindowEx(aoltoolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter$ = Asc(letter)
letter2$ = Asc(letter2)
Call PostMessage(menu, WM_CHAR, letter$, 0&)
Call PostMessage(menu, WM_CHAR, letter2$, 0&)
End Sub

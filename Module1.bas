Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#If Win32 Then

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else

Public Declare Function ShellExecute Lib "shell.dll" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If
Public Const SW_SHOWNORMAL = 1

Declare Sub ReleaseCapture Lib "User32" ()
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public fso As New FileSystemObject
Public strm As TextStream

Public Const blue As String = &HFF0000
Public Const Black As String = &H80000012
Public Const Grey As String = &H808080

Public Const ProgramName As String = "BartNet MSN Fun"
Public Const ProgramVersion As String = "1.0.0"

Public Sub SetHelpView(ByVal What As String, ByVal Which As String)
    With Form1
        If What = "Hide" Then
            Select Case Which
                Case "Auto Message"
                    .txtHelp.Visible = False
                Case "Blocker"
                    .txtHelp.Visible = False
                Case "Crasher"
                    .txtHelp.Visible = False
                Case "Logger"
                    .txtHelp.Visible = False
                Case "MultiMessage Sender"
                    .txtHelp.Visible = False
                Case "NickName Popups"
                    .txtHelp.Visible = False
                Case "NickName Scroller"
                    .txtHelp.Visible = False
                Case "Send IM"
                    .txtHelp.Visible = False
                Case "Talk Offline"
                    .txtHelp.Visible = False
                Case "Welcoming Message"
                    .txtHelp.Visible = False
                Case "About"
                    .lblAbout.Visible = False
            End Select
        Else
            Select Case Which
                Case "Auto Message"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & _
                        "Auto Message will automatically send a message to any user who says anything to you.  There is a 90 second timeout so that the user doesn't get overloaded with messages." & vbCrLf & vbCrLf & _
                        "This is usefull for when you are away and want to let people know that you will be back soon." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "1) Check the Enable Auto Message checkbox to activate." & vbCrLf & "2) Select whether you want your status to be set as Busy or Away." & vbCrLf & "3) Type the message you want to be send." & vbCrLf & vbCrLf & "optional :" & vbCrLf & vbCrLf & "4) Check the Include Estimated Time Of Return checkbox." & vbCrLf & "5) Now choose between Countdown or Normal.  With normal it will include the time of terurn the way you enter it into the textbox, with countdown it will take off 1 minute every minute so that it is 0 when you return." & vbCrLf & vbCrLf & "Done, for more info select send feedback in the combobox above and ask whatever questions you have."
                Case "Blocker"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "Blocker will block users.  Normal block is the block you know, BartNet block will simply block any messages coming from the users on the BartNet block list." & vbCrLf & vbCrLf & "This comes in handy if someone is trying to crash you, as it will not allow you to be crashed by the users in the BartNet block list." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "Simple, just select a user from any list and press the appropriate button." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "Crasher"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "Crasher will send a very long message every millisecond to te selected user.  Because there will be so many messages coming in, the other user's MSN Messenger will eventually freeze and they will be forced to use Ctrl + Alt + Delete to end the program, logging them off in the process." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "Just select any user from the list and press the Crash button.  When the user is crashed the crash history list will be updated." & vbCrLf & vbCrLf & "BE PATIENT ! Not all users will crash quickly, it depends on a number of things." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "Logger"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "Logger will log most of the incoming data, such as when a user changes his / her nickname or status, when you recieve an email, etc." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "Check the Enable Log checkbox to log the data, check the Save Logs To Text File On Exit checkbox to save the data to Log.txt (found in the folder of the program)." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "MultiMessage Sender"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "MultiMessage Sender will send a message to a user a certain amount of times." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "1) Type the message to be send." & vbCrLf & "2) Select the user to send it to." & vbCrLf & "3) Type how many times the message is to be send (must be a number between 1 and 10000)." & vbCrLf & "4) Press GO to send the message." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "NickName Popups"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "NickName Popups will generate nickname popups (like the ones you get when someone signs in).  You choose what each popup will say." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "1) Type the text to popup in the textbox, start a new line for each popup. The top line will be the topmost popup." & vbCrLf & "2) Press the Popup Now button to popup." & vbCrLf & vbCrLf & "Press the Show People You Are Online button to generate 8 popups with your current nickname, this is shure to let everybody know you are online." & vbCrLf & vbCrLf & "DO NOT USE POPUPS TOO MANY TIMES IN A ROW, IT COULD SLOW DOWN YOUR MSN." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "NickName Scroller"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "NickName Scroller will change your nickname automatically every certain period of time." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "1) Type a different nickname in each textbox.  The top will be your first nickname, and so on." & vbCrLf & "2) Check the Reset NickName on Stop checkbox if you want your nickname to automatically be changed back to it's original when you stop the nickname scroller." & vbCrLf & "3) Type the time between each nickname change (in milliseconds).  Don't make it too small as it will slow down your MSN.  A good time is once every 4 or 5 seconds." & vbCrLf & "4) Press Start to start." & vbCrLf & vbCrLf & "Check the Set Nick As Time checkbox to have the time as nickname." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "Send IM"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "Send IM will simply open an Instant Message window.  It comes in handy if you need to say something to somebosy while using BartNet MSN Fun." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "1) Select a user." & vbCrLf & "2) Press the Send IM button." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "Talk Offline"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "Talk offline will allow you to appear offline and still send instant messages !" & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "Select Appear Offline to appear offline and Appear Online to appear online." & vbCrLf & vbCrLf & "WARNING : WHEN USING THIS ALL USERS WILL BE UNBLOCKED AUTOMATICALLY." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "Welcoming Message"
                    .txtHelp.Visible = True
                    
                    .txtHelp.Text = "What does it do :" & vbCrLf & vbCrLf & "Welcoming Message will automatically send a certain message to any user that comes online." & vbCrLf & vbCrLf & "How to use :" & vbCrLf & vbCrLf & "1) Check the Enable Welcoming Message checkbox to activate." & vbCrLf & "2) In the textbox enter the text you want to be send." & vbCrLf & vbCrLf & "For more info select send feedback in the combobox above and ask whatever questions you have."
                Case "About"
                    .lblAbout.Visible = True
            End Select
        End If
    End With
End Sub

Public Sub SetView(ByVal What As String, ByVal Which As String)
    With Form1
        If What = "Hide" Then
            Select Case Which
                Case "AutoMessage"
                    .ckAutoMessage.Visible = False
                    .ckIncludeTime.Visible = False
                    .Picture1.Visible = False
                    .Picture2.Visible = False
                    .lblAway.Visible = False
                    .lblBusy.Visible = False
                    .lblToSend.Visible = False
                    .txtAutoMessage.Visible = False
                    .txtCountDown.Visible = False
                    .txtNormal.Visible = False
                    .lblCheckAutoMessage.Visible = False
                    .lblIncludeTime.Visible = False
                    .lblCountDown.Visible = False
                    .lblNormal.Visible = False
                Case "Blocker"
                    .lblNormalBlocked.Visible = False
                    .lstNormalBlocked.Visible = False
                    .lblNotBlocked.Visible = False
                    .lstNotBlocked.Visible = False
                    .lblBartNetBlocked.Visible = False
                    .lstBartNetBlocked.Visible = False
                    .cmdNormalAllow.Visible = False
                    .cmdNormalBlock.Visible = False
                    .cmdBartNetBlock.Visible = False
                    .cmdBartNetAllow.Visible = False
                    .cmdNormalAllowAll.Visible = False
                    .cmdNormalBlockAll.Visible = False
                    .cmdBartNetBlockAll.Visible = False
                    .cmdBartNetAllowAll.Visible = False
                Case "Crasher"
                    .t1.Visible = False
                    .cmdCrash.Visible = False
                    .lblCurrentlyCrashing.Visible = False
                    .lblCrashHistory.Visible = False
                    .lstCrashHistory.Visible = False
                Case "Help"
                    .lblSelectCategory.Visible = False
                    .cboHelp.Visible = False
                    .cmdHelpGO.Visible = False
                
                    SetHelpView "Hide", "Auto Message"
                    SetHelpView "Hide", "Blocker"
                    SetHelpView "Hide", "Crasher"
                    SetHelpView "Hide", "Logger"
                    SetHelpView "Hide", "MultiMessage Sender"
                    SetHelpView "Hide", "NickName Popups"
                    SetHelpView "Hide", "NickName Scroller"
                    SetHelpView "Hide", "Send IM"
                    SetHelpView "Hide", "Talk Offline"
                    SetHelpView "Hide", "Welcoming Message"
                    SetHelpView "Hide", "About"
                Case "Logger"
                    .ckLog.Visible = False
                    .lblLog.Visible = False
                    .ckSaveLog.Visible = False
                    .lblSaveLog.Visible = False
                    .txtLog.Visible = False
                Case "MultiMessageSender"
                    .lblSendWhat.Visible = False
                    .txtSendWhat.Visible = False
                    .lblToWho.Visible = False
                    .t2.Visible = False
                    .lblHowManyTimes.Visible = False
                    .txtHowManyTimes.Visible = False
                    .cmdGo.Visible = False
                Case "NickNamePopups"
                    .txtPopup.Visible = False
                    .lblPopup.Visible = False
                    .cmdPopup.Visible = False
                    .cmdShowPeopleYouAreOnline.Visible = False
                Case "NickNameScroller"
                    .txtScroll(0).Visible = False
                    .txtScroll(1).Visible = False
                    .txtScroll(2).Visible = False
                    .txtScroll(3).Visible = False
                    .txtScroll(4).Visible = False
                    .txtScroll(5).Visible = False
                    .ckNickTime.Visible = False
                    .lblNickTime.Visible = False
                    .ckResetNick.Visible = False
                    .lblResetNick.Visible = False
                    .Label1.Visible = False
                    .Label2.Visible = False
                    .txtTime.Visible = False
                    .cmdScroll.Visible = False
                Case "SendIM"
                    .t3.Visible = False
                    .cmdSendIM.Visible = False
                Case "TalkOffline"
                    .optAppearOnline.Visible = False
                    .lblAppearOnline.Visible = False
                    .optAppearOffline.Visible = False
                    .lblAppearOffline.Visible = False
                Case "WelcomingMessage"
                    .ckWelcomingMessage.Visible = False
                    .lblEnableWelcomingMessage.Visible = False
                    .lblMessage.Visible = False
                    .txtWelcomingMessage.Visible = False
            End Select
        Else
            Select Case Which
                Case "AutoMessage"
                    .ckAutoMessage.Visible = True
                    .ckIncludeTime.Visible = True
                    .Picture1.Visible = True
                    .Picture2.Visible = True
                    .lblAway.Visible = True
                    .lblBusy.Visible = True
                    .lblToSend.Visible = True
                    .txtAutoMessage.Visible = True
                    .txtCountDown.Visible = True
                    .txtNormal.Visible = True
                    .lblCheckAutoMessage.Visible = True
                    .lblIncludeTime.Visible = True
                    .lblCountDown.Visible = True
                    .lblNormal.Visible = True
                Case "Blocker"
                    .lblNormalBlocked.Visible = True
                    .lstNormalBlocked.Visible = True
                    .lblNotBlocked.Visible = True
                    .lstNotBlocked.Visible = True
                    .lblBartNetBlocked.Visible = True
                    .lstBartNetBlocked.Visible = True
                    .cmdNormalAllow.Visible = True
                    .cmdNormalBlock.Visible = True
                    .cmdBartNetBlock.Visible = True
                    .cmdBartNetAllow.Visible = True
                    .cmdNormalAllowAll.Visible = True
                    .cmdNormalBlockAll.Visible = True
                    .cmdBartNetBlockAll.Visible = True
                    .cmdBartNetAllowAll.Visible = True
                Case "Crasher"
                    .t1.Visible = True
                    .cmdCrash.Visible = True
                    .lblCurrentlyCrashing.Visible = True
                    .lblCrashHistory.Visible = True
                    .lstCrashHistory.Visible = True
                Case "Help"
                    .lblSelectCategory.Visible = True
                    .cboHelp.Visible = True
                    .cmdHelpGO.Visible = True
                    
                    .cmdHelpGO_Click
                Case "Logger"
                    .ckLog.Visible = True
                    .lblLog.Visible = True
                    .ckSaveLog.Visible = True
                    .lblSaveLog.Visible = True
                    .txtLog.Visible = True
                Case "MultiMessageSender"
                    .lblSendWhat.Visible = True
                    .txtSendWhat.Visible = True
                    .lblToWho.Visible = True
                    .t2.Visible = True
                    .lblHowManyTimes.Visible = True
                    .txtHowManyTimes.Visible = True
                    .cmdGo.Visible = True
                Case "NickNamePopups"
                    .txtPopup.Visible = True
                    .lblPopup.Visible = True
                    .cmdPopup.Visible = True
                    .cmdShowPeopleYouAreOnline.Visible = True
                Case "NickNameScroller"
                    .txtScroll(0).Visible = True
                    .txtScroll(1).Visible = True
                    .txtScroll(2).Visible = True
                    .txtScroll(3).Visible = True
                    .txtScroll(4).Visible = True
                    .txtScroll(5).Visible = True
                    .ckNickTime.Visible = True
                    .lblNickTime.Visible = True
                    .ckResetNick.Visible = True
                    .lblResetNick.Visible = True
                    .Label1.Visible = True
                    .Label2.Visible = True
                    .txtTime.Visible = True
                    .cmdScroll.Visible = True
                Case "SendIM"
                    .t3.Visible = True
                    .cmdSendIM.Visible = True
                Case "TalkOffline"
                    .optAppearOnline.Visible = True
                    .lblAppearOnline.Visible = True
                    .optAppearOffline.Visible = True
                    .lblAppearOffline.Visible = True
                Case "WelcomingMessage"
                    .ckWelcomingMessage.Visible = True
                    .lblEnableWelcomingMessage.Visible = True
                    .lblMessage.Visible = True
                    .txtWelcomingMessage.Visible = True
            End Select
        End If
    End With
End Sub

Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hWnd, &HA1, 2, 0&)
End Sub

Sub main()
On Error GoTo Y
    
    With Form1
        .Width = 7500
        .Height = 4995
        
        Set strm = fso.OpenTextFile(App.Path & "\Defaults.BartNet", ForReading)
        
        .Top = strm.ReadLine
        .Left = strm.ReadLine
        
        .imgMinimize.Top = 100
        .imgClose.Top = 100
        
        .imgClose.Left = .Width - .imgClose.Width - 130
        .imgMinimize.Left = .imgClose.Left - .imgMinimize.Width - 130
        
        .lblAutoMessage.Top = 120
        .lblBlocker.Top = 480
        .lblCrasher.Top = 840
        .lblHelp.Top = 1200
        .lblLogger.Top = 1560
        .lblMultiMessageSender.Top = 1920
        .lblNickNamePopups.Top = 2280
        .lblNickNameScroller.Top = 2640
        .lblSendIM.Top = 3000
        .lblTalkOffline.Top = 3360
        .lblWelcomingMessage.Top = 3720
        .lblStatus.Top = 4080
        
        .lblAutoMessage.Left = 120
        .lblBlocker.Left = 120
        .lblCrasher.Left = 120
        .lblHelp.Left = 120
        .lblLogger.Left = 120
        .lblMultiMessageSender.Left = 120
        .lblNickNamePopups.Left = 120
        .lblNickNameScroller.Left = 120
        .lblSendIM.Left = 120
        .lblTalkOffline.Left = 120
        .lblWelcomingMessage.Left = 120
        .lblStatus.Left = 120
        
        .lblAutoMessage.BackStyle = 0
        .lblBlocker.BackStyle = 0
        .lblCrasher.BackStyle = 0
        .lblHelp.BackStyle = 0
        .lblLogger.BackStyle = 0
        .lblMultiMessageSender.BackStyle = 0
        .lblNickNamePopups.BackStyle = 0
        .lblNickNameScroller.BackStyle = 0
        .lblSendIM.BackStyle = 0
        .lblTalkOffline.BackStyle = 0
        .lblWelcomingMessage.BackStyle = 0
        .lblStatus.BackStyle = 0
        
        .lblAutoMessage.ForeColor = Black
        .lblBlocker.ForeColor = Black
        .lblCrasher.ForeColor = Black
        .lblHelp.ForeColor = Black
        .lblLogger.ForeColor = Black
        .lblMultiMessageSender.ForeColor = Black
        .lblNickNamePopups.ForeColor = Black
        .lblNickNameScroller.ForeColor = Black
        .lblSendIM.ForeColor = Black
        .lblTalkOffline.ForeColor = Black
        .lblWelcomingMessage.ForeColor = Black
        .lblStatus.ForeColor = Black
        
        .ckAutoMessage.Top = 515
        .optAway.Top = 120
        .optBusy.Top = 480
        .lblCheckAutoMessage.Top = 515
        .lblAway.Top = 1080
        .lblBusy.Top = 1440
        .lblToSend.Top = 1800
        .txtAutoMessage.Top = 2160
        .lblIncludeTime.Top = 2880
        .ckIncludeTime.Top = 2880
        .optCountDown.Top = 140
        .optNormal.Top = 500
        .lblCountDown.Top = 3240
        .lblNormal.Top = 3600
        .txtCountDown.Top = 3205
        .txtNormal.Top = 3565
        .Picture1.Top = 3100
        .Picture2.Top = 960
        
        .ckAutoMessage.Left = 2040
        .optAway.Left = 60
        .optBusy.Left = 60
        .lblCheckAutoMessage.Left = 2400
        .lblAway.Left = 3000
        .lblBusy.Left = 3000
        .lblToSend.Left = 2640
        .txtAutoMessage.Left = 2640
        .ckIncludeTime.Left = 2040
        .lblIncludeTime.Left = 2400
        .optCountDown.Left = 100
        .optNormal.Left = 100
        .lblCountDown.Left = 3000
        .lblNormal.Left = 3000
        .txtCountDown.Left = 4680
        .txtNormal.Left = 4680
        .Picture1.Left = 2550
        .Picture2.Left = 2580
        
        .Caption = ProgramName
        
        .ckAutoMessage.Value = strm.ReadLine
        .optAway.Value = strm.ReadLine
        .optBusy.Value = strm.ReadLine
        .ckIncludeTime.Value = strm.ReadLine
        .optCountDown.Value = strm.ReadLine
        .optNormal.Value = strm.ReadLine
        .txtCountDown.Text = strm.ReadLine
        .txtNormal.Text = strm.ReadLine
        
        .lblNormalBlocked.Top = 120
        .lstNormalBlocked.Top = 480
        .lblNotBlocked.Top = 1560
        .lstNotBlocked.Top = 1920
        .lblBartNetBlocked.Top = 3000
        .lstBartNetBlocked.Top = 3360
        .cmdNormalAllow.Top = 1200
        .cmdNormalBlock.Top = 1560
        .cmdBartNetBlock.Top = 2640
        .cmdBartNetAllow.Top = 3000
        .cmdNormalAllowAll.Top = 600
        .cmdNormalBlockAll.Top = 1920
        .cmdBartNetBlockAll.Top = 2280
        .cmdBartNetAllowAll.Top = 3480
        
        .lblNormalBlocked.Left = 1920
        .lstNormalBlocked.Left = 1920
        .lblNotBlocked.Left = 1920
        .lstNotBlocked.Left = 1920
        .lblBartNetBlocked.Left = 1920
        .lstBartNetBlocked.Left = 1920
        .cmdNormalAllow.Left = 3840
        .cmdNormalBlock.Left = 3360
        .cmdBartNetBlock.Left = 3360
        .cmdBartNetAllow.Left = 3840
        .cmdNormalAllowAll.Left = 5520
        .cmdNormalBlockAll.Left = 5520
        .cmdBartNetBlockAll.Left = 5520
        .cmdBartNetAllowAll.Left = 5520
        
        .t1.Top = 120
        .cmdCrash.Top = 1560
        .lblCurrentlyCrashing.Top = 1920
        .lblCrashHistory.Top = 2280
        .lstCrashHistory.Top = 2640
        
        .t1.Left = 1920
        .cmdCrash.Left = 1920
        .lblCurrentlyCrashing.Left = 1920
        .lblCrashHistory.Left = 1920
        .lstCrashHistory.Left = 1920
        
        .lstCrashHistory.ColumnHeaders(1).Width = .lstCrashHistory.Width / 4 * 2 - 15
        .lstCrashHistory.ColumnHeaders(2).Width = .lstCrashHistory.Width / 4 - 15
        .lstCrashHistory.ColumnHeaders(3).Width = .lstCrashHistory.Width / 4 - 15
        
        .ckLog.Top = 120
        .lblLog.Top = 120
        .ckSaveLog.Top = 120
        .lblSaveLog.Top = 120
        .txtLog.Top = 480
        
        .ckLog.Left = 1920
        .lblLog.Left = 2160
        .ckSaveLog.Left = 3960
        .lblSaveLog.Left = 4200
        .txtLog.Left = 1920
        
        .ckLog.Value = strm.ReadLine
        .ckSaveLog.Value = strm.ReadLine
        
        .lblSendWhat.Top = 120
        .txtSendWhat.Top = 480
        .lblToWho.Top = 1200
        .t2.Top = 1560
        .lblHowManyTimes.Top = 3360
        .txtHowManyTimes.Top = 3330
        .cmdGo.Top = 3720
        
        .lblSendWhat.Left = 1920
        .txtSendWhat.Left = 1920
        .lblToWho.Left = 1920
        .t2.Left = 1920
        .lblHowManyTimes.Left = 1920
        .txtHowManyTimes.Left = 3480
        .cmdGo.Left = 1920
        
        .txtHowManyTimes.Text = strm.ReadLine
        
        .txtPopup.Top = 120
        .lblPopup.Top = 480
        .cmdPopup.Top = 3120
        .cmdShowPeopleYouAreOnline.Top = 3600
        
        .txtPopup.Left = 1920
        .lblPopup.Left = 4440
        .cmdPopup.Left = 4440
        .cmdShowPeopleYouAreOnline.Left = 4440
        
        .txtScroll(0).Top = 360
        .txtScroll(1).Top = 720
        .txtScroll(2).Top = 1080
        .txtScroll(3).Top = 1440
        .txtScroll(4).Top = 1800
        .txtScroll(5).Top = 2160
        .ckNickTime.Top = 2640
        .lblNickTime.Top = 2640
        .ckResetNick.Top = 3000
        .lblResetNick.Top = 3000
        .Label1.Top = 3360
        .Label2.Top = 3360
        .txtTime.Top = 3325
        .cmdScroll.Top = 3600
        
        .txtScroll(0).Left = 3240
        .txtScroll(1).Left = 3240
        .txtScroll(2).Left = 3240
        .txtScroll(3).Left = 3240
        .txtScroll(4).Left = 3240
        .txtScroll(5).Left = 3240
        .ckNickTime.Left = 3240
        .lblNickTime.Left = 3480
        .ckResetNick.Left = 3240
        .lblResetNick.Left = 3480
        .Label1.Left = 1920
        .Label2.Left = 5040
        .txtTime.Left = 3840
        .cmdScroll.Left = 6240
        
        .txtScroll(0).Text = strm.ReadLine
        .txtScroll(1).Text = strm.ReadLine
        .txtScroll(2).Text = strm.ReadLine
        .txtScroll(3).Text = strm.ReadLine
        .txtScroll(4).Text = strm.ReadLine
        .txtScroll(5).Text = strm.ReadLine
        .ckNickTime.Value = strm.ReadLine
        .ckResetNick.Value = strm.ReadLine
        .txtTime.Text = strm.ReadLine
        
        .t3.Top = 120
        .cmdSendIM.Top = 3480
        
        .t3.Left = 1920
        .cmdSendIM.Left = 6600
        
        .optAppearOnline.Top = 1770
        .lblAppearOnline.Top = 1770
        .optAppearOffline.Top = 2250
        .lblAppearOffline.Top = 2250
        
        .optAppearOnline.Left = 3060
        .lblAppearOnline.Left = 3345
        .optAppearOffline.Left = 3060
        .lblAppearOffline.Left = 3345
        
        .optAppearOnline.Value = True
        
        .ckWelcomingMessage.Top = 840
        .lblEnableWelcomingMessage.Top = 840
        .lblMessage.Top = 1920
        .txtWelcomingMessage.Top = 2280
        
        .ckWelcomingMessage.Left = 2400
        .lblEnableWelcomingMessage.Left = 2640
        .lblMessage.Left = 2400
        .txtWelcomingMessage.Left = 2400
        
        .ckWelcomingMessage.Value = strm.ReadLine
        
        .cboHelp.AddItem "Auto Message"
        .cboHelp.AddItem "Blocker"
        .cboHelp.AddItem "Crasher"
        .cboHelp.AddItem "Logger"
        .cboHelp.AddItem "MultiMessage Sender"
        .cboHelp.AddItem "NickName Popups"
        .cboHelp.AddItem "NickName Scroller"
        .cboHelp.AddItem "Send IM"
        .cboHelp.AddItem "Talk Offline"
        .cboHelp.AddItem "Welcoming Message"
        .cboHelp.AddItem "--------------------------------------"
        .cboHelp.AddItem "Report Errors"
        .cboHelp.AddItem "Visit BartNet Online"
        .cboHelp.AddItem "Send Feedback"
        .cboHelp.AddItem "About"
        
        .lblSelectCategory.Top = 515
        .cboHelp.Top = 480
        .cmdHelpGO.Top = 480
        
        .lblSelectCategory.Left = 1920
        .cboHelp.Left = 3480
        .cmdHelpGO.Left = 5640
        
        .txtHelp.Top = 840
        
        .txtHelp.Left = 1920
        
        .lblAbout.Top = 840
        
        .lblAbout.Left = 1920
        
        .lblAbout.Caption = ProgramName & vbCrLf & vbCrLf & vbCrLf & "Created by Bart De Moitié" & vbCrLf & vbCrLf & vbCrLf & "Visit my web site for funny pictures, online games and much more :" & vbCrLf & vbCrLf & "http://www.bartnet.freeservers.com" & vbCrLf & vbCrLf & "You can always email me or add me to your MSN Messenger contacts list :" & vbCrLf & vbCrLf & "BartDeMoitie@msn.com" & vbCrLf & vbCrLf & vbCrLf & "Copyright 2002 to the BartNet Corp."
        
        If .ckLog.Value <> 1 Then
            .ckSaveLog.Enabled = False
            .lblSaveLog.ForeColor = Grey
            .txtLog.Enabled = False
        Else
            .ckSaveLog.Enabled = True
            .lblSaveLog.ForeColor = Black
            .txtLog.Enabled = True
        End If
        
        If .ckAutoMessage.Value <> 1 Then
            .optAway.Enabled = False
            .optBusy.Enabled = False
            .txtAutoMessage.Enabled = False
            .ckIncludeTime.Enabled = False
            .lblAway.ForeColor = Grey
            .lblBusy.ForeColor = Grey
            .lblToSend.ForeColor = Grey
            .lblIncludeTime.ForeColor = Grey
            
            .optCountDown.Enabled = False
            .optNormal.Enabled = False
            .txtCountDown.Enabled = False
            .txtNormal.Enabled = False
            .lblCountDown.ForeColor = Grey
            .lblNormal.ForeColor = Grey
            
            .timAutoMessage.Enabled = False
        Else
            .optAway.Enabled = True
            .optBusy.Enabled = True
            .txtAutoMessage.Enabled = True
            .ckIncludeTime.Enabled = True
            .lblAway.ForeColor = Black
            .lblBusy.ForeColor = Black
            .lblToSend.ForeColor = Black
            .lblIncludeTime.ForeColor = Black
        End If
        
        If .ckAutoMessage.Value = Checked Then
        
        Else
            If .ckIncludeTime.Value <> Checked Then
                If .optCountDown.Value = True Then
                    .txtCountDown.Enabled = True
                    .txtNormal.Enabled = False
                    .lblCountDown.ForeColor = Black
                    .lblNormal.ForeColor = Grey
                    
                    .timAutoMessage.Enabled = True
                Else
                    .txtCountDown.Enabled = False
                    .txtNormal.Enabled = True
                    .lblCountDown.ForeColor = Grey
                    .lblNormal.ForeColor = Black
                    
                    .timAutoMessage.Enabled = False
                End If
                .optCountDown.Enabled = True
                .optNormal.Enabled = True
            Else
                .txtCountDown.Enabled = False
                .txtNormal.Enabled = False
                .optCountDown.Enabled = False
                .optNormal.Enabled = False
                .lblCountDown.ForeColor = Grey
                .lblNormal.ForeColor = Grey
                
                .timAutoMessage.Enabled = False
            End If
        End If
        
        If .ckNickTime.Value <> 1 Then
            .txtScroll(0).Enabled = True
            .txtScroll(1).Enabled = True
            .txtScroll(2).Enabled = True
            .txtScroll(3).Enabled = True
            .txtScroll(4).Enabled = True
            .txtScroll(5).Enabled = True
            .txtTime.Enabled = True
        Else
            .txtScroll(0).Enabled = False
            .txtScroll(1).Enabled = False
            .txtScroll(2).Enabled = False
            .txtScroll(3).Enabled = False
            .txtScroll(4).Enabled = False
            .txtScroll(5).Enabled = False
            .txtTime.Enabled = False
        End If
        
        If .ckWelcomingMessage.Value <> 1 Then
            .txtWelcomingMessage.Enabled = False
            .lblMessage.Enabled = False
        Else
            .txtWelcomingMessage.Enabled = True
            .lblMessage.Enabled = True
        End If
        
        If .ckIncludeTime.Enabled = True Then
            If .ckIncludeTime.Value <> 1 Then
                .optCountDown.Enabled = False
                .optNormal.Enabled = False
                .lblCountDown.ForeColor = Grey
                .lblNormal.ForeColor = Grey
                .txtCountDown.Enabled = False
                .txtNormal.Enabled = False
            Else
                If .optCountDown.Value = True Then
                    .lblCountDown.ForeColor = Black
                    .lblNormal.ForeColor = Grey
                    .txtCountDown.Enabled = True
                    .txtNormal.Enabled = False
                Else
                    .lblCountDown.ForeColor = Grey
                    .lblNormal.ForeColor = Black
                    .txtCountDown.Enabled = False
                    .txtNormal.Enabled = True
                End If
            End If
        Else
            .optCountDown.Enabled = False
            .optNormal.Enabled = False
            .txtCountDown.Enabled = False
            .txtNormal.Enabled = False
            .lblCountDown.ForeColor = Grey
            .lblNormal.ForeColor = Grey
        End If
        
        Dim abc As Byte
        Dim cba As String
        
        abc = 0
        
        Do Until strm.AtEndOfStream
            cba = strm.ReadLine
            If cba = "<<*****END*****>>" Then
                abc = 1
            Else
                If cba = "**<<<<<END>>>>>**" Then
                    abc = 2
                Else
                    If cba = "*****<<END>>*****" Then
                        abc = 3
                    Else
                        Select Case abc
                            Case 0
                                .txtPopup.Text = .txtPopup.Text & cba & vbCrLf
                            Case 1
                                .txtSendWhat.Text = .txtSendWhat.Text & cba & vbCrLf
                            Case 2
                                .txtWelcomingMessage.Text = .txtWelcomingMessage.Text & cba & vbCrLf
                            Case 3
                                .txtAutoMessage.Text = .txtAutoMessage.Text & cba & vbCrLf
                        End Select
                    End If
                End If
            End If
        Loop
        
        .txtPopup.Text = Mid(.txtPopup.Text, 1, Len(.txtPopup.Text) - 2)
        .txtSendWhat.Text = Mid(.txtSendWhat.Text, 1, Len(.txtSendWhat.Text) - 2)
        '.txtWelcomingMessage.Text = Mid(.txtWelcomingMessage.Text, 1, Len(.txtWelcomingMessage.Text) - 2)
        .txtAutoMessage.Text = Mid(.txtAutoMessage.Text, 1, Len(.txtAutoMessage.Text) - 2)
        
        strm.Close
        
        .LoadTheFormForReal
        
        .Show
    End With
    
    App.Title = ProgramName
    
    SetView "Hide", "AutoMessage"
    SetView "Hide", "Blocker"
    SetView "Hide", "Crasher"
    SetView "Hide", "Help"
    SetView "Hide", "Logger"
    SetView "Hide", "MultiMessageSender"
    SetView "Hide", "NickNamePopups"
    SetView "Hide", "NickNameScroller"
    SetView "Hide", "SendIM"
    SetView "Hide", "TalkOffline"
    SetView "Hide", "WelcomingMessage"
    
    SetHelpView "Hide", "Auto Message"
    SetHelpView "Hide", "Blocker"
    SetHelpView "Hide", "Crasher"
    SetHelpView "Hide", "Logger"
    SetHelpView "Hide", "MultiMessage Sender"
    SetHelpView "Hide", "NickName Popups"
    SetHelpView "Hide", "NickName Scroller"
    SetHelpView "Hide", "Send IM"
    SetHelpView "Hide", "Talk Offline"
    SetHelpView "Hide", "Welcoming Message"
    SetHelpView "Hide", "About"
    
    Exit Sub
    
Y:
    Set strm = fso.CreateTextFile(App.Path & "\Defaults.BartNet", False)
    With strm
        .WriteLine (Screen.Height - Form1.Height) / 2
        .WriteLine (Screen.Width - Form1.Width) / 2
        
        .WriteLine "1"
        .WriteLine "true"
        .WriteLine "false"
        .WriteLine "1"
        .WriteLine "false"
        .WriteLine "true"
        .WriteLine "16:40"
        .WriteLine "18:40"
        .WriteLine "1"
        .WriteLine "1"
        .WriteLine "500"
        .WriteLine "BelgiumBoy_007 :)"
        .WriteLine "BelgiumBoy_007 :p"
        .WriteLine "BelgiumBoy_007 :d"
        .WriteLine "BelgiumBoy_007 (6)"
        .WriteLine "BelgiumBoy_007 (h)"
        .WriteLine "BelgiumBoy_007 (a)"
        .WriteLine "0"
        .WriteLine "1"
        .WriteLine "9000"
        .WriteLine "0"
        .WriteLine "popup 1"
        .WriteLine "popup 2"
        .WriteLine "popup 3"
        .WriteLine "<<*****END*****>>"
        .WriteLine ":@"
        .WriteLine ":@"
        .WriteLine "**<<<<<END>>>>>**"
        .WriteLine "G'day"
        .WriteLine "*****<<END>>*****"
        .WriteLine "I'm not here at the moment, I will be back soon."
        
        .Close
    End With
    main
End Sub

Public Sub SaveValues()
    Set strm = fso.OpenTextFile(App.Path & "\Defaults.BartNet", ForWriting)
    With strm
        .WriteLine Form1.Top
        .WriteLine Form1.Left
        .WriteLine Form1.ckAutoMessage.Value
        .WriteLine Form1.optAway.Value
        .WriteLine Form1.optBusy.Value
        .WriteLine Form1.ckIncludeTime.Value
        .WriteLine Form1.optCountDown.Value
        .WriteLine Form1.optNormal.Value
        .WriteLine Form1.txtCountDown.Text
        .WriteLine Form1.txtNormal.Text
        .WriteLine Form1.ckLog.Value
        .WriteLine Form1.ckSaveLog.Value
        .WriteLine Form1.txtHowManyTimes.Text
        .WriteLine Form1.txtScroll(0).Text
        .WriteLine Form1.txtScroll(1).Text
        .WriteLine Form1.txtScroll(2).Text
        .WriteLine Form1.txtScroll(3).Text
        .WriteLine Form1.txtScroll(4).Text
        .WriteLine Form1.txtScroll(5).Text
        .WriteLine Form1.ckNickTime.Value
        .WriteLine Form1.ckResetNick.Value
        .WriteLine Form1.txtTime.Text
        .WriteLine Form1.ckWelcomingMessage.Value
        .WriteLine Form1.txtPopup.Text
        .WriteLine "<<*****END*****>>"
        .WriteLine Form1.txtSendWhat.Text
        .WriteLine "**<<<<<END>>>>>**"
        .WriteLine Form1.txtWelcomingMessage.Text
        .WriteLine "*****<<END>>*****"
        .WriteLine Form1.txtAutoMessage.Text
        
        .Close
    End With
End Sub

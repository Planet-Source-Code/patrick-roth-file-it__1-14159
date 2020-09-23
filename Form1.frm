VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form File_Outlook_Items 
   Caption         =   "File Outlook Items"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleMode       =   0  'User
   ScaleTop        =   100
   ScaleWidth      =   5985
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   5400
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Access DB- Testing Only"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar MessageProgress 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Sent Items"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Messages to Process"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "File_Outlook_Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public MyOlApp As New Outlook.Application
 Public MyOlMessage As Outlook.MailItem
 Public MyOlSpace As Outlook.NameSpace
 Public MyFolder As Outlook.MAPIFolder
 Public MyOlMessageFolder As Outlook.MAPIFolder
 Public MoveToFolder As Outlook.MAPIFolder
 
 Public MsgCount, MyItem As Integer
 Public MyText As String
 Public currentmessage As Integer
 
Dim db As DAO.Database
Dim rsMsg As DAO.Recordset
Dim wrkODBC As DAO.Workspace

Public genie As IAgentCtlCharacterEx

Private Sub Command1_Click()
    Call ImportMessages("Sent Items")
    File_Outlook_Items.Hide
    Unload File_Outlook_Items
End
End Sub

'======================================================================
'FUNCTION: ParseRecipients
'
'Purpose: Check a MAPI message for a specific type of recipient and
'         return a semicolon delimited list of recipients. For
'         instance, if this function is called using the MapiTo
'         constant, this function will return a semicolon delimited
'         list of all recipients on the 'TO' line of the message.
'======================================================================

Function ParseRecipients(objMessage As Outlook.MailItem)
    Dim RecipientCount As Long
    Dim Recipient As Object
    Dim TheSender As Object
    Dim ReturnString As String
    Dim EmailName As String
    Dim result As String
    Dim messagecopy As Outlook.MailItem
    
    RecipientCount = objMessage.Recipients.Count
    Set Recipient = objMessage.Recipients(RecipientCount)
    If RecipientCount > 0 Then
        ReturnString = objMessage.Recipients(1).Address
    End If
    result = UCase((StripQuote(ReturnString)))
    'these are all the email addresses that we use
    'I check this because we also want to keep copies of what
    'we they send back to us - ie, didn't work or worked or whatever
    If result = UCase("/O=GPS/OU=GPSDNS/CN=RESOURCES/CN=TEAMS/CN=DYNTOOLS") _
    Or result = UCase("dexsupport@gps.com") _
    Or result = UCase("dexsupport@GreatPlains.com") _
    Or result = UCase("dexsuprt@gps.com") _
    Or result = UCase("dexsuprt@GreatPlains.com") _
    Or result = UCase("dexterity_support@gps.com") _
    Or result = UCase("dexterity_support@GreatPlains.com") _
    Or result = UCase("dyntools@GPS.com") _
    Or result = UCase("dyntools@gpsdns.gps.com") _
    Or result = UCase("dyntools@GreatPlains.com") _
    Or result = UCase("tdexteri@gps.com") _
    Or result = UCase("tdexteri@GreatPlains.com") Then
        'only way I could figure out how to get the sender address was to
        'make a copy with a reply.  Now the sender is the recipient and
        'I can get that easily enough
        Set messagecopy = objMessage.Reply
        ReturnString = StripQuote(messagecopy.Recipients(1).AddressEntry.Address)
        messagecopy.Delete
        GoTo mylabel
    End If

    For RecipientCount = 1 To objMessage.Recipients.Count
       
        Set Recipient = objMessage.Recipients(RecipientCount)
    
        EmailName = StripQuote(Recipient.Name)
        If EmailName = ReturnString Then
            ReturnString = StripQuote(Recipient.Address)
        End If
    Next
    
mylabel:
    If Len(ReturnString) > 0 Then
        ReturnString = Left(Trim(ReturnString), Len(ReturnString))
        ParseRecipients = ReturnString
    Else
        ParseRecipients = Null
    End If
End Function


'======================================================================
'SUB: WriteMessage
'
'Purpose: Adds message information to fields in the table through the
'         the recordset opened in the ImportMessages Sub. This

'         procedure is called from the RetrieveMessage Sub when it is
'         time to write information to the table.
'======================================================================

Sub WriteMessage(objMessage As Object, foldername As String, _
                 InfoStore As String, returnfolder As String)
    Dim RetVal
    Dim iString As String

    txt = "select * from Contacts where Email = '" & ParseRecipients(objMessage) & "'"
    Set rsMsg = db.OpenRecordset(txt, dbOpenDynamic, dbExecDirect, dbOptimistic)
    
    On Error Resume Next
    rsMsg.MoveFirst
  
    'if we can't find out where it belongs, ie no address logged then
    'set to nothing
    If rsMsg.EOF Then
        returnfolder = ""
    Else
        returnfolder = rsMsg!CompanyName
    End If
End Sub

'======================================================================
'SUB: RetrieveMessage
'
'Purpose: Loop through the Messages collection of each Folder of the

'         specified information store(s) and calls the WriteMessage Sub
'         to write individual messages to the table. This procedure is
'         called by the ImportMessages Sub.
'======================================================================

Sub RetrieveMessage(objInfoStore As Object, foldername As Variant)
    Dim objFoldersColl As Outlook.MAPIFolder
    Dim objFolder As Outlook.MAPIFolder
    Dim objMessage As Outlook.MailItem
    Dim FoundMessage As Outlook.MailItem
    Dim targetfolder As String
    Dim olTarget As Outlook.MAPIFolder
    Dim olNewTarget As Outlook.MAPIFolder
    Dim messagecount As Long
    Dim currentmessagecount As Long
    Dim olRootTarget As Outlook.MAPIFolder
    
    On Error GoTo moveErr
    
    Set MyOlSpace = MyOlApp.GetNamespace("MAPI")
    Set objFolder = MyOlSpace.GetDefaultFolder(olFolderSentMail)
    
    currentmessagecount = 1
     
    If objFolder.Name = foldername Then
        File_Outlook_Items.Show
        
        'here is where we choose where the messages to file are sitting,
        'probably in the send items folder
        MsgBox "Choose the folder where the messages to file are"
        Set olRootTarget = MyOlSpace.PickFolder
       
       'here is where we choose the root folder of where all the sub folders
       'are. In our case, our folder was Dyntools with all the company folders
       'set within it
        File_Outlook_Items.Show
        MsgBox "Choose the Dyntools Sent Items as the destination"
        Set objFolder = MyOlSpace.PickFolder
  
        messagecount = olRootTarget.Items.Count
        If messagecount < 1 Then messagecount = 1
        MessageProgress.Min = 0
        If messagecount <> 0 Then
            MessageProgress.Max = messagecount
        End If
        
        For currentmessagecount = messagecount To 1 Step -1
 
            Set FoundMessage = olRootTarget.Items(currentmessagecount)
            MessageProgress.Value = currentmessagecount
            Call WriteMessage(FoundMessage, _
            olRootTarget.Name, objInfoStore.Name, targetfolder)
            
            'we found matching address because the target folder not empty
            'so move it to the destination
            If targetfolder > "" Then
                Set olTarget = objFolder
                targetfolder = Trim(targetfolder)
                Set olTarget = olTarget.Folders.Item(targetfolder)
                FoundMessage.Move olTarget
            End If
        Next
        
    End If
    
    GoTo getout:
moveErr:
    File_Outlook_Items.Show
   If targetfolder <> "" Then
    Call moveerror(FoundMessage.Subject, targetfolder)
   End If
   
   Resume Next
getout:

End Sub


'======================================================================

'SUB: ImportMessage
'
'Purpose: Opens a MAPI session through OLE automation and opens a
'         recordset based on the Messages table. Then, this procedure
'         checks to see if it needs to import messages from top level
'         folders in ALL information stores, or just a specific
'         information store. Based upon this, the procedure will call
'         the RetrieveMessage sub for the specified information stores.
'======================================================================

Sub ImportMessages(Optional foldername As Variant, _
                   Optional InfoStoreName As Variant)
    Dim objMapi As Object
    Dim objFoldersColl As Object
    Dim objInfoStore As Object
    Dim RetVal
    Dim foldercount As Integer
    
    On Error GoTo helpme
    
    'set the sql connection string
    strConnect = "ODBC;DSN=OutlookDB;UID=sa;PWD="
    
    'Check1 is a testing check. if checked, then open an access db directly
    'instead
    If Check1 Then
        Set db = OpenDatabase("d:\outlook.mdb")
    Else
        If wrkODBC Is Nothing Then
            Set wrkODBC = DBEngine.CreateWorkspace("ODBC", "", "", dbUseODBC)
            Set db = wrkODBC.OpenDatabase("SQL Dynamics", dbDriverComplete, False, strConnect)
        End If
    End If
    


    Set objMapi = CreateObject("Mapi.Session")


'In the following line, replace the ProfileName argument with a valid
'profile. If you omit the ProfileName argument, Microsoft Exchange will
'prompt you for your profile.

    objMapi.Logon ProfileName:=""

    'Loop through each InfoStore in the MAPI session and determine if
    'we should read in messages from ALL InfoStores or just a specified

    'InfoStore. InfoStores include a user's personal store files
    '(.PST Files), Network stores, and Public Folders.

        For Each objInfoStore In objMapi.InfoStores
            If Not IsMissing(InfoStoreName) Then
                If objInfoStore.Name = InfoStoreName Then
                    Call RetrieveMessage(objInfoStore, foldername)
                    Exit For
                End If
                Exit Sub
            Else
                Call RetrieveMessage(objInfoStore, foldername)
                Exit Sub
            End If
        Next
    objMapi.Logoff  ' Log out of the MAPI session.
    Set objMapi = Nothing
    db.Close  ' Close the Database.
    Set db = Nothing
    Exit Sub
helpme:
    File_Outlook_Items.Show
    MsgBox Err.Description
    Call myerror

End Sub


Private Function StripQuote(instring As String) As String
    Dim x As Integer
    Dim i As Integer
    
    For x = 1 To Len(instring)
        If Mid(instring, x, 1) = ";" Or Mid(instring, x, 1) = "," Then
            Exit Function
        End If
        If Mid(instring, x, 1) <> "'" Then
            StripQuote = StripQuote & Mid(instring, x, 1)
        End If
    Next
End Function





Private Sub Command2_Click()
    If Command2.Caption = "Help" Then
        Command2.Caption = "Done"
             Set Agent1 = CreateObject("Agent.Control.1")
             Agent1.Connected = True
             Agent1.Characters.Load "Genie", "Genie.Acs"
             Set genie = Agent1.Characters("Genie")
             genie.Show
             genie.Top = 110
             genie.Left = 620


             genie.Speak ("Welcome to the Tools Outlook email filing program.")
             genie.Play ("Greet")
             genie.Play ("RestPose")
             genie.Speak ("I'm here to speed up the email filing process by filing emails with known email addresses.")
             genie.Play ("Explain")
             genie.Speak ("Press the 'Process Sent Items' button to start.")
             genie.Play ("RestPose")
             genie.Speak ("First select your profile, the default profile will probably work fine.")
             genie.Speak ("Then choose the folder where the items to be filed are.")
             genie.Speak ("This will likely be your Sent Items folder or the Tools Sent Items folder.")
             genie.Speak ("Next choose the folder where the items are to be filed.  This will be the Tools Sent Items folder.")
             genie.Speak ("And then Poof!")
             genie.Play ("DoMagic1")
             genie.Play ("DoMagic2")
             genie.Play ("RestPose")
             genie.Speak ("The emails are moved to the correct folder.")

             genie.Speak ("Congratulations, you hopefully have saved a signifigant amount of time.")
             genie.Play ("Congratulate")
             genie.Play ("RestPose")
             genie.Speak ("After the process is finished, the items remaining are new address emails")
             genie.Speak ("File these manually in the Tools Sent Items subfolders and run the Initial.exe program.")
             genie.Play ("RestPose")
             genie.Play ("RestPose")
             genie.Play ("Idle3_1")
             genie.Play ("Idle3_2")
             While Command2.Caption = "Done"
                DoEvents
             Wend
        
    Else
        Command2.Caption = "Help"
        genie.Stop
        genie.Hide
        
    End If
    

End Sub

Private Sub myerror()
    If Command2.Caption = "Help" Then
        Command2.Caption = "Done"
             Set Agent1 = CreateObject("Agent.Control.1")
             Agent1.Connected = True
             Agent1.Characters.Load "Genie", "Genie.Acs"
             Set genie = Agent1.Characters("Genie")
             genie.Show
             genie.Top = 110
             genie.Left = 620
             
             genie.Speak ("You are here because the connection to INTLDEV2 could not be reached")
             genie.Speak ("The most likely cause is not having the DSN setup.  Make sure that you have a SQL Driver DSN named OutlookDB.")
             genie.Speak ("It needs the Initial Database set to OutlookDB as well.")
      
             While Command2.Caption = "Done"
                DoEvents
             Wend
        
    Else
        Command2.Caption = "Help"
        genie.Stop
        genie.Hide
    End If

End Sub


Private Sub moveerror(mystring As String, theaddress As String)
    If Command2.Caption = "Help" Then
        Command2.Caption = "Done"
             Set Agent1 = CreateObject("Agent.Control.1")
             Agent1.Connected = True
             Agent1.Characters.Load "Genie", "Genie.Acs"
             Set genie = Agent1.Characters("Genie")
             genie.Show
             genie.Top = 110
             genie.Left = 620
             
              genie.Speak ("I'm having a problem with something.  Most likely it is because the email")
              genie.Speak ("I'm trying to move has a known email address but the folder where it belongs")
              genie.Speak ("has been deleted.  Or very possibly has spaces before or after the name.")
              genie.Speak ("The Subject of the email in question was " & mystring)
              genie.Speak ("The email was moving to " + theaddress)
              genie.Speak ("The email will be moved to the Tools Sent Items folder.")
              genie.Speak ("Sorry.  Hit the 'Done' button and we shall continue.")
              
          
             While Command2.Caption = "Done"
                DoEvents
             Wend
        
    Else
        Command2.Caption = "Help"
        genie.Stop
        genie.Hide
    End If

End Sub
Private Sub Form_Load()
File_Outlook_Items.Left = 3500
File_Outlook_Items.Top = 2000
End Sub

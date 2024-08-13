
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Const WM_PASTE = &H302
Const WM_UNDO = &H304

Dim strLanguage As String

Dim boolCancel As Boolean

Dim intDelay As Integer

Private Sub MDIForm_Initialize()

  Call InitCommonControls

End Sub

Private Sub MDIForm_Load()
On Error GoTo ErrHld

  Dim sTmp As String
  
  Me.Caption = "Textpad"

  intLRT = 1

  If ActiveForm Is Nothing Then
  
    LoadNewDoc

  End If

  mnuToolbar_Click
  mnuStatusBar_Click
  
  oSVoice = 1

  sTmp = GetSetting(App.Title, "Setting", "Volume", 80)

  If IsNumeric(sTmp) = False Then
    oSVolume = 80
  Else
    oSVolume = CInt(sTmp)
  End If
  
  objCount = 0

 If Not Len(Command$) = 0 Then
 
   If Len(Dir$(App.Path & "\Run.ini", vbNormal)) <> 0 Then
   
      mdiMain.LoadRun Replace(App.Path & "\Run.ini", Chr$(34), "")
   
   End If

 ElseIf Len(Command) < 0 Then

   mdiMain.LoadRun Replace(Command$, Chr$(34), "")

 End If
  
 intDelay = 1
 
 strLanguage = "VBScript"

Exit Sub
ErrHld:

 Unload Me

End Sub

Private Sub MDIForm_Resize()
On Error GoTo ErrHld

Exit Sub
ErrHld:

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo ErrHld
  
  Call SaveSetting(App.Title, "Setting", "Width", Me.Width)
  Call SaveSetting(App.Title, "Setting", "Height", Me.Height)
  
  Call SaveSetting(App.Title, "Setting", "WindowState", Me.WindowState)

  Unload frmAbout
  Unload frmDocument
  Unload frmEncode
  Unload frmTip
  Unload frmWebBrowser

  Dim i As Integer
  
  For i = 0 To iWebbrowser
  
    If Not fWebbrowser(i) Is Nothing Then
  
      Unload fWebbrowser(i)
  
    End If

  Next
  
  For i = 0 To iDocument
  
    If Not fDocument(i) Is Nothing Then
  
      fDocument(i).SetFocus
      
      DoEvents

   

      Unload fDocument(i)
  
    End If

  Next
  
Exit Sub
ErrHld:

End Sub

Private Sub mnuAboutDialog_Click()

  frmAbout.Show 1, Me

End Sub

Private Sub mnuAddress_Click()
On Error GoTo ErrHld

  strLastURL = InputBox("Enter Address URL: ", "Notice", "http://www.w3school.com")

  ActiveForm.wB1.Navigate strLastURL

Exit Sub
ErrHld:

  LoadNewWeb
  
  DoEvents

  ActiveForm.wB1.Navigate strLastURL

End Sub

Private Sub mnuArrangeIcon_Click()
 
 Me.Arrange vbArrangeIcons
 
End Sub

Private Sub mnuBack_Click()
On Error GoTo ErrHld

  If ActiveForm Is Nothing Then
    Exit Sub
  End If
  
  ActiveForm.wB1.GoBack

Exit Sub
ErrHld:

End Sub

Private Sub mnuCascade_Click()

  Me.Arrange 0 + vbArrangeIcons

End Sub

Private Sub mnuClose_Click()

 If ActiveForm Is Nothing Then
 
   Exit Sub
 
 End If
 
 sBar1.Panels(1).Text = "Close..."

 Unload ActiveForm
 
 timStatus.Enabled = True
 
End Sub

Private Sub mnuCloseAll_Click()
On Error GoTo ErrHld

  Dim i As Integer
  
  sBar1.Panels(1).Text = "Close windows..."

  For i = 0 To iDocument
  
    If Not fDocument(i) Is Nothing Then
  
      Unload fDocument(i)
  
    End If

  Next
  
  For i = 0 To iWebbrowser
  
    If Not fWebbrowser(i) Is Nothing Then
  
      Unload fWebbrowser(i)
  
    End If

  Next
  
  timStatus.Enabled = True

Exit Sub
ErrHld:
  
  timStatus.Enabled = True

End Sub

Private Sub mnuCopy_Click()

  If ActiveForm Is Nothing Then
    Exit Sub
  End If
  
  Call Clipboard.Clear
  Call Clipboard.SetText(ActiveForm.rtfText.SelText)
  
  ActiveForm.rtfText.SelStart = 0

End Sub

Public Function LoadRun(Optional ByVal sPath As String = "") As String
On Error GoTo ErrHld

  Dim strLine As String

  Dim i As Long
  
  Dim fFreeFile As Integer
  
  Dim bIO As Boolean

  Dim strCount As String

  Dim strLast As String
  
  fFreeFile = FreeFile
  
  strCount = 0
                   
  intSRT = iDocument
             
  intLRT = 1

  Open sPath For Input As #fFreeFile
  
     Do Until EOF(fFreeFile)
     
       Line Input #fFreeFile, strLine
     
       If InStr(1, strLine, "[Run]", vbBinaryCompare) <> 0 Then
       
         bIO = True
       
       End If
       
       If (bIO = True) Then
       
         If InStr(1, strLine, "Last=", vbBinaryCompare) <> 0 Then
         
           strLast = strLine
         
         End If
         
         If InStr(1, strLine, "Count=", vbBinaryCompare) <> 0 Then
         
          strCount = Replace(strLine, "Count=", "")
          
          If IsNumeric(strCount) = False Then
            Exit Function
          End If
         
         End If
         
         If InStr(1, strLine, "Pkg", vbBinaryCompare) <> 0 Then
           
           For i = 1 To CInt(strCount)
               
             If InStr(1, strLine, "Pkg" & i & "=", vbBinaryCompare) <> 0 Then

               strLine = Replace(strLine, "Pkg" & i & "=", "")
    
               If Len(strLine) = 0 Then
                 Exit Function
               End If
         
               If Not strLine = "" Then

                 LoadNewDoc
                 DoEvents

                 ActiveForm.rtfText.LoadFile strLine
                 
  ReDim Preserve LoadRunTag(intLRT)
                 
  LoadRunTag(intLRT) = strLine
                 
  intLRT = (intLRT + 1)
  
                 ActiveForm.Caption = strLine
    
                 ActiveForm.Tag = strLine
  
                 ActiveForm.WindowState = vbMaximized
   
               Else
               
                 LoadNewDoc
                 DoEvents

                 ActiveForm.rtfText.LoadFile strLine
                 
  ReDim Preserve LoadRunTag(intLRT)
                 
  LoadRunTag(intLRT) = strLine
                 
  intLRT = (intLRT + 1)
                   
                 ActiveForm.Caption = strLine
    
                 ActiveForm.Tag = strLine
  
                 ActiveForm.WindowState = vbMaximized

               End If

             End If
   
           Next

         End If

       End If
     
     Loop
  
  Close #fFreeFile

Exit Function
ErrHld:

End Function


Private Sub mnuCRun_Click()
On Error GoTo ErrHld

   With cDiag1
   
     .CancelError = True
     
     .DialogTitle = "Create Run..."
     
     .FileName = ""
     
     .Filter = "INI Files (*.ini)|*.ini"

     .ShowSave
     
     If Not Len(.FileName) = 0 Then
     
       CreateRun .FileName
     
     End If
   
   End With

Exit Sub
ErrHld:

End Sub

Public Function CreateRun(Optional sPath As String = "") As String
On Error Resume Next

   Dim i As Long
   Dim m As Long
  
   Dim fFreeFile As Integer
   
   Dim bIO As Boolean
   
   fFreeFile = FreeFile
   
   Open sPath For Output As #fFreeFile
       
     Print #fFreeFile, "[Run]"
     Print #fFreeFile, "Count=" & dlgDocumentCount

     For i = 1 To dlgDocumentCount
            
       If Not fDocument(i).Tag = "" Then
                
        If bIO = False Then
    
          Print #fFreeFile, "Last=" & fDocument(i).Tag
             
          bIO = True

        End If

        Print #fFreeFile, "Pkg" & i & "=" & fDocument(i).Tag
        
        ReDim Preserve LoadRunTag(intLRT + 1)
                 
        LoadRunTag(intLRT + 1) = strLine

        intLRT = (intLRT + 1)

       End If

      Next i

   Close #fFreeFile

Exit Function
ErrHld:

End Function

Private Sub mnuCut_Click()

  If ActiveForm Is Nothing Then
    Exit Sub
  End If

  Call Clipboard.Clear
  Call Clipboard.SetText(ActiveForm.rtfText.SelText)
  
  ActiveForm.rtfText.SelText = ""

End Sub

Private Sub mnuDbgStart_Click()
On Error GoTo ErrHld

  Dim strData As String
  
  boolCancel = False
  
  intDelay = 1
  
  strData = ""
  
  Do Until intDelay >= 5000
  
   DoEvents
  
    If boolCancel = True Then
      timStatus.Enabled = True
      Exit Sub
    End If
  
    sBar1.Panels(1).Text = "Delay before script debug, " & (intDelay / 1000) & " seconds, please wait..."
    
    intDelay = (intDelay + 10)
  
  Loop
  
  If boolCancel = True Then
   timStatus.Enabled = True
   Exit Sub
  End If

  ScrCtrl1.Language = strLanguage

  strData = strData & ""
  strData = strData & ActiveForm.rtfText.Text
  strData = strData & vbCrLf & vbCrLf

  ScrCtrl1.ExecuteStatement strData
  
  timStatus.Enabled = True
   
Exit Sub
ErrHld:

End Sub

Private Sub mnuDbgStop_Click()
On Error GoTo ErrHld

  boolCancel = True

  ScrCtrl1.Reset

Exit Sub
ErrHld:

End Sub

Private Sub mnuDecrypt_Click()
On Error GoTo ErrHld

  Dim i As Long
    
  Dim sInput As String

  frmEncode.Show 1, Me

  sInput = DecodeString2(ActiveForm.rtfText.Text, sPassword)

  ActiveForm.rtfText.Text = sInput

Exit Sub
ErrHld:

End Sub

Private Sub mnuEncrypt_Click()
On Error GoTo ErrHld

  Dim i As Long
    
  Dim sInput As String

  frmEncode.Show 1, Me

  sInput = EncodeString2(ActiveForm.rtfText.Text, sPassword)

  ActiveForm.rtfText.Text = sInput

Exit Sub
ErrHld:

End Sub

Private Sub mnuExit_Click()

  Unload Me

End Sub

Private Sub mnuFind_Click()

  frmFind.Show , mdiMain

End Sub

Private Sub mnuFNext_Click()

  Call FindNext(strFind$)

End Sub

Private Sub mnuFont_Click()
On Error GoTo ErrHld

  If ActiveForm Is Nothing Then
    Exit Sub
  End If
  
  With cDiag1
  
    .CancelError = True
    
    .FontName = GetSetting(App.Title, "Setting", "FontName", ActiveForm.Font.Name)
    .FontSize = GetSetting(App.Title, "Setting", "FontSize", ActiveForm.Font.Size)
    
    .FontUnderline = GetSetting(App.Title, "Setting", "FontUnderline", ActiveForm.Font.Underline)
    .FontStrikethru = GetSetting(App.Title, "Settting", "FontStrikethru", ActiveForm.Font.Strikethrough)
    
    .FontBold = GetSetting(App.Title, "Setting", "FontBold", ActiveForm.Font.Bold)
    .FontItalic = GetSetting(App.Title, "Setting", "FontItalic", ActiveForm.Font.Italic)

    .ShowFont
    
    ActiveForm.Font.Name = .FontName
    ActiveForm.Font.Size = .FontSize

    ActiveForm.Font.Bold = .FontBold
    ActiveForm.Font.Italic = .FontItalic
    
    ActiveForm.Font.Underline = .FontUnderline
    ActiveForm.Font.Strikethrough = .FontStrikethru
    
    SaveSetting App.Title, "Setting", "FontName", .FontBold
    SaveSetting App.Title, "Setting", "FontSize", .FontSize
        
    SaveSetting App.Title, "Setting", "FontUnderline", .FontUnderline
    SaveSetting App.Title, "Setting", "FontStrikethru", .FontStrikethru
  
    SaveSetting App.Title, "Setting", "FontBold", .FontBold
    SaveSetting App.Title, "Setting", "FontItalic", .FontItalic

  End With

Exit Sub
ErrHld:

End Sub

Private Sub mnuForward_Click()
On Error GoTo ErrHld

  If ActiveForm Is Nothing Then
    Exit Sub
  End If
  
  ActiveForm.wB1.GoForward

Exit Sub
ErrHld:

End Sub

Private Sub mnuHelp_Click()
On Error GoTo ErrHld

  Dim dblHelp As Double
  
  dblHelp = Shell("hh.exe " & App.Path & "\" & App.Title & ".chm", vbNormalFocus)

Exit Sub
ErrHld:

  Call MsgBox("No help file found!", vbExclamation, "Notice")

End Sub

Private Sub mnuInsert_Click()
On Error GoTo ErrHld

  With cDiag1
  
    .CancelError = True
    
    .DialogTitle = "Insert Picture..."
    
    .FileName = ""
    
    .DefaultExt = ""
    
    .Filter = "BMP files (*.bmp)|*.bmp|JPEG files (*.jpg)|*.jpg|GIF files (*.gif)|*.gif"
    
    .ShowOpen
    
    If Not Len(.FileName) = 0 Then

      ActiveForm.rtfText.OLEObjects.Add , "Obj" & objCount, .FileName

    End If

  End With

Exit Sub
ErrHld:

End Sub

Private Sub mnuLanguage_Click()
On Error GoTo ErrHld

  strLanguage = InputBox("Enter Language: ", "Notice", "VBScript")

  If IsNumeric(strLanguage) = True Then
  
    strLanguage = "VBScript"
    
  ElseIf Len(strLanguage) = 0 Then
  
    strLanguage = "VBScript"
  
  End If

Exit Sub
ErrHld:

  strLanguage = "VBScript"

End Sub

Private Sub mnuLoadViewer_Click()

  Unload frmRunView

  frmRunView.Show 1, Me

End Sub

Private Sub mnuLRun_Click()
On Error GoTo ErrHld

  If ActiveForm Is Nothing Then
    LoadNewDoc
  End If

  With cDiag1
  
    .CancelError = True
    
    .DialogTitle = "Load Run..."
    
    .DefaultExt = ""
    
    .Filter = "INI Files (*.ini)|*.ini"

    .ShowOpen
    
    If Len(.FileName) <> 0 Then
    
      Call LoadRun(.FileName)
    
    End If
  
  End With

Exit Sub
ErrHld:

End Sub

Private Sub mnuNew_Click()

   sBar1.Panels(1).Text = "New Document..."

   LoadNewDoc

   timStatus.Enabled = True

End Sub

Private Sub mnuNWin_Click()

  LoadNewWeb

End Sub

Private Sub mnuNWin1_Click()

  LoadNewDoc

End Sub

Private Sub mnuOpen_Click()
On Error GoTo ErrHld

   If ActiveForm Is Nothing Then

     LoadNewDoc

     Exit Sub
    
   End If
   
   sBar1.Panels(1).Text = "Open..."

   With cDiag1
   
     .CancelError = True
     
     .DialogTitle = "Open..."
     
     .Filter = "RTF files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|ENC files (*.enc)|*.enc|Bitmap files (*.bmp)|*.bmp|All files (*.*)|*.*"
     
     .FileName = ""
     
     .DefaultExt = ""
     
     .ShowOpen

     sBar1.Panels(1).Text = "Loading file..."
       
     If Not (Len(.FileName) = 0) Then

       LoadNewDoc

       If InStr(1, .FileName, ".txt", vbBinaryCompare) <> 0 Then

         Call ActiveForm.rtfText.LoadFile(.FileName, rtfText)
         
         DoEvents
         
         ActiveForm.Caption = Split(.FileName, "\", -1)(UBound(Split(.FileName, "\", -1)))
         ActiveForm.Tag = .FileName
         
       ElseIf InStr(1, .FileName, ".rtf", vbBinaryCompare) <> 0 Then

         Call ActiveForm.rtfText.LoadFile(.FileName, rtfRTF)
         
         DoEvents
         
         ActiveForm.Caption = Split(.FileName, "\", -1)(UBound(Split(.FileName, "\", -1)))
         ActiveForm.Tag = .FileName
         
       ElseIf InStr(1, .FileName, ".enc", vbBinaryCompare) <> 0 Then

         Call ActiveForm.rtfText.LoadFile(.FileName, rtfText)
         
         DoEvents

         ActiveForm.Caption = Split(.FileName, "\", -1)(UBound(Split(.FileName, "\", -1)))
         ActiveForm.Tag = .FileName
       
       ElseIf InStr(1, .FileName, ".bmp", vbBinaryCompare) <> 0 Then

         If ActiveForm Is Nothing Then
         
           LoadNewDoc
           
         End If

         ActiveForm.rtfText.OLEObjects.Add , "Obj" & objCount, .FileName
         
         objCount = (objCount + 1)
         
         DoEvents

       Else

         ActiveForm.rtfText.LoadFile .FileName
         
         DoEvents
         
         ActiveForm.Caption = Split(.FileName, "\", -1)(UBound(Split(.FileName, "\", -1)))
         ActiveForm.Tag = .FileName
         
       End If

     End If
   
   End With
     
   timStatus.Enabled = True
     
Exit Sub
ErrHld:

End Sub

Private Sub mnuPaste_Click()

 If ActiveForm Is Nothing Then
   Exit Sub
 End If
 
 'Call SendMessage(ActiveForm.rtfText.hwnd, WM_PASTE, 0, 0)
 
 ActiveForm.rtfText.SelText = Clipboard.GetText

End Sub

Private Sub mnuPause_Click()
On Error GoTo ErrHld

  oSpeak.Pause

Exit Sub
ErrHld:

End Sub

Private Sub mnuPrint_Click()
On Error GoTo ErrHld

  Dim i As Integer
    
  With cDiag1
  
    .CancelError = True
    
    .Filter = (cdlPDReturnDC + cdlPDNoPageNums + cdlPDAllPages)

    .ShowPrinter
        
    sBar1.Panels(1).Text = "Printing..."

    For i = 1 To Printer.Copies

      ActiveForm.rtfText.SelPrint Printer.hDC
    
    Next

    Printer.EndDoc
    
    timStatus.Enabled = True


  
  End With

Exit Sub
ErrHld:

End Sub

Private Sub mnuRefresh_Click()
On Error GoTo ErrHld

  If ActiveForm Is Nothing Then
    Exit Sub
  End If
  
  ActiveForm.wB1.Refresh

Exit Sub
ErrHld:

End Sub

Private Sub mnuResume_Click()
On Error GoTo ErrHld

 oSpeak.Resume
 
Exit Sub
ErrHld:

End Sub

Private Sub mnuSave_Click()
On Error GoTo ErrHld

 If ActiveForm Is Nothing Then
   Exit Sub
 End If
 
   If (Len(ActiveForm.Tag) = 0) Then
   
     sBar1.Panels(1).Text = "Save as..."
   
     mnuSaveAs_Click
     
     timStatus.Enabled = True
     
     Exit Sub
   
   End If
   
   sBar1.Panels(1).Text = "Save..."
   
   If InStr(1, ActiveForm.Caption, ".txt", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag

   ElseIf InStr(1, ActiveForm.Caption, ".rtf", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfRTF
     
     ActiveForm.Caption = ActiveForm.Tag
     
   ElseIf InStr(1, ActiveForm.Caption, ".enc", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   ElseIf InStr(1, ActiveForm.Caption, ".htm", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   ElseIf InStr(1, ActiveForm.Caption, ".html", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
   
   ElseIf InStr(1, ActiveForm.Caption, ".hta", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
        
   ElseIf InStr(1, ActiveForm.Caption, ".mta", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
   
   ElseIf InStr(1, ActiveForm.Caption, ".css", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   ElseIf InStr(1, ActiveForm.Caption, ".vbs", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
    ElseIf InStr(1, ActiveForm.Caption, ".js", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   Else

     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   End If

   timStatus.Enabled = True
Exit Sub
ErrHld:

End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo ErrHld

 If ActiveForm Is Nothing Then
   Exit Sub
 End If

 With cDiag1
 
   .CancelError = True
   
   .DialogTitle = "Save as..."
   
   .FileName = ""
   
   .Filter = "RTF files (*.rtf)|*.rtf|ENC files (*.enc)|*.enc|Text files (*.txt)|*.txt|All files (*.*)|*.*"
   
   .ShowSave
    
   sBar1.Panels(1).Text = "Save as..."

   If Not (Len(.FileName) = 0) Then

     If InStr(1, .FileName, ".txt", vbBinaryCompare) <> 0 Then

       ActiveForm.rtfText.SaveFile .FileName, rtfText

     ElseIf InStr(1, .FileName, ".rtf", vbBinaryCompare) <> 0 Then
     
       ActiveForm.rtfText.SaveFile .FileName, rtfRTF
       
     ElseIf InStr(1, .FileName, ".enc", vbBinaryCompare) <> 0 Then
     
       Dim i As Long
    
       Dim sInput As String

       frmEncode.Show 1, Me

       sInput = EncodeString2(ActiveForm.rtfText.Text, sPassword)

       ActiveForm.rtfText.Text = sInput
       
       ActiveForm.rtfText.SaveFile .FileName, rtfText
         
   ElseIf InStr(1, ActiveForm.Caption, ".htm", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   ElseIf InStr(1, ActiveForm.Caption, ".html", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
   
   ElseIf InStr(1, ActiveForm.Caption, ".hta", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
        
   ElseIf InStr(1, ActiveForm.Caption, ".mta", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
   
   ElseIf InStr(1, ActiveForm.Caption, ".css", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
   ElseIf InStr(1, ActiveForm.Caption, ".vbs", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
    ElseIf InStr(1, ActiveForm.Caption, ".js", vbBinaryCompare) <> 0 Then
      
     ActiveForm.rtfText.SaveFile ActiveForm.Tag, rtfText
     
     ActiveForm.Caption = ActiveForm.Tag
     
     Else
     
       ActiveForm.rtfText.SaveFile .FileName, rtfText

     End If

     ActiveForm.Caption = Split(.FileName, "\", -1)(UBound(Split(.FileName, "\", -1)))
     
     ActiveForm.Tag = .FileName

   End If
      
   timStatus.Enabled = True

 End With

Exit Sub
ErrHld:

End Sub

Private Sub mnuSpellChk_Click()
On Error GoTo ErrHld

  Call DictionaryLoad

Exit Sub
ErrHld:

End Sub

Private Sub mnuStart_Click()
'On Error GoTo ErrHld
  
  Set oSpeak = New SpVoice
  
  Set oSpeak.Voice = oSpeak.GetVoices().Item(GetSetting(App.Title, "Setting", "Voice", 0))

  sTmp = GetSetting(App.Title, "Setting", "Volume", 80)

  If IsNumeric(sTmp) = False Then
  
    oSVolume = 80
    
    SaveSetting App.Title, "Setting", "Volume", oSVolume
    
  Else
    oSVolume = CInt(sTmp)
  End If

  oSpeak.Volume = oSVolume
  
  DoEvents
  
  If ActiveForm.Tag = "WEBDOC" Then
    
    oSpeak.Speak ActiveForm.wB1.Document.body.outerText, SVSFlagsAsync
  
  Else
    
    oSpeak.Speak ActiveForm.rtfText.Text, SVSFlagsAsync
  
  End If

  DoEvents
  
  sBar1.Panels(1).Text = "Speak..."
  
  timStatus.Enabled = True

Exit Sub
ErrHld:

End Sub

Private Sub mnuStop_Click()
On Error GoTo ErrHld

  oSpeak.Pause
    
  DoEvents
  
  oSpeak.Speak "", SVSFlagsAsync
  
  Set oSpeak = Nothing
 
Exit Sub
ErrHld:
   
  Set oSpeak = Nothing
 
End Sub

Private Sub mnuTipOfTheDay_Click()
On Error GoTo ErrHld

  frmTip.Show
  
  frmTip.ZOrder 0

Exit Sub
ErrHld:

End Sub

Private Sub mnuTitleHoriztonal_Click()

  Me.Arrange 2

End Sub

Private Sub mnuTitleVertical_Click()

  Me.Arrange 1

End Sub

Private Sub mnuToolbar_Click()

  If tBar1.Visible = True Then
    tBar1.Visible = False
  ElseIf tBar1.Visible = False Then
    tBar1.Visible = True
  End If
  
  SaveSetting App.Title, "Setting", "Toolbar", tBar1.Visible
 
  mnuToolbar.Checked = IIf(tBar1.Visible, GetSetting(App.Title, "Setting", "Toolbar", True), GetSetting(App.Title, "Setting", "Toolbar", False))
  
End Sub

Private Sub mnuStatusBar_Click()

  If sBar1.Visible = False Then
    sBar1.Visible = True
  ElseIf sBar1.Visible = True Then
    sBar1.Visible = False
  End If
  
  SaveSetting App.Title, "Setting", "Statusbar", sBar1.Visible

  mnuStatusBar.Checked = IIf(sBar1.Visible, GetSetting(App.Title, "Setting", "Statusbar", True), GetSetting(App.Title, "Setting", "Statusbar", False))
  
End Sub

Private Sub mnuUndo_Click()
On Error GoTo ErrHld

  If ActiveForm Is Nothing Then
    Exit Sub
  End If

  Call SendMessage(ActiveForm.rtfText.hwnd, WM_UNDO, 0, 0)

Exit Sub
ErrHld:

End Sub

Private Sub mnuVoices_Click()

  Set oSpeak = New SpVoice

  SaveSetting App.Title, "Setting", "Voice", InputBox("Enter Voice (0-" & oSpeak.GetVoices().Count & "): ", "Notice", 0)

End Sub

Private Sub mnuVol_Click()

  SaveSetting App.Title, "Setting", "Volume", InputBox("Enter Volume (1-100): ", "Notice")

End Sub

Private Sub tBar1_ButtonClick(ByVal Button As ComctlLib.Button)

  Select Case Button.Key
  
    Case "New": LoadNewDoc
    Case "Open": mnuOpen_Click
    Case "Save": mnuSave_Click
    Case "Print": mnuPrint_Click
    Case "Cut": mnuCut_Click
    Case "Copy": mnuCopy_Click
    Case "Paste": mnuPaste_Click
    Case "Bold"
    
       ActiveForm.rtfText.SelBold = IIf(ActiveForm.rtfText.SelBold, False, True)

    Case "Italic"
    
       ActiveForm.rtfText.SelItalic = IIf(ActiveForm.rtfText.SelItalic, False, True)
       
    Case "Underline"
    
       ActiveForm.rtfText.SelUnderline = IIf(ActiveForm.rtfText.SelUnderline, False, True)
    
    Case "AlignLeft"
    
       ActiveForm.rtfText.SelAlignment = rtfLeft

    Case "AlignCenter"
    
       ActiveForm.rtfText.SelAlignment = rtfCenter

    Case "AlignRight"
    
       ActiveForm.rtfText.SelAlignment = rtfRight

    Case "Font"
    
       mnuFont_Click
       
    Case "Spell Check"

       mnuSpellChk_Click

    Case "Help"
    
       mnuHelp_Click

  End Select

End Sub

Private Sub timStatus_Timer()

  sBar1.Panels(1).Text = "Status"
  
  timStatus.Enabled = False
  
End Sub

Public Sub FindNext(ByVal strText As String)

  Where = InStr(LastWhere, ActiveForm.rtfText.Text, strText$, vbTextCompare)

  If Where Then
  
    ActiveForm.SetFocus
    ActiveForm.rtfText.SelStart = (Where - 1)
    ActiveForm.rtfText.SelLength = (Len(strText$) + 1)

    LastWhere = (Where + (Len(strText$) + 1))

  End If

End Sub

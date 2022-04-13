VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExpAndPub 
   Caption         =   "Export & Publish"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4140
   OleObjectBlob   =   "frmExpAndPub.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExpAndPub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
Dim txtFixed As Scripting.TextStream, Contents() As String
Set fso = New Scripting.FileSystemObject

  If Dir(Environ("USERPROFILE") & "\Documents\CustomTools-FixedExport.txt") <> "" Then
    Set txtFixed = fso.OpenTextFile(Environ("USERPROFILE") & "\Documents\CustomTools-FixedExport.txt", ForReading, True)
        Contents = Split(txtFixed.ReadAll, " --- ")
        If Contents(0) = "True" Then
            cbUseFixed = True
        Else
            cbUseFixed = False
        End If
        FixedPath = Contents(1)
    txtFixed.Close
  End If
  
  SettingsHelp.Caption = "Fixed Export exports another duplicate picture to a location you specify. Useful for having a designated batch folder for pics, you never need to search for folders in Photoshop. But you will need to remember to delete old files before running a new batch. Enter a location/specify to use or not use Fixed Export and click 'Save Settings' to remember these options next time you open Export & Publish."

  EPMultiPage.Value = 0
  If Documents.Count > 0 Then ProdCall = ProdFldr
  
    If ProdCall = "N/A" Or ProdCall = "" Then
        ExportPath = "S:\XI-Online\"
        PublishPath = "S:\XI-Online\"
        ExportPathCZ = "S:\XI-Online\"
    ElseIf ProdCall = "Not Found" Then
        ExportPath = fso.GetParentFolderName(fso.GetParentFolderName(fso.GetParentFolderName(ActiveDocument.FullFileName)))
        PublishPath = fso.GetParentFolderName(fso.GetParentFolderName(fso.GetParentFolderName(ActiveDocument.FullFileName)))
        ExportPathCZ = fso.GetParentFolderName(fso.GetParentFolderName(fso.GetParentFolderName(ActiveDocument.FullFileName)))
    Else
        ExportPath = ProdCall
        PublishPath = ProdCall
        ExportPathCZ = ProdCall
    End If
    
    With ExportDrop
        .AddItem "First"
        .AddItem "Last"
        .ListWidth = 48
    End With
    With ExportDropCZ
        .AddItem "First"
        .AddItem "Last"
        .ListWidth = 48
    End With
    With PublishDrop
        .AddItem "First"
        .AddItem "Last"
        .AddItem "Mugs"
        .AddItem "JCN"
        .AddItem "VSB"
        .AddItem "LVSB"
        .AddItem "Coolies"
        .ListWidth = 48
    End With

    If Documents.Count > 0 Then
        If ActiveDocument.Properties("Type", 1) = "Laser" Or ActiveDocument.Properties("Type", 1) = "DTG" Then
            optPNG = True
            ExportDrop = "Last"
            cbPublish = False
        ElseIf ActiveDocument.Properties("Template", 1) = "3RCO" Then
            optPNG = True
        ElseIf ActiveDocument.Properties("Template", 1) = "BDSC" Then
            cbExport = False
        ElseIf ActiveDocument.Properties("Template", 1) = "Coolie" Or ActiveDocument.Properties("Template", 1) = "Slim Coolie" Then
            PublishDrop = "Coolies"
        ElseIf ActiveDocument.Properties("Template", 1) = "Mug" Then
            PublishDrop = "Mugs"
        ElseIf ActiveDocument.Properties("Template", 1) = "JCN" Then
            PublishDrop = "JCN"
        ElseIf ActiveDocument.Properties("Template", 1) = "VSB" Then
            PublishDrop = "VSB"
        ElseIf ActiveDocument.Properties("Template", 1) = "LVSB" Then
            PublishDrop = "LVSB"
        End If
        
        ExportDropCZ = ExportDrop
        optJPEGCZ = optJPEG
        optPNGCZ = optPNG
    End If
    
    If InStr(1, ActiveDocument.Title, "CZ") <> 0 Then
        EPMultiPage.Value = 1
    End If
    
frmEPOpen = True
Set fso = Nothing
End Sub
Private Sub EPMultiPage_Change()
    If EPMultiPage.Value = 1 Then ExpAndPub.ShowCZText
End Sub
Private Sub btnHelp_Click()
If btnHelp = True Then
    If PublishDrop = "First" Then
        PDFHelp.Caption = "Publishes the first page of the document, using the document name as the file name."
        PDFHelp.Visible = True
    ElseIf PublishDrop = "Last" Then
        PDFHelp.Caption = "Publishes the last page of the document, using the document name as the file name."
        PDFHelp.Visible = True
    ElseIf PublishDrop = "Mugs" Then
        PDFHelp.Caption = "Publishes the document using the document name as the file name, then sets the 'SKU' layer to not printable and exports again with the '15M-' prefix."
        PDFHelp.Visible = True
    ElseIf PublishDrop = "JCN" Then
        PDFHelp.Caption = "Creates a folder with the document name in the location specified, then publishes each page with the respective '-UP' suffix into that folder."
        PDFHelp.Visible = True
    ElseIf PublishDrop = "VSB" Then
        PDFHelp.Caption = "Creates a folder with the document name in the location specified, then publishes each page (except 1) with the respective '-UP' suffix into that folder."
        PDFHelp.Visible = True
    ElseIf PublishDrop = "LVSB" Then
        PDFHelp.Caption = "Creates a folder with the document name in the location specified, then publishes each page with the respective '-UP' suffix into that folder."
        PDFHelp.Visible = True
    End If
Else
    PDFHelp.Visible = False
End If
End Sub
Private Sub cbExport_Click()
  If cbExport = True Then
    ExportPathBtn.Enabled = True
    optJPEG.Enabled = True: optPNG.Enabled = True
    ExportPath.Enabled = True: ExportPath.BackColor = &H80000005
    ExportDrop.Enabled = True: ExportDrop.BackColor = &H80000005
    PicSuffBox.Enabled = True: PicSuffBox.BackColor = &H80000005
  Else
    ExportPathBtn.Enabled = False
    optJPEG.Enabled = False: optPNG.Enabled = False
    ExportPath.Enabled = False: ExportPath.BackColor = &H80000016
    ExportDrop.Enabled = False: ExportDrop.BackColor = &H80000016
    PicSuffBox.Enabled = False: PicSuffBox.BackColor = &H80000016
  End If
End Sub
Private Sub cbPublish_Click()
  If cbPublish = True Then
    PublishPathBtn.Enabled = True
    PublishPath.Enabled = True: PublishPath.BackColor = &H80000005
    PublishDrop.Enabled = True: PublishDrop.BackColor = &H80000005
    PDFSuffBox.Enabled = True: PDFSuffBox.BackColor = &H80000005
  Else
    PublishPathBtn.Enabled = False
    PublishPath.Enabled = False: PublishPath.BackColor = &H80000016
    PublishDrop.Enabled = False: PublishDrop.BackColor = &H80000016
    PDFSuffBox.Enabled = False: PDFSuffBox.BackColor = &H80000016
  End If
End Sub
Private Sub btnOpenPicFldr_Click()
    If Right(btnOpenPicFldr.Caption, 8) = "[Custom]" Then
        shell "explorer.exe /root," & ExportPathCZ, vbNormalFocus
    Else
        shell "explorer.exe /root," & ExportPath, vbNormalFocus
    End If
End Sub
Private Sub btnOpenPubFldr_Click()
    shell "explorer.exe /root," & PublishPath, vbNormalFocus
End Sub
Private Sub ExportPathBtn_Click()
Dim CFPath As String
    CFPath = EPChooseFolder("Export")
    If CFPath <> "" Then ExportPath = CFPath
End Sub
Private Sub ExportPathBtnCZ_Click()
Dim CFPath As String
    CFPath = EPChooseFolder("CZ")
    If CFPath <> "" Then ExportPathCZ = CFPath
End Sub
Private Sub PublishPathBtn_Click()
Dim CFPath As String
    CFPath = EPChooseFolder("Publish")
    If CFPath <> "" Then PublishPath = CFPath
End Sub
Private Sub FixedPathBtn_Click()
Dim CFPath As String
    CFPath = EPChooseFolder("Fixed")
    If CFPath <> "" Then FixedPath = CFPath
End Sub
Private Sub SaveSettingsBtn_Click()
Dim txt As Scripting.TextStream, Contents() As String
Set fso = New Scripting.FileSystemObject

If FixedPath.Value = "" Then MsgBox "Enter a location": Exit Sub
If Dir(FixedPath.Value, vbDirectory) = "" Then MsgBox "Location Not Found": Exit Sub

Set txt = fso.OpenTextFile(Environ("USERPROFILE") & "\Documents\CustomTools-FixedExport.txt", ForWriting, True)
txt.Write cbUseFixed & " --- " & FixedPath.Value

MsgBox "Settings Saved"
End Sub
Private Sub ExportPath_Change()
    ExportPath.ControlTipText = ExportPath.Text
End Sub
Private Sub PublishPath_Change()
    PublishPath.ControlTipText = PublishPath.Text
End Sub
Private Sub ExportPathCZ_Change()
    ExportPathCZ.ControlTipText = ExportPathCZ.Text
End Sub
Private Sub btnRefreshCZ_Click()
    ExpAndPub.ShowCZText
End Sub
Private Sub btnAddSelCZ_Click()
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "No Selection"
    Else
        ExpAndPub.AddSelCZ
    End If
End Sub

Private Sub btnCurDocCZ_Click()
    If StartCheckCZ = True Then ExpAndPub.EPGoCZ False
End Sub
Private Sub btnAllDocCZ_Click()
    If StartCheckCZ = True Then ExpAndPub.EPGoCZ True
End Sub
Private Sub btnCurDoc_Click()
    If StartCheck = True Then ExpAndPub.EPGo False
End Sub
Private Sub btnAllDoc_Click()
    If StartCheck = True Then ExpAndPub.EPGo True
End Sub


Private Sub UserForm_Terminate()
    frmEPOpen = False
End Sub

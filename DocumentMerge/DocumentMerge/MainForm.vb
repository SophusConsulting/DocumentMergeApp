Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports MFilesAPI
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports Microsoft.Office.Interop

Public Class MainForm

    Const ONLY_DOC = False 'set to true to allow concatenation only of .DOC files (not .TXT etc.)
    Const TXT_SUFFIX = "MasterDoc.PDF"

    Dim word, fs, folderpath, outdocname, folder, outdoc, combo2
    Dim D1 = "D", f = "f", t = "t", dash = "-", b = "B", o = "o", a = "a", r = "r", d = "d", Sp = " ", A2 = "A", g = "g", e2 = "e", n = "n"

    'GetConfigData class
    Dim configData As New GetConfigData

    Dim strArg, strArg2, boardAgendaName As String
    Dim objVault = getVaultConectionObject()

    Sub ParseCommandLine()
        ProgressBar.PerformStep()
        Dim combo As String = ComboBox1.SelectedItem
        folderpath = configData.ServerURL & ComboBox1.SelectedItem.ToString
        ProgressBar.PerformStep()
        outdocname = folderpath + TXT_SUFFIX
        ProgressBar.PerformStep()
    End Sub

    Public Function logProcess(ByVal Process As String)
        Dim directory As Object
        directory = CreateObject("Scripting.FileSystemObject")
        Dim strFile As String
        Dim sw As StreamWriter

        Try
            If directory.FolderExists(configData.ServerURL & configData.ServerURLProcessLog) Then
                strFile = configData.ServerURL & configData.ServerURLProcessLog & DateTime.Today.ToString("dd-MM-yyyy") & ".txt"
            Else
                directory.CreateFolder(configData.ServerURL & configData.ServerURLProcessLog)
                strFile = configData.ServerURL & configData.ServerURLProcessLog & DateTime.Today.ToString("dd-MM-yyyy") & ".txt"
            End If

            If (Not File.Exists(strFile)) Then
                sw = File.CreateText(strFile)
                sw.WriteLine("Start Log Process for today ")
            Else
                sw = File.AppendText(strFile)
            End If

            sw.WriteLine("Process at -- " & DateTime.Now & " : " & Process)
            sw.Close()
        Catch ex2 As IOException
            MsgBox(configData.writingLogError)
        End Try
    End Function

    Private Function getPropertyIDByAlias(ByVal PropertyAlias As String) As Integer
        'Get the property id by alias
        Try
            Dim MatterPropertyAlias = objVault.PropertyDefOperations.GetPropertyDefIDByAlias(PropertyAlias)
            If MatterPropertyAlias = -1 Or MatterPropertyAlias = 0 Then
                Throw New System.Exception("Alias Not Found")
            Else
                Return MatterPropertyAlias
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString & ": " & PropertyAlias)
            logProcess(ex.Message.ToString & PropertyAlias)
            Return Nothing
        End Try
    End Function
    Private Function getVaultConectionObject() As Vault
        'CONNECTION TO VAULT
        Dim objMFClient As New MFilesAPI.MFilesClientApplication()
        Dim objVaultCons As MFilesAPI.VaultConnections = objMFClient.GetVaultConnections
        Dim properties As MFilesAPI.PropertyDef
        Dim objVault As MFilesAPI.Vault = Nothing
        For Each objVaultCon As MFilesAPI.VaultConnection In objVaultCons
            If objVaultCon.Name = configData.DocumentVault Then
                objVault = objVaultCon.BindToVault(Me.Handle, True, True)
                'MessageBox.Show("Connection to Vault Successfully")
                logProcess("Connection to Vault Successfully")
                Exit For
            End If
        Next
        Return objVault
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.FormBorderStyle = FormBorderStyle.Fixed3D


        Try
            strArg = Command().ToString()

            strArg = "Draft - Board Agenda - 9850: Test Matter (Sylvia Lara)"
            boardAgendaName = strArg
            strArg2 = strArg
            'strArg = strArg.Replace(":", "")
            logProcess("Board Agenda " & strArg)
            Dim word As String
            word = strArg
            Dim word2 As String = word.Replace("Draft - Board Agenda - ", "")
            Dim final() As String = word2.Split(":")

            strArg = final(0)

            ProgressBar.Visible = False
            ComboBox1.Hide()
            ToolTip1.SetToolTip(ComboBox1, configData.selectMatterError)
            Me.CenterToScreen()
            Dim oFSO, oFolder, oSubFolder, i
            oFSO = CreateObject("Scripting.FileSystemObject")
            oFolder = oFSO.GetFolder(configData.ServerURL)

            oSubFolder = oFolder.SubFolders

            For Each i In oSubFolder

                If (strArg = i.name) Then
                    ComboBox1.Items.Add(strArg)
                    ComboBox1.SelectedIndex = 0
                    Exit For
                End If
            Next

            If (ComboBox1.Items.Count = 0) Then
                logProcess(configData.boardAgendaNotReadyError)
                MessageBox.Show(configData.boardAgendaNotReadyError)
                Me.Close()
            Else
                Me.Show()
                Me.Focus()
                ProgressBar.Visible = True
                '____________________________________________________________________________
                'it starts the main proccess
                '____________________________________________________________________________

                If ComboBox1.SelectedItem = "" Then
                    MessageBox.Show(configData.selectMatterError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Close()
                Else
                    ProgressBar.PerformStep()
                    Dim Header As String = ComboBox1.SelectedItem.ToString
                    Dim combo2 = Header.TrimStart(D1, r, a, f, t, Sp, dash, Sp)
                    Dim outdocname3 = configData.ServerURL & ComboBox1.SelectedItem.ToString & "\Final -" & combo2 & ".docx"
                    Dim outdocname4 = configData.ServerURL & ComboBox1.SelectedItem.ToString & "\Final -" & combo2 & ".pdf"


                    Dim fso = CreateObject("Scripting.FileSystemObject")

                    If (fso.FileExists(outdocname3)) Then
                        fso.DeleteFile(outdocname3)
                        If (fso.FileExists(outdocname4)) Then
                            fso.DeleteFile(outdocname4)
                        End If
                    End If

                    ProgressBar.PerformStep()
                    Dim oFS = CreateObject("Scripting.FileSystemObject")
                    Dim oFolders = oFS.GetFolder(configData.ServerURL & ComboBox1.SelectedItem.ToString)
                    If (oFolders.Files.Count) < 2 Then
                        logProcess(configData.stateMergerWorkFlowError)
                        MessageBox.Show(configData.stateMergerWorkFlowError)
                        Me.Close()
                    End If
                    'Progress Bar
                    ProgressBar.Style = ProgressBarStyle.Continuous
                    ProgressBar.PerformStep()
                    Try

                        FirstCleanup()
                        ProgressBar.PerformStep()

                        Call ParseCommandLine()

                        StartServers()
                        ProgressBar.PerformStep()

                        Process()
                        ProgressBar.PerformStep()

                        'Cleanup()
                        ProgressBar.PerformStep()

                        ConvertToPDF()
                        ProgressBar.PerformStep()
                        createwindowfile()

                        Me.Close()

                    Catch ex As Exception
                        Dim strFile As String = configData.ServerURL & DateTime.Today.ToString("dd-MM-yyyy") & ".txt"
                        Dim sw As StreamWriter
                        Try
                            If (Not File.Exists(strFile)) Then
                                sw = File.CreateText(strFile)
                                sw.WriteLine("Start Error Log for today")
                            Else
                                sw = File.AppendText(strFile)
                            End If
                            sw.WriteLine("Error Message Occured at-- " & DateTime.Now & " : " & ex.ToString)
                            sw.Close()
                        Catch ex2 As IOException
                            MsgBox(configData.writingLogError)
                        End Try
                        MessageBox.Show(configData.seeLogPopError & vbCrLf & ex.Message)
                        Close()
                    End Try


                End If
            End If
        Catch ex As Exception
            Dim directory As Object
            directory = CreateObject("Scripting.FileSystemObject")
            Dim strFile As String
            Dim sw As StreamWriter
            Try

                If directory.FolderExists(configData.ServerURL & configData.ServerURLProcessLog) Then
                    strFile = configData.ServerURL & configData.ServerURLProcessLog & DateTime.Today.ToString("dd-MM-yyyy") & ".txt"
                Else
                    directory.CreateFolder(configData.ServerURL & configData.ServerURLProcessLog)
                    strFile = configData.ServerURL & configData.ServerURLProcessLog & DateTime.Today.ToString("dd-MM-yyyy") & ".txt"
                End If

                If (Not File.Exists(strFile)) Then
                    sw = File.CreateText(strFile)
                    sw.WriteLine("Start Error Log for today")
                Else
                    sw = File.AppendText(strFile)
                End If
                sw.WriteLine("Error Message Occured at-- " & DateTime.Now & " : " & ex.ToString)
                sw.Close()
            Catch ex2 As IOException
                MsgBox("Error writing to log file.")
            End Try
            MessageBox.Show("An Error ocurred while merging the document, see error log for more details")
            Close()
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    '—————————————-
    Sub StartServers()

        '– Start Word
        ProgressBar.PerformStep()
        word = CreateObject("Word.Application")
        'word.Visible = true
        ProgressBar.PerformStep()
        fs = CreateObject("Scripting.FileSystemObject")
        folder = fs.GetFolder(folderpath)
        ProgressBar.PerformStep()

        logProcess("Starting Servers")
    End Sub

    Private Sub createwindowfile()
        'CONNECTION TO VAULT
        Dim objMFClient As New MFilesAPI.MFilesClientApplication()
        Dim objVaultCons As MFilesAPI.VaultConnections = objMFClient.GetVaultConnections
        Dim properties As MFilesAPI.PropertyDef
        Dim objVault As MFilesAPI.Vault = Nothing
        For Each objVaultCon As MFilesAPI.VaultConnection In objVaultCons
            If objVaultCon.Name = configData.DocumentVault Then
                objVault = objVaultCon.BindToVault(Me.Handle, True, True)
                'MessageBox.Show("Connection to Vault Successfully")
                logProcess("Connection to Vault Successfully")
                Exit For
            End If
        Next
        ' Construct parameters for document card.
        Dim oObjectCreationInfo As New MFilesAPI.ObjectCreationInfo
        oObjectCreationInfo.SetObjectType(MFilesAPI.MFBuiltInObjectType.MFBuiltInObjectTypeDocument, False)
        Dim oSourceFiles As New MFilesAPI.SourceObjectFiles
        Dim oSourceFile As New MFilesAPI.SourceObjectFile
        Dim oSourceFiles2 As New MFilesAPI.SourceObjectFiles
        Dim oSourceFile2 As New MFilesAPI.SourceObjectFile
        Dim D1 = "D", f = "f", t = "t", dash = "-", b = "B", o = "o", a = "a", r = "r", d = "d", Sp = " ", A2 = "A", g = "g", e2 = "e", n = "n"

        Dim Header As String = ComboBox1.SelectedItem.ToString
        Dim combo2 = Header.TrimStart(D1, r, a, f, t, Sp, dash, Sp)
        Dim MatterDescrip = Header.TrimStart(D1, r, a, f, t, Sp, dash, Sp, b, o, a, r, d, Sp, A2, g, e2, n, d, a, Sp, dash, Sp)


        Dim originals As String = combo2
        Dim MatterDescription As String = MatterDescrip
        Dim Ind As Integer
        Ind = MatterDescription.IndexOf(" ")
        Dim cnt = MatterDescription
        cnt = cnt.Count()
        Dim combo3 = originals

        oSourceFiles2.AddFile("TestDocument", "docx", configData.ServerURL & ComboBox1.SelectedItem.ToString + "\Final -" & combo2 & ".docx")
        oObjectCreationInfo.SetSourceFiles(oSourceFiles2)
        oObjectCreationInfo.SetDisableObjectCreation(True)
        oObjectCreationInfo.SetSingleFileDocument(True, False)

        oSourceFiles.AddFile("TestDocument", "pdf", configData.ServerURL & ComboBox1.SelectedItem.ToString + "\Final -" & combo2 & ".pdf")
        oObjectCreationInfo.SetSourceFiles(oSourceFiles)
        oObjectCreationInfo.SetDisableObjectCreation(True)
        oObjectCreationInfo.SetSingleFileDocument(True, False)



        ' We are searching for objects from class "Matter" (ID is 89 in Sample Vault).
        Dim iClass, BoardAClass As Integer
        iClass = 4 ' In this case, this identifies the class "Matter".

        'Search for Class ID based on object type given
        iClass = MF_FindClassID(objVault, "Matter")
        BoardAClass = MF_FindClassID(objVault, configData.BAObject)

        Dim iproperty As Integer
        'Create a search condition for the object class.

        Dim oSearchCondition As New SearchConditions

        Dim SC1 As New SearchCondition
        SC1.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
        SC1.Expression.DataPropertyValuePropertyDef = MFilesAPI.MFBuiltInPropertyDef.MFBuiltInPropertyDefClass
        SC1.TypedValue.SetValue(MFilesAPI.MFDataType.MFDatatypeLookup, iClass)
        oSearchCondition.Add(-1, SC1)


        Dim SC2 As New SearchCondition
        SC2.ConditionType = MFConditionType.MFConditionTypeEqual
        SC2.Expression.DataPropertyValuePropertyDef = getPropertyIDByAlias("Matter Number")
        SC2.ConditionType = MFConditionType.MFConditionTypeContains
        SC2.TypedValue.SetValue(MFDataType.MFDatatypeText, MatterDescription)
        oSearchCondition.Add(-1, SC2)

        'Dim oSearchCondition As MFilesAPI.SearchCondition = New MFilesAPI.SearchCondition
        'oSearchCondition.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
        'oSearchCondition.Expression.DataPropertyValuePropertyDef = MFilesAPI.MFBuiltInPropertyDef.MFBuiltInPropertyDefClass
        'oSearchCondition.TypedValue.SetValue(MFilesAPI.MFDataType.MFDatatypeLookup, iClass)

        'Invoke the search operation.
        Dim oObjectVersions As MFilesAPI.ObjectSearchResults = objVault.ObjectSearchOperations.SearchForObjectsByConditions(oSearchCondition, MFSearchFlags.MFSearchFlagNone, False)

        'Dim oObjectVersions As MFilesAPI.ObjectSearchResults = objVault.ObjectSearchOperations.SearchForObjectsByCondition(oSearchCondition, False)

        'Removes the word Board Agenda for verification
        Dim Header1 As String = ComboBox1.SelectedItem.ToString
        Dim combo = Header1.TrimStart(D1, r, a, f, t, Sp, dash, Sp, b, o, a, r, d, Sp, A2, g, e2, n, d, a, Sp, dash, Sp)

        Dim original As String = combo
        Dim Index As Integer
        Index = original.IndexOf(" ")
        Dim count = original
        count = count.Count()
        combo = original

        'Simply process the search results.
        For Each oObjectVersion As MFilesAPI.ObjectVersion In oObjectVersions

            'Get the document class id
            Dim documentClassID = MF_FindClassID(objVault, "Documents")

            ' Resolve the object type.
            Dim oObjType As MFilesAPI.ObjType
            oObjType = objVault.ObjectTypeOperations.GetObjectType(oObjectVersion.ObjVer.Type)

            ' Output the result.
            Dim mattertitle As String
            mattertitle = oObjectVersion.Title

            Dim word As String
            word = mattertitle
            Dim word2 As String = word.Replace("Draft - Board Agenda - ", "")
            Dim final() As String = word2.Split(":")


            Dim mattertitlefinal = final(0)

            If (mattertitlefinal = combo) Then
                'MessageBox.Show("Title of " + oObjType.NameSingular + ": " + oObjectVersion.Title)
                ' Create property definitions
                Dim oPropertyValues As MFilesAPI.PropertyValues = New MFilesAPI.PropertyValues

                ' Add 'Name and Title' property by creating a new PropertyValue object.
                Dim oPropertyValue1 As MFilesAPI.PropertyValue = New MFilesAPI.PropertyValue
                oPropertyValue1.PropertyDef = MFilesAPI.MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle
                oPropertyValue1.TypedValue.SetValue(MFilesAPI.MFDataType.MFDatatypeText, "Final - Board Agenda -" & mattertitle)
                oPropertyValues.Add(0, oPropertyValue1)

                'Add 'Class' property
                Dim oPropertyValue2 As MFilesAPI.PropertyValue = New MFilesAPI.PropertyValue
                oPropertyValue2.PropertyDef = MFilesAPI.MFBuiltInPropertyDef.MFBuiltInPropertyDefClass
                oPropertyValue2.TypedValue.SetValue(MFilesAPI.MFDataType.MFDatatypeLookup, documentClassID)  ' is the Documents Class
                oPropertyValues.Add(1, oPropertyValue2)


                ' Add 'Matter' property
                Dim MatterPropertyAlias = objVault.PropertyDefOperations.GetPropertyDefIDByAlias("Matter")
                Dim oPropertyValue3 As MFilesAPI.PropertyValue = New MFilesAPI.PropertyValue
                oPropertyValue3.PropertyDef = MatterPropertyAlias
                oPropertyValue3.TypedValue.SetValue(MFilesAPI.MFDataType.MFDatatypeMultiSelectLookup, oObjectVersion.DisplayID) 'This is the Matter ID 
                oPropertyValues.Add(2, oPropertyValue3)
                'Show document card for the new customer object.
                Dim oObjectWinResult As MFilesAPI.ObjectWindowResult
                oObjectWinResult = objVault.ObjectOperations.ShowPrefilledNewObjectWindow(CUInt(Handle), MFilesAPI.MFObjectWindowMode.MFObjectWindowModeInsertSourceFiles, oObjectCreationInfo, oPropertyValues)

                'Check if the creation of the document was cancelled.
                If (oObjectWinResult.Result = MFilesAPI.MFObjectWindowResultCode.MFObjectWindowResultCodeCancel) Then
                    MsgBox(configData.mergeCanceled)
                    logProcess(configData.mergeCanceled)
                    Dim outdocname3 = configData.ServerURL & ComboBox1.SelectedItem.ToString + "\Final -" & combo2 & ".docx"
                    Dim outdocname4 = configData.ServerURL & ComboBox1.SelectedItem.ToString + "\Final -" & combo2 & ".pdf"
                    Dim fso = CreateObject("Scripting.FileSystemObject")

                    If (fso.FileExists(outdocname3)) Then
                        fso.DeleteFile(outdocname3)
                        'ProgressBar1.PerformStep()
                        If (fso.FileExists(outdocname4)) Then
                            fso.DeleteFile(outdocname4)
                        End If
                    End If


                Else

                    '------------------------------------
                    'Searches for Board Agenda Temp Object
                    Dim SC4 As New SearchCondition
                    Dim searchconditions As New SearchConditions
                    SC4.ConditionType = MFConditionType.MFConditionTypeEqual
                    SC4.Expression.DataPropertyValuePropertyDef = getPropertyIDByAlias("Board Agenda Name")
                    SC4.ConditionType = MFConditionType.MFConditionTypeContains
                    SC4.TypedValue.SetValue(MFDataType.MFDatatypeText, boardAgendaName)
                    searchconditions.Add(-1, SC4)

                    'Invoke the search operation.
                    Dim oObjectVersionsBA As MFilesAPI.ObjectSearchResults = objVault.ObjectSearchOperations.SearchForObjectsByConditions(searchconditions, MFSearchFlags.MFSearchFlagNone, False)

                    ' Create a new object.b b 
                    Dim oObjVerAndProps As MFilesAPI.ObjectVersionAndProperties
                    Dim oObjVerAndProps2 As MFilesAPI.ObjectVersionAndProperties
                    oObjVerAndProps = objVault.ObjectOperations.CreateNewObject(MFilesAPI.MFBuiltInObjectType.MFBuiltInObjectTypeDocument, oObjectWinResult.Properties, oSourceFiles)
                    oObjVerAndProps2 = objVault.ObjectOperations.CreateNewObject(MFilesAPI.MFBuiltInObjectType.MFBuiltInObjectTypeDocument, oObjectWinResult.Properties, oSourceFiles2)


                    ' Check the document object in.
                    objVault.ObjectOperations.CheckIn(oObjVerAndProps.ObjVer)
                    objVault.ObjectOperations.CheckIn(oObjVerAndProps2.ObjVer)

                    MessageBox.Show(configData.mergeSuccess)
                    logProcess(configData.mergeSuccess)

                    For Each SearchConditionBA As MFilesAPI.ObjectVersion In oObjectVersionsBA
                        If (strArg2 = SearchConditionBA.Title) Then

                            'Destroy the Draft Board Agenda Object
                            objVault.ObjectOperations.DestroyObject(SearchConditionBA.OriginalObjID, True, -1)

                            ''Gets the entire static route to the desktop, for C:\Users\user\Desktop\
                            'Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & ServerURL & ComboBox1.SelectedItem.ToString

                            ''Directory path
                            Dim path As String = configData.ServerURL & ComboBox1.SelectedItem.ToString

                            ''Detele the folder with all the documents and folders in it
                            System.IO.Directory.Delete(path, True)
                            Exit For
                        End If
                    Next

                End If
                Exit For
            End If
        Next

    End Sub
    ' Helper function to find the class ID by a class name.
    Function MF_FindClassID(
            ByRef oVault As MFilesAPI.Vault,
            ByVal szClassName As String) As Integer

        ' Set the search conditions for the value list item.
        Dim oScValueListItem As New MFilesAPI.SearchCondition
        oScValueListItem.Expression.SetValueListItemExpression(
        MFilesAPI.MFValueListItemPropertyDef.MFValueListItemPropertyDefName,
        MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone)
        oScValueListItem.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
        oScValueListItem.TypedValue.SetValue(MFilesAPI.MFDataType.MFDatatypeText, szClassName)
        Dim arrSearchConditions As New MFilesAPI.SearchConditions
        arrSearchConditions.Add(-1, oScValueListItem)

        ' Search for the value list item.
        Dim results As MFilesAPI.ValueListItemSearchResults
        results = oVault.ValueListItemOperations.SearchForValueListItemsEx(MFilesAPI.MFBuiltInValueList.MFBuiltInValueListClasses, arrSearchConditions)
        If results.Count > 0 Then
            ' Found.
            MF_FindClassID = results(1).ID
        Else
            ' Not found.
            MF_FindClassID = -1
        End If
    End Function

    Sub DeleteOldOutput()
        ProgressBar.PerformStep()
        If fs.FileExists(outdocname) Then

            fs.DeleteFile(outdocname)

        End If
        ProgressBar.PerformStep()
    End Sub


    Sub ConvertToPDF()
        Dim myfile
        Dim D1 = "D", f = "f", t = "t", dash = "-", b = "B", o = "o", a = "a", r = "r", d = "d", Sp = " ", A2 = "A", g = "g", e2 = "e", n = "n"

        Dim Header As String = ComboBox1.SelectedItem.ToString
        Dim combo2 = Header.TrimStart(D1, r, a, f, t, Sp, dash, Sp)
        ProgressBar.PerformStep()
        'If (op = 1) Then
        '    myfile = configData.ServerURL + "9850" & "\" & "1" & configData.AgendaItemDR
        'ElseIf (op = 2) Then
        '    myfile = configData.ServerURL + "9850" & "\" & "5" & configData.StaffArgumentDR
        'ElseIf (op = 3) Then
        myfile = configData.ServerURL + ComboBox1.SelectedItem.ToString & "\" & "Final -" & combo2 & ".docx"
        'End If


        Dim objDoc, objFile, objFSO, objWord, strFile, strHTML

        Const wdFormatPDF = 17

        ' Create a File System object
        objFSO = CreateObject("Scripting.FileSystemObject")

        ' Create a Word object
        objWord = CreateObject("Word.Application")
        ProgressBar.PerformStep()
        With objWord
            ' True: make Word visible; False: invisible
            .Visible = False

            ' Check if the Word document exists
            If objFSO.FileExists(myfile) Then
                objFile = objFSO.GetFile(myfile)
                strFile = objFile.Path
            Else

                logProcess(configData.fileOpenError)
                MessageBox.Show(configData.fileOpenError & vbCrLf)
                'Close Word
                .Quit

            End If
            ProgressBar.PerformStep()
            ' Build the fully qualified HTML file name
            strHTML = objFSO.BuildPath(objFile.ParentFolder,
                      objFSO.GetBaseName(objFile) & ".pdf")
            ProgressBar.PerformStep()

            ' Open the Word document
            .Documents.Open(strFile)

            ' Make the opened file the active document
            objDoc = .ActiveDocument

            ' Save as HTML
            objDoc.SaveAs(strHTML, wdFormatPDF)

            ' Close the active document
            objDoc.Close

            ' Close Word
            .Quit
        End With
        logProcess("Create PDF document " & "Final -" & combo2 & ".pdf")

    End Sub

    Sub ProcessFile(filename, insertBreak)
        Dim doc
        doc = word.Documents.Open(filename)
        'word.Visible = True
        word.Selection.WholeStory
        word.Selection.Copy
        outdoc.Activate
        ProgressBar.PerformStep()
        'If insertBreak Then word.Selection.InsertBreak(Type:=WdBreakType.wdPageBreak)

        word.Selection.PasteAndFormat(Type:=WdPasteOptions.wdKeepSourceFormatting)

        'word.Selection.Paste 'use this one so that it works for Word2000 too
        'word.Visible = False
        doc.Close
        ProgressBar.PerformStep()
        doc = Nothing

    End Sub

    Sub Process()
        Dim D1 = "D", f1 = "f", t = "t", dash = "-", b = "B", o = "o", a = "a", r = "r", d = "d", Sp = " ", A2 = "A", g = "g", e2 = "e", n = "n"

        Dim Header As String = ComboBox1.SelectedItem.ToString
        Dim combo2 = Header.TrimStart(D1, r, a, f1, t, Sp, dash, Sp)

        ProgressBar.PerformStep()
        DeleteOldOutput()

        Dim f

        ProgressBar.PerformStep()
        Dim wrdApp As New Word.Application
        Dim docNew As Word.Document
        docNew = wrdApp.Documents.Add
        Dim outdocname2 = configData.ServerURL & ComboBox1.SelectedItem.ToString & "\" & "Final -" & combo2 & ".docx"

        Dim count2 As Integer = 1
        Dim cnt As Integer = 0
        For Each s In folder.Files
            cnt = cnt + 1
        Next

        'Dim objectList As ArrayList = Nothing

        'For Each d As ArrayList In folder.Files

        'Next

        'If (objectList.Contains("PD Agenda Item Staff Argument (General)")) Then
        '    MessageBox.Show("It contains")
        'End If

        Dim objFSO = CreateObject("Scripting.FileSystemObject")

        For Each c In folder.Files
            If (c.name = "1_" & configData.AgendaItemGeneral) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemGeneral
            ElseIf (c.name = "1_" & configData.AgendaItemEff) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemEff
            ElseIf (c.name = "1_" & configData.AgendaItemDR) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemDR
            ElseIf (c.name = "1_" & configData.AgendaItemHaywood) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemHaywood
            ElseIf (c.name = "1_" & configData.AgendaItemMembership) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemMembership
            ElseIf (c.name = "1_" & configData.AgendaItemReeval) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemReeval
            ElseIf (c.name = "2_" & configData.CoverSheetA) Then
                objFSO.GetFile(c.path).Name = configData.CoverSheetA
            ElseIf (c.name = "3_" & configData.Placeholder1) Then
                objFSO.GetFile(c.path).Name = configData.Placeholder1
            ElseIf (c.name = "4_" & configData.CoverSheetB) Then
                objFSO.GetFile(c.path).Name = configData.CoverSheetB
            ElseIf (c.name = "5_" & configData.StaffArgumentGeneral) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentGeneral
            ElseIf (c.name = "5_" & configData.StaffArgumentEff) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentEff
            ElseIf (c.name = "5_" & configData.StaffArgumentDR) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentDR
            ElseIf (c.name = "5_" & configData.StaffArgumentHaywood) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentHaywood
            ElseIf (c.name = "5_" & configData.StaffArgumentMembership) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentMembership
            ElseIf (c.name = "5_" & configData.StaffArgumentReeval) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentReeval
            ElseIf (c.name = "6_" & configData.CoverSheetC) Then
                objFSO.GetFile(c.path).Name = configData.CoverSheetC
            ElseIf (c.name = "7_" & configData.Placeholder2) Then
                objFSO.GetFile(c.path).Name = configData.Placeholder2
            End If
        Next


        For Each c In folder.Files
            If (c.name = configData.AgendaItemGeneral) Then
                objFSO.GetFile(c.path).Name = "1_" & configData.AgendaItemGeneral
            ElseIf (c.name = configData.AgendaItemEff) Then
                objFSO.GetFile(c.path).Name = "1_" & configData.AgendaItemEff
            ElseIf (c.name = configData.AgendaItemDR) Then
                objFSO.GetFile(c.path).Name = "1_" & configData.AgendaItemDR
            ElseIf (c.name = configData.AgendaItemHaywood) Then
                objFSO.GetFile(c.path).Name = "1_" & configData.AgendaItemHaywood
            ElseIf (c.name = configData.AgendaItemMembership) Then
                objFSO.GetFile(c.path).Name = "1_" & configData.AgendaItemMembership
            ElseIf (c.name = configData.AgendaItemReeval) Then
                objFSO.GetFile(c.path).Name = "1_" & configData.AgendaItemReeval
            ElseIf (c.name = configData.CoverSheetA) Then
                objFSO.GetFile(c.path).Name = "2_" & configData.CoverSheetA
            ElseIf (c.name = configData.Placeholder1) Then
                objFSO.GetFile(c.path).Name = "3_" & configData.Placeholder1
            ElseIf (c.name = configData.CoverSheetB) Then
                objFSO.GetFile(c.path).Name = "4_" & configData.CoverSheetB
            ElseIf (c.name = configData.StaffArgumentGeneral) Then
                objFSO.GetFile(c.path).Name = "5_" & configData.StaffArgumentGeneral
            ElseIf (c.name = configData.StaffArgumentEff) Then
                objFSO.GetFile(c.path).Name = "5_" & configData.StaffArgumentEff
            ElseIf (c.name = configData.StaffArgumentDR) Then
                objFSO.GetFile(c.path).Name = "5_" & configData.StaffArgumentDR
            ElseIf (c.name = configData.StaffArgumentHaywood) Then
                objFSO.GetFile(c.path).Name = "5_" & configData.StaffArgumentHaywood
            ElseIf (c.name = configData.StaffArgumentMembership) Then
                objFSO.GetFile(c.path).Name = "5_" & configData.StaffArgumentMembership
            ElseIf (c.name = configData.StaffArgumentReeval) Then
                objFSO.GetFile(c.path).Name = "5_" & configData.StaffArgumentReeval
            ElseIf (c.name = configData.CoverSheetC) Then
                objFSO.GetFile(c.path).Name = "6_" & configData.CoverSheetC
            ElseIf (c.name = configData.Placeholder2) Then
                objFSO.GetFile(c.path).Name = "7_" & configData.Placeholder2
            End If
        Next
        Dim first = True

        For Each f In folder.Files
            Dim test = f.path

            ' objectList.Add(f.name)

            Dim myRange As Range
            If (Not ONLY_DOC) Then
                If first Then
                    outdoc = word.Documents.Add
                    outdoc.SaveAs(outdocname2)

                    'ConvertToPDF(1)
                    ProcessFile(f.path, False)

                    With word.ActiveDocument.Sections(count2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        End With
                    End With
                    first = False
                    count2 = count2 + 1

                ElseIf (count2 = 2) Then

                    'word.Visible = True
                    'word.Paragraphs(word.Paragraphs.Count).Range
                    'word.Selection.MoveEnd

                    word.Selection.InsertBreak(Type:=WdBreakType.wdSectionBreakNextPage)
                    ProcessFile(f.path, True)


                    With word.ActiveDocument.Sections(2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            .LinkToPrevious = False
                        End With
                    End With

                    With word.ActiveDocument.Sections(2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            .Range.Text = String.Empty
                        End With
                    End With
                    count2 = count2 + 1

                    'word.Visible = True
                    'With word.ActiveDocument.Sections(2)
                    '    word.Selection.Range.InsertAlignmentTab(WdVerticalAlignment.wdAlignVerticalCenter)
                    'End With

                    'word.Selection.MoveLeft(nit:=WdUnits.wdCharacter, Count:=10, Extend:=1)

                    With word.ActiveDocument.Sections(2)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment a")
                        End With

                    End With

                    With word.ActiveDocument.Sections(2)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment a")
                        End With

                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter

                        End With
                    End With

                    With word.Selection.Font
                        .name = "Arial"
                        .Size = 12
                        .Color = WdColor.wdColorBlack
                        .Bold = True
                    End With

                    With word.ActiveDocument.Sections(2)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="The proposed decision")
                        End With
                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                        End With

                        With word.Selection.Font
                            .name = "Arial"
                            .Size = 12
                            .Color = WdColor.wdColorBlack
                            .Bold = True
                        End With
                    End With

                ElseIf (count2 > 2 And count2 < 4) Then

                    'Set the cursor at the end of file
                    word.Selection.MoveEnd
                    word.Selection.EndKey(Unit:=WdUnits.wdStory)
                    'Insert Page break
                    word.Selection.InsertBreak(Type:=WdBreakType.wdSectionBreakNextPage)
                    ProcessFile(f.path, True)
                    count2 = count2 + 1
                    'word.Visible = True

                ElseIf (count2 = 4) Then

                    'Set the cursor at the end of file
                    word.Selection.MoveEnd
                    'Insert Page break
                    word.Selection.InsertBreak(Type:=WdBreakType.wdSectionBreakNextPage)
                    ProcessFile(f.path, True)
                    count2 = count2 + 1
                    'word.Visible = True

                    With word.ActiveDocument.Sections(4)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment b")
                        End With

                    End With

                    With word.ActiveDocument.Sections(4)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment b")
                        End With

                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter

                        End With
                    End With

                    With word.Selection.Font
                        .name = "Arial"
                        .Size = 12
                        .Color = WdColor.wdColorBlack
                        .Bold = True
                    End With

                    With word.ActiveDocument.Sections(4)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="STAFF’S ARGUMENT")
                        End With
                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                        End With

                        With word.Selection.Font
                            .name = "Arial"
                            .Size = 12
                            .Color = WdColor.wdColorBlack
                            .Bold = True
                        End With
                    End With

                ElseIf (count2 = 5) Then
                    'ConvertToPDF(2)
                    'Set the cursor at the end of file
                    word.Selection.MoveEnd
                    word.Selection.EndKey(Unit:=WdUnits.wdStory)
                    'Insert Page break
                    word.Selection.InsertBreak(Type:=WdBreakType.wdSectionBreakNextPage)
                    'Insert second file
                    ProcessFile(f.path, True)

                    With word.ActiveDocument.Sections(count2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            .PageNumbers.RestartNumberingAtSection = True
                            .PageNumbers.StartingNumber = 1
                            'Dim test4 = WdInformation.wdNumberOfPagesInDocument
                        End With
                    End With

                    count2 = count2 + 1

                ElseIf (count2 = 6) Then
                    'Set the cursor at the end of file
                    word.Selection.MoveEnd
                    'Insert Page break
                    word.Selection.InsertBreak(Type:=WdBreakType.wdSectionBreakNextPage)
                    ProcessFile(f.path, True)

                    With word.ActiveDocument.Sections(count2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            .LinkToPrevious = False
                        End With
                        With word.ActiveDocument.Sections(count2)
                            With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                                .Range.Text = String.Empty
                                .Range.InsertAfter(".")
                            End With
                        End With
                    End With
                    count2 = count2 + 1

                    With word.ActiveDocument.Sections(2)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment c")
                        End With

                    End With

                    With word.ActiveDocument.Sections(2)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment c")
                        End With

                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment c")
                        End With

                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="Attachment c")
                        End With

                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter

                        End With
                    End With

                    With word.Selection.Font
                        .name = "Arial"
                        .Size = 12
                        .Color = WdColor.wdColorBlack
                        .Bold = True
                    End With

                    With word.ActiveDocument.Sections(2)
                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="RESPONDENT(S) ARGUMENT(S)")
                        End With

                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                        End With

                        With word.Selection.Font
                            .name = "Arial"
                            .Size = 12
                            .Color = WdColor.wdColorBlack
                            .Bold = True
                        End With

                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="(S)")
                        End With

                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                        End With

                        With word.Selection.Font
                            .name = "Arial"
                            .Size = 12
                            .Color = WdColor.wdColorRed
                            .Bold = True
                        End With

                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="(S)")
                        End With

                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                        End With

                        With word.Selection.Font
                            .name = "Arial"
                            .Size = 12
                            .Color = WdColor.wdColorRed
                            .Bold = True
                        End With


                        With word.Selection.Find
                            .Forward = True
                            .ClearFormatting()
                            .MatchWholeWord = True
                            .MatchCase = False
                            .Wrap = WdFindWrap.wdFindContinue
                            .Execute(FindText:="(None Submitted)")
                        End With
                        With word.Selection
                            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                        End With

                        With word.Selection.Font
                            .name = "Arial"
                            .Size = 12
                            .Color = WdColor.wdColorRed
                            .Bold = True
                        End With
                    End With
                ElseIf (count2 = 7) Then
                    'Set the cursor at the end of file
                    word.Selection.MoveEnd
                    word.Selection.EndKey(Unit:=WdUnits.wdStory)
                    'Insert Page break
                    word.Selection.InsertBreak(Type:=WdBreakType.wdSectionBreakNextPage)

                    'Insert second file
                    ProcessFile(f.path, True)
                    word.Selection.WholeStory
                    word.ActiveDocument.Fields.Unlink


                    With word.ActiveDocument.Sections(count2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            .LinkToPrevious = False
                        End With
                    End With

                    With word.ActiveDocument.Sections(count2)
                        With .Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                            .Range.Text = String.Empty
                            .Range.InsertAfter(".")

                        End With
                    End With

                    word.Visible = True



                    'word.Selection.InsertAlignmentTab(WdAlignmentTabAlignment.wdCenter)


                    'word.Selection.Alignment = WdAlignmentTabAlignment.wdCenter

                    'With word.ActiveDocument.Sections(count2)
                    '    .Name = "Arial"
                    '    .Alignment = WdAlignmentTabAlignment.wdCenter
                    'End With


                    'word.Visible = True
                    word.Selection.WholeStory
                    With word.Selection.Font
                        .name = "Arial"
                    End With
                    'word.Visible = True
                    'save the document
                    outdoc.Close(WdSaveOptions.wdSaveChanges, False, False)
                    word.Quit()
                    outdoc = Nothing
                    word = Nothing
                    count2 = 0
                ElseIf (count2 = 0) Then
                    Exit For
                End If
            End If


        Next
        ProgressBar.PerformStep()


        For Each c In folder.Files
            If (c.name = "1_" & configData.AgendaItemGeneral) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemGeneral
            ElseIf (c.name = "1_" & configData.AgendaItemEff) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemEff
            ElseIf (c.name = "1_" & configData.AgendaItemDR) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemDR
            ElseIf (c.name = "1_" & configData.AgendaItemHaywood) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemHaywood
            ElseIf (c.name = "1_" & configData.AgendaItemMembership) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemMembership
            ElseIf (c.name = "1_" & configData.AgendaItemReeval) Then
                objFSO.GetFile(c.path).Name = configData.AgendaItemReeval
            ElseIf (c.name = "2_" & configData.CoverSheetA) Then
                objFSO.GetFile(c.path).Name = configData.CoverSheetA
            ElseIf (c.name = "3_" & configData.Placeholder1) Then
                objFSO.GetFile(c.path).Name = configData.Placeholder1
            ElseIf (c.name = "4_" & configData.CoverSheetB) Then
                objFSO.GetFile(c.path).Name = configData.CoverSheetB
            ElseIf (c.name = "5_" & configData.StaffArgumentGeneral) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentGeneral
            ElseIf (c.name = "5_" & configData.StaffArgumentEff) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentEff
            ElseIf (c.name = "5_" & configData.StaffArgumentDR) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentDR
            ElseIf (c.name = "5_" & configData.StaffArgumentHaywood) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentHaywood
            ElseIf (c.name = "5_" & configData.StaffArgumentMembership) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentMembership
            ElseIf (c.name = "5_" & configData.StaffArgumentReeval) Then
                objFSO.GetFile(c.path).Name = configData.StaffArgumentReeval
            ElseIf (c.name = "6_" & configData.CoverSheetC) Then
                objFSO.GetFile(c.path).Name = configData.CoverSheetC
            ElseIf (c.name = "7_" & configData.Placeholder2) Then
                objFSO.GetFile(c.path).Name = configData.Placeholder2
            End If
        Next

    End Sub

    'Sub Cleanup()
    '    ProgressBar.PerformStep()
    '    outdoc = Nothing
    '    word = Nothing
    '    folder = Nothing
    '    fs = Nothing
    '    ProgressBar.PerformStep()

    '    logProcess("Second cleandup")
    'End Sub

    Sub FirstCleanup()
        ProgressBar.PerformStep()
        outdoc = Nothing
        'word.Close()
        'word.Application.Quit(False)
        word = Nothing
        logProcess("First Cleanup")
    End Sub

    Sub Main(ByVal cmdArgs() As String)
        Console.WriteLine("main process for exe with argurments")

    End Sub

End Class




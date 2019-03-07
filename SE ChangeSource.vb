Imports System.IO
Imports System.Runtime.InteropServices


Module Module1
    'TextBox1: 5 PLCS Selected item from ListBox1 (Original model linked to .dft)
    '	orig: getSelected
    '	UI: addNewAsmtoDft
    'TextBox3: 3 PLCS P/N entered by user
    '	orig: getSelected
    '	UI: folderChoose


    'TextBox6: 2 PLCS User-chosen folder Changed to ClassGV.ChosenFolder
    '	orig: folderChoose
    '	UI: getSelectedPartFromAsm
    'TextBox2: 1 PLC Path of original linked model  CO
    'TextBox4: 1 PLC Full name of new link  CO
    'TextBox5: 2 PLCS Full name of active document Changed to FullNameActiveDoc variable
    '	orig: folderChoose
    'TextBox7: 1 PLC Full name of new .dft  CO
    'SUBS:
    '	Main
    '	GetSelected
    '	CreateNewModelFile
    '	folderChoose
    '	GetAsmInfo
    '	GetSelectedPartFromAsm
    '	addNewAsmtoDft
    '
    '
    Public OrigModLinkedToDft As String

    Public Sub Main() 'Populates ListBox1
        Dim objApp As SolidEdgeFramework.Application
        Dim objDoc As SolidEdgeDraft.DraftDocument
        Dim objModelLinks As SolidEdgeDraft.ModelLinks
        Dim objModelLink As SolidEdgeDraft.ModelLink
        Dim objModelLinkApp As SolidEdgeFramework.Application
        Dim Int As Integer

        objApp = GetObject(, "SolidEdge.Application")
        objDoc = objApp.ActiveDocument
        objModelLinks = objDoc.ModelLinks

        Int = objModelLinks.Count

        For Each item In objModelLinks
            ClassGV.Nm = item.filename
            ClassGV.IndRef = item.IndexReference
            Form1.ListBox1.Items.Add(ClassGV.Nm)
            'Form1.ListBox1.Items.Add(ClassGV.IndRef & objApp.ActiveDocument.path)
        Next

        Form1.ListBox1.Items.Add(Int & " Link(s)")

        ' Release objects
        objApp = Nothing
        objDoc = Nothing
        objModelLinks = Nothing
        objModelLink = Nothing
        objModelLinkApp = Nothing
    End Sub



    Public Sub GetSelected()
        Dim newfilename As String, FileExist As String, FileExist2 As String
        Dim objModelLink As SolidEdgeDraft.ModelLink
        Dim objApp As SolidEdgeFramework.Application
        Dim objDoc As SolidEdgeDraft.DraftDocument
        Dim objModelLinks As SolidEdgeDraft.ModelLinks
        Dim varExt As String
        Dim oDocs As SolidEdgeFramework.Documents
        Dim odoc As SolidEdgeFramework.SolidEdgeDocument
        Dim objAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing

        objApp = GetObject(, "SolidEdge.Application")
        
        Form1.Cursor = Cursors.WaitCursor

        objDoc = objApp.ActiveDocument

        If Form1.ListBox1.SelectedIndex >= 0 Then
            'Form1.TextBox1.Text = Form1.ListBox1.SelectedItem.ToString() 'Original linked model - full name w.suffix - check file type
            OrigModLinkedToDft = Form1.ListBox1.SelectedItem.ToString()
            varExt = System.IO.Path.GetExtension(Form1.ListBox1.SelectedItem.ToString()) '
        Else
            MsgBox("Nothing selected. Make selection and try again.")
            Exit Sub
        End If

        'Path = System.IO.Path.GetDirectoryName(Form1.TextBox1.Text) 'Just the path

        newfilename = Form1.TextBox3.Text 'Desired PART NUMBER for new model file

        'Build new file name
        ClassGV.newlink = objApp.ActiveDocument.path & "\" & newfilename & varExt 'Create full name with path - need to add appropriate suffix
        'Form1.TextBox4.Text = ClassGV.newlink 'Write full new name to textbox4

        objModelLinks = objDoc.ModelLinks
        objModelLink = objModelLinks.Item(Index:=ClassGV.IndRef)

        ClassGV.IndRef = objModelLink.IndexReference
        oDocs = objApp.Documents

        FileExist2 = (If(IO.File.Exists(OrigModLinkedToDft), "YesFile", "NoFile")) 'Check to see if it exists

        Try
            If FileExist2 = "YesFile" Then
                File.Copy(OrigModLinkedToDft, ClassGV.newlink) 'save selected file as newlink
            Else
                MsgBox("File " & OrigModLinkedToDft & " does not exist")
                Exit Sub
            End If
        Catch ex As Exception
            ' Display the message.
            Console.WriteLine(ex.Message)
        End Try

        'IF LINK IS AN ASSEMBLY
        If varExt = ".asm" Then 'if .asm, open, but don't link yet
            odoc = oDocs.Open(ClassGV.newlink)
            Call GetAsmInfo()
        Else
            For Each objModelLink In objModelLinks
                If objModelLink.FileName = OrigModLinkedToDft Then
                    objModelLink.ChangeSource(ClassGV.newlink)
                    Console.WriteLine(objModelLink.FileName) 'SHOULD REPORT NEW LINK NAME
                    CreateObject("WScript.Shell").Popup("Link Updated", 2, "Status")
                    Application.Exit()
                End If
            Next
        End If

        'get the right part to replace

        FileExist = (If(IO.File.Exists(ClassGV.newlink), "YesFile", "NoFile")) 'Check to see if it exists

        Try
            If FileExist = "YesFile" Then
                'MsgBox("New Model file " & ClassGV.newlink & " exists")
            Else
                MsgBox("File " & ClassGV.newlink & " does not exist")
            End If

        Catch ex As Exception
            ' Display the message.
            Console.WriteLine(ex.Message)
        End Try

        objApp.Visible = True
        Form1.Cursor = Cursors.Default


    End Sub




    ' Bring up a dialog to chose a folder path
    'build new file name from new p/n, selected path, and add ".dft"
    'check to see if it exists
    'If Not, Copy active .dft file To New .dft fil And open
    Public Sub folderChoose()
        Dim objApp As SolidEdgeFramework.Application
        Dim folderBrowserDialog1 As New FolderBrowserDialog
        Dim newPN As String, folderName As String
        Dim objDoc As SolidEdgeDraft.DraftDocument
        Dim oDocs As SolidEdgeFramework.Documents
        Dim fullNameActiveDoc As String, fileexist As String
        Dim varDocType As String

        objApp = GetObject(, "SolidEdge.Application")
        Try
            Form1.Cursor = Cursors.WaitCursor

            If objApp.Documents.Count > 0 Then
                varDocType = objApp.ActiveDocumentType.ToString
                If varDocType <> "igDraftDocument" Then
                    MsgBox(varDocType)
                    Exit Sub
                End If
            ElseIf objApp.Documents.Count = 0 Then
                MsgBox("No document open")
                Exit Sub
            End If
            objDoc = objApp.ActiveDocument
        Catch ex As Exception
            MsgBox(ex.Message & "   Make sure a .dft is open.")
        End Try

        oDocs = objApp.Documents

        If ClassGV.PathOK = "Change" Then
            folderBrowserDialog1.SelectedPath = My.Settings.LastPath
            ' Show the FolderBrowserDialog.
            Dim result As DialogResult = folderBrowserDialog1.ShowDialog()

            If (result = DialogResult.OK) Then
                My.Settings.LastPath = folderBrowserDialog1.SelectedPath
                Form1.Cursor = Cursors.Default
                Exit Sub
            End If
        End If

        newPN = Form1.TextBox3.Text
            folderName = My.Settings.LastPath
            Form1.TextBox6.Text = folderName
            ClassGV.ChosenFolder = folderName 'was Form1.TextBox6.Text
        ClassGV.fullnewname = folderName & "\" & newPN & ".dft"
        fullNameActiveDoc = objApp.ActiveDocument.path & "\" & objApp.ActiveDocument.name  'was Form1.TextBox5.Text
        fileexist = (If(IO.File.Exists(ClassGV.fullnewname), "YesFile", "NoFile")) 'Check to see if it exists
            Try
                If fileexist = "YesFile" Then
                    MsgBox("File already exists")
                    Exit Sub
                Else
                    'MsgBox("File does not exist")
                    System.IO.File.Copy(fullNameActiveDoc, ClassGV.fullnewname) 'copies active .dft to new fullnewname
                End If
            Catch ex As Exception
                ' Display the message.
                Console.WriteLine(ex.Message)
            End Try
            ClassGV.fullnewname = ClassGV.fullnewname.Replace(":", ":\")
            'oDocs.Close() - be more specific. Close the old .dft document.
            objApp.DoIdle()
            objDoc = oDocs.Open(ClassGV.fullnewname)

        objApp.Visible = True
        Form1.Cursor = Cursors.Default

        Call Main()
    End Sub


    Public Sub GetAsmInfo() 'Populate ListBox2
        Dim objApp As SolidEdgeFramework.Application
        Dim objDoc As SolidEdgeAssembly.AssemblyDocument
        Dim objOccurrences As SolidEdgeAssembly.Occurrences
        Dim varDocType As String

        objApp = GetObject(, "SolidEdge.Application")

        If objApp.Documents.Count > 0 Then
            varDocType = objApp.ActiveDocumentType.ToString
            If varDocType <> "igAssemblyDocument" Then
                MsgBox(varDocType)
                Exit Sub
            End If
        ElseIf objApp.Documents.Count = 0 Then
            MsgBox("No document open")
            Exit Sub
        End If

        objDoc = objApp.ActiveDocument
        objOccurrences = objDoc.Occurrences

        For Each Item In objOccurrences
            Form1.ListBox2.Items.Add(Item.partfilename)
        Next
    End Sub



    Public Sub GetSelectedPartFromAsm()
        Dim objApp As SolidEdgeFramework.Application
        Dim objDoc As SolidEdgeAssembly.AssemblyDocument
        Dim objDftDoc As SolidEdgeDraft.DraftDocument
        Dim oDocs As SolidEdgeFramework.Documents
        Dim objOccurrences As SolidEdgeAssembly.Occurrences

        Dim SelectedPartToReplace As String
        Dim newModelForAsm As String
        Dim varext As String

        objApp = GetObject(, "SolidEdge.Application")
        
        Form1.Cursor = Cursors.WaitCursor

        objDoc = objApp.ActiveDocument
        oDocs = objApp.Documents
        '      objModelLinks = objDoc.ModelLinks

        objOccurrences = objDoc.Occurrences
        If Form1.ListBox2.SelectedIndex >= 0 Then
            SelectedPartToReplace = Form1.ListBox2.SelectedItem.ToString() 'Original linked model
            ' MsgBox(SelectedPartToReplace)
        Else
            MsgBox("Nothing selected. Make selection and try again.")
            Exit Sub
        End If
        For Each item In objOccurrences
            If item.partfilename = SelectedPartToReplace Then
                varext = System.IO.Path.GetExtension(SelectedPartToReplace) '
                newModelForAsm = (ClassGV.ChosenFolder & "\" & Form1.TextBox3.Text & varext) 'ClassGV.ChosenFolder was Form1.TextBox6.Text
                'MsgBox(newModelForAsm)
                System.IO.File.Copy(SelectedPartToReplace, newModelForAsm) 'Create copy of model to be used in .asm
                objApp.DoIdle()
                item.Replace(NewOccurrenceFileName:=newModelForAsm, ReplaceAll:=False)
                objApp.DoIdle()
                objDoc.Save() 'save new .asm
                objDoc.Close() 'close new .asm
                objDftDoc = oDocs.Item(ClassGV.fullnewname).Activate 'activates new .dft file
                objApp.DoIdle()

                'now in .dft env

                Call addNewAsmtoDft()
                Exit Sub
            End If

        Next
    End Sub
	
	
	
    Public Sub addNewAsmtoDft()
        Dim objModelLink As SolidEdgeDraft.ModelLink
        Dim objApp As SolidEdgeFramework.Application
        Dim objDoc As SolidEdgeDraft.DraftDocument
        Dim objModelLinks As SolidEdgeDraft.ModelLinks
        Dim oDocs As SolidEdgeFramework.Documents
        Dim objAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing

        objApp = GetObject(, "SolidEdge.Application")
        oDocs = objApp.Documents
        objDoc = oDocs.Item(ClassGV.fullnewname).Activate 'activates new .dft file

        objDoc = objApp.ActiveDocument
        objModelLinks = objDoc.ModelLinks
        objModelLink = objModelLinks.Item(Index:=ClassGV.IndRef)

        ClassGV.IndRef = objModelLink.IndexReference

        For Each objModelLink In objModelLinks
            If objModelLink.FileName = OrigModLinkedToDft Then
                objModelLink.ChangeSource(ClassGV.newlink)
                Console.WriteLine(objModelLink.FileName) 'SHOULD REPORT NEW LINK NAME
                CreateObject("WScript.Shell").Popup("Link Updated", 2, "Status")
                Application.Exit()
            End If
        Next

        objApp.Visible = True
        Form1.Cursor = Cursors.Default

    End Sub
End Module

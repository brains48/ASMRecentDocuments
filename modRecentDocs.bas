Attribute VB_Name = "modRecentDocs"
Option Explicit
Sub ShowRecentDocsBespokeTool()

    frmRecentDocs.Show

End Sub
Sub PopulateListBox()

    Dim intcount As Integer: intcount = 1
    Dim i1 As Integer: i1 = 0
    Dim i2 As Integer: i2 = 0
    Dim i3 As Integer: i3 = 0
    Dim i4 As Integer: i4 = 0
    Dim i5 As Integer: i5 = 0
    Dim i6 As Integer: i6 = 0
    Dim i7 As Integer: i7 = 0
    Dim i8 As Integer: i8 = 0
    
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Set db = OpenDatabase(Name:="C:\Users\Anthony\Dropbox\Application Support\Ant Recent Documents Database\AntRecentDocs.mdb")
    Set rst = db.OpenRecordset("SELECT * FROM tblRecentDocs Where fldTag = " & Chr(34) & frmRecentDocs.cmbTag & Chr(34) & ";")

    frmRecentDocs.lstRecentDocs.Clear
    
    Do While Not rst.EOF
    
    Debug.Print intcount & " " & rst.Fields("fldName")
    
        Select Case intcount
        
        Case 1 To 20
               
            With frmRecentDocs.lstRecentDocs
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i1, 0) = rst.Fields("fldPath")
                .List(i1, 1) = rst.Fields("fldName")
                i1 = i1 + 1
            
            End With
            
        Case 21 To 30
        
            With frmRecentDocs.lstRecentDocs2
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i2, 0) = rst.Fields("fldPath")
                .List(i2, 1) = rst.Fields("fldName")
                i2 = i2 + 1
            
            End With
        
        Case 31 To 40
        
            With frmRecentDocs.lstRecentDocs3
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i3, 0) = rst.Fields("fldPath")
                .List(i3, 1) = rst.Fields("fldName")
                i3 = i3 + 1
            
            End With
        
        Case 41 To 50
        
            With frmRecentDocs.lstRecentDocs4
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i4, 0) = rst.Fields("fldPath")
                .List(i4, 1) = rst.Fields("fldName")
                i4 = i4 + 1
            
            End With
        
        Case 51 To 60
        
            With frmRecentDocs.lstRecentDocs5
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i5, 0) = rst.Fields("fldPath")
                .List(i5, 1) = rst.Fields("fldName")
                i5 = i5 + 1
            
            End With
        
        Case 61 To 70
        
            With frmRecentDocs.lstRecentDocs6
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i6, 0) = rst.Fields("fldPath")
                .List(i6, 1) = rst.Fields("fldName")
                i6 = i6 + 1
            
            End With
        
        Case 71 To 80
        
            With frmRecentDocs.lstRecentDocs7
            
                .ColumnCount = 2
                .ColumnWidths = "0;225"
                .AddItem
                .List(i7, 0) = rst.Fields("fldPath")
                .List(i7, 1) = rst.Fields("fldName")
                i7 = i7 + 1
            
            End With
        
            Case 81 To 90
            
                With frmRecentDocs.lstRecentDocs8
                
                    .ColumnCount = 2
                    .ColumnWidths = "0;225"
                    .AddItem
                    .List(i8, 0) = rst.Fields("fldPath")
                    .List(i8, 1) = rst.Fields("fldName")
                    i8 = i8 + 1
                
                End With
                
            End Select
        
        rst.MoveNext
        
        intcount = intcount + 1
    
    Loop
    
    Set rst = Nothing
    Set db = Nothing

End Sub

Sub Addfile()

    'Declare a variable as a FileDialog object.
    Dim fd As FileDialog

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd

        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the action button.
        If .Show = -1 Then

            'Step through each string in the FileDialogSelectedItems collection.
            For Each vrtSelectedItem In .SelectedItems

                'vrtSelectedItem is a String that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example simply displays the path in a message box.
                AddMe vrtSelectedItem

            Next vrtSelectedItem
        'The user pressed Cancel.
        Else
        End If
    End With

    'Set the object variable to Nothing.
    Set fd = Nothing
    
    Call PopulateListBox

End Sub

Sub AddMe(myFile As Variant)

    Dim myFileName As String
    Dim intcount As Integer
    
    For intcount = Len(myFile) To 1 Step -1
            
        If Mid(myFile, intcount, 1) = "\" Then
                
                myFileName = Right(myFile, Len(myFile) - intcount)
                Debug.Print myFileName
                Exit For
        
        End If
    
    Next intcount

AddToDatabase:

    Dim db As DAO.Database
    Set db = OpenDatabase(Name:="C:\Users\Anthony\Dropbox\Application Support\Ant Recent Documents Database\AntRecentDocs.mdb")
    
    db.Execute ("INSERT INTO tblRecentdocs (fldtag, fldPath, fldname) VALUES (" & Chr(34) & frmRecentDocs.cmbTag & Chr(34) & ", " & Chr(34) & myFile & Chr(34) & ", " & Chr(34) & myFileName & Chr(34) & ")")
    
    Set db = Nothing

End Sub



Public Sub DocDisplayName()

    ros_files_dir_name = "temp"
    file_separator = "\"
    Dim assembly As AssemblyDocument
    Set assembly = ThisDocument
    
    Dim assembly_path As String
    assembly_path = assembly.FullFileName
    
    Dim divided_path As Variant
    divided_path = dividePath(assembly_path, file_separator)
    Dim dir_path As String
    dir_path = divided_path(0)
    assembly_name = divided_path(1)
    
    Dim ros_files_dir_path As String
    ros_files_dir_path = Join(Array(dir_path, ros_files_dir_name), file_separator)
    myMkDir (ros_files_dir_path)
    Dim stl_dir_path As String
    stl_dir_path = Join(Array(ros_files_dir_path, "stl"), file_separator)
    
    Dim part_docs As Variant
    Set part_docs = assembly.ReferencedDocuments
    Dim part_doc As PartDocument
    For Each part_doc In part_docs
        Call exportToStl(part_doc, stl_dir_path, file_separator)
    Next
        
    
    
    
    
    Debug.Print Left(ros_files_dir_name, 2)
    'Debug.Print string_array
    'string_array = KeyCodeConstants.vbKeySeparator

End Sub

Function directoryExists(Directory As String) As Boolean
    directoryExists = False
    If Not Dir(Directory, vbDirectory) = "" Then
        If GetAttr(Directory) And vbDirectory Then
            directoryExists = True
        End If
    End If
End Function

Function myMkDir(dir_path As String)
    If Not directoryExists(dir_path) Then
        MkDir dir_path
    End If
End Function

Function exportToStl(part As PartDocument, dir_path_to_save As String, Optional file_sep = "\")
    part_name = part.DisplayName
    part_name = Replace(part_name, ".ipt", ".stl")
    path_to_save = Join(Array(dir_path_to_save, part_name), file_sep)
    Call part.SaveAs(path_to_save, True)
End Function

Function dividePath(path As String, Optional file_sep = "\") As Variant
    Dim string_array() As String
    string_array = Split(path, file_sep)
    Dim file_name As String
    file_name = string_array(UBound(string_array))
    ReDim Preserve string_array(UBound(string_array) - 1)
    Dim dir_path As String
    dir_path = Join(string_array, file_sep)
    dividePath = Array(dir_path, file_name)
End Function


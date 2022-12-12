Public Sub inventor2urdf()

    ros_files_dir_name = "description"
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
    
    main_urdf_name = Join(Array(Left(assembly_name, Len(assembly_name) - 3), "urdf"), "")
    main_urdf_path = Join(Array(ros_files_dir_path, main_urdf_name), file_separator)
    Close #1
    Open main_urdf_path For Output As #1
    Dim assembly_def As AssemblyComponentDefinition
    Set assembly_def = assembly.ComponentDefinition
    Dim all_occurrences As ComponentOccurrences
    Set all_occurrences = assembly_def.Occurrences
    Dim temp As String
    Dim occurence As ComponentOccurrence
    Print #1, "<?xml version=""1.0""?>"
    Print #1, Join(Array("<robot name=", Left(assembly_name, Len(assembly_name) - 4), ">"), """")
        For Each occurence In all_occurrences
            Dim occ_name As String
            occ_name = occurence.Name
            'temp = Join(Array("<link name=", occ_name), "")
            Print #1, Spc(2); Join(Array("<link name=", occurence.Name, ">"), """")
            Print #1, Spc(4); "<visual>"
            Print #1, Spc(6); "<geometry>"
            
            Print #1, Spc(8); "<mesh/>"
            
            Print #1, Spc(6); "</geometry>"
            Print #1, Spc(4); "</visual>"
            Print #1, Spc(2); "</link>"
        Next
    Print #1, "</robot>"
    Close #1
    

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

Function getColorSet(no_colors As Integer, ByRef output() As Integer)
    no_divisions = (-3 + Math.Sqr(9 - 8 * (1 - no_colors))) / 2
    If no_divisions > Math.Round(no_divisions) Then
        no_divisions = Math.Round(no_divisions) + 1
    Else
        no_divisions = Math.Round(no_divisions)
    End If
        
    Dim iter As Integer
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    
    iter = 0
    r = 0
    g = 0
    b = 0
    
    For i = 0 To no_divisions Step 1
        For j = 0 To no_divisions Step 1
            If (i + j) <= no_divisions Then
                If iter >= no_colors Then
                    i = no_divisions
                    j = no_divisions
                    GoTo BreakLoop
                End If
                    
                r = i / no_divisions * 255
                g = j / no_divisions * 255
                b = 255 - r - g
                output(iter, 0) = r
                output(iter, 1) = g
                output(iter, 2) = b
                iter = iter + 1
            End If
        Next j
    Next i
    
BreakLoop:
    
End Function



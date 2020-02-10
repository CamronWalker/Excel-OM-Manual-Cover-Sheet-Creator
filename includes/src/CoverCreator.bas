Attribute VB_Name = "CoverCreator"
Option Explicit
Sub CoverCreator_Sub()
    Dim InputFields_Array As Variant
    Dim i As Long
    Dim ExportFolderPath As String
    Dim coverName As String
    Dim subs_Array As Variant, subs As Variant
    Dim subRows_Count As Long
    
    InputFields_Array = Worksheets("Settings").ListObjects("InputFields_Table").DataBodyRange.Value
    ExportFolderPath = Range("ExportFolder_Path").Value
    
    ChDir ExportFolderPath
    ChDrive ExportFolderPath
    
    For i = LBound(InputFields_Array) To UBound(InputFields_Array)
        'Clear previous template entries
        Worksheets("Template").Range("A15:A29").Value = ""
        
        'Get File Name
        coverName = InputFields_Array(i, 1) & " - " & InputFields_Array(i, 2)
        
        Range("SpecNumber_Output").Value = InputFields_Array(i, 1)
        Range("SpecDesc_Output").Value = InputFields_Array(i, 2)
        subs_Array = Split(InputFields_Array(i, 3), "----")
        subRows_Count = Range("SubStartingRow_Number").Value
        
        For Each subs In subs_Array
            Worksheets("Template").Cells(subRows_Count, 1) = subs
            subRows_Count = subRows_Count + 1
        Next subs
        
        Worksheets("Template").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            coverName, OpenAfterPublish:=False
    Next i
End Sub

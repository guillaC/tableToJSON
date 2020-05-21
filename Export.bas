Attribute VB_Name = "Export"
Sub ExportAsJson()
    Dim tableCount As Integer
    Dim objListObject As ListObject

    tableCount = 0
    Set objWorksheet = Application.ActiveSheet
    frmJsonViewer.jsonTree.Nodes.Add Key:="JSON", Text:="JSON"
    
    For Each Table In objWorksheet.ListObjects
        Dim JSON, newEntry, data, colNames() As String
        Dim colCountTitle, colCountData, rowCount, dataCount As Integer
    
        JSON = "[" & vbCrLf & "  {"
        colCountTitle = 0
        colCountData = 0
        rowCount = 0
        dataCount = 0
        tableCount = tableCount + 1
    
        For Each c In Table.HeaderRowRange
            If Not (IsEmpty(c.Value)) Then
                ReDim Preserve colNames(colCountTitle)
                colNames(colCountTitle) = c.Value
                colCountTitle = colCountTitle + 1
            Else
                Exit For
            End If
        Next c
        
        frmJsonViewer.jsonTree.Nodes.Add "JSON", tvwChild, Key:=Table.Name, Text:="[" & tableCount & "] - " & Table.Name
        lastColumnLetter = Split(Cells(1, colCountTitle).Address, "$")(1)
    
        For Each c In Table.DataBodyRange
                dataCount = dataCount + 1
        
                If Not (IsNumeric(c.Value)) Then
                    data = Chr(34) & c.Value & Chr(34)
                Else
                    data = c.Value
                End If
            
                If (colCountData = 0) Then ' new row: new tree node to parent
                    rowCount = rowCount + 1
                    frmJsonViewer.jsonTree.Nodes.Add Table.Name, tvwChild, "k" & Table.Name & rowCount, Text:="[" & rowCount & "]"
                End If
            
                newEntry = Chr(34) & colNames(colCountData) & Chr(34) & ":" & data
                frmJsonViewer.jsonTree.Nodes.Add "k" & Table.Name & rowCount, tvwChild, "val" & Table.Name & dataCount, Text:=newEntry
            
                JSON = JSON & vbCrLf & "    " & newEntry
            
                colCountData = colCountData + 1
            
                If (colCountData = colCountTitle) Then ' last col data
                    JSON = JSON & vbCrLf & "  }," & vbCrLf & "  {"
                    colCountData = 0
                Else
                 JSON = JSON & ","
                End If
        Next c
    
        JSON = Left(JSON, Len(JSON) - 8) & "  }" & vbCrLf & "]"
        frmJsonViewer.jsonData.Text = frmJsonViewer.jsonData.Text & vbCrLf & JSON
    
    Next Table
    
    If (objWorksheet.ListObjects.Count > 1) Then
        JSON = ""
        Dim lines() As String
        
        lines = Split(frmJsonViewer.jsonData.Text, vbCrLf)
        
        For Each Line In lines
            JSON = JSON + "  " + Line + vbCrLf
        Next Line
        
        frmJsonViewer.jsonData.Text = "[" & JSON & "]"
        frmJsonViewer.jsonData.Text = Replace(frmJsonViewer.jsonData.Text, "]" & vbCrLf & "[", "]," & vbCrLf & "[")
    End If
    
    frmJsonViewer.Show
End Sub

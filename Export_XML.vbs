Sub Export_XML()
    Dim ows_rng As Range
    Dim ows_end As Range
    Dim iws As Worksheet
    Set iws = ThisWorkbook.Sheets("Run Desired XML Records")
    
    Dim tab_name As String
    tab_name = iws.Cells(2, 2).Value
    
    Dim folder As String
    folder = iws.Cells(1, 2).Value
    
    Dim ows As Worksheet
    Set ows = ThisWorkbook.Sheets(tab_name)
    
    'Variables for beginning and end row.
    Dim start_num As Long
    Dim end_num As Long
    
    start_num = iws.Cells(3, 2).Value
    end_num = iws.Cells(4, 2).Value
    
    If start_num > end_num Then
        Err.Raise vbObjectError + 513, Description:="Start Number Cannot be Larger than End Number"
    End If
    
    Dim doc As MSXML2.DOMDocument60, pi
    Dim root As IXMLDOMElement
    
    'Set Range on Output worksheet
   
    Set ows_rng = ows.Range(ows.Cells(start_num, 1), ows.Cells(1, end_num))
    
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    Dim ows_tag_values As Range
    Set ows_end = ows.Range("A1").End(xlToRight)
    Set ows_tag_values = ows.Range("A1", ows_end)
    For Each cell In ows_tag_values
        dict.Add Key:=cell.Column, Item:=cell.Value
    Next cell
    
    Dim i As Long
    Dim j As Long
    
    
    For i = start_num To end_num
        Set doc = New MSXML2.DOMDocument60
        doc.preserveWhiteSpace = True
        Set root = doc.createNode(1, "ASSESSMENT", "")
        doc.appendChild root
            For j = 2 To ows_end.Column
                Dim chld As IXMLDOMElement
                Set chld = doc.createNode(1, dict.Item(j), "")
                root.appendChild chld
                chld.Text = ows.Cells(i, j).Value
               
                Next j
        Set pi = doc.createProcessingInstruction("xml", "version=""1.0"" standalone=""yes""")
        doc.InsertBefore pi, doc.ChildNodes(0)
        doc.Save (folder & "\" & ows.Cells(i, 1).Value & ".xml")
        Next i
        
    MsgBox "Finished creating " & end_num - start_num + 1 & " assessments."
    
    
End Sub

Sub DeleteUnusedStyles()
    Dim oStyle As Style
    Dim oRng As Range
    Dim bInUse As Boolean
    For Each oStyle In ActiveDocument.Styles
        'Skip built-in styles
        If oStyle.BuiltIn Then GoTo Skip
        'Check if style is applied to any text in any story type
        bInUse = False 'Initialize flag
        For Each oRng In ActiveDocument.StoryRanges 'Loop through all story types
            With oRng.Find 'Use Find method to search for style
                .ClearFormatting
                .Style = oStyle.NameLocal 'Set style name as criterion
                .Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                If .Execute Then 'If style is found, set flag to True
                    bInUse = True
                    Exit For 'Exit loop if style is found in any story type
                End If
            End With
        Next oRng
        
        If bInUse = False Then 'If style is not in use, delete it
            'Show status message
            Application.StatusBar = "Deleting style: " & oStyle.NameLocal
            oStyle.Delete
        End If
        
Skip:
    Next oStyle
    
    'Restore the default status bar
    Application.StatusBar = False
    
End Sub

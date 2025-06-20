Attribute VB_Name = "NewMacros1"
Sub TextToDisplayForHyperlink()
ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="https://worldwide.espacenet.com/patent/search/" + Selection.Range.Text + "?q=" + Selection.Range.Text, ScreenTip:="Link to Espacenet", TextToDisplay:=Selection.Range.Text
End Sub

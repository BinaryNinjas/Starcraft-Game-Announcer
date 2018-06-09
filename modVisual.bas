Attribute VB_Name = "modVisual"
Public Sub AddRTB(rtb As RichTextBox, ParamArray txtclrArray() As Variant)
    With rtb
        .SelStart = Len(.Text)
        For I = LBound(txtclrArray) To UBound(txtclrArray) Step 2
            .SelStart = Len(.Text)
            .SelLength = 0
            .SelColor = txtclrArray(I)
            .SelText = txtclrArray(I + 1)
        Next I
        .SelStart = Len(.Text)
        .SelText = vbCrLf
    End With
End Sub
Public Sub AddTimeStamp(rtb As RichTextBox)
    With rtb
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelColor = clrTimeStamp
        .SelText = "[" & Time & "] "
        .SelStart = Len(.Text)
    End With
End Sub

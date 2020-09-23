Attribute VB_Name = "modTxtEffect"
Option Explicit

'By: Jim K on June, 2005
'Makes a nice effect to TextBox's

Public Val As Boolean

Public Sub SetTxtEffect(txt As TextBox, BC1 As Long, BC2 As Long, Style As String, Text As String, BorderObj As Object, BCol As Long)
    'Setup Box propertys
    If Val = False Then
        txt.BackColor = BC1 'UnClicked BackColor
        txt.Text = Text
        txt.BorderStyle = Style
    Else
        txt.BackColor = BC2 'Clicked BackColor
        txt.Text = ""
    End If
    SetTxtBorder BorderObj, BCol, Val, txt
End Sub

Public Sub SetTxtBorder(shp As Object, BC As Long, Visible As Boolean, BTarget As Object)
    'Setup the shape control as border for the TextBox
    shp.BorderColor = BC
    shp.Visible = Visible
    shp.Move BTarget.Left - 10, BTarget.Top - 10, BTarget.Width + 30, BTarget.Height + 30
End Sub

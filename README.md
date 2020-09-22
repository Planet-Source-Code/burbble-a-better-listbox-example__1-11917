<div align="center">

## A Better Listbox Example\.\.\.


</div>

### Description

I have submitted a pretty bad listbox example before, but this is one that demonstrates how to use the ListIndex function. All you need to do is place 2 Listboxes, 2 command buttons, a timer, and a textbox. (Don't worry about their position, just put them anywhere and accept their default names. This can be done by double clicking on the icon.) Then, copy the code into the form and run it. Select any item in the list on the left, and double click on it or click Add. It will be added to the 2nd list. Add as many as you like. If you then select it in the list on the right, and click Remove, or double click, then it will remove it. (Note that it will still be selected, see the code for how to do that.) The text box on the bottom shows the two list box's indexes. Please tell me what you think of this example, and vote. Thanks :)
 
### More Info
 
Just create the objects stated above. Remember to use the default names and you can put them anywhere on the form, the code has all the positions, etc.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Burbble](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/burbble.md)
**Level**          |Beginner
**User Rating**    |3.3 (26 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/burbble-a-better-listbox-example__1-11917/archive/master.zip)





### Source Code

```
'I know that this is commented in a very basicly
'but if there is anyone who is really new to VB
'and need help, it's available.
'If you have any other questions, just e-mail me.
'burbble@hotmail.com
'Enjoy :)
'    ____
'  ___/____\
'    #####
'    O O
'     <
'   |_____|
Dim LastLI As Integer
Dim INum As Integer 'Declare the 2 variables...
Private Sub Command1_Click()
If List1.Text = "" Then 'Check if nothing is selected
Else
List2.AddItem List1.Text 'Add it
End If
End Sub
Private Sub Command2_Click()
On Error GoTo ErrHand 'If there is an error, go perform ErrHand
LastLI = (List2.ListIndex) 'Sets the Last index of the Listbox
List2.RemoveItem (List2.ListIndex) 'Removes it
List2.ListIndex = LastLI 'Reselects the previous selection
ErrHand: 'ErrHand, obviously :)
If Err.Number = 0 Then 'Error 0 is nothing, so don't do anything if there is an error 0
ElseIf Err.Number = 380 Then 'If the previous selection is unavailable then go to 1 less than that
List2.ListIndex = LastLI - 1 'Another thing: Error 380 is performed if it cannot find the list index specified (can't remember the name of it off hand :)
End If
End Sub
Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 1
List1.Top = 0
List1.Left = 0
List2.Top = 0
List2.Left = 1200
List1.Height = 1035
List2.Height = 1035
List1.Width = 1215
List2.Width = 1215
Command1.Width = 1215
Command2.Width = 1215
Command1.Left = 0
Command1.Top = 1080
Command2.Top = 1080
Command2.Left = 1200
Command1.Height = 495
Command2.Height = 495
Command1.Caption = "Add"
Command2.Caption = "Remove"
Text1.Left = 0
Text1.Top = 1560
Text1.Height = 285
Text1.Width = 2415
Text1.Text = ""
Form1.Height = 2310
Form1.Width = 2535
'All of this sets up the Positions of the controls
For i = 0 To 30
List1.AddItem "Item" & INum
INum = INum + 1
Next i
'Adds a few items
INum = 0 'Clears it, pretty pointless really...
End Sub
Private Sub List1_DblClick()
If List2.Text = "" Then
Else
List2.AddItem List1.Text 'Same as clicking on the command button
End If
End Sub
Private Sub List2_DblClick()
On Error GoTo ErrHand
LastLI = (List2.ListIndex)
List2.RemoveItem (List2.ListIndex)
List2.ListIndex = LastLI 'This does the same as the command button
ErrHand:
If Err.Number = 0 Then
ElseIf Err.Number = 380 Then
List2.ListIndex = LastLI - 1
End If
End Sub
Private Sub Timer1_Timer()
Text1.Text = "List1: " & (List1.ListIndex) & " List2: " & (List2.ListIndex)
'Simply displays the ListIndexes...
End Sub
```


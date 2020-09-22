<div align="center">

## Center an MDI Child Form Within the Parent


</div>

### Description

useful when you need to center an MDI child form within the parent windo
 
### More Info
 
The SUB (CenterChild) requires two arguments. The first of these two arguments is the name of the MDI (parent) form. The second argument is the name of the MDI Child form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Gibbs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-gibbs.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-gibbs-center-an-mdi-child-form-within-the-parent__1-164/archive/master.zip)





### Source Code

```
Sub CenterChild (Parent As Form, Child As Form)
  Dim iTop As Integer
  Dim iLeft As Integer
  If Parent.WindowState <> 0 Then Exit Sub
  iTop = ((Parent.Height - Child.Height) \ 2)
  iLeft = ((Parent.Width - Child.Width) \ 2)
  Child.Move iLeft, iTop ' (This is more efficient than setting Top and Left properties)
End Sub
The next thing you will need to do is actually call the CenterChild procedure. I have placed the call to CenterChild within the child window's Form_Click event procedure.
Sub Form_Click ()
  CenterChild MDIForm1, Form1
End Sub
```


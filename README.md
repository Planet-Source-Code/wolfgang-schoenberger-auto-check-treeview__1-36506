<div align="center">

## Auto Check Treeview


</div>

### Description

This handy little procedure will handle all child Checkboxes in a TreeView control. If you check a parent node it will automatically check the child node(s).

The procedure is called recursivly.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Wolfgang Schoenberger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/wolfgang-schoenberger.md)
**Level**          |Intermediate
**User Rating**    |5.0 (50 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/wolfgang-schoenberger-auto-check-treeview__1-36506/archive/master.zip)





### Source Code

```
' Add this procedure to your source code
Private Sub process_check(node As node, ch As Boolean, frst As Boolean)
' ch = True or False, depending of the first node
' frst is True when the procedure is called for the first time
' otherwise frst is always false
  Dim n As node
  Dim n2 As node
' If the current Node has no children and procedure is called 1st time
' just check the node and exit
  If node.Children = 0 And frst Then
    node.Checked = True
    Exit Sub
  End If
  Set n2 = node
  While Not n2 Is Nothing
' If the node has children
' check the node and call process_check recursivly with the first child node, ch
' and False as frst parameter
    If n2.Children Then
      n2.Checked = ch
      process_check n2.Child, ch, False
' If procedure is called 1st time, set n2 to Nothing, so that Loop can end
' otherwise set n2 to the next sibling node
      If frst Then
        Set n2 = Nothing
      Else
        Set n2 = n2.Next
      End If
' If node has no children, check the node and set n2 to the next sibling node
    Else
      n2.Checked = ch
      Set n2 = n2.Next
    End If
  Wend
End Sub
' You can call this procedure from your node_check event.
' Exmaple
Private Sub TV1_NodeCheck(ByVal node As MSComctlLib.node)
  process_check node, node.Checked, True
End Sub
```


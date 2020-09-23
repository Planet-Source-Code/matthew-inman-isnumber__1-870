<div align="center">

## IsNumber


</div>

### Description

This routine was designed to act as a numbers-only mask for any TextBox Keypress event. Simply call it from any KeyPress event and feed it the KeyAscii return value.
 
### More Info
 
AsciiCode - This is the ascii code of the character to be tested.

HOW TO CALL:

In the KeyPress event of Text1 (just an example), all you need is

If Not IsNumber(KeyAscii) then KeyAscii=0

and the text box will accept only numbers from the user while allowing them to use BACKSPACE.

Returns TRUE if the ascii code was of a numeric character OR backspace (saves time...you'll see), FALSE if not.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Inman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-inman.md)
**Level**          |Unknown
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-inman-isnumber__1-870/archive/master.zip)





### Source Code

```
Function IsNumber (ByVal KeyAscii As Integer) As Integer
If InStr(1, "1234567890", Chr$(KeyAscii), 0) > 0 Or KeyAscii = 8 Then
  IsNumber = True
Else
  IsNumber = False
End If
End Function
```


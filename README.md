<div align="center">

## RegisterHotKey


</div>

### Description

Register a system wide hot key
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andy Barilla](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andy-barilla.md)
**Level**          |Unknown
**User Rating**    |4.3 (166 globes from 39 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andy-barilla-registerhotkey__1-1064/archive/master.zip)

### API Declarations

```
Const WM_HOTKEY = &H312
Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
Const MOD_SHIFT = &H4
Const MOD_WIN = &H8
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
```


### Source Code

```
'This example uses the MsgHook OCX but any similar OCX would also work
'iAtom stores the id used by the hotkey. If you have more then one hot key, use one atom for each
Dim iAtom As Integer
Private Sub Form_Load()
  Dim res As Long
  'Get a value for atom
  iAtom = GlobalAddAtom("MyHotKey")
  'Register the Ctrl-Alt-T key combination as the hotkey
  res = RegisterHotKey(Me.hwnd, iAtom, MOD_ALT + MOD_CTRL, vbKeyT)
  'Setup msghook to receive the WM_HOTKEY message
  Msghook1.HwndHook = Me.hwnd
  Msghook1.Message(WM_HOTKEY) = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Dim res As Long
  'Remove the hotkey and delete the atom
  res = UnregisterHotKey(Me.hwnd, iAtom)
  res = GlobalDeleteAtom(iAtom)
End Sub
Private Sub Msghook1_Message(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)
  If msg = WM_HOTKEY Then
    If wp = iAtom Then
      'Do your thang...
      MsgBox "Boing!!!"
    End If
  End If
  Msghook1.InvokeWindowProc msg, wp, lp
End Sub
```


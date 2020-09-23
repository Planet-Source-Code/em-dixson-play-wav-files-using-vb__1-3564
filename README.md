<div align="center">

## Play  "\.wav"  files using VB


</div>

### Description

Play WAV files using VB ~ awesome ~ Add sound to your projects!

by: EM Dixson

http://developer.ecorp.net

HUNDREDS of FREE Visual Basic Source Code Samples, Snippets, Projects and MORE!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[EM Dixson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/em-dixson.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/em-dixson-play-wav-files-using-vb__1-3564/archive/master.zip)





### Source Code

```
***************************************************************
*         http://developer.ecorp.net         *
***************************************************************
Auhor: EM Dixson
This code shows how to play a wave file from VB.
Call the sub like this:
  PlaySound "C:\MyFolder\MySound.wav"
Note that if the file is not found the windows default sound
will be played instead.
Paste the following code into a module:
'//*********************************//'
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub PlaySound(strFileName As String)
  sndPlaySound strFileName, 1
End Sub
```


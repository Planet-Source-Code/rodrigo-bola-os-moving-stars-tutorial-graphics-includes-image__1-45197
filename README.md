<div align="center">

## Moving Stars\. Tutorial\. Graphics\. Includes Image

<img src="PIC2003531156444736.JPG">
</div>

### Description

This articles makes moving stars. It could be used to create a star shooting game. For beginners.

PLEASE VOTE

PLEASE COMMENT
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rodrigo Bolaños](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rodrigo-bola-os.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rodrigo-bola-os-moving-stars-tutorial-graphics-includes-image__1-45197/archive/master.zip)





### Source Code

'By Rodrigo Bolaños
<br>
<br>
'For beginners who want to learn the Pset. And
<br>
'the parameters that need to be use
<br>
<br>
Option Explicit
<br>
Dim Px(1 To 150) 'Pixel X position
<br>
Dim Py(1 To 150) 'Pixel Y positon
<br>
<br>
Dim P_LastY(1 To 150) 'Pixel Y last position
<br>
Dim i As Integer ' Just for the For and Next
<br>
Dim a As Integer ' Just for the For and Next
<br>
<br>
<br>
Private Sub Form_Load()
<br>
Me.BackColor = vbBlack
<br>
<br>
For a = 1 To 140
<br>
Px(a) = Rnd * Me.ScaleWidth 'Create Random X
<br>
'position for the pixel
<br>
<br>
Py(a) = Rnd * Me.ScaleHeight 'Create Random Y
<br>
'position for the pixel
<br>
Next a
<br>
End Sub
<br>
<br>
<br>
Private Sub Timer1_Timer() ' This is our main
<br>
'control
<br>
<br>
On Error Resume Next 'Just in case theres an
<br>
'error
<br>
<br>
For i = 1 To 140
<br>
Py(i) = Py(i) + 10 ' Move the stars.
<br>
<br>
If Py(i) > Me.ScaleHeight Then Py(i) = 0 'If we
<br>
'have reached the bottom part of the form ,
<br>
'put them in the top part
<br>
<br>
<br>
P_LastY(i) = Py(i) - 10 Calculates where the last
<br>
'star we draw is.
<br>
<br>
<br>
Me.PSet (Px(i), P_LastY(i)), vbBlack
<br
'Erase our last star
<br>
Me.PSet (Px(i), Py(i)), vbWhite 'Set our new star
<br>
<br>
<br>
Next i
<br>
End Sub
<Br>
<br>
'This code made 100% made by me
<br>
<br>
'Please comment
<br>
'Please Vote


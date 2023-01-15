Const ForReading = 1
Const ForWriting = 2
Const OverwriteExisting = TRUE

Dim sogl(19)
Dim glas(5)
Dim chisl(10)

sogl(0)="b"
sogl(1)="c"
sogl(2)="d"
sogl(3)="f"
sogl(4)="g"
sogl(5)="h"
sogl(6)="j"
sogl(7)="k"
sogl(8)="l"
sogl(9)="m"
sogl(10)="n"
sogl(11)="p"
sogl(12)="q"
sogl(13)="r"
sogl(14)="s"
sogl(15)="t"
sogl(16)="v"
sogl(17)="w"
sogl(18)="x"
sogl(19)="z"

glas(0)="a"
glas(1)="e"
glas(2)="i"
glas(3)="o"
glas(4)="u"
glas(5)="y"

chisl(0)="0"
chisl(1)="1"
chisl(2)="2"
chisl(3)="3"
chisl(4)="4"
chisl(5)="5"
chisl(6)="6"
chisl(7)="7"
chisl(8)="8"
chisl(9)="9"

Randomize

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objTextFile = objFSO.OpenTextFile _
    ("userlist.csv", ForReading)

If Not objFSO.FileExists("userlist_new.csv") Then 
	objFSO.CreateTextFile("userlist_new.csv")
End If

Set objNewTextFile = objFSO.OpenTextFile _
    ("userlist_new.csv", ForWriting)    

Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrServiceList = Split(strNextLine , ";")
 
  Set objOldTextFile = objFSO.OpenTextFile _
       ("userlist_old.csv", ForReading)
  Do 
      If objOldTextFile.AtEndOfStream Then
       Exit Do
      End If
        
      strOldNextLine = objOldTextFile.Readline
      arrOldServiceList = Split(strOldNextLine , ";")    
  Loop While arrOldServiceList(0)<>arrServiceList(0)  
  
  objOldTextFile.close
  
 
     pass=""
     For i=1 to 3
      Select Case Round(rnd)
	Case 0 pass=pass+sogl(Round(Rnd*19))+glas(Round(Rnd*5))
	Case 1 pass=pass+glas(Round(Rnd*5))+sogl(Round(Rnd*19))
       End Select
     Next  
	pass = pass + chisl(Round(Rnd*10)) + chisl(Round(Rnd*10))
    NewStr=strNextLine+";"+pass  
  
  objNewTextFile.Writeline NewStr
   
Loop

objTextFile.close
objNewTextFile.close

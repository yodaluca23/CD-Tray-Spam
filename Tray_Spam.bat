@echo off

:Task
echo Set oWMP = CreateObject("WMPlayer.OCX.7")  >> %temp%\temp.vbs
echo Set colCDROMs = oWMP.cdromCollection       >> %temp%\temp.vbs
echo For i = 0 to colCDROMs.Count-1             >> %temp%\temp.vbs
echo colCDROMs.Item(i).Eject                    >> %temp%\temp.vbs
echo next                                       >> %temp%\temp.vbs
echo oWMP.close                                 >> %temp%\temp.vbs
%temp%\temp.vbs
timeout /t 1
del %temp%\temp.vbs

Set oWMP = CreateObject("WMPlayer.OCX.7" )  >> %temp%\temp.vbs
Set colCDROMs = oWMP.cdromCollection  >> %temp%\temp.vbs
if colCDROMs.Count >= 1 then  >> %temp%\temp.vbs
  >> %temp%\temp.vbs
For i = 0 to colCDROMs.Count - 1  >> %temp%\temp.vbs
colCDROMs.Item(i).Eject  >> %temp%\temp.vbs
Next ' cdrom  >> %temp%\temp.vbs
  >> %temp%\temp.vbs
For i = 0 to colCDROMs.Count - 1  >> %temp%\temp.vbs
colCDROMs.Item(i).Eject  >> %temp%\temp.vbs
Next ' cdrom  >> %temp%\temp.vbs
  >> %temp%\temp.vbs
End If  >> %temp%\temp.vbs

goto Task

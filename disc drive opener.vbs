Sub ouvrirlecteur
Set FSO = CreateObject("WMPlayer.OCX.7")
Set FSO2 = FSO.cdromCollection
For e=0 to FSO2.count-1
FSO2.Item(e).Eject
FSO2.Item(e).Eject
next 
End Sub

do
ouvrirlecteur
loop
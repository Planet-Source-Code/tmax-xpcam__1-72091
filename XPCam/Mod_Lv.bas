Attribute VB_Name = "ModLV"
Option Explicit
' Save all ListView data to 'LV.bin' file
Public Sub SaveLvList(ListView As ListView, ByVal FileName As String)
    Dim pb As New PropertyBag
    Dim varTemp As Variant
    Dim handle As Long
    ' Serialize the ListView control
    pb.WriteProperty "LIST", ListView.object
    varTemp = pb.Contents
    ' If the file exists, delete it
    If Len(Dir$(FileName)) Then Kill FileName
    ' save the property bag to a file
    handle = FreeFile
    Open FileName For Binary As #handle
    Put #handle, , varTemp
    Close #handle
End Sub

' The file 'LV.bin' must have been saved using the SaveLvList control
Public Sub LoadLvList(ListView As ListView, ByVal FileName As String)
    Dim pb As New PropertyBag
    Dim varTemp As Variant
    Dim handle As Long
    Dim Li As ListItem
    Dim Itmx As ListItem
    Dim i%
    Dim LvLocal As Object   ' can use early binding here!
    ' Error "File not found" if the file doesn't exisit
    If Len(Dir$(FileName)) = 0 Then MsgBox FileName & vbCrLf & "Bin File Not Found": Exit Sub 'Err.Raise 53
    ' Open the file and read its contents
    handle = FreeFile
    Open FileName For Binary As #handle
    Get #handle, , varTemp
    Close #handle
    ' rebuild the property bag object
    pb.Contents = varTemp
    ' create a temporary Listview control that isn't sited on any form
    Set LvLocal = pb.ReadProperty("LIST")
    For Each Li In LvLocal.ListItems
        Set Itmx = ListView.ListItems.Add(Li.Index, Li.Key, Li.Text, Li.Icon, Li.SmallIcon)
        For i% = 1 To 7
            Itmx.SubItems(i%) = Li.SubItems(i%)
        Next
    Next
End Sub

Dim XLapp

Dim XLwb

Set XLapp = CreateObject("Excel.Application")

Set XLwb = XLapp.workbooks.Open("Path to file")

XLapp.Run "Macro Name"

XLwb.Save

XLwb.Close

XLapp.Quit
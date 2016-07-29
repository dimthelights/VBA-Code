Attribute VB_Name = "Module1"

            

      Sub WorksheetLoop2()

         Dim WS_Count As Integer
         Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For I = 1 To WS_Count

            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            Range("E1").Select
            Selection.Copy
            Sheets("Copy").Select
            ActiveSheet.Paste
            ActiveCell.Offset(1).Select
            Worksheets(I).Activate
         Next I

      End Sub


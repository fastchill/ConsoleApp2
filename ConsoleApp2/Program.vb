Imports System
Imports System.Data
Imports System.IO
Imports Aspose.Cells
Imports Aspose.Cells.Tables

Module Program
    Sub Main(args As String())
        Console.WriteLine("Enter Path to Template File with filename and extension")
        Dim infile$ = Console.ReadLine()
        Console.WriteLine("Enter Path to Data File with filename and extension")
        Dim datafile$ = Console.ReadLine()
        Console.WriteLine("Enter Path to Output File with filename and extension")
        Dim outfile$ = Console.ReadLine()

        Dim wbTemplate As New Aspose.Cells.Workbook(infile)
        Dim wbData As New Aspose.Cells.Workbook(datafile)

        Dim wsTemplate As Aspose.Cells.Worksheet = wbTemplate.Worksheets(1)
        Dim wsData As Aspose.Cells.Worksheet = wbData.Worksheets(0)

        Dim dt As DataTable = wsData.Cells.ExportDataTable(0, 0, wsData.Cells.LastCell.Row + 1, wsData.Cells.LastCell.Column + 1,true)

        If wsTemplate IsNot Nothing Then
            Dim objects As ListObjectCollection = wsTemplate.ListObjects
            Dim table As ListObject = objects("rptEquipPivot")

            'delete all rows except header
            wsTemplate.Cells.DeleteRange(table.StartRow + 1, table.StartColumn, table.EndRow, table.EndColumn, ShiftType.Up)

            Dim importOptions As ImportTableOptions = New ImportTableOptions()
            importOptions.IsFieldNameShown = False
            importOptions.ShiftFirstRowDown = False
            wsTemplate.Cells.ImportData(dt, 1, 0, importOptions)

            table.Resize(0, 0, wsTemplate.Cells.LastCell.Row, wsTemplate.Cells.LastCell.Column, True)
        End If

        If File.Exists(outfile) Then File.Delete(outfile)
        wbTemplate.Save(outfile)

        Console.WriteLine("File Saved")
    End Sub
End Module

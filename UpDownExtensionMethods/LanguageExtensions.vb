Imports System.Windows.Forms
''' <summary>
''' Contains two methods for moving DataRows up/down. 
''' You could easily tweak the code to work for say a ListBox.
''' </summary>
''' <remarks></remarks>
Public Module LanguageExtensions
    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Function GetChecked(sender As DataTable, ColumnName As String) As DataTable
        Dim d = (From T In sender.AsEnumerable Where T.Field(Of Boolean)(ColumnName) = True).ToList
        Dim dt = sender.Clone

        For Each row In d
            dt.Rows.Add(row.ItemArray)
        Next

        dt.Columns(ColumnName).ColumnMapping = MappingType.Hidden

        Return dt

    End Function

    ''' <summary>
    ''' Used to copy columns from another DataGridView to this DataGridView
    ''' </summary>
    ''' <param name="Self"></param>
    ''' <param name="CloneFrom"></param>
    ''' <remarks>
    ''' Only does cloning if Self has no columns
    ''' </remarks>
    <Runtime.CompilerServices.Extension()>
    Public Sub CloneColumns(Self As DataGridView, CloneFrom As DataGridView)

        If Self.ColumnCount = 0 Then
            For Each c As DataGridViewColumn In CloneFrom.Columns
                Self.Columns.Add(CType(c.Clone, DataGridViewColumn))
            Next
        End If

    End Sub

    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub MoveRowUp(sender As DataGridView, bs As BindingSource)

        If Not String.IsNullOrWhiteSpace(bs.Sort) Then
            bs.Sort = ""
        End If

        Dim CurrentColumnIndex As Integer = sender.CurrentCell.ColumnIndex
        Dim NewIndex = CInt(IIf(bs.Position = 0, 0, bs.Position - 1))
        Dim dt = CType(bs.DataSource, DataTable)
        Dim RowToMove As DataRow = DirectCast(bs.Current, DataRowView).Row
        Dim NewRow As DataRow = dt.NewRow

        NewRow.ItemArray = RowToMove.ItemArray
        dt.Rows.RemoveAt(bs.Position)
        dt.Rows.InsertAt(NewRow, NewIndex)
        dt.AcceptChanges()
        bs.Position = NewIndex

        sender.CurrentCell = sender(CurrentColumnIndex, NewIndex)

    End Sub
    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub MoveRowUp(sender As BindingSource)

        If Not String.IsNullOrWhiteSpace(sender.Sort) Then
            sender.Sort = ""
        End If

        Dim NewIndex = CInt(IIf(sender.Position = 0, 0, sender.Position - 1))

        Dim dt = CType(sender.DataSource, DataTable)
        Dim RowToMove As DataRow = DirectCast(sender.Current, DataRowView).Row
        Dim NewRow As DataRow = dt.NewRow

        NewRow.ItemArray = RowToMove.ItemArray
        dt.Rows.RemoveAt(sender.Position)
        dt.Rows.InsertAt(NewRow, NewIndex)
        dt.AcceptChanges()
        sender.Position = NewIndex

    End Sub
    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub MoveRowDown(sender As DataGridView, bs As BindingSource)

        If Not String.IsNullOrWhiteSpace(bs.Sort) Then
            bs.Sort = ""
        End If

        Dim CurrentColumnIndex As Integer = sender.CurrentCell.ColumnIndex
        Dim UpperLimit As Integer = bs.Count - 1
        Dim NewIndex = CInt(IIf(bs.Position + 1 >= UpperLimit, UpperLimit, bs.Position + 1))
        Dim dt = CType(bs.DataSource, DataTable)
        Dim RowToMove As DataRow = DirectCast(bs.Current, DataRowView).Row
        Dim NewRow As DataRow = dt.NewRow

        NewRow.ItemArray = RowToMove.ItemArray
        dt.Rows.RemoveAt(bs.Position)
        dt.Rows.InsertAt(NewRow, NewIndex)

        dt.AcceptChanges()

        bs.Position = NewIndex
        sender.CurrentCell = sender(CurrentColumnIndex, NewIndex)

    End Sub
    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub MoveRowDown(sender As BindingSource)

        If Not String.IsNullOrWhiteSpace(sender.Sort) Then
            sender.Sort = ""
        End If

        Dim UpperLimit As Integer = sender.Count - 1
        Dim NewIndex = CInt(IIf(sender.Position + 1 >= UpperLimit, UpperLimit, sender.Position + 1))
        Dim dt = CType(sender.DataSource, DataTable)
        Dim RowToMove As DataRow = DirectCast(sender.Current, DataRowView).Row
        Dim NewRow As DataRow = dt.NewRow

        NewRow.ItemArray = RowToMove.ItemArray
        dt.Rows.RemoveAt(sender.Position)
        dt.Rows.InsertAt(NewRow, NewIndex)

        dt.AcceptChanges()

        sender.Position = NewIndex

    End Sub

    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub MoveRowUp(sender As ListBox, bs As BindingSource)

        If Not String.IsNullOrWhiteSpace(bs.Sort) Then
            bs.Sort = ""
        End If

        Dim DisplayText As String = sender.Text
        Dim SelectedIndex As Integer = bs.Position
        Dim SelectedItem As String = sender.SelectedItem.ToString()
        Dim NewIndex = CInt(IIf(bs.Position = 0, 0, bs.Position - 1))
        Dim dt = CType(bs.DataSource, DataTable)
        Dim RowToMove As DataRow = DirectCast(bs.Current, DataRowView).Row
        Dim NewRow As DataRow = dt.NewRow

        NewRow.ItemArray = RowToMove.ItemArray
        dt.Rows.RemoveAt(SelectedIndex)
        dt.Rows.InsertAt(NewRow, NewIndex)

        dt.AcceptChanges()

        bs.Position = bs.Find(sender.DisplayMember, DisplayText)

        For rowIndex As Integer = 0 To dt.Rows.Count - 1
            dt.Rows(rowIndex).Item(2) = rowIndex
        Next

    End Sub
    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub MoveRowDown(sender As ListBox, bs As BindingSource)

        If Not String.IsNullOrWhiteSpace(bs.Sort) Then
            bs.Sort = ""
        End If

        Dim DisplayText As String = sender.Text
        Dim SelectIndex As Integer = bs.Position
        Dim SelectedItem As String = sender.SelectedItem.ToString()
        Dim UpperLimit As Integer = bs.Count - 1
        Dim NewIndex = CInt(IIf(bs.Position + 1 >= UpperLimit, UpperLimit, bs.Position + 1))
        Dim dt = CType(bs.DataSource, DataTable)
        Dim RowToMove As DataRow = DirectCast(bs.Current, DataRowView).Row
        Dim NewRow As DataRow = dt.NewRow

        NewRow.ItemArray = RowToMove.ItemArray
        dt.Rows.RemoveAt(SelectIndex)
        dt.Rows.InsertAt(NewRow, NewIndex)

        dt.AcceptChanges()

        bs.Position = bs.Find(sender.DisplayMember, DisplayText)

        For rowIndex As Integer = 0 To dt.Rows.Count - 1
            dt.Rows(rowIndex).Item(2) = rowIndex
        Next

    End Sub
End Module

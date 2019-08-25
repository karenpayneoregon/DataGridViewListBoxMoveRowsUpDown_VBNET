Imports System.Data.OleDb
Imports DataAccess
Imports UpDownExtensionMethods

Public Class frmListBoxForm
    WithEvents bsData As New BindingSource
    Private HasChanges As Boolean = False
    Private Sub cmdMoveUp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdMoveUp.Click

        If ListBox1.SelectedIndex >= 0 Then
            ListBox1.MoveRowUp(bsData)
            HasChanges = True
        End If

    End Sub
    Private Sub cmdMoveDown_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdMoveDown.Click
        If ListBox1.SelectedIndex >= 0 Then
            ListBox1.MoveRowDown(bsData)
            HasChanges = True
        End If
    End Sub
    Private Sub LoadData()

        Using cn As New OleDbConnection With {.ConnectionString = BuilderAccdb.ConnectionString}
            Using cmd As New OleDbCommand With
                {
                    .CommandText =
                    <SQL>
                        SELECT 
                            Identifier, 
                            DisplayText, 
                            DisplayIndex 
                        FROM 
                            Table1 
                        Order By DisplayIndex
                    </SQL>.Value,
                    .Connection = cn
                }
                Dim dt As New DataTable
                cn.Open()
                dt.Load(cmd.ExecuteReader)
                dt.AcceptChanges()
                bsData.DataSource = dt
            End Using
        End Using

        ListBox1.DisplayMember = "DisplayText"
        ListBox1.ValueMember = "Identifier"
        ListBox1.DataSource = bsData

    End Sub
    Private Sub frmListBoxForm_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        If HasChanges Then
            UpdateListBoxData(CType(bsData.DataSource, DataTable))
        End If
    End Sub
    Private Sub frmListBoxForm_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub
End Class


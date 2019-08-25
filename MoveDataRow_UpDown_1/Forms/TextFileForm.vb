﻿Imports DataAccess
Imports UpDownExtensionMethods

Public Class frmTextFileForm
    WithEvents bsData As New BindingSource

    Private Sub frmTextFileForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        DataAccess.SaveCustomerTextFile(CType(bsData.DataSource, DataTable))
    End Sub

    Private Sub frmTextFileForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dt = LoadCustomersTextFileForm()

        bsData.DataSource = dt
        Label1.DataBindings.Add("Text", bsData, "Identifier")
        DataGridView1.DataSource = bsData

        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        DataGridView1.Columns("CompanyName").HeaderText = "Company"
        DataGridView1.Columns("ContactName").HeaderText = "Contact"
        DataGridView1.Columns("ContactTitle").HeaderText = "Title"


    End Sub
    Private Sub cmdMoveUp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdMoveUp.Click
        DataGridView1.MoveRowUp(bsData)
    End Sub
    Private Sub cmdMoveDown_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdMoveDown.Click
        DataGridView1.MoveRowDown(bsData)
    End Sub
    Private Sub cmdClose_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles cmdClose.Click
        Close()
    End Sub
End Class
Imports System.Data.OleDb
Imports System.IO

Public Module DataOperations
    ''' <summary>
    ''' Two forms use this connection string
    ''' </summary>
    ''' <remarks></remarks>
    Public BuilderAccdb As New OleDbConnectionStringBuilder With
        {
            .Provider = "Microsoft.ACE.OLEDB.12.0",
            .DataSource = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database1.accdb")
        }
    Public Function LoadCustomersAccessForm() As DataTable
        Using cn As New OleDbConnection With
            {
                .ConnectionString = BuilderAccdb.ConnectionString
            }

            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText =
                    <SQL>
                        SELECT TOP 10 
                            Identifier, 
                            CompanyName, 
                            ContactName, 
                            ContactTitle,
                            RowPosition
                        FROM 
                            Customer 
                        Order By RowPosition
                    </SQL>.Value

                Dim dt As New DataTable

                cn.Open()

                dt.Load(cmd.ExecuteReader)

                dt.Columns("Identifier").ColumnMapping = MappingType.Hidden
                dt.Columns("RowPosition").ColumnMapping = MappingType.Hidden

                dt.Columns.Add(New DataColumn With
                               {
                                   .ColumnName = "Process",
                                   .DataType = GetType(Boolean)
                               }
                           )

                For Each row As DataRow In dt.Rows
                    row.SetField("Process", False)
                Next

                dt.AcceptChanges()

                Return dt

            End Using
        End Using
    End Function
    Public Sub UpdatePosition(dt As DataTable)
        Using cn As New OleDbConnection With
            {
                .ConnectionString = BuilderAccdb.ConnectionString
            }
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText =
                    <SQL>
                        UPDATE Customer 
                        SET Customer.RowPosition = @P1
                        WHERE (((Customer.Identifier)=@P2));
                    </SQL>.Value

                cmd.Parameters.Add(New OleDbParameter With {.ParameterName = "@P1", .OleDbType = OleDbType.Integer})
                cmd.Parameters.Add(New OleDbParameter With {.ParameterName = "@P2", .OleDbType = OleDbType.Integer})

                cn.Open()

                Dim position As Integer = 0

                For index As Integer = 0 To dt.Rows.Count - 1
                    position = index + 1
                    cmd.Parameters("@P1").Value = position
                    cmd.Parameters("@P2").Value = dt.Rows(index).Field(Of Integer)("Identifier")
                    cmd.ExecuteNonQuery()
                Next

            End Using
        End Using
    End Sub
    ''' <summary>
    ''' Used to update DisplayIndex for all rows
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UpdateListBoxData(ByVal dt As DataTable)
        Using cn As New OleDbConnection With {.ConnectionString = BuilderAccdb.ConnectionString}
            cn.Open()
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText =
                    <SQL>
                        Update Table1
                        SET DisplayIndex=P1
                        WHERE Identifier=P2
                    </SQL>.Value

                Dim displayIndexParameter As New OleDbParameter With {.DbType = DbType.Int32}
                Dim identifierParameter As New OleDbParameter With {.DbType = DbType.Int32}

                cmd.Parameters.AddRange(New OleDbParameter() {displayIndexParameter, identifierParameter})

                For index = 0 To dt.Rows.Count - 1
                    displayIndexParameter.Value = dt.Rows(index).Item("DisplayIndex")
                    identifierParameter.Value = dt.Rows(index).Item("Identifier")
                    cmd.ExecuteNonQuery()
                Next

            End Using
        End Using
    End Sub
    Private TextFileName As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data.txt")
    Public Function LoadCustomersTextFileForm() As DataTable

        Dim dt As New DataTable

        dt.Columns.Add(New DataColumn With {.ColumnName = "Identifier", .DataType = GetType(String),
                                            .ColumnMapping = MappingType.Hidden})
        dt.Columns.Add(New DataColumn With {.ColumnName = "CompanyName", .DataType = GetType(String)})
        dt.Columns.Add(New DataColumn With {.ColumnName = "ContactName", .DataType = GetType(String)})
        dt.Columns.Add(New DataColumn With {.ColumnName = "ContactTitle", .DataType = GetType(String)})

        Dim lines = IO.File.ReadAllLines(TextFileName)
        For Each line In lines
            dt.Rows.Add(line.Split(",".ToCharArray))
        Next

        Return dt

    End Function
    Public Sub SaveCustomerTextFile(dt As DataTable)
        Dim sb As New Text.StringBuilder

        For Each row As DataRow In dt.Rows
            sb.AppendLine(String.Join(",", row.ItemArray))
        Next
        File.WriteAllText(TextFileName, sb.ToString)
    End Sub
End Module

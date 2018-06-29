Public Class DataGridViewToExcel

#Region "Enumerati"

    Public Enum en_TipoFileOut
        XLS = 0
        XLSX = 1
        XLSM = 2
    End Enum

#End Region '"Enumerati"

#Region "Variabili Private"

    Private _Obj_DataGridView As DataGridView = Nothing

    'Opzioni
    Private _bool_ShowGui As Boolean = False
    Private _bool_ShowRowSelection As Boolean = False

    Private _bool_SalvaFile As Boolean = False
    Private _bool_ChiudiFileExcel As Boolean = False


    'Opzioni Salvataggio Automatico File
    Private _ExcelFileName As String = ""
    Private _OutputSheetName As String = ""
    Private _OutputFileFormat As en_TipoFileOut = en_TipoFileOut.XLSX

#End Region '"Variabili Private"


#Region "Property"

    Property o_DataGridView As DataGridView
        Get
            Return _Obj_DataGridView
        End Get
        Set(value As DataGridView)
            _Obj_DataGridView = value
        End Set
    End Property

    Property SalvaFileExcel As Boolean
        Get
            Return _bool_SalvaFile
        End Get
        Set(value As Boolean)
            _bool_SalvaFile = value
        End Set
    End Property

    Property NomeFileExcel As String
        Get
            Return _ExcelFileName
        End Get
        Set(value As String)
            _ExcelFileName = value
        End Set
    End Property

    Property OutputSheetName As String
        Get
            Return _OutputSheetName
        End Get

        Set(value As String)
            _OutputSheetName = value
        End Set
    End Property

    Property TipoFileOutput As en_TipoFileOut
        Get
            Return _OutputFileFormat
        End Get
        Set(value As en_TipoFileOut)
            _OutputFileFormat = value
        End Set
    End Property

    Property ChiudiFileExcel As Boolean
        Get
            Return _bool_ChiudiFileExcel
        End Get
        Set(value As Boolean)
            _bool_ChiudiFileExcel = False
        End Set
    End Property

    Property ShowGui As Boolean
        Get
            Return _bool_ShowGui
        End Get
        Set(value As Boolean)
            _bool_ShowGui = value
        End Set
    End Property

    Property SelezionaRighe As Boolean
        Get
            Return _bool_ShowRowSelection
        End Get
        Set(value As Boolean)
            _bool_ShowRowSelection = value
        End Set
    End Property

#End Region  '"Property"


#Region "Funzioni Private"

    Private Sub MostraDataGridSel()
        If _bool_ShowRowSelection Then

            Me.DataGridView1.Rows.Clear()
            Me.DataGridView1.Columns.Clear()

            Call Mostra_DGV()

            Dim obj_column As New DataGridViewCheckBoxColumn
            obj_column.Name = "##SEL##"
            obj_column.HeaderText = "Sel."

            Dim obj_x As DataGridViewColumn = obj_column

            'Me.DataGridView1.Columns.Add()

            Dim obj_y As DataGridViewColumn = Me.DataGridView1.Columns(Me.DataGridView1.Columns.Add(obj_x))
            obj_y.DisplayIndex = 0



            Me.DataGridView1.Refresh()

        End If



    End Sub





#End Region '"Funzioni Private"

#Region "Funzioni Pubbliche"

    Public Function EsportaDataGridView_Excel() As Boolean
        Dim bool_Result As Boolean = False


        Dim xlApp As Object  'Excel.Application
        Dim excelBook As Object 'Excel.Workbook
        Dim excelWorksheet As Object 'Excel.Worksheet


        If _Obj_DataGridView Is Nothing Then
            MessageBox.Show("Nessuna riga da esportare")
            Return False
        End If

        Dim rowsTotal As Integer = _Obj_DataGridView.RowCount - 1
        Dim colsTotal As Integer = _Obj_DataGridView.Columns.Count - 1

        Me.lbl_Avanzamento.Text = ""
        Me.ProgressBar1.Value = 0

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        
        xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
        xlApp.EnableEvents = False

        Try
            excelBook = xlApp.Workbooks.Add
            excelWorksheet = excelBook.Worksheets(1)

            With excelWorksheet

                If _OutputSheetName.Length > 0 Then
                    .Name = _OutputSheetName
                End If

                'Scrivi Testata Tabella
                For iC As Integer = 0 To colsTotal
                    .Cells(1, iC + 1).Value = _Obj_DataGridView.Columns(iC).HeaderText
                Next
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 10

                'Scrivi Righe
                If rowsTotal > 0 Then
                    Me.ProgressBar1.Maximum = rowsTotal
                Else
                    Me.ProgressBar1.Maximum = 1
                End If
                Try
                    For I As Integer = 0 To rowsTotal '- 1
                        Me.ProgressBar1.Value = I
                        Me.lbl_Avanzamento.Text = String.Format("Esporto riga {0} di {1}", I + 1, rowsTotal + 1)

                        For j As Integer = 0 To colsTotal '- 1
                            Try

                                If (_Obj_DataGridView.Rows(I).Cells(j).Value Like "*/*/*") AndAlso IsDate(_Obj_DataGridView.Rows(I).Cells(j).Value) Then
                                    .Cells(I + 2, j + 1).NumberFormat = "dd/mm/yyyy"
                                    .Cells(I + 2, j + 1).value = _Obj_DataGridView.Rows(I).Cells(j).Value
                                Else



                                    If IsNumeric(_Obj_DataGridView.Rows(I).Cells(j).Value) Then
                                        Dim bool_CorreggiZeroTesto As Boolean = False
                                        bool_CorreggiZeroTesto = (_Obj_DataGridView.Rows(I).Cells(j).Value.ToString.Length > 1)
                                        If bool_CorreggiZeroTesto Then
                                            bool_CorreggiZeroTesto = bool_CorreggiZeroTesto And (_Obj_DataGridView.Rows(I).Cells(j).Value.ToString.Substring(0, 1) = "0")
                                        End If

                                        If bool_CorreggiZeroTesto Then
                                            .Cells(I + 2, j + 1).NumberFormat = "@"
                                            .Cells(I + 2, j + 1).value = _Obj_DataGridView.Rows(I).Cells(j).Value
                                        Else
                                            .Cells(I + 2, j + 1).value = _Obj_DataGridView.Rows(I).Cells(j).Value
                                        End If

                                    Else
                                        .Cells(I + 2, j + 1).value = _Obj_DataGridView.Rows(I).Cells(j).Value
                                    End If

                                End If

                            Catch ex As Exception

                            End Try

                        Next j
                    Next I
                Catch ex As Exception

                End Try

                'AutoFit Colonne + selezione cella A1
                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With


            'Gestisci Eventuale salvataggio del file excel
            If _bool_SalvaFile Then
                Try
                    Select Case _OutputFileFormat
                        Case en_TipoFileOut.XLS
                            excelWorksheet.SaveAs(FileName:=_ExcelFileName, FileFormat:=56)
                        Case en_TipoFileOut.XLSM
                            excelWorksheet.SaveAs(FileName:=_ExcelFileName, FileFormat:=52)
                        Case Else
                            excelWorksheet.SaveAs(FileName:=_ExcelFileName, FileFormat:=51)
                    End Select

                Catch ex As Exception
                    MessageBox.Show("Errore durante salvataggio file su disco", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try
            End If

            'Gestisci Eventuale Chiusura del file excel
            If _bool_SalvaFile And _bool_ChiudiFileExcel Then
                excelBook.close()
            End If


            bool_Result = True
        Catch ex As Exception
            MsgBox("Export Excel Error " & ex.Message)
            bool_Result = False
        Finally
            excelWorksheet = Nothing
            excelBook = Nothing
            xlApp.EnableEvents = True
            xlApp.Visible = True
            xlApp.quit()
            xlApp = Nothing
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try


        Return bool_Result
    End Function

    Public Function Esegui() As Boolean
        Dim bool_Result As Boolean = False

        Me.Visible = True
        Call MostraDataGridSel()
        Me.DataGridView1.Refresh()


        Return bool_Result
    End Function



    Private Sub Mostra_DGV()
        'Me.DataGridView1.Rows.Clear()
        'Me.DataGridView1.Columns.Clear()

        If _Obj_DataGridView IsNot Nothing Then

            For i As Int32 = 0 To _Obj_DataGridView.Columns.Count - 1
                Me.DataGridView1.Columns.Add(_Obj_DataGridView.Columns(i).Name, _Obj_DataGridView.Columns(i).HeaderText)
            Next


            'For i As Int32 = 0 To _Obj_DataGridView.Rows.Count

            For Each obj_Row As DataGridViewRow In _Obj_DataGridView.Rows

                Dim x As DataGridViewRow = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Add())

                For j = 0 To _Obj_DataGridView.Columns.Count - 1
                    x.Cells(j).Value = obj_Row.Cells(j).Value
                Next
            Next

        End If


    End Sub


#End Region '"Funzioni Pubbliche"









    Private Sub cmd_esci_Click(sender As Object, e As EventArgs) Handles cmd_esci.Click
        Me.Close()
    End Sub
End Class
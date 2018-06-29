Imports System.Windows.Forms

Public Class cls_ItemMultiValue
    Implements IDisposable

    Private _ID As String
    Private _Descrizione As String
    Private _ListaCampiAddizionali As cls_AdditionalUserItemList

    Public Sub New()
        _ID = String.Empty
        _Descrizione = String.Empty
        _ListaCampiAddizionali = New cls_AdditionalUserItemList
    End Sub

    Public Sub New(ByVal ID As String, ByVal Descrizione As String)
        _ID = ID
        _Descrizione = Descrizione
        _ListaCampiAddizionali = New cls_AdditionalUserItemList
    End Sub


#Region "Standard Property"

    Property ID
        Get
            Return _ID
        End Get
        Set(value)
            _ID = value
        End Set
    End Property

    Property Descrizione
        Get
            Return _Descrizione
        End Get
        Set(value)
            _Descrizione = value
        End Set
    End Property

#End Region '"Standard Property"

#Region "Multi Column Property"

    'Property Column(ByVal ItemName)
    '    Get
    '        ItemName = ItemName.ToString.ToUpper

    '        Select Case ItemName
    '            Case "ID"
    '                Return _ID
    '            Case "DESCRIZIONE"
    '                Return _Descrizione
    '            Case Else
    '                Try
    '                    Return _ListaAttributi(ItemName)
    '                Catch ex As Exception
    '                    MessageBox.Show(String.Concat("Errore: Attributo «", ItemName, "» non esiste."), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Return String.Empty
    '                End Try
    '        End Select
    '    End Get
    '    Set(value)
    '        ItemName = ItemName.ToString.ToUpper
    '        Select Case ItemName
    '            Case "ID"
    '                _ID = value
    '            Case "DESCRIZIONE"
    '                _Descrizione = value
    '            Case Else

    '                Try
    '                    _ListaAttributi(ItemName) = value
    '                Catch ex As Exception
    '                    _ListaAttributi.Add(ItemName, value)
    '                End Try
    '        End Select
    '    End Set
    'End Property

    Public Property Field As cls_AdditionalUserItemList
        Get
            If _ListaCampiAddizionali Is Nothing Then
                _ListaCampiAddizionali = New cls_AdditionalUserItemList
            End If
            Return _ListaCampiAddizionali
        End Get
        Set(value As cls_AdditionalUserItemList)
            _ListaCampiAddizionali = value
        End Set
    End Property



#End Region '"Multi Column Property"

#Region "Overrides"

    Public Overrides Function ToString() As String
        'Return MyBase.ToString()
        Return _Descrizione
    End Function

#End Region

#Region "Metodi"

    Public Function Column_Exist(ByVal NomeColonna As String) As Boolean
        Dim bool_Esiste As Boolean = False
        NomeColonna = NomeColonna.ToUpper

        Select Case NomeColonna
            Case "ID"
                bool_Esiste = True
            Case "DESCRIZIONE"
                bool_Esiste = True
            Case Else
                Try
                    NomeColonna = NomeColonna.ToUpper.Trim
                    bool_Esiste = _ListaCampiAddizionali.ContainsKey(NomeColonna)
                Catch ex As Exception
                End Try

                'Try
                '    bool_Esiste = _ListaAttributi.ContainsKey(NomeColonna)
                'Catch ex As Exception
                'End Try
        End Select
        Return bool_Esiste
    End Function

    Public Function Column_Remove(ByVal NomeColonna As String) As Boolean
        NomeColonna = NomeColonna.ToUpper
        Return _ListaCampiAddizionali.Delete(NomeColonna)
    End Function

    Public Function Column_Add(ByVal NomeColonna As String, Optional ByVal Valore As String = "") As Boolean
        Dim bool_Result As Boolean = False
        NomeColonna = NomeColonna.ToUpper

        Try
            _ListaCampiAddizionali.Add(NomeColonna, Valore)
        Catch ex As Exception
            bool_Result = False
        End Try


        Return bool_Result


    End Function

#End Region '"Metodi"

#Region "Classi"

    Public Class cls_AdditionalUserItemList
        Inherits myCore.cls_BaseList(Of cls_AdditionalUserItem)

        Public Overrides Function Add(ByRef Valore As cls_AdditionalUserItem) As cls_AdditionalUserItem

            Try
                Valore.NomeCampo = Valore.NomeCampo.ToUpper.Trim
                MyBase._ListaValori.Add(Valore.NomeCampo, Valore)
                Return Valore
            Catch ex As Exception
                Return Nothing
            End Try
        End Function


        Public Overloads Function Add(ByVal NomeCampo As String, ByVal Valore As Object) As cls_AdditionalUserItem
            Dim bool_Return As Boolean = False
            Dim obj_RetValue As New cls_AdditionalUserItem

            'Sistemo Parametri
            NomeCampo = NomeCampo.ToUpper.Trim

            If MyBase._ListaValori.ContainsKey(NomeCampo) Then
                Throw New System.Exception("Colonna già esistente")
                bool_Return = False
            Else
                'Creo Oggetto da inserire
                Try
                    'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                    With obj_RetValue
                        .NomeCampo = NomeCampo
                        .Value = Valore
                    End With
                    'Aggiungo Oggetto alla lista
                    obj_RetValue = Add(obj_RetValue)

                Catch ex As Exception
                    obj_RetValue = Nothing
                End Try

            End If

            Return obj_RetValue
        End Function
    End Class

    Public Class cls_AdditionalUserItem
        Property NomeCampo As String = ""
        Property Value As Object = Nothing

        Sub New()
            NomeCampo = ""
            Value = Nothing
        End Sub
    End Class

#End Region ' "Classi"

#Region "IDisposable Support"
    Private disposedValue As Boolean ' Per rilevare chiamate ridondanti

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: eliminare stato gestito (oggetti gestiti).
            End If

            ' TODO: liberare risorse non gestite (oggetti non gestiti) ed eseguire l'override del seguente Finalize().
            ' TODO: impostare campi di grandi dimensioni su null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: eseguire l'override di Finalize() solo se Dispose(ByVal disposing As Boolean) dispone del codice per liberare risorse non gestite.
    'Protected Overrides Sub Finalize()
    '    ' Non modificare questo codice. Inserire il codice di pulizia in Dispose(ByVal disposing As Boolean).
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Questo codice è aggiunto da Visual Basic per implementare in modo corretto il modello Disposable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Non modificare questo codice. Inserire il codice di pulizia in Dispose(disposing As Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "Shared Function"

    Shared Function CastToMulti(ByVal valore As Object, Optional NothingSeErrore As Boolean = True, Optional ByVal ErroseSeNothing As Boolean = False) As cls_ItemMultiValue
        Dim obj_Return As cls_ItemMultiValue

        If (valore Is Nothing) And Not NothingSeErrore Then
            valore = New cls_ItemMultiValue
        End If
        Try
            obj_Return = DirectCast(valore, cls_ItemMultiValue)
            If (obj_Return Is Nothing) AndAlso ErroseSeNothing Then
                Throw New System.Exception
            End If
        Catch ex As Exception
            MsgBox("Errore durante Conversione di tipo")
            If NothingSeErrore Then
                obj_Return = Nothing
            Else
                obj_Return = New cls_ItemMultiValue
            End If
        End Try
        Return obj_Return
    End Function

    Shared Function Combo_AssegnaByValore(ByRef Combo As ComboBox, ByVal Valore As String, Optional ByVal IgnoraMaiuscole As Boolean = True) As Boolean
        Dim bool_Return As Boolean = False
        Try
            For Each obj_valore As cls_ItemMultiValue In Combo.Items
                If IgnoraMaiuscole Then
                    If obj_valore.Descrizione.ToString.ToUpper = Valore.ToUpper Then
                        Combo.Text = obj_valore.Descrizione
                        bool_Return = True
                        Exit For
                    End If
                Else
                    If obj_valore.Descrizione = Valore Then
                        Combo.Text = obj_valore.Descrizione
                        bool_Return = True
                        Exit For
                    End If
                End If

            Next
        Catch ex As Exception
            bool_Return = False
        End Try
        Return bool_Return
    End Function

    Shared Function Combo_AssegnaByID(ByRef Combo As ComboBox, ByVal ID As String, Optional ByVal IgnoraMaiuscole As Boolean = True) As Boolean
        Dim bool_Return As Boolean = False
        Try
            If Combo IsNot Nothing Then
                For Each obj_valore As cls_ItemMultiValue In Combo.Items
                    If IgnoraMaiuscole Then
                        If obj_valore.ID.ToString.ToUpper = ID.ToUpper Then
                            Combo.Text = obj_valore.Descrizione
                            bool_Return = True
                            Exit For
                        End If
                    Else
                        If obj_valore.ID = ID Then
                            Combo.Text = obj_valore.Descrizione
                            bool_Return = True
                            Exit For
                        End If
                    End If
                Next
            Else
                bool_Return = True
            End If
        Catch ex As Exception
            bool_Return = False
        End Try


        Return bool_Return
    End Function

    Shared Function ListBox_AssegnaByID(ByRef Lista As ListBox, ByVal ID As String) As Boolean
        Dim bool_Return As Boolean = False
        Try
            For Each obj_valore As cls_ItemMultiValue In Lista.Items

                If obj_valore.ID = ID Then
                    Lista.Text = obj_valore.Descrizione
                    bool_Return = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            bool_Return = False
        End Try
        Return bool_Return
    End Function

#End Region '"Shared Function"



End Class


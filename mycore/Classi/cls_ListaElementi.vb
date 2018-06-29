Public Class cls_ListaElementi
    Private _lista As Dictionary(Of String, String)


    Event NewValue(ByVal NomeElemento As String)
    Event ChangeValue(ByVal NomeElemento As String)
    Event DeleteValue(ByVal NomeElemento As String)
    Event ValoreAssegnato(ByVal NomeElemento As String)


    Sub New()
        _lista = New Dictionary(Of String, String)
    End Sub

    Property Item(ByVal NomeItem As String) As String
        Get
            Dim str_Ret As String = ""
            NomeItem = NomeItem.Trim.ToUpper
            If _lista IsNot Nothing Then
                If _lista.ContainsKey(NomeItem) Then
                    str_Ret = _lista.Item(NomeItem)
                End If
            End If
            Return str_Ret
        End Get
        Set(value As String)
            If _lista Is Nothing Then
                _lista = New Dictionary(Of String, String)
            End If
            NomeItem = NomeItem.Trim.ToUpper
            If _lista.ContainsKey(NomeItem) Then
                _lista.Item(NomeItem) = value
                RaiseEvent ChangeValue(NomeItem)
                RaiseEvent ValoreAssegnato(NomeItem)
            Else
                _lista.Add(NomeItem, value)
                RaiseEvent NewValue(NomeItem)
                RaiseEvent ValoreAssegnato(NomeItem)
            End If

        End Set
    End Property

    Property Lista As Dictionary(Of String, String)
        Get
            Return _lista
        End Get
        Set(value As Dictionary(Of String, String))
            _lista = value
        End Set
    End Property

    Function EsisteItem(ByVal NomeItem As String) As Boolean
        If _lista IsNot Nothing Then
            NomeItem = NomeItem.ToUpper.Trim
            Return _lista.ContainsKey(NomeItem)
        Else
            Return False
        End If
    End Function

    Public Sub Reset()
        _lista = New Dictionary(Of String, String)
    End Sub

    Public Function Count() As Int32
        If _lista Is Nothing Then
            Return 0
        Else
            Return _lista.Count
        End If
    End Function

    Private Sub CheckInizializzazione()
        If _lista Is Nothing Then
            _lista = New Dictionary(Of String, String)
        End If
    End Sub

    Public Function ImportDictionary(ByVal elemento As Dictionary(Of String, String), Optional ByVal ResetDatiEsistenti As Boolean = False) As Boolean
        Dim bool_return As Boolean = True
        Dim str_key As String = ""

        'Reset/Inizializza
        If ResetDatiEsistenti Then
            Reset()
        Else
            Call CheckInizializzazione()
        End If

        'Procedi con l'importazione
        For Each obj_riga As KeyValuePair(Of String, String) In elemento

            str_key = obj_riga.Key.ToUpper.Trim
            If _lista.ContainsKey(str_key) Then
                'escludi o aggiorna ?!?!?
                bool_return = False
            Else
                Item(str_key) = obj_riga.Value
            End If
        Next

        Return bool_return
    End Function

    Public Function Clona() As cls_ListaElementi
        Dim obj_elemento As New cls_ListaElementi

        If _lista IsNot Nothing Then
            For Each obj_item As KeyValuePair(Of String, String) In _lista
                obj_elemento.Item(obj_item.Key) = obj_item.Value
            Next
        End If
        Return obj_elemento
    End Function


End Class

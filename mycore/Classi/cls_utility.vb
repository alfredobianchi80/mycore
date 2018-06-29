Public Class cls_utility

    Public Shared Function ProssimaSequenza(ByVal SequenzaAttuale As Int32, Optional ByVal PassoSequenza As Int32 = 10) As Int32
        Dim lng_NuovaSequenza As Int32 = 0

        Try
            lng_NuovaSequenza = CInt(SequenzaAttuale / PassoSequenza)
            lng_NuovaSequenza += 1
            lng_NuovaSequenza = lng_NuovaSequenza * PassoSequenza
        Catch ex As Exception
            lng_NuovaSequenza = -1
        End Try

        Return lng_NuovaSequenza

    End Function


    Public Shared Function DBNullIfEmpty(ByVal valore As String) As Object
        If valore.Length = 0 Then
            Return DBNull.Value
        Else
            Return valore
        End If
    End Function

#Region "Funzioni per RETRO-Compatibilità"

    Shared Function ListBox_AssegnaByID(ByRef Lista As ListBox, ByVal ID As String) As Boolean
        Return cls_ItemMultiValue.ListBox_AssegnaByID(Lista, ID)
    End Function


    Shared Function ComboMulti_AssegnaByValore(ByRef Combo As ComboBox, ByVal Valore As String, Optional ByVal IgnoraMaiuscole As Boolean = True) As Boolean
        Return cls_ItemMultiValue.Combo_AssegnaByValore(Combo, Valore, IgnoraMaiuscole)
    End Function

    Shared Function ComboMulti_AssegnaByID(ByRef Combo As ComboBox, ByVal ID As String, Optional ByVal IgnoraMaiuscole As Boolean = True) As Boolean
        Return cls_ItemMultiValue.Combo_AssegnaByID(Combo, ID, IgnoraMaiuscole)
    End Function

#End Region '"Funzioni per RETRO-Compatibilità"
    


End Class

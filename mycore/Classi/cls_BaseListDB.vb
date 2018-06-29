Public MustInherit Class cls_BaseListDB(Of T)
    Inherits cls_BaseList(Of T)

    Shared Function CreaFromDB(ByRef DBConnection As System.Data.Common.DbConnection, ByRef Valore As T) As T

    End Function

    Public MustOverride Function SaveToDB(ByRef DBConnection As Common.DbConnection, ByVal Valore As T) As Boolean

    'Shared MustOverride function 


End Class



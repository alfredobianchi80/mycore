Imports System.Reflection

'*****************************************************************************************************
'***** Code By dukearena
'***** http://forum.masterdrive.it/visual-basic-net-18/deepclone-manuale-52248/#post217226
'*****
'*****
'***** USAGE EXAMPLE:
'*****
'*****     Public Function Clone() As Object Implements System.ICloneable.Clone
'*****         Dim CloneClass As New Clonatore(Of Object)
'*****         Return CloneClass.DeepClone(Me)
'*****     End Function
'*****
'*****************************************************************************************************

Public Class cls_Clonatore(Of T)
    ' Controlla se un'oggetto supporta una determinata interfaccia
    Private Function CanSupportInterface(ByVal IInterfaccia As eInterfacce, ByRef FInfo As Object) As Boolean
        Dim ICloneType As Type = FInfo.GetType.GetInterface(IInterfaccia.ToString, True)
        Return ICloneType IsNot Nothing
    End Function

    Private Enum eInterfacce
        ICloneable
        IEnumerable
        IList
        IDictionary
    End Enum

    Public Function DeepClone(ByRef OggettoPadre As T) As T
        'Creo una nuova istanza
        Dim Risultato As T = Activator.CreateInstance(OggettoPadre.GetType)
        'Ottengo un array che conterrà i campi della nuova istanza
        Dim Fields() As System.Reflection.FieldInfo = Risultato.GetType.GetFields()

        Dim i As Int32 = 0
        For Each FI As System.Reflection.FieldInfo In OggettoPadre.GetType.GetFields()
            'Controllo se l'oggetto supporta il clone
            If CanSupportInterface(eInterfacce.ICloneable, FI) Then
                'Ottengo l'interfaccia ICloneable dall'oggetto
                Dim IClone As ICloneable = FI.GetValue(OggettoPadre)
                'Utilizzo il metodo clone per settare il nuovo valore al campo
                Fields(i).SetValue(Risultato, IClone.Clone())
            Else
                'Se il campo non suppporta ICloneable allora lo setto e basta
                Fields(i).SetValue(Risultato, FI.GetValue(OggettoPadre))
            End If

            'Controllo se supporta IEnumerable. Se si, devo enumerare tutti gli indici e controllare se usano ICloneable
            If CanSupportInterface(eInterfacce.IEnumerable, FI) Then
                'Ottengo IEnumerable dal campo
                Dim IEnum As IEnumerable = FI.GetValue(OggettoPadre)

                Dim j As Int32 = 0
                If CanSupportInterface(eInterfacce.IList, FI) Then
                    'Ottengo l'Interfaccia IList
                    Dim Lista As IList = Fields(i).GetValue(Risultato)
                    For Each Obj As Object In IEnum
                        'Controllo se l'oggetto corrente supporta ICloneable Interface
                        'ICloneType = Obj.GetType.GetInterface("ICloneable", True)
                        If CanSupportInterface(eInterfacce.ICloneable, Obj) Then
                            'Se possiede ICloneable, eseguiamo Clone
                            Dim Clone As ICloneable = Obj
                            Lista(j) = Clone.Clone
                            'NOTA: Se l'oggetto non possiede ICloneable allora sarà lo stesso oggetto della lista originario
                        End If
                        j += 1
                    Next

                ElseIf CanSupportInterface(eInterfacce.IDictionary, FI) Then
                    Dim Dic As IDictionary = Fields(i).GetValue(Risultato)
                    j = 0
                    For Each DE As DictionaryEntry In IEnum
                        'Controllo se è possibile usare ICloneable
                        If CanSupportInterface(eInterfacce.ICloneable, DE.Value) Then
                            Dim Clone As ICloneable = DE.Value
                            Dic(DE.Key) = Clone.Clone
                        End If
                        j += 1
                    Next
                End If
            End If
            i += 1
        Next
        Return Risultato
    End Function
End Class

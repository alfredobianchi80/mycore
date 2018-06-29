Public Class cls_INIFIle
    Inherits Dictionary(Of String, Gruppo)


#Region "SottoClassi"

    Public Class Chiave
        Private m_nome As String
        Private m_valore As String

        Public Property Nome() As String
            Get
                Return m_nome
            End Get
            Set(ByVal value As String)
                m_nome = value
            End Set
        End Property

        Public Property Valore() As String
            Get
                Return m_valore
            End Get
            Set(ByVal value As String)
                m_valore = value
            End Set
        End Property

        Public Sub New(ByVal nomeChiave As String, ByVal valoreChiave As String)

            m_nome = nomeChiave
            m_valore = valoreChiave

        End Sub
    End Class

    Public Class Gruppo
        Inherits List(Of Chiave)

        'Rappresenta un gruppo coerente di coppie "chiave=valore"

        Private m_nome As String

        Public Sub New(ByVal nome As String)
            m_nome = nome
        End Sub

        Public Property Nome() As String
            Get
                Return m_nome
            End Get
            Set(ByVal value As String)
                m_nome = value
            End Set
        End Property

        Public Function GetChiaveByNome(ByVal nomeChiave As String) As Chiave

            For Each C As Chiave In Me
                If C.Nome = nomeChiave Then
                    Return C
                End If
            Next
            Return Nothing

        End Function

    End Class

#End Region '"SottoClassi"

    Private m_nomeFileCompleto As String


#Region "Costruttore"
    Public Sub New(ByVal nomeFileCompleto As String)
        m_nomeFileCompleto = nomeFileCompleto
    End Sub

    Public Sub New()
        m_nomeFileCompleto = ""
    End Sub

#End Region '"Costruttore"


    Public Sub Carica(Optional ByVal nomefile As String = "")
        Dim str_NomeFile As String = ""
        Dim str_key As String = ""
        Dim str_LineaLetta As String = ""

        If nomefile.Length > 0 Then
            str_NomeFile = nomefile
        Else
            str_NomeFile = m_nomeFileCompleto
        End If

        Me.Clear()

        Using TR As New System.IO.StreamReader(str_NomeFile, System.Text.Encoding.GetEncoding(1252))
            Dim str_LineaFile As String = ""
            Dim n As String = ""
            Dim v As String = ""
            Dim nv As String()
            Dim G As Gruppo = Nothing
            Do Until str_LineaFile Is Nothing
                Try
                    str_LineaFile = TR.ReadLine()
                    str_LineaLetta = Trim(str_LineaFile)
                    If str_LineaLetta = "" Then Continue Do
                    If str_LineaLetta(0) = ";" Then Continue Do 'Gestisci linee di commento

                    If str_LineaLetta(0) = "[" And str_LineaLetta(str_LineaLetta.Length - 1) = "]" Then
                        str_key = str_LineaLetta.Substring(1, str_LineaLetta.Length - 2)
                        If Not Me.ContainsKey(str_key) Then
                            G = New Gruppo(str_key)
                            Me.Add(str_key, G)
                        End If
                    Else
                        nv = str_LineaLetta.Split("=")
                        If nv.Length > 1 Then 'Gestisci casi strani
                            n = Trim(nv(0))
                            v = Trim(nv(1))
                        End If
                        G.Add(New Chiave(n, v))
                    End If

                Catch ex As Exception
                    MessageBox.Show("errore lettura file ini")
                End Try
            Loop
            TR.Close()
        End Using

    End Sub

    Public Sub SalvaNuovo(Optional ByVal nomefile As String = "")
        Dim str_NomeFile As String = ""
        If nomefile.Length > 0 Then
            str_NomeFile = nomefile
        Else
            str_NomeFile = m_nomeFileCompleto
        End If

        If Me.Count = 0 Then Exit Sub
        Using TW As New System.IO.StreamWriter(str_NomeFile, False, System.Text.Encoding.GetEncoding(1252))
            For Each G As Gruppo In Me.Values
                TW.WriteLine(String.Concat("[", G.Nome, "]"))
                For Each C As Chiave In G
                    TW.WriteLine(String.Concat(C.Nome, "=", C.Valore))
                Next
                TW.WriteLine("")
            Next
            TW.Flush()
            TW.Close()
        End Using

    End Sub

    Public Function GetGruppoByNome(ByVal nomeGruppo As String) As Gruppo

        For Each G As Gruppo In Me.Values
            If G.Nome = nomeGruppo Then
                Return G
            End If
        Next
        Return Nothing

    End Function

    Private Function SalvaParametro(ByVal Gruppo As String, ByVal NomeParametro As String, ByVal Valore As String) As Boolean
        Dim bool_Risultato As Boolean = False

        'Trova Gruppo nel file



        Return bool_Risultato
    End Function


End Class

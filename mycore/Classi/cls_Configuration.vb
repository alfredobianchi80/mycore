Public Class cls_Configuration

#Region "Variabili"

    Event ClasseNonInizializzata()

    Private _ElencoParametri As Dictionary(Of String, String)
    Private _Inizializzato As Boolean = False
    Private _Error As Boolean = False
    Private _NoExistCode As String = "#@####&&&&55545##@#@#@#"

#End Region '"Variabili"

#Region "Enumerati"

    Public Enum TipoFileConfig
        XML = 1
        INI = 2
    End Enum

#End Region '"Enumerati"

#Region "Costruttori"

    Public Sub New()
        _ElencoParametri = New Dictionary(Of String, String)
        _Inizializzato = True
    End Sub

#End Region '"Costruttori"

#Region "Property"

    ReadOnly Property NoExistCode
        Get
            Return _NoExistCode
        End Get
    End Property

    ReadOnly Property isError
        Get
            Return _Error
        End Get

    End Property

    Property Item(ByVal NomeParametro As String)
        Get
            _Error = False
            NomeParametro = NomeParametro.ToUpper
            'Se la classe non è inizializzata --> Errore
            If Not _Inizializzato Then
                MessageBox.Show(String.Concat("Classe non inizializzata. Parametro ", NomeParametro, " non accessibile"), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _Error = True
                RaiseEvent ClasseNonInizializzata()
                Return ""
            Else
                'Verifica se il parametro esiste, in caso contrario ritorna vuoto + Errore (?!?)
                If _ElencoParametri.ContainsKey(NomeParametro) Then
                    Return _ElencoParametri(NomeParametro)
                Else
                    _Error = True
                    Return _NoExistCode
                End If

            End If
        End Get

        Set(value)
            _Error = False
            NomeParametro = NomeParametro.ToUpper
            'Se la classe non è inizializzata --> Errore
            If Not _Inizializzato Then
                MessageBox.Show(String.Concat("Classe non inizializzata. Parametro ", NomeParametro, " non accessibile"), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _Error = True
                RaiseEvent ClasseNonInizializzata()
            Else
                If _ElencoParametri.ContainsKey(NomeParametro) = False Then
                    _ElencoParametri.Add(NomeParametro, value)
                Else
                    _ElencoParametri(NomeParametro) = value
                End If

            End If
        End Set
    End Property

    Public Sub ClearErr()
        _Error = False
    End Sub

#End Region '"Property"

#Region "Funzioni"

    Public Function CaricaFromFile(ByVal TipoFile As TipoFileConfig, ByVal nomeFile As String, Optional ByVal appendi As Boolean = False)
        Dim bool_Risultato As Boolean = False
        Dim str_OutPutMSG As String = ""

        If _Error Then
            Return False
        End If

        If System.IO.File.Exists(nomeFile) = False Then
            MessageBox.Show(String.Concat("il file «", nomeFile, "» non esiste."))
            Return False
        End If

        'Gestisci Caricamento in base al tipo indicato
        Select Case TipoFile
            Case TipoFileConfig.XML
                bool_Risultato = LoadFrom_XML(nomeFile, appendi)

            Case TipoFileConfig.INI
                bool_Risultato = LoadFrom_INI(nomeFile, appendi)
            Case Else
                str_OutPutMSG = "Tipo di file di configurazione non supportato"
                bool_Risultato = False

        End Select

        'Termina e Esci
        If str_OutPutMSG.Length > 0 Then
            MessageBox.Show(str_OutPutMSG)
        End If

        Return bool_Risultato
    End Function

#End Region '"Funzioni"

#Region "Altro"

    Private Function LoadFrom_INI(ByVal nomefile As String, Optional ByVal Appendi As Boolean = False) As Boolean
        Dim bool_Risultato As Boolean = False
        Dim obj_ini As New cls_INIFile()
        Dim str_key As String

        If Not (Appendi) Then
            _ElencoParametri.Clear()
        End If

        Try
            'Leggi file INI
            obj_ini.Carica(nomefile)
            bool_Risultato = True
        Catch ex As Exception
            bool_Risultato = False
        End Try

        If bool_Risultato Then
            Try
                'Trasforma Oggetto Letto in variabili
                For Each obj_Gruppo As KeyValuePair(Of String, cls_INIFile.Gruppo) In obj_ini

                    For Each obj_Chiave As cls_INIFile.Chiave In obj_Gruppo.Value
                        str_key = String.Concat(obj_Gruppo.Value.Nome, ".", obj_Chiave.Nome).ToUpper
                        _ElencoParametri.Add(str_key, obj_Chiave.Valore)
                    Next
                Next
            Catch ex As Exception
                bool_Risultato = False
            End Try
        End If


        '*** x debug : INIZIO
        'MessageBox.Show("config.INIT=«" & _ElencoParametri("CONFIG.INIT") & "»")
        'obj_ini.SalvaNuovo("D:\saved.ini")
        '*** x debug : FINE

        bool_Risultato = True
        Return bool_Risultato
    End Function

    Private Function LoadFrom_XML(ByVal nomefile As String, Optional ByVal Appendi As Boolean = False) As Boolean
        Dim bool_Risultato As Boolean = False
        Dim obj_xml As New Object
        Dim str_key As String = ""

        If Not (Appendi) Then
            _ElencoParametri.Clear()
        End If

        Try
            'Leggi file XML
            'obj_xml.Carica(nomefile)
            bool_Risultato = True
        Catch ex As Exception
            bool_Risultato = False
        End Try

        If bool_Risultato Then
            Try
                'Trasforma Oggetto Letto in variabili
                'For Each obj_Gruppo As KeyValuePair(Of String, cls_INIFile.Gruppo) In obj_ini
                '  For Each obj_Chiave As cls_INIFile.Chiave In obj_Gruppo.Value
                '        str_key = String.Concat(obj_Gruppo.Value.Nome, ".", obj_Chiave.Nome)
                '        _ElencoParametri.Add(str_key, obj_Chiave.Valore)
                '   Next
                'Next
            Catch ex As Exception
                bool_Risultato = False
            End Try
        End If

        Return bool_Risultato
    End Function

#End Region '"Altro"

End Class

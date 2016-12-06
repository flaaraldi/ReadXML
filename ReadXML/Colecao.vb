Imports System.Runtime.InteropServices
<Microsoft.VisualBasic.ComClass()> <System.Serializable()> Public Class Colecao

    Inherits System.Collections.DictionaryBase
    <System.NonSerialized()> <System.Runtime.InteropServices.DispId(1)> Dim xitem As New Propriedades
    Function lengh() As Integer
        Return Me.Dictionary.Count
    End Function
    Public Sub Add(ByVal key As String, ByVal value As Propriedades)

        Me.Dictionary.Add(key.ToLower, value)
    End Sub
    Default Public Property Item(ByVal key As String) As Propriedades

        Get

            If CType(Me.Dictionary(key.ToLower), Propriedades) Is Nothing Then
                xitem.nome = ""
                xitem.atributo = ""
                xitem.valor = ""
                Return xitem
            Else
                Return CType(Me.Dictionary(key.ToLower), Propriedades)
            End If

        End Get
        Set(ByVal value As Propriedades)
            'If IsNumeric(key) And key >= 0 Then
            If Not String.IsNullOrEmpty(key) Then
                Me.Dictionary(key.ToLower) = value
            Else
                Throw New IndexOutOfRangeException()
            End If
        End Set
    End Property
    Function BuscaTag(ByVal xtag As String, ByVal xcolecao As Colecao) As Integer
        Dim xaux As Integer
        xaux = 0

        For Each xitem As DictionaryEntry In xcolecao ' procura tag 
            Dim item() As String = xitem.Key.ToString.Split(":")
            If item(1).ToLower.Equals(xtag.ToLower) Then
                xaux = xaux + 1
            End If
        Next

        Return xaux
    End Function
End Class


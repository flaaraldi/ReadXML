Imports System.Runtime.InteropServices
<Microsoft.VisualBasic.ComClass()> <System.Serializable()> Public Class Propriedades
    Public m_valor As String
    Public m_nome As String
    Public m_atributo As String
    Public Property valor() As String
        Get
            Return m_valor
        End Get
        Set(ByVal value As String)
            m_valor = value
        End Set
    End Property
    Public Property atributo() As String
        Get
            Return m_atributo
        End Get
        Set(ByVal value As String)
            m_atributo = value
        End Set
    End Property
    Public Property nome() As String
        Get
            Return m_nome
        End Get
        Set(ByVal value As String)
            m_nome = value
        End Set
    End Property
End Class

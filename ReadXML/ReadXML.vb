Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
<Microsoft.VisualBasic.ComClass()> <System.Serializable()> Public Class ReadXML

    Public Function LeArquivoXMLPorTag(ByVal arquivo As String, ByVal xtag As String) As Colecao()
        Return pLeArquivoXMLPorTag(arquivo, xtag, False)
    End Function

    Public Function LeArquivoXMLPorTagNTags(ByVal arquivo As String, ByVal xtag As String) As Colecao()
        Return pLeArquivoXMLPorTag(arquivo, xtag, True)
    End Function

    Private Function pLeArquivoXMLPorTag(ByVal arquivo As String, ByVal xtag As String, ByVal bTagsRepete As Boolean) As Colecao()

        Dim reader As XmlNodeReader = Nothing
        Dim doc As New XmlDocument()
        doc.Load(arquivo)
        reader = New XmlNodeReader(doc)

        Dim node As XmlNode
        Dim itens As Integer
        Dim Item As Integer
        Dim colecao(0) As Colecao
        Dim xcolecao As New Colecao
        Dim xitens As New Propriedades
        Dim tag As String
        Dim tagP As String
        Dim colecaoaux As New Colecao 'coleção com todos os valores de nods, para pode buscar a tag do item correto, pois podem ter varia tags do msm valor
        Dim itemaux As Integer
        Dim itenx As New Propriedades
        Dim i As Integer = 0
        Dim contNFref As Integer = 0
        Dim contRefNF As Integer = 0
        Dim contRefNFP As Integer = 0
        Dim contTagRepete As Integer = 1
        itemaux = 0
        Item = 0
        itens = 0
        tagP = ""
        tag = ""

        Do While (reader.Read()) ' leio ateh o final do arquivo

            If reader.NodeType = XmlNodeType.Element Then ' se for um elemento adicono a coleção auxliar
                itemaux = itemaux + 1
                itenx = New Propriedades
                itenx.nome = reader.Name
                colecaoaux.Add("aux:" & reader.Name & ":" & itemaux, itenx) ' adiciono a colecao

                If reader.Name = xtag Then ' achei a tag procurada leio todos os nós q estão dentro dela
                    'exemplo, "det" , onde cada det possui um produto 
                    ReDim Preserve colecao(Item)
                    reader.Read() 'leio o próximo nó que esta dentro da det
                    If reader.NodeType = XmlNodeType.Element Then ' exemplo o prod
                        itemaux = itemaux + 1
                        itenx = New Propriedades
                        itenx.nome = reader.Name
                        colecaoaux.Add("aux:" & reader.Name & ":" & itemaux, itenx) ' adiciono a colecao
                    End If

                    xcolecao = New Colecao
                    'na coleção para identificar , nome da tag pai: tag filha exemplo : ("prod:cProd")
                    'no caso do lote pode existir mais de um lote para cada produto
                    ' ("med:nlote:1") e gravo o total de itens de lote para poder adicionar depois  (med:totalLote)

                    Do While reader.Name <> xtag Or reader.NodeType <> XmlNodeType.EndElement ' passo por todos elementos até encontrar a tag final  
                        If reader.NodeType = XmlNodeType.Element Then

                            If reader.MoveToContent = XmlNodeType.Element Then

                                itens = xcolecao.BuscaTag(reader.Name, colecaoaux) 'para achar qual item, pois pode ter inumeras tag com msm nome
                                'node = doc.GetElementsByTagName(reader.Name)

                                node = doc.GetElementsByTagName(reader.Name).Item(itens - 1) ' crio node
                                tagP = node.ParentNode.Name 'nome tag pai
                                If reader.Name = "nLote" Then
                                    i = i + 1
                                End If


                            Else
                                node = Nothing
                            End If
                            tag = reader.Name ' nome tag filha
                        End If

                        If reader.NodeType = XmlNodeType.Text Then 'se o proximo eh um eletemnto com valores do elmento anterior , gravo
                            xitens = New Propriedades
                            xitens.nome = tag

                            If tagP.Equals("NFref") Then
                                contNFref += 1
                            ElseIf tagP.Equals("refNF") Then
                                contRefNF += 1
                            ElseIf tagP.Equals("refNFP") Then
                                contRefNFP += 1
                                'ElseIf bTagsRepete Then
                                'contTagRepete += 1
                            End If

                            xitens.valor = reader.Value ' valor do elemento
                            If tagP = "med" Then
                                xcolecao.Add(tagP & ":" & tag & ":" & i, xitens) ' adiciono a colecao
                            ElseIf tagP = "NFref" Then
                                xcolecao.Add(tagP & ":" & tag & ":" & contNFref, xitens) ' adiciono a colecao
                            ElseIf tagP = "refNF" Then
                                xcolecao.Add(tagP & ":" & tag & ":" & contRefNF, xitens) ' adiciono a colecao
                            ElseIf tagP = "refNFP" Then
                                xcolecao.Add(tagP & ":" & tag & ":" & contRefNFP, xitens) ' adiciono a colecao
                            Else
                                If bTagsRepete Then
                                    contTagRepete = 1
                                    'Se já existir, então começa a iterar
                                    If xcolecao.Item(tagP & ":" & tag & ":" & contTagRepete).valor <> "" Then
                                        Do While xcolecao.Item(tagP & ":" & tag & ":" & contTagRepete + 1).valor <> "" 'Enquanto o próximo não for vazio, itera
                                            contTagRepete += 1
                                        Loop
                                        xcolecao.Add(tagP & ":" & tag & ":" & contTagRepete + 1, xitens) ' adiciono na proxima posição da colecao, que é vazio
                                    Else
                                        xcolecao.Add(tagP & ":" & tag & ":1", xitens) ' adiciono a colecao
                                    End If
                                Else
                                    If xcolecao.Item(tagP & ":" & tag).valor = "" Then
                                        xcolecao.Add(tagP & ":" & tag, xitens) ' adiciono a colecao
                                    End If
                                End If
                            End If

                        End If
                        reader.Read() 'próximo objeto

                        If reader.NodeType = XmlNodeType.Element Then
                            itemaux = itemaux + 1
                            itenx = New Propriedades
                            itenx.nome = reader.Name
                            colecaoaux.Add("aux:" & reader.Name & ":" & itemaux, itenx) ' adiciono a colecao
                        End If

                    Loop
                    If i > 0 Then
                        xitens = New Propriedades
                        xitens.nome = "totalLote"
                        xitens.valor = i ' valor do elemento
                        xcolecao.Add("med:totallote", xitens) ' adiciono a colecao
                    End If



                    colecao(Item) = xcolecao
                    Item = Item + 1
                    i = 0
                End If
            End If
        Loop

        'para teste
        'For i As Integer = 0 To colecao.Count - 1

        ' ListBox1.Items.Add("prodcprod:" & i & "  " & "  codigo:" & colecao(i).Item("prod:cprod").valor)
        'Next


        Return colecao


    End Function

    Public Function LeAtributosXMLPorTagEspecifica(ByVal arquivo As String, ByVal xtag As String, ByVal itemTag As Integer) As Colecao
        Dim colecao As New Colecao
        Dim doc As New XmlDocument()
        Dim node As XmlNode
        Dim ok As Boolean = False

        doc.Load(arquivo)
        node = doc.GetElementsByTagName(xtag).Item(itemTag)

        If Not node Is Nothing Then

            For Each xatributo As XmlAttribute In node.Attributes
                Dim xitens As New Propriedades
                xitens.nome = node.Name
                xitens.atributo = xatributo.Value
                colecao.Add(xtag & ":" & xatributo.Name, xitens)
            Next

        Else
            Dim xitens As New Propriedades
            xitens.nome = ""
            xitens.atributo = ""
            colecao.Add(xtag, xitens)
        End If

        Return colecao
    End Function

    Private Sub ctrl_le_arq_xml_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim xcol() As Colecao = LeArquivoXMLPorTag("C:\35150407626135000106550010001060021980835750-nfe.xml", "ide")
        Dim xNota As Colecao = LeAtributosXMLPorTagEspecifica("C:\35150407626135000106550010001060021980835750-nfe.xml", "infNFe", 0)

        MsgBox(xcol.Count - 1)
        ''

    End Sub
End Class

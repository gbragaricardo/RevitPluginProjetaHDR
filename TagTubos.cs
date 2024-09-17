using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using System.Windows;
using Autodesk.Revit.DB.Plumbing;
using System.Windows.Documents;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;


namespace ProjetaHDR
{
    [Transaction(TransactionMode.Manual)]
    internal class TagTubos : IExternalCommand
    {
        //create global variables
        UIApplication _uiapp;
        Autodesk.Revit.ApplicationServices.Application _app;
        UIDocument _uidoc;
        Document _doc;
        View _vistaAtiva;

        #region MAIN

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Creating App and doc objects.
            _uiapp = commandData.Application;
            _app = _uiapp.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;

            _vistaAtiva = _doc.ActiveView; // Variavel global para a vista ativa

            using (Transaction transacao = new Transaction(_doc, "Detalhar tubos"))
            {

                bool linhaDeChamada = false; // Variavel booleana declarada com false para negar a opção de linha de chamada

                TagOrientation tagOrientacao = TagOrientation.Horizontal; // Variavel declarada com a orientação trabalhada pelo revit como Horizontal para todas as tags



                string detalheEscolhido = InputEscolhaDetalhe();

                if (detalheEscolhido != null)
                {

                    ElementId IdTipoTag = ObterIdTags(detalheEscolhido);

                    // Instanciação da lista de tubos (Element) da vista ativa
                    IList<Element> tubosNaVista = new List<Element>(SelecionarTubosNaVista());

                    // Instanciação da lista de tubos (Element) acima de 20cm de comprimento
                    IList<Element> filtradosPorComprimento = FiltrarPorComprimento(tubosNaVista);

                    // Instanciação da lista de tubos (Element) com remoção dos tubos verticais
                    IList<Element> filtradorPorComprimentoEVerticaisRemovidos = RemoverVerticais(filtradosPorComprimento);

                    // Instanciação da lista de orientações (string) dos tubos em planta baixa (Horizontal, Vertical, Ou Diagonais)
                    IList<string> orientacoesEmPlanta = VerificarOrientacaoEmPlanta(filtradorPorComprimentoEVerticaisRemovidos);

                    // Instanciação da lista XYZ dos pontos de inserção das Tags
                    IList<XYZ> pontosDeInsercaoTags = PontoInsercaoTag(filtradorPorComprimentoEVerticaisRemovidos, orientacoesEmPlanta, detalheEscolhido);



                    transacao.Start(); // Inicio da transação que faz modificações do Projeto/Arquivo

                    // Uso da variavel estatica que remove as tags do mesmo ID da que será utilizada que estão nos mesmos tubos que serão utilizados
                    RemoverTagsExistentes(_doc, _vistaAtiva, filtradorPorComprimentoEVerticaisRemovidos, IdTipoTag);

                    // Criação das Tags
                    CriarTags(filtradorPorComprimentoEVerticaisRemovidos, IdTipoTag, _vistaAtiva.Id, linhaDeChamada, tagOrientacao, pontosDeInsercaoTags);

                    SetarValorAoParametroInclinacao(filtradorPorComprimentoEVerticaisRemovidos);

                    transacao.Commit(); // Fim da transação que faz modificações do Projeto/Arquivo
                }
            }

            return Result.Succeeded; //Retorno padrao do metodo execute
        }
        #endregion


        #region InputUsuario
        public string InputEscolhaDetalhe()
        {
            string tipoEscolhido = "";

            // Cria uma nova TaskDialog
            TaskDialog entradaTipoTag = new TaskDialog("Escolha a TAG");

            // Define a instrução principal do diálogo
            entradaTipoTag.MainInstruction = "Selecione uma opção";

            // Define os botões comuns (OK e Cancel)
            TaskDialogCommonButtons buttons = TaskDialogCommonButtons.Cancel;
            entradaTipoTag.CommonButtons = buttons;

            // Adiciona os comandos de links (opções) que o usuário pode escolher
            entradaTipoTag.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Diâmetro");
            entradaTipoTag.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Inclinação ");

            // Exibe o diálogo e captura a resposta do usuário
            TaskDialogResult result = entradaTipoTag.Show();

            if (result == TaskDialogResult.Cancel || result == TaskDialogResult.Close)
            {
                // Se o usuário clicar em Cancel, encerra o script retornando null ou lançando uma exceção
                return null;
            }

            switch (result)
            {
                case TaskDialogResult.CommandLink1:
                    tipoEscolhido = "Diametro";
                    break;
                case TaskDialogResult.CommandLink3:
                    tipoEscolhido = "Inclinacao";
                    break;
                case TaskDialogResult.Cancel:
                    tipoEscolhido = "Cancelado";
                    break;

            }

            return tipoEscolhido;
        }
        #endregion


        #region ObterIdTags

        public ElementId ObterIdTags(string detalheEscolhido)
        {

            ElementId idTag = null;
            IList<Element> tiposDeTag = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_PipeTags).WhereElementIsElementType().ToElements();

            foreach (Element tipoDeTag in tiposDeTag)
            {
                if (tipoDeTag.Name == detalheEscolhido)
                {
                    idTag = tipoDeTag.Id;
                }
            }

            if (idTag == null)
            {
                TaskDialog.Show("Erro", $"Tag com nome {detalheEscolhido} não encontrada");
                return null;
            }
            else
            {
                return idTag;
            }

        }

        #endregion


        #region SelecionarTubosNaVista

        /// <summary>
        /// Seleciona todos os tubos da vista ativa
        /// </summary>
        /// <returns>Lista de Element</returns>

        public IList<Element> SelecionarTubosNaVista()
        {
            _vistaAtiva = _doc.ActiveView;

            IList<Element> tubosNaVista = new FilteredElementCollector(_doc, _vistaAtiva.Id).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements();


            return tubosNaVista;
        }
        #endregion


        #region FiltrarPorComprimento

        /// <summary>
        /// Filtra tubos com relação ao seu comprimento
        /// </summary>
        /// <param name="tubosNaVista"></param> Recebe tubos selecionados da vista ativa
        /// <returns>Lista de Elemnent</returns>

        public IList<Element> FiltrarPorComprimento(IList<Element> tubosNaVista)
        {
            IList<Element> filtrados = new List<Element>(); // Instancia uma lista vazia que será o retorno do tipo lista de Element

            foreach (Element tubo in tubosNaVista) // Percorre cada tubo da vista ativa
            {
                Parameter parametroComprimento = tubo.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH); // parametroComprimento recebe o parametro incorporado "comprimento" do revit
                double valorComprimentoMetros = parametroComprimento.AsDouble() * 0.3048; // Recebe o valor do parametro comprimento convertendo para metros pois o revit utiliza pés :3

                if (valorComprimentoMetros > .2) // Condiciona tubos acima de 20 centimetros.
                {
                    filtrados.Add(tubo); // Adiciona a lista cada tubo da vista ativa maior que 20 centimetros
                }
            }

            return filtrados;
        }
        #endregion


        #region RemoverVerticais

        /// <summary>
        /// Filtra para que os tubos verticais da vista ativa sejam descartados
        /// </summary>
        /// <param name="tubosNaVistaFiltrados"></param> Tubos ativos na vista que já foram filtrados pelo comprimento
        /// <returns> Lista de Element</returns>

        public IList<Element> RemoverVerticais(IList<Element> tubosNaVistaFiltrados)
        {
            IList<Element> filtrados = new List<Element>(); // Instancia uma lista vazia que será o retorno do tipo lista de Element

            foreach (Element tubo in tubosNaVistaFiltrados) // Percorre todos os tubos recebidos pelo parametro tubosNaVistaFiltrados
            {
                Location tuboLocation = tubo.Location; // Obtem a localização no modelo de cada tubo 

                LocationCurve localCurvaTubo = (LocationCurve)tuboLocation; // a partir da localização dos tubos Atribui uma variavel da classe LocationCurve para uso dos metodos apropriados

                Curve curvaTubo = localCurvaTubo.Curve; // Atribui uma variavel da classe Curve a partir de LocationCurve

                XYZ pontoInicial = curvaTubo.GetEndPoint(0); // Obtem o ponto inicial da curva
                XYZ pontoFinal = curvaTubo.GetEndPoint(1); // Obtem o ponto final da curva

                // Condiciona os tubos que tem variação no Eixo Y ou Eixo X, removendo assim tubos verticais pois só variam em Z

                if (Math.Round(pontoInicial.X) != Math.Round(pontoFinal.X) || Math.Round(pontoInicial.Y) != Math.Round(pontoFinal.Y))
                {
                    filtrados.Add(tubo); // Adiciona a lista de retorno os tubos que não sao verticais
                }
            }

            return filtrados;
        }
        #endregion


        #region VerificarOrientacaoEmPlanta

        /// <summary>
        /// Verifica em planta quais tubos estão na horizontal, vertical ou diagonal,
        /// diferenciando em Diagonal positiva e negativa, necessário para que o deslocamento da tag seja aplicado corretamente
        /// </summary>
        /// <param name="tubosNaVistaFiltrados"></param> Tubos filtrados tanto por comprimento quanto pela remoção dos verticais
        /// <returns> Lista de string</returns>

        public IList<string> VerificarOrientacaoEmPlanta(IList<Element> tubosNaVistaFiltrados)
        {

            IList<string> orientacaoTubos = new List<string>();  // Instancia uma lista vazia que será o retorno do tipo lista de string

            double margemDePrecisao = 0.01; // declara uma variavel double que será a margem de precisão utilizada

            foreach (Element tubo in tubosNaVistaFiltrados) // Percorre cada tubo recebido como parametro na lista tubosNaVistaFiltrados
            {

                Location tuboLocation = tubo.Location; // Obtem a localização no modelo de cada tubo
                LocationCurve localCurvaTubo = (LocationCurve)tuboLocation; // a partir da localização dos tubos Atribui uma variavel da classe LocationCurve para uso dos metodos apropriados
                Curve curvaTubo = localCurvaTubo.Curve; // Atribui uma variavel da classe Curve a partir de LocationCurve

                XYZ pontoInicial = curvaTubo.GetEndPoint(0); // Recebe o ponto inicial de cada curva
                XYZ pontoFinal = curvaTubo.GetEndPoint(1); // Recebe o ponto final de cada curva

                double deltaX = Math.Abs(pontoFinal.X - pontoInicial.X); // Recebe o delta do Eixo X, utilizando o valor absoluto dos pontos em X e convertendo em double 
                double deltaY = Math.Abs(pontoFinal.Y - pontoInicial.Y); // Recebe o delta do Eixo Y, utilizando o valor absoluto dos pontos em Y e convertendo em double 

                if (pontoInicial != null && pontoFinal != null) // Verifica que os pontos iniciais e finais existem
                {

                    // Verifica se o valor absoluto do delta x - delta y é menor que a margem de precisao
                    // Pois assim é possível verificar se percorrem em um angulo de 45° pois percorrem a mesma quantidade no eixo x e y. Portanto sao diagonais

                    if (Math.Abs((deltaX - deltaY)) < margemDePrecisao)
                    {

                        // Verifica se o produto de cada Delta (X e Y) é positivo (Maior que zero)
                        // Pois assim é possivel constatar que o tubo está numa posição similar à " / "
                        // Caso contrario o tubos estará em posição mais proxima a uma barra invertida " \ "

                        if ((pontoFinal.X - pontoInicial.X) * (pontoFinal.Y - pontoInicial.Y) > 0)
                        {
                            orientacaoTubos.Add("Diagonal Positiva"); //Adiciona a orientação do tubo a lista
                        }
                        else
                        {
                            orientacaoTubos.Add("Diagonal Negativa"); //Adiciona a orientação do tubo a lista
                        }

                    }

                    // Caso o tubos não seja em Diagonal e percorra uma valor maior no Eixo X em relãção ao eixo Y, ele será classificado como Horizontal

                    else if (deltaX > deltaY)
                    {
                        orientacaoTubos.Add("Horizontal"); //Adiciona a orientação do tubo a lista
                    }

                    // Caso o tubos não seja em Diagonal e percorra uma valor maior no Eixo Y em relãção ao Eixo X, ele será classificado como Vertical

                    else
                    {
                        orientacaoTubos.Add("Vertical"); //Adiciona a orientação do tubo a lista
                    }

                }

                // Else com classificação N/A apenas por segurança

                else
                {
                    orientacaoTubos.Add("N/A"); //Adiciona a orientação nula do tubo a lista
                }

            }

            return orientacaoTubos;
        }
        #endregion


        #region PontoInsercaoTag

        /// <summary>
        /// Relaciona a orientação que cada tubo recebeu e define o será o ponto de inserção da tag
        /// </summary>
        /// <param name="tubosNaVistaFiltrados"></param> Tubos filtados da vista ativa
        /// <param name="orientacoes"></param> Orientação definida na função VerificaOrientacaoEmPlanta
        /// <returns> Lista XYZ </returns>

        public IList<XYZ> PontoInsercaoTag(IList<Element> tubosNaVistaFiltrados, IList<string> orientacoes, string detalheEscolhido)
        {

            IList<XYZ> pontoMedio = new List<XYZ>();        // Instancia uma lista vazia de XYZ que receberá o ponto médio de cada tubo
            IList<XYZ> pontoDeInsercaoTag = new List<XYZ>(); // Instancia uma lista vazia de XYZ que será o retorno, com todos os pontos de insereção de tag para cada tubo
            IList<double> deslocamentos = new List<double>(); // Instancia uma lista vazia de XYZ que receberá o deslocamento de cada tag
                                                              // essa lista será somada ao ponto médio para gerar o ponto de inserção

            foreach (Element tubo in tubosNaVistaFiltrados) // Percorre cada tubo
            {

                Location tuboLocation = tubo.Location; // Obtem a localização no modelo de cada tubo 

                // a partir da localização dos tubos Atribui uma variavel da classe LocationCurve para uso dos metodos apropriados
                LocationCurve localCurvaTubo = (LocationCurve)tuboLocation;
                Curve curvaTubo = localCurvaTubo.Curve; // Atribui uma variavel da classe Curve a partir de LocationCurve

                XYZ pontoMedioCurva = curvaTubo.Evaluate(0.5, true); //Encontra o ponto médio da curva do tubo que está sendo iterado no momento
                pontoMedio.Add(pontoMedioCurva); // Adiciona o valor XYZ do ponto médio a lista do ponto médio

                Parameter parametroDiametroExterno = tubo.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER); // Atribui a uma variavel Parameter o parametro incorporado do Revit Diametro Externo
                double valorDiametroExternoPes = parametroDiametroExterno.AsDouble(); // Atribui a uma variavel double o valor do parametro diametro externo
                deslocamentos.Add(valorDiametroExternoPes / 2); // Adiciona a lista de deslocamentos o valor do diametro externo divido por 2
            }

            for (int i = 0; i < tubosNaVistaFiltrados.Count; i++)
            {
                if (orientacoes[i] == "Horizontal") // Condiciona para apenas tubos Horizontais
                {
                    if (detalheEscolhido == "Diametro")
                    {
                        // Ponto de inserção da tag recebe o valor de deslocamento em Y e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(0, deslocamentos[i], 0) + pontoMedio[i]);
                    }
                    else if (detalheEscolhido == "Inclinacao")
                    {
                        // Ponto de inserção da tag recebe o valor de deslocamento em Y e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(0, -deslocamentos[i], 0) + pontoMedio[i]);
                    }

                }

                else if (orientacoes[i] == "Vertical") // Condiciona para apenas tubos Verticais
                {
                    if (detalheEscolhido == "Diametro")
                    {
                        // Ponto de inserção da tag recebe o valor negativo do deslocamento em X, e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(-deslocamentos[i], 0, 0) + pontoMedio[i]);
                    }
                    else if (detalheEscolhido == "Inclinacao")
                    {
                        // Ponto de inserção da tag recebe o valor negativo do deslocamento em X, e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(deslocamentos[i], 0, 0) + pontoMedio[i]);
                    }
                }

                else if (orientacoes[i] == "Diagonal Positiva") // Condiciona para apenas tubos Diagonal Positiva
                {
                    if (detalheEscolhido == "Diametro")
                    {
                        // Ponto de inserção da tag recebe o valor negativo do deslocamento em X, Positivo em Y, e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(-deslocamentos[i], deslocamentos[i], 0) + pontoMedio[i]);
                    }
                    else if (detalheEscolhido == "Inclinacao")
                    {
                        // Ponto de inserção da tag recebe o valor negativo do deslocamento em X, Positivo em Y, e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(deslocamentos[i], -deslocamentos[i], 0) + pontoMedio[i]);
                    }

                }

                else if (orientacoes[i] == "Diagonal Negativa")
                {
                    if (detalheEscolhido == "Diametro")
                    {
                        // Ponto de inserção da tag recebe o valor do deslocamento em X e Y, e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(deslocamentos[i], deslocamentos[i], 0) + pontoMedio[i]);
                    }
                    else if (detalheEscolhido == "Inclinacao")
                    {
                        // Ponto de inserção da tag recebe o valor do deslocamento em X e Y, e soma tudo ao ponto médio
                        pontoDeInsercaoTag.Add(new XYZ(-deslocamentos[i], -deslocamentos[i], 0) + pontoMedio[i]);
                    }

                }

                else
                {
                    pontoDeInsercaoTag = pontoMedio; // Caso nao atenda a nenhuma condição o ponto de inserção da tag recebe o ponto 
                }
            }

            return pontoDeInsercaoTag; // Retorna a lista de pontos XYZ
        }
        #endregion


        #region RemoverTagsExistentes

        /// <summary>
        /// Remove tags existentes que são do tipo escolhido e estão sobre tubos que vão ser inseridas novas tags
        /// </summary>
        /// <param name="doc"></param> // Documento do projeto atual
        /// <param name="vistaAtiva"></param> // Vista Ativa
        /// <param name="tubosNaVistaFiltrados"></param> // Tubos que receberão novas tags
        /// <param name="IdTipoTag"></param> // ID do tipo de tag escolhido

        public static void RemoverTagsExistentes(Document doc, View vistaAtiva, IList<Element> tubosNaVistaFiltrados, ElementId IdTipoTag)
        {
            // Lista de Element que recebe todas as Tags na vista ativa
            IList<Element> tagsTotalVista = new FilteredElementCollector(doc, vistaAtiva.Id).OfClass(typeof(IndependentTag)).WhereElementIsNotElementType().ToElements();

            //Lista de ElementId vazia que receberá o ID de todos os tubos recebidos por parametro
            IList<ElementId> idTubos = new List<ElementId>();

            // Declaração da variavel do tipo IndependentTag
            IndependentTag tagIndependente;

            foreach (Element tubo in tubosNaVistaFiltrados) // Percorre todos os tubos recebidos
            {
                idTubos.Add(tubo.Id); // Adiciona o id de cada tubo a lista idTubos
            }

            foreach (Element tag in tagsTotalVista) // Percorre cada tag da vista ativa
            {
                if (tag.GetTypeId() == IdTipoTag) // Verifica se a tag iterada corresponde as tags que serão colocadas
                {

                    tagIndependente = (IndependentTag)tag; // Converte a Tag iterada no momento de Element para IndependentTag, para que sejam usados metodos apropriados

                    //Declara uma variavel de uma Lista ISet de ElementId que receberá o ElementId de cada tubo hospedeiro da Tag Iterada no momento
                    ISet<ElementId> idTubosComTag = tagIndependente.GetTaggedLocalElementIds();

                    if (idTubosComTag.Any(idTubos.Contains)) // Verifica se o Id de tubos Tageados está na lista dos tubos que receberão novas Tags
                    {
                        doc.Delete(tag.Id); // Deleta a Tag
                    }

                }


            }

        }
        #endregion


        #region SetarValorAoParametroInclinacao

        public void SetarValorAoParametroInclinacao(IList<Element> tubosNaVista)
        {


            foreach (Element tubo in tubosNaVista)
            {
                Parameter parametroInclinacao = tubo.LookupParameter("PRJ HDR: Inclinacao Tag");
                Parameter parametroAbreviaturaSistema = tubo.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM);
                Parameter parametroClassificacaoSistema = tubo.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                Parameter parametroDiametro = tubo.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

                string abreviaturaSistema = parametroAbreviaturaSistema.AsString();
                string classificacaoSistema = parametroClassificacaoSistema.AsString();
                double diametroEmMilimetros = parametroDiametro.AsDouble() * 304.8;

                if (abreviaturaSistema == "ESG")
                {
                    if (classificacaoSistema == "Sanitário")
                    {

                        if (diametroEmMilimetros <= 75)
                        {
                            parametroInclinacao.Set("2%");

                        }
                        else
                        {
                            parametroInclinacao.Set("1%");
                        }
                    }

                    else
                    {
                        parametroInclinacao.Set("1%");
                    }
                }
            }

        }
        #endregion


        #region CriarTags

        /// <summary>
        /// Cria tags do tipo IndependentTag nos tubos da vista ativa e locais predefinidos
        /// </summary>
        /// <param name="tubosNaVistaFiltrados"></param> // Tubos da vista ativa filtrados por comprimento e colunas removidas
        /// <param name="IdTipoTag"></param> // id do tipo da tag que será inserida
        /// <param name="IdVistaAtiva"></param> // Id da vista ativa
        /// <param name="linhaDeChamada"></param> // Opção de inserir linha de chamda na tag
        /// <param name="Orientacao"></param> // Orientação da tag que o revit utiliza, vertical ou horizontal
        /// <param name="pontosInsercao"></param> // pontos de inserição XYZ que as tags serao colocadas

        public void CriarTags(
            IList<Element> tubosNaVistaFiltrados,
            ElementId IdTipoTag,
            ElementId IdVistaAtiva,
            bool linhaDeChamada,
            TagOrientation Orientacao,
            IList<XYZ> pontosInsercao
            )
        {

            _vistaAtiva = _doc.ActiveView;

            // Instancia uma lista vazia de reference pois o método de criar tags utilizada a class Reference para os elementos a serem tageados
            IList<Reference> tubos = new List<Reference>();

            foreach (Element tubo in tubosNaVistaFiltrados) // Percorre cada Element (tubo)
            {
                Reference referenciaTubo = new Reference(tubo); // Converte a class Element de cada tubo para Reference
                tubos.Add(referenciaTubo); // Adiciona os tubos convertidos a lista de Reference que será utilziada no método de criar as tags

            }

            for (int i = 0; i < tubosNaVistaFiltrados.Count; i++) // Itera sob o tamanho da lista de tubos para que cada tubo receba uma tag
            {
                IndependentTag.Create( // Método que cria as tags
                    _doc,
                    IdTipoTag,
                    IdVistaAtiva,
                    tubos[i],
                    linhaDeChamada,
                    Orientacao,
                    pontosInsercao[i]);
            }

        }
        #endregion
    }
}

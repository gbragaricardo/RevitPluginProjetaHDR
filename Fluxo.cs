using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;


namespace ProjetaHDR
{
    [Transaction(TransactionMode.Manual)]
    internal class Fluxo : IExternalCommand
    {
        //create global variables
        UIApplication _uiapp;
        Autodesk.Revit.ApplicationServices.Application _app;
        UIDocument _uidoc;
        Document _doc;
        View _vistaAtiva;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Creating App and doc objects.
            _uiapp = commandData.Application;
            _app = _uiapp.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;
            _vistaAtiva = _doc.ActiveView;

           


                bool linhaDeChamada = false;
                TagOrientation orientacaoTag = TagOrientation.Horizontal;

                IList<Element> tubosNaVista = SelecionarTubosNaVista();

                IList<Element> tubosNaVistaFiltradosPorComprimento = FiltrarPorComprimento(tubosNaVista);

                IList<Element> FiltradorPorComprimentoEVerticais = RemoverVerticais(tubosNaVistaFiltradosPorComprimento);

                IList<string> orientacoes = VerificarOrientacaoEmPlanta(FiltradorPorComprimentoEVerticais);

                IList<XYZ> pontosDeInsercao = PontoInsercaoTag(FiltradorPorComprimentoEVerticais);

                IList<string> direcoes = VerificarDirecaoDoFluxo(FiltradorPorComprimentoEVerticais);

                IList<ElementId> idsTags = ObterIdTags(direcoes);


            using (Transaction transacao = new Transaction(_doc, "Fluxo"))
            {
                transacao.Start();

                RemoverTagsExistentes(_doc, _vistaAtiva, FiltradorPorComprimentoEVerticais, idsTags);

                CriarTags(FiltradorPorComprimentoEVerticais, idsTags, _vistaAtiva.Id, linhaDeChamada, orientacaoTag, pontosDeInsercao);

                transacao.Commit();
            }

            return Result.Succeeded;
        }


        #region VerificarDirecaoDoFluxo

        public IList<string> VerificarDirecaoDoFluxo(IList<Element> tubosNaVista)
        {

            IList<string> direcaoSeta = new List<string>();
            string direcaoTemp = "Direita";

            Selection selecao = _uidoc.Selection;

            foreach (Element tuboAsElement in tubosNaVista)
            {

                selecao.SetElementIds(new List<ElementId>  { tuboAsElement.Id });

                Pipe tuboAsPipe = tuboAsElement as Pipe;

                if (tuboAsPipe != null)
                {
                    ConnectorManager gerenciadorDeConectores = tuboAsPipe.ConnectorManager;

                    if (gerenciadorDeConectores != null)
                    {

                        // Inicializa variáveis para armazenar o conector de entrada e de saída
                        Connector entradaTubo = null;
                        Connector saidaTubo = null;

                        foreach (Connector conector in gerenciadorDeConectores.Connectors)
                        {

                            FlowDirectionType parametroTipoDeFluxo = conector.Direction;
                            string tipoDeFluxo = parametroTipoDeFluxo.ToString();

                            switch (tipoDeFluxo)
                            {
                                case "In":
                                    entradaTubo = conector;
                                    break;

                                case "Out":
                                    saidaTubo = conector;
                                    break;

                                case "Bidirectional":
                                    break;


                            }
                        }

                        if (entradaTubo != null && saidaTubo != null)
                        {
                            XYZ direcaoFluxo = saidaTubo.Origin - entradaTubo.Origin;
                            if (Math.Round(direcaoFluxo.X, 3) != 0)
                            {
                                if (Math.Round(direcaoFluxo.X, 3) > 0)
                                {
                                    direcaoTemp = "Direita";
                                    direcaoSeta.Add(direcaoTemp);
                                }
                                else
                                {
                                    direcaoTemp = "Esquerda";
                                    direcaoSeta.Add(direcaoTemp);
                                }
                            }

                            else if (Math.Round(direcaoFluxo.Y, 3) > 0)
                            {
                                direcaoTemp = "Direita";
                                direcaoSeta.Add(direcaoTemp);
                            }

                            else
                            {
                                direcaoTemp = "Esquerda";
                                direcaoSeta.Add(direcaoTemp);
                            }
                        }

                        else
                        {
                            direcaoSeta.Add(direcaoTemp);
                        }


                    }

                }

                selecao.SetElementIds(new List<ElementId>());
            }

            return direcaoSeta;
        }

        #endregion


        #region ObterIdTags

        public IList<ElementId> ObterIdTags(IList<string> direcaoDoFluxo)
        {

            IList<ElementId> idTag = new List<ElementId>();
            ElementId setaEsquerda = null;
            ElementId setaDireita = null;


            IList<Element> tiposDeTag = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_PipeTags).WhereElementIsElementType().ToElements();

            foreach (Element tipoDeTag in tiposDeTag)
            {
                if (tipoDeTag.Name == "Esquerda")
                {
                    setaEsquerda = tipoDeTag.Id;
                }
                else if( tipoDeTag.Name == "Direita")
                {
                    setaDireita = tipoDeTag.Id;
                }

            }

            if (setaEsquerda != null && setaDireita != null)
            {
                foreach (string direcao in direcaoDoFluxo)
                {
                    if (direcao == "Esquerda")
                    {
                        idTag.Add(setaEsquerda);
                    }
                    else
                    {
                        idTag.Add(setaDireita);
                    }
                }
            }

            return idTag;
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

        public IList<XYZ> PontoInsercaoTag(IList<Element> tubosNaVistaFiltrados)
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

           
            return pontoMedio; // Retorna a lista de pontos XYZ
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

        public static void RemoverTagsExistentes(Document doc, View vistaAtiva, IList<Element> tubosNaVistaFiltrados, IList<ElementId> IdTipoTag)
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
            
            // Percorre cada tag da vista ativa
            for (int i = 0; i < tagsTotalVista.Count; i++)
            {
                if (IdTipoTag.Contains(tagsTotalVista[i].GetTypeId())) // Verifica se a tag iterada corresponde as tags que serão colocadas
                {

                    tagIndependente = (IndependentTag)tagsTotalVista[i]; // Converte a Tag iterada no momento de Element para IndependentTag, para que sejam usados metodos apropriados

                    //Declara uma variavel de uma Lista ISet de ElementId que receberá o ElementId de cada tubo hospedeiro da Tag Iterada no momento
                    ISet<ElementId> idTubosComTag = tagIndependente.GetTaggedLocalElementIds();

                    if (idTubosComTag.Any(idTubos.Contains)) // Verifica se o Id de tubos Tageados está na lista dos tubos que receberão novas Tags
                    {
                        doc.Delete(tagsTotalVista[i].Id); // Deleta a Tag
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
        IList<ElementId> IdTipoTag,
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
                    IdTipoTag[i],
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

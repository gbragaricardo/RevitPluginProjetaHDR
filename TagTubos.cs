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
        ElementId idTipoTag = new ElementId(12292844);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Creating App and doc objects.
            _uiapp = commandData.Application;
            _app = _uiapp.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;

            _vistaAtiva = _doc.ActiveView;

            using (Transaction transacao = new Transaction(_doc, "Nome da Transação"))
            {

                ElementId IdTipoTag = new ElementId(12292844);
                bool linhaDeChamada = false;
                TagOrientation tagOrientacao = TagOrientation.Horizontal;

                IList<Element> tubosNaVista = new List<Element>(SelecionarTubosNaVista());

                IList<Element> filtradosPorComprimento = FiltrarPorComprimento(tubosNaVista);

                IList<Element> filtradorPorComprimentoEVerticaisRemovidos = RemoverVerticais(filtradosPorComprimento);

                IList<string> orientacoesEmPlanta = VerificarOrientacaoEmPlanta(filtradorPorComprimentoEVerticaisRemovidos);

                IList<XYZ> pontosDeInsercaoTags = PontoInsercaoTag(filtradorPorComprimentoEVerticaisRemovidos, orientacoesEmPlanta);


                transacao.Start();

                RemoverTagsExistentes(_doc, _vistaAtiva, filtradorPorComprimentoEVerticaisRemovidos, idTipoTag);
                CriarTags(filtradorPorComprimentoEVerticaisRemovidos, IdTipoTag, _vistaAtiva.Id, linhaDeChamada, tagOrientacao, pontosDeInsercaoTags);

                transacao.Commit();
            }

            return Result.Succeeded;
        }

        #region SelecionarTubosNaVista
        public IList<Element> SelecionarTubosNaVista()
        {
            _vistaAtiva = _doc.ActiveView;

            IList<Element> tubosNaVista = new FilteredElementCollector(_doc, _vistaAtiva.Id).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements();


            return tubosNaVista;
        }
        #endregion


        #region FiltrarPorComprimento
        public IList<Element> FiltrarPorComprimento(IList<Element> tubosNaVista)
        {
            IList<Element> filtrados = new List<Element>();

            foreach (Element tubo in tubosNaVista)
            {
                Parameter parametroComprimento = tubo.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                double valorComprimentoMetros = parametroComprimento.AsDouble() * 0.3048;

                if (valorComprimentoMetros > .2)
                {
                    filtrados.Add(tubo);
                }
            }

            return filtrados;
        }
        #endregion


        #region RemoverVerticais
        public IList<Element> RemoverVerticais(IList<Element> tubosNaVistaFiltrados)
        {
            IList<Element> filtrados = new List<Element>();

            foreach (Element tubo in tubosNaVistaFiltrados)
            {
                Location tuboLocation = tubo.Location;
                LocationCurve localCurvaTubo = (LocationCurve)tuboLocation;
                Curve curvaTubo = localCurvaTubo.Curve;
             
                XYZ pontoInicial = curvaTubo.GetEndPoint(0);
                XYZ pontoFinal = curvaTubo.GetEndPoint(1);

                if (Math.Round(pontoInicial.X) != Math.Round(pontoFinal.X) || Math.Round(pontoInicial.Y) != Math.Round(pontoFinal.Y))
                {
                    filtrados.Add(tubo);
                }
            }

            return filtrados ;
        }
        #endregion


        #region VerificarOrientacaoEmPlanta
        public IList<string> VerificarOrientacaoEmPlanta(IList<Element> tubosNaVistaFiltrados)
        {
            
            IList<string> orientacaoTubos = new List<string>();

            double margemDePrecisao = 0.01;

            foreach (Element tubo in tubosNaVistaFiltrados)
            {

                Location tuboLocation = tubo.Location;
                LocationCurve localCurvaTubo = (LocationCurve)tuboLocation;
                Curve curvaTubo = localCurvaTubo.Curve;

                XYZ pontoInicial = curvaTubo.GetEndPoint(0);
                XYZ pontoFinal = curvaTubo.GetEndPoint(1);

                double deltaX = Math.Abs(pontoFinal.X - pontoInicial.X);
                double deltaY = Math.Abs(pontoFinal.Y - pontoInicial.Y);

                if (pontoInicial != null && pontoFinal != null)
                {

                    if (Math.Abs((deltaX - deltaY)) < margemDePrecisao)
                    {

                        if ((pontoFinal.X - pontoInicial.X) * (pontoFinal.Y - pontoInicial.Y) > 0)
                        {
                            orientacaoTubos.Add("Diagonal Positiva");
                        }
                        else
                        {
                            orientacaoTubos.Add("Diagonal Negativa");
                        }

                    }

                    else if (deltaX > deltaY)
                    {
                        orientacaoTubos.Add("Horizontal");
                    }

                    else
                    {
                        orientacaoTubos.Add("Vertical");
                    }

                }
                else
                {
                    orientacaoTubos.Add("N/A");
                }

            }

            return orientacaoTubos;
        }
        #endregion


        #region PontoInsercaoTag
        public IList<XYZ> PontoInsercaoTag( IList<Element> tubosNaVistaFiltrados, IList<string> orientacoes)
        {

            IList<XYZ> pontoMedio = new List<XYZ>();
            IList<XYZ> pontoDeInsercaoTag = new List<XYZ>();
            IList<double> deslocamentos = new List<double>();

            foreach (Element tubo in tubosNaVistaFiltrados)
            {

                Location tuboLocation = tubo.Location;
                LocationCurve localCurvaTubo = (LocationCurve)tuboLocation;
                Curve curvaTubo = localCurvaTubo.Curve;

                XYZ pontoMedioCurva = curvaTubo.Evaluate(0.5, true);
                pontoMedio.Add(pontoMedioCurva);

                Parameter parametroDiametroExterno = tubo.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                double valorDiametroExternoPes = parametroDiametroExterno.AsDouble();
                deslocamentos.Add(valorDiametroExternoPes / 2);
            }

            for (int i = 0; i < tubosNaVistaFiltrados.Count; i++)
            {
                if (orientacoes[i] == "Horizontal")
                {
                    pontoDeInsercaoTag.Add(new XYZ(0, deslocamentos[i], 0) + pontoMedio[i]);
                }

                else if (orientacoes[i] == "Vertical")
                {
                    pontoDeInsercaoTag.Add(new XYZ(-deslocamentos[i], 0, 0) + pontoMedio[i]);
                }

                else if (orientacoes[i] == "Diagonal Positiva")
                {
                    pontoDeInsercaoTag.Add(new XYZ(-deslocamentos[i], deslocamentos[i], 0) + pontoMedio[i]);
                }

                else if (orientacoes[i] == "Diagonal Negativa")
                {
                    pontoDeInsercaoTag.Add(new XYZ(deslocamentos[i], deslocamentos[i], 0) + pontoMedio[i]);
                }

                else
                {
                    pontoDeInsercaoTag = pontoMedio;
                }
            }

            return pontoDeInsercaoTag;
        }
        #endregion


        #region RemoverTagsExistentes
        public static void RemoverTagsExistentes(Document doc, View vistaAtiva, IList<Element> tubosNaVistaFiltrados, ElementId IdTipoTag)
        {

            IList<Element> tagsTotalVista = new FilteredElementCollector(doc, vistaAtiva.Id).OfClass(typeof(IndependentTag)).WhereElementIsNotElementType().ToElements();
            IList<ElementId> idTubos = new List<ElementId>();
            IndependentTag tagIndependente;

            foreach (Element tubo in tubosNaVistaFiltrados)
            {
                idTubos.Add(tubo.Id);
            }
            
            foreach(Element tag in tagsTotalVista)
            {
                tagIndependente = (IndependentTag)tag;

                ISet<ElementId> idTubosComTag = tagIndependente.GetTaggedLocalElementIds();

                if (idTubosComTag.Any(idTubos.Contains))
                {
                    doc.Delete(tag.Id);
                }

            }
            
        }
        #endregion


            #region CriarTags
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

            IList<Reference> tubos = new List<Reference>();

            foreach (Element tubo in tubosNaVistaFiltrados)
            {
                Reference referenciaTubo = new Reference(tubo);
                tubos .Add(referenciaTubo);
                
            }

            for (int i = 0; i < tubosNaVistaFiltrados.Count; i++)
            {
                IndependentTag.Create(
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

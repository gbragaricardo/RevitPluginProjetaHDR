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
    internal class BotaoDebug : IExternalCommand
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

            using (Transaction transacao = new Transaction(_doc, "Nome da Transação"))
            {
                transacao.Start();


                VerificarDirecaoDoFluxo(_doc, SelecionarTubosNaVista());



                transacao.Commit();
            }
                 

           

            return Result.Succeeded;
        }



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



        #region VerificarDirecaoDoFluxo

        public IList<string> VerificarDirecaoDoFluxo(Document doc, IList<Element> tubosNaVista)
        {

            IList<string> direcaoSeta = new List<string>();
            string direcaoTemp = "Direita";

            Selection selecao = _uidoc.Selection;

            foreach (Element tuboAsElement in tubosNaVista)
            {

                //selecao.SetElementIds(new List<ElementId>  { tuboAsElement.Id });

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

                //selecao.SetElementIds(new List<ElementId>());
            }

            return direcaoSeta;
        }
    }
    #endregion


}

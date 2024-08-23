using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
                Pipe tubo = _doc.GetElement(new ElementId(20816845)) as Pipe;
                ObterFluxoDaTubulacao(SelecionarTubosNaVista());
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



        #region GetFlowDirectionFromPipe

        /// <summary>
        /// Seleciona todos os tubos da vista ativa
        /// </summary>
        /// <returns>Lista de Element</returns>

        public void ObterFluxoDaTubulacao(IList<Element> tubosNaVista)
        {



            foreach (Element tuboAsElement in tubosNaVista)
            {
                Pipe tuboAsPipe = tuboAsElement as Pipe;



                // Obtenha os conectores do tubo
                ConnectorManager gerenciadorDeConector = tuboAsPipe.ConnectorManager;

                if (gerenciadorDeConector != null)
                {
                    // Supondo que o tubo tem dois conectores, um de entrada e um de saída
                    Connector entradaTubo = null;
                    Connector saidaTubo = null;

                    foreach (Connector conector in gerenciadorDeConector.Connectors)
                    {
                        if (conector.IsConnected)
                        {
                            // Verifique se o conector está conectado a outro elemento
                            foreach (Connector refConector in conector.AllRefs)
                            {
                                if (refConector.ConnectorType == ConnectorType.End)
                                {
                                    if (entradaTubo == null)
                                    {
                                        entradaTubo = conector;
                                    }
                                    else
                                    {
                                        saidaTubo = conector;
                                    }
                                }
                            }
                        }
                    }

                    if (entradaTubo != null && saidaTubo != null)
                    {
                        // Direção do fluxo é do conectorIn para o conectorOut
                        XYZ direcaoFluxo = saidaTubo.Origin - entradaTubo.Origin;

                        // Exiba a direção do fluxo
                        TaskDialog.Show("Flow Direction", $"Pipe ID: {tuboAsPipe.Id}\nFlow Direction: {direcaoFluxo}");
                    }
                    else
                    {
                        TaskDialog.Show("Flow Direction", "Não foi possível determinar a direção do fluxo.");
                    }
                }
                else
                {
                    TaskDialog.Show("Error", "Não foi possível acessar os conectores do tubo.");
                }
            }
        }
        #endregion
    }
}

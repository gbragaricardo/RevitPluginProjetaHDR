using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;


namespace ProjetaHDR
{
    [Transaction(TransactionMode.Manual)]
    internal class LuvasEsgPluv : IExternalCommand
    {
        //create global variables
        UIApplication _uiapp;
        Autodesk.Revit.ApplicationServices.Application _app;
        UIDocument _uidoc;
        Document _doc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Creating App and doc objects.
            _uiapp = commandData.Application;
            _app = _uiapp.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;

            int instanciasAlteradas;

            using (Transaction transacao = new Transaction(_doc, "Nome da Transação"))
            {

                transacao.Start();

                instanciasAlteradas = InserirSistemaFamiliaAninhada(SelecionarConexoes(), "Abreviatura do sistema", "PRJ HDR: Sistema");

                transacao.Commit();
            }
                

            TaskDialog.Show("Retorno", $"{instanciasAlteradas} Luvas Modificadas");

            return Result.Succeeded;
        }

        #region SelecionarConexoes
        public IList<Element> SelecionarConexoes()
        {
            IList<Element> conexoesGeral = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements();
            
            return conexoesGeral;
        }
        #endregion



        #region InserirSistemaFamiliaAninhada
        public int InserirSistemaFamiliaAninhada(IList<Element> instancias, string nomeParametroHospedeiro, string nomeParametroAninhado)
        {
            FamilyInstance familiaHospedeira;

            int contador = 0;

            foreach (Element element in instancias)
            {
                FamilyInstance conexaoInstancia = element as FamilyInstance;
        
                if (conexaoInstancia != null)
                {
                    familiaHospedeira = conexaoInstancia.SuperComponent as FamilyInstance;

                    if (familiaHospedeira != null)
                    {
                        Parameter parametroHospedeiro = familiaHospedeira.LookupParameter(nomeParametroHospedeiro);

                        if (parametroHospedeiro != null)
                        {
                            string valorParametroHospedeiro = parametroHospedeiro.AsString();

                            if (valorParametroHospedeiro != "")
                            {
                                Parameter parametroAninhado = conexaoInstancia.LookupParameter(nomeParametroAninhado);

                                if (parametroAninhado != null)
                                {
                                    parametroAninhado.Set(valorParametroHospedeiro);
                                    contador++;
                                }
                            }                                
                        }
                        
                    }
                }
            }

            return contador;
        }
        #endregion
    }
}

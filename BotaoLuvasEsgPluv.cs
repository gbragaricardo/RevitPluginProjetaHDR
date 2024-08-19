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

        #region MAIN
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Creating App and doc objects.
            _uiapp = commandData.Application;
            _app = _uiapp.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;
            int instanciasAlteradas;

            using (Transaction transacao = new Transaction(_doc, "Luvas ESG / PLUV"))
            {

                transacao.Start(); //Inicia Mudanças no arquivo

                //Executa Funções para selecionar conexoes aninhadas e atribuir o valor do parametro hospedeiro à PJR HDR: Sistema

                instanciasAlteradas = InserirSistemaFamiliaAninhada(SelecionarConexoes(), "Abreviatura do sistema", "PRJ HDR: Sistema");

                transacao.Commit(); //Finaliza Mudanças no arquivo
            }
                

            TaskDialog.Show("Retorno", $"{instanciasAlteradas} Luvas Modificadas"); //Retorna a quantidade de instancias que receberam valor no parametro

            return Result.Succeeded; //Retorno padrao do metodo execute
        }
        #endregion


        #region SelecionarConexoes

        /// <summary>
        /// Seleciona todas as conexoes de tubo no projeto
        /// </summary>
        /// <returns>Lista de Elements</returns>

        public IList<Element> SelecionarConexoes()
        {
            IList<Element> conexoesGeral = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements();
            
            return conexoesGeral;
        }
        #endregion



        #region InserirSistemaFamiliaAninhada

        /// <summary>
        /// Executa Funções para selecionar conexoes aninhadas e atribuir o valor do parametro hospedeiro ao parametro aninhado
        /// </summary>
        /// <param name="instancias"></param> --- Lista de Elements conexoes de tubo 
        /// <param name="nomeParametroHospedeiro"></param> --- Parametro inserido em familia hospedeira em que será buscado o seu valor
        /// <param name="nomeParametroAninhado"></param> --- Parametro inserido em familia aninhada em que receberá o valor
        /// <returns>Quantidade de instancias modificadas</returns>

        public int InserirSistemaFamiliaAninhada(IList<Element> instancias, string nomeParametroHospedeiro, string nomeParametroAninhado)
        {
            FamilyInstance familiaHospedeira; // Variavel declarada como family instance para que seja utilizados metodos como SuperComponent

            int contador = 0; // Contador de instancias modificadas

            foreach (Element element in instancias) // Percorre cada conexao de tubo do projeto
            {
                FamilyInstance conexaoInstancia = element as FamilyInstance; // Converte Element para FamilyInstance
        
                if (conexaoInstancia != null) // Garante que a conversão funcionou
                {
                    familiaHospedeira = conexaoInstancia.SuperComponent as FamilyInstance; // Obtém a familia hospedeira de cada conexao da lista 

                    if (familiaHospedeira != null) // Garante que há uma familia hospedeira
                    {
                        Parameter parametroHospedeiro = familiaHospedeira.LookupParameter(nomeParametroHospedeiro); // Atribui em "parametroHospedeiro" parametro referencia na familia hospedeira

                        if (parametroHospedeiro != null) // Garante que há o parametro na familia hospedeira
                        {
                            string valorParametroHospedeiro = parametroHospedeiro.AsString(); // Recebe o valor do parametro atribuido

                            if (valorParametroHospedeiro != "") // Garante que há algum valor para que apenas seja contabilizado mudanças onde há valor nos parametros
                            {
                                Parameter parametroAninhado = conexaoInstancia.LookupParameter(nomeParametroAninhado); // Atribui em "parametroAninhado" parametro destino na familia alvo 

                                if (parametroAninhado != null) // Garante que há o parametro aninhado
                                {
                                    parametroAninhado.Set(valorParametroHospedeiro); // Copia o valor do parametro hospedeiro para o parametro na familia aninhada
                                    contador++; // Contabiliza a mudança no parametro aninhado
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

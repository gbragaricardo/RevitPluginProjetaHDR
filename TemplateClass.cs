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
    internal class Template : IExternalCommand
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

            using (Transaction transacao = new Transaction(_doc, "Nome da Transação"))
            {

                transacao.Start();

                

                transacao.Commit();
            }
                

            TaskDialog.Show("Title", "Content");

            return Result.Succeeded;
        }

        #region funcaopadrao
        public void funcaopadrao()
        {
            
        }
        #endregion



        
    }
}

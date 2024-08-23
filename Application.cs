using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

namespace ProjetaHDR
{
    /// <summary>
    /// Implements the Revit add-in interface IExternalApplication
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Application : IExternalApplication
    {
        /// <summary>
        /// Implements the on Shutdown event
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Implements the OnStartup event
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        public Result OnStartup(UIControlledApplication application)
        {
            RibbonPanel panelTabelas = CreateRibbonPanelTabelas(application);
            RibbonPanel panelDetalhamento = CreateRibbonPanelDetalhamento(application);
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            #region BotaoLuvasEsgPluv
            // Cria os dados do botão
            PushButtonData luvasEsgPluvData = new PushButtonData(
                "Luvas ESG / PLUV",            // Nome interno do botão
                "Luvas\nESG / PLUV",            // Texto exibido no botão
                thisAssemblyPath,          // Caminho do assembly onde o comando está localizado
                "ProjetaHDR.LuvasEsgPluv"    // Nome completo da classe de comando
            );

            // Adiciona o botão ao painel e verifica se a adição foi bem-sucedida
            PushButton luvasEsgPluvButton = panelTabelas.AddItem(luvasEsgPluvData) as PushButton;

            // Se o botão foi criado com sucesso
            if (luvasEsgPluvButton != null)
            {
                // Define uma dica (tooltip) que aparecerá quando o usuário passar o mouse sobre o botão
                luvasEsgPluvButton.ToolTip = "Adiciona o valor do parametro Abreviatura do sistema em familias aninhadas com o valor da familia hospedeira";

                // Define o caminho para o ícone do botão
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", "luvas.ico");

                // Cria a imagem do ícone
                Uri uri = new Uri(iconPath);
                BitmapImage bitmap = new BitmapImage(uri);

                // Define a imagem como o ícone do botão
                luvasEsgPluvButton.LargeImage = bitmap;
            }
            #endregion


            #region BotaoDebug
            // Cria os dados do botão
            PushButtonData debugData = new PushButtonData(
                "BotaoDebug",            // Nome interno do botão
                "BotaoDebug",            // Texto exibido no botão
                thisAssemblyPath,          // Caminho do assembly onde o comando está localizado
                "ProjetaHDR.BotaoDebug"    // Nome completo da classe de comando
            );

            // Adiciona o botão ao painel e verifica se a adição foi bem-sucedida
            PushButton debugButton = panelTabelas.AddItem(debugData) as PushButton;

            // Se o botão foi criado com sucesso
            if (debugButton != null)
            {
                // Define uma dica (tooltip) que aparecerá quando o usuário passar o mouse sobre o botão
                debugButton.ToolTip = "Auxilio a testar metodos";

                // Define o caminho para o ícone do botão
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", "debug.ico");

                // Cria a imagem do ícone
                Uri uri = new Uri(iconPath);
                BitmapImage bitmap = new BitmapImage(uri);

                // Define a imagem como o ícone do botão
                debugButton.LargeImage = bitmap;
            }
            #endregion


            #region BotaoTagTubos
            // Cria os dados do botão
            PushButtonData tagTubosData = new PushButtonData(
                "Detalhar Tubos",            // Nome interno do botão
                "Detalhar\nTubos",            // Texto exibido no botão
                thisAssemblyPath,          // Caminho do assembly onde o comando está localizado
                "ProjetaHDR.TagTubos"    // Nome completo da classe de comando
            );

            // Adiciona o botão ao painel e verifica se a adição foi bem-sucedida
            PushButton tagTubosButton = panelDetalhamento.AddItem(tagTubosData) as PushButton;

            // Se o botão foi criado com sucesso
            if (tagTubosButton != null)
            {
                // Define uma dica (tooltip) que aparecerá quando o usuário passar o mouse sobre o botão
                tagTubosButton.ToolTip = "Adiciona Tag aos tubos em planta baixa";

                // Define o caminho para o ícone do botão
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", "tagstubos.ico");

                // Cria a imagem do ícone
                Uri uri = new Uri(iconPath);
                BitmapImage bitmap = new BitmapImage(uri);

                // Define a imagem como o ícone do botão
                tagTubosButton.LargeImage = bitmap;
            }
            #endregion


            return Result.Succeeded;
        }

        /// <summary>
        /// Função que cria o RibbonPanel
        /// </summary>
        /// <param name="aplicativo"></param>
        /// <returns></returns>
        public RibbonPanel CreateRibbonPanelTabelas(UIControlledApplication aplicativo)
        {
            string tab = "Projeta HDR";
            RibbonPanel ribbonPanel = null;

            try
            {
                aplicativo.CreateRibbonTab(tab);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            try
            {
                ribbonPanel = aplicativo.CreateRibbonPanel(tab, "Tabelas");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return ribbonPanel;
        }

        public RibbonPanel CreateRibbonPanelDetalhamento(UIControlledApplication aplicativo)
        {
            string tab = "Projeta HDR";
            RibbonPanel ribbonPanel = null;

            try
            {
                aplicativo.CreateRibbonTab(tab);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            try
            {
                ribbonPanel = aplicativo.CreateRibbonPanel(tab, "Detalhamento");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return ribbonPanel;
        }
    }
}

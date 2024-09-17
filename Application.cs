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
    
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Application : IExternalApplication
    {
        
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Cria a aba "Projeta HDR" uma vez
            string tab = "Projeta HDR";
            try
            {
                application.CreateRibbonTab(tab);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Aba '{tab}' já existe ou houve um erro ao tentar criá-la: {ex.Message}");
            }

            // Cria os painéis
            RibbonPanel panelTabelas = CriarPainel(application, tab, "Tabelas");
            RibbonPanel panelDetalhamento = CriarPainel(application, tab, "Detalhamento");         

            // Caminho do assembly
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            #region BotaoFamiliasAninhadas
            // Cria os dados do botão
            PushButtonData familiasAninhadasData = new PushButtonData(
                "Familias Aninhadas",            // Nome interno do botão
                "Familias\nAninhadas",            // Texto exibido no botão
                thisAssemblyPath,          // Caminho do assembly onde o comando está localizado
                "ProjetaHDR.FamiliasAninhadas"    // Nome completo da classe de comando
            );

            // Adiciona o botão ao painel e verifica se a adição foi bem-sucedida
            PushButton familiasAninhadasButton = panelTabelas.AddItem(familiasAninhadasData) as PushButton;

            
            if (familiasAninhadasButton != null)
            {
                // Define uma dica (tooltip) que aparecerá quando o usuário passar o mouse sobre o botão
                familiasAninhadasButton.ToolTip = "Adiciona o valor do parametro Abreviatura do sistema em familias aninhadas com o valor da familia hospedeira";

                // Define o caminho para o ícone do botão
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", "aninhadas.ico");

                // Cria a imagem do ícone
                Uri uri = new Uri(iconPath);
                BitmapImage bitmap = new BitmapImage(uri);

                // Define a imagem como o ícone do botão
                familiasAninhadasButton.LargeImage = bitmap;
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


            #region BotaoFluxo
            // Cria os dados do botão
            PushButtonData fluxoData = new PushButtonData(
                "Fluxo",            // Nome interno do botão
                "Fluxo\n(Em Teste)",            // Texto exibido no botão
                thisAssemblyPath,          // Caminho do assembly onde o comando está localizado
                "ProjetaHDR.Fluxo"    // Nome completo da classe de comando
            );

            // Adiciona o botão ao painel e verifica se a adição foi bem-sucedida
            PushButton fluxoButton = panelDetalhamento.AddItem(fluxoData) as PushButton;

            // Se o botão foi criado com sucesso
            if (fluxoButton != null)
            {
                // Define uma dica (tooltip) que aparecerá quando o usuário passar o mouse sobre o botão
                fluxoButton.ToolTip = "Adiciona a seta de fluxo a tubulações";

                // Define o caminho para o ícone do botão
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", "fluxo.ico");

                // Cria a imagem do ícone
                Uri uri = new Uri(iconPath);
                BitmapImage bitmap = new BitmapImage(uri);

                // Define a imagem como o ícone do botão
                fluxoButton.LargeImage = bitmap;
            }
            #endregion

            return Result.Succeeded;
        }

        
        public RibbonPanel CriarPainel(UIControlledApplication application, string nomeAba, string nomePainel)
        {
            RibbonPanel ribbonPanel = null;
            try
            {
                ribbonPanel = application.CreateRibbonPanel(nomeAba, nomePainel);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Erro ao criar o painel '{nomePainel}': {ex.Message}");
            }

            return ribbonPanel;
        }
    }
}

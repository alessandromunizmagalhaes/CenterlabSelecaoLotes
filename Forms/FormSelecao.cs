using SAPbouiCOM;
using SAPHelper;
using System;

namespace CenterlabSelecaoLotes
{
    abstract class FormSelecao : SAPHelper.Form
    {
        public abstract string MtxLotesDisponiveisUID { get; }
        public abstract string MtxLotesSelecionadosUID { get; }
        public abstract string BtnDeselecionarUID { get; }
        public abstract string BtnSelecaoAutomaticaUID { get; }

        public override void OnAfterFormLoad(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            using (var formCOM = new FormCOM(FormUID))
            {
                if (Program._selecaoAutomatica)
                {
                    SelecionarLotes(formCOM.Form);
                    Program._selecaoAutomatica = false;
                }
            }
        }

        private void SelecionarLotes(SAPbouiCOM.Form form)
        {
            try
            {
                var mtxItens = GetMatrix(form, "3");
                var mtxLotesDisponiveis = GetMatrix(form, MtxLotesDisponiveisUID);
                var mtxLotesSelecionados = GetMatrix(form, MtxLotesSelecionadosUID);
                var btnDeselecionar = form.Items.Item(BtnDeselecionarUID);

                for (int i = 1; i <= mtxItens.VisualRowCount; i++)
                {
                    // selecionando a linha
                    mtxItens.Columns.Item("1").Cells.Item(i).Click();

                    for (int j = 1; j <= mtxLotesSelecionados.VisualRowCount; j++)
                    {
                        mtxLotesSelecionados.Columns.Item("1").Cells.Item(j).Click();
                        btnDeselecionar.Click();
                    }

                    // fazendo a ordenação da data
                    mtxLotesDisponiveis.Columns.Item("15").TitleObject.Click(BoCellClickType.ct_Double);

                    form.Items.Item(BtnSelecaoAutomaticaUID).Click();
                }

                form.Visible = true;
                form.Items.Item("1").Click();
                form.Close();
            }
            catch (Exception e)
            {
                Dialogs.PopupError("Erro interno. Erro ao selecionar os lotes automaticamente.\nErro: " + e.Message);
            }
        }
    }
}

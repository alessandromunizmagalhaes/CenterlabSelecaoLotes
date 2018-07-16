using SAPbouiCOM;
using SAPHelper;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CenterlabSelecaoLotes
{
    static class Program
    {
        public static bool _selecaoAutomatica = false;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ConectarComSAP();

            DeclararEventos();

            Dialogs.Info("Addon de Seleção de Lotes iniciado.");

            // deixa a aplicação ativa
            System.Windows.Forms.Application.Run();
        }

        private static void ConectarComSAP()
        {
            try
            {
                SAPConnection.GetDICompany();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                System.Windows.Forms.Application.Exit();
            }
        }


        private static void DeclararEventos()
        {
            var eventFilters = new EventFilters();
            eventFilters.Add(BoEventTypes.et_MENU_CLICK);

            try
            {
                var formPedidoVenda = new FormPedidoVenda();
                var formSelecaoLotes = new FormSelecaoLotes();
                var formSelecaoSeries = new FormSelecaoSeries();

                FormEvents.DeclararEventos(eventFilters, new List<MapEventsToForms>() {
                    new MapEventsToForms(BoEventTypes.et_ITEM_PRESSED, formPedidoVenda),
                });
                FormEvents.DeclararEventos(eventFilters, new List<MapEventsToForms>() {
                    new MapEventsToForms(BoEventTypes.et_FORM_LOAD, new List<SAPHelper.Form>(){
                        formSelecaoLotes,
                        formSelecaoSeries
                    }),
                });
            }
            catch (Exception e)
            {
                Dialogs.PopupError("Erro ao declarar eventos de formulário.\nErro: " + e.Message);
            }

            try
            {
                Global.SBOApplication.SetFilter(eventFilters);
            }
            catch (Exception e)
            {
                Dialogs.PopupError("Erro ao setar eventos declarados da aplicação.\nErro: " + e.Message);
            }

            Global.SBOApplication.ItemEvent += FormEvents.ItemEvent;
        }
    }

    enum ItemAdministrado
    {
        Nenhum, PorLote, PorSerie
    }
}

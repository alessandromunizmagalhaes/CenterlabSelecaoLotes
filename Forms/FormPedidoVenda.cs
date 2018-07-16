using SAPbouiCOM;
using SAPHelper;
using System;
using System.Collections.Generic;

namespace CenterlabSelecaoLotes
{
    class FormPedidoVenda : SAPHelper.Form
    {
        public override string FormType { get { return ((int)FormTypes.PedidoDeVenda).ToString(); } }

        public override void OnBeforeItemPressed(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            using (var formCOM = new FormCOM(FormUID))
            {
                var form = formCOM.Form;
                if (pVal.ItemUID == "1" && form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    var mtx = GetMatrix(form, "38");
                    var firstItemCode = mtx.Columns.Item("1").Cells.Item(1).Specific.Value;
                    var firstItemType = ItemAdministradoPor(firstItemCode);
                    var nextRow = GetNextRow(firstItemCode, firstItemType, mtx);

                    Program._selecaoAutomatica = true;
                    AbrirFormDeSelecao(form, mtx, 1);
                    if (nextRow > 0)
                    {
                        Program._selecaoAutomatica = true;
                        AbrirFormDeSelecao(form, mtx, nextRow);
                    }
                }
            }
        }

        private int GetNextRow(string firstItemCode, ItemAdministrado firstItemType, Matrix mtx)
        {
            int res = -1;
            var items = new List<string>();
            if (mtx.RowCount > 2)
            {
                switch (firstItemType)
                {
                    case ItemAdministrado.PorLote:
                        items = GetItemsPorAdministracao(ItemAdministrado.PorSerie, mtx);
                        break;
                    case ItemAdministrado.PorSerie:
                        items = GetItemsPorAdministracao(ItemAdministrado.PorLote, mtx);
                        break;
                    case ItemAdministrado.Nenhum:
                    default:
                        break;
                }
            }

            if (items.Count > 0)
            {
                res = GetNextRowByItens(items, mtx);
            }

            return res;
        }

        private int GetNextRowByItens(List<string> items, Matrix mtx)
        {
            int res = -1;
            for (int i = 2; i < mtx.RowCount; i++)
            {
                if (items.Contains(mtx.Columns.Item("1").Cells.Item(i).Specific.Value))
                {
                    res = i;
                    break;
                }
            }

            if (res < 0)
            {
                throw new Exception("Nenhum dos itens administrados foram encontrados.");
            }

            return res;
        }

        private List<string> GetItemsPorAdministracao(ItemAdministrado administracaoAProcura, Matrix mtx)
        {
            List<string> res = new List<string>();
            string itemcodes = GetItemCodeList(mtx);
            using (var rsCOM = new RecordSet())
            {
                var campo = string.Empty;
                if (administracaoAProcura == ItemAdministrado.PorLote)
                {
                    campo = "ManBtchNum";
                }
                else if (administracaoAProcura == ItemAdministrado.PorSerie)
                {
                    campo = "ManSerNum";
                }
                else
                {
                    throw new Exception("Tipo de Administração de Item não encontrado.");
                }

                var rs = rsCOM.DoQuery($"SELECT {campo}, ItemCode FROM OITM WHERE ItemCode IN ({itemcodes})");
                while (!rs.EoF)
                {
                    if (rs.Fields.Item(campo).Value == "Y")
                    {
                        res.Add(rs.Fields.Item("ItemCode").Value);
                    }

                    rs.MoveNext();
                }
            }
            return res;
        }

        private static string GetItemCodeList(Matrix mtx)
        {
            var itemcodes = string.Empty;
            for (int i = 2; i < mtx.RowCount; i++)
            {
                itemcodes += ",'" + mtx.Columns.Item("1").Cells.Item(i).Specific.Value + "'";
            }
            itemcodes = itemcodes.Remove(0, 1);
            return itemcodes;
        }

        private void AbrirFormDeSelecao(SAPbouiCOM.Form form, Matrix mtx, int row)
        {
            try
            {
                form.Freeze(true);
                form.Items.Item("112").Click();
                mtx.Columns.Item("11").Cells.Item(row).Click();
                form.Freeze(false);
                Global.SBOApplication.ActivateMenuItem("5896");
            }
            catch (Exception e)
            {
                Dialogs.PopupError("Erro interno. Não foi possível selecionar os lotes automaticamente.\nErro: " + e.Message);
            }
        }

        private ItemAdministrado ItemAdministradoPor(string itemcode)
        {
            using (var rsCOM = new RecordSet())
            {
                var rs = rsCOM.DoQuery($"SELECT ManBtchNum, ManSerNum FROM OITM WHERE ITEMCODE = '{itemcode}'");
                if (rs.Fields.Item("ManBtchNum").Value == "Y")
                {
                    return ItemAdministrado.PorLote;
                }
                else if (rs.Fields.Item("ManSerNum").Value == "Y")
                {
                    return ItemAdministrado.PorSerie;
                }
                else
                {
                    return ItemAdministrado.Nenhum;
                }
            }
        }
    }
}

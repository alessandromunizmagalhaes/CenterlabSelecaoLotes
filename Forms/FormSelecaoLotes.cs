namespace CenterlabSelecaoLotes
{
    class FormSelecaoLotes : FormSelecao
    {
        public override string FormType { get { return "42"; } }
        public override string MtxLotesDisponiveisUID { get { return "4"; } }
        public override string MtxLotesSelecionadosUID { get { return "5"; } }
        public override string BtnDeselecionarUID { get { return "47"; } }
        public override string BtnSelecaoAutomaticaUID { get { return "16"; } }
    }
}
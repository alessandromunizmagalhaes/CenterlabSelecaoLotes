namespace CenterlabSelecaoLotes
{
    class FormSelecaoSeries : FormSelecao
    {
        public override string FormType { get { return "25"; } }
        public override string MtxLotesDisponiveisUID { get { return "5"; } }
        public override string MtxLotesSelecionadosUID { get { return "55"; } }
        public override string BtnDeselecionarUID { get { return "6"; } }
        public override string BtnSelecaoAutomaticaUID { get { return "4"; } }
    }
}

using AzulBusiness.Interfaces;
using System.Collections.Generic;

namespace AzulBusiness
{
    public class LayoutBusiness : ILayoutBusiness
    {

        public string GetTableTitle(string system, string flow, string pnrx)
        {
            var dictionaryTableTitle = GetTableReferenceByTitle();
            string tableTitle;

            if (dictionaryTableTitle.TryGetValue($"{system}-{flow}-{pnrx}", out tableTitle))
            {
                return tableTitle;
            }

            return "Table Not Found";
        }


        public Dictionary<string, string> GetTableReferenceByTitle()
        {
            var tableTitles = new Dictionary<string, string> {
                #region PNR1 
                {"1", "VOO_ORIGINAL_IDA_PNR1"},
                {"2", "VOO_ORIGINAL_VOLTA_PNR1"},
                {"3", "MOVIMENTACAO_DE_VOOS_IDA_PNR1"},
                {"4", "MOVIMENTACAO_DE_VOOS_VOLTA_PNR1"},
                {"5", "TRECHO_FINAL_IDA_PNR1"},
                {"6", "TRECHO_FINAL_VOLTA_PNR1"},
                {"ss-bd-PNR1", "ESCLARECIMENTO_SKYSPEED_BREAKDOWN_PNR1"},
                {"ss-rs-PNR1", "ESCLARECIMENTO_SKYSPEED_SUMMARY_PNR1"},
                {"8", "INTEGRIDADE_OPERACIONAL_CICLOS_PNR1"},
                {"9", "INTEGRIDADE_OPERACIONAL_ALTERNADOS_PNR1"},
                {"10", "INTEGRIDADE_OPERACIONAL_ATRASOS_PNR1"},
                {"11", "INTEGRIDADE_OPERACIONAL_CANCELADO_PNR1"},
                {"12", "ESCLARECIMENTO_SKYSPEED_PAYMENTS_PNR1"},
                {"ss-ps-PNR1", "DADOS_DO_PAYMENTS_PNR1"},
                {"md-default-PNR1", "ESCLARECIMENTO_TUDOAZUL_MIDDLEWARE_PNR1"},
                {"ss-cs-PNR1", "ESCLARECIMENTO_SKYSPEED_COMMENTS_PNR1"},
                {"ss-rr-PNR1", "ESCLARECIMENTO_SKYSPEED_RETRIEVER_PNR1"},
                {"rm-default-PNR1", "ESCLARECIMENTO_REDEMET_PNR1"},
                {"ac-default-PNR1", "ESCLARECIMENTO_ANAC_PNR1"},
                {"fo-default-PNR1", "ESCLARECIMENTO_FLIGHT_UTILITIES_PNR1"},
                #endregion
                #region PNR2 
                {"13", "VOO_ORIGINAL_IDA_PNR2"},
                {"14", "VOO_ORIGINAL_VOLTA_PNR2"},
                {"15", "MOVIMENTACAO_DE_VOOS_IDA_PNR2"},
                {"16", "MOVIMENTACAO_DE_VOOS_VOLTA_PNR2"},
                {"17", "TRECHO_FINAL_IDA_PNR2"},
                {"18", "TRECHO_FINAL_VOLTA_PNR2"},
                {"ss-bd-PNR2", "ESCLARECIMENTO_SKYSPEED_BREAKDOWN_PNR2"},
                {"ss-rs-PNR2", "ESCLARECIMENTO_SKYSPEED_SUMMARY_PNR2"},
                {"19", "INTEGRIDADE_OPERACIONAL_CICLOS_PNR2"},
                {"20", "INTEGRIDADE_OPERACIONAL_ALTERNADOS_PNR2"},
                {"21", "INTEGRIDADE_OPERACIONAL_ATRASOS_PNR2"},
                {"22", "INTEGRIDADE_OPERACIONAL_CANCELADO_PNR2"},
                {"ss-ps-PNR2", "ESCLARECIMENTO_SKYSPEED_PAYMENTS_PNR2"},
                {"23", "DADOS_DO_PAYMENTS_PNR2"},
                {"md-default-PNR2", "ESCLARECIMENTO_TUDOAZUL_MIDDLEWARE_PNR2"},
                {"ss-cs-PNR2", "ESCLARECIMENTO_SKYSPEED_COMMENTS_PNR2"},
                {"ss-rr-PNR2", "ESCLARECIMENTO_SKYSPEED_RETRIEVER_PNR2"},
                {"rm-default-PNR2", "ESCLARECIMENTO_REDEMET_PNR2"},
                {"ac-default-PNR2", "ESCLARECIMENTO_ANAC_PNR2"},
                {"fo-default-PNR2", "ESCLARECIMENTO_FLIGHT_UTILITIES_PNR2"},
                #endregion
                #region PNR3 
                {"24", "VOO_ORIGINAL_IDA_PNR3"},
                {"25", "VOO_ORIGINAL_VOLTA_PNR3"},
                {"26", "MOVIMENTACAO_DE_VOOS_IDA_PNR3"},
                {"27", "MOVIMENTACAO_DE_VOOS_VOLTA_PNR3"},
                {"28", "TRECHO_FINAL_IDA_PNR3"},
                {"29", "TRECHO_FINAL_VOLTA_PNR3"},
                {"ss-bd-PNR3", "ESCLARECIMENTO_SKYSPEED_BREAKDOWN_PNR3"},
                {"ss-rs-PNR3", "ESCLARECIMENTO_SKYSPEED_SUMMARY_PNR3"},
                {"30", "INTEGRIDADE_OPERACIONAL_CICLOS_PNR3"},
                {"31", "INTEGRIDADE_OPERACIONAL_ALTERNADOS_PNR3"},
                {"32", "INTEGRIDADE_OPERACIONAL_ATRASOS_PNR3"},
                {"33", "INTEGRIDADE_OPERACIONAL_CANCELADO_PNR3"},
                {"ss-ps-PNR3", "ESCLARECIMENTO_SKYSPEED_PAYMENTS_PNR3"},
                {"34", "DADOS_DO_PAYMENTS_PNR3"},
                {"md-default-PNR3", "ESCLARECIMENTO_TUDOAZUL_MIDDLEWARE_PNR3"},
                {"ss-cs-PNR3", "ESCLARECIMENTO_SKYSPEED_COMMENTS_PNR3"},
                {"ss-rr-PNR3", "ESCLARECIMENTO_SKYSPEED_RETRIEVER_PNR3"},
                {"rm-default-PNR3", "ESCLARECIMENTO_REDEMET_PNR3"},
                {"ac-default-PNR3", "ESCLARECIMENTO_ANAC_PNR3"},
                {"fo-default-PNR3", "ESCLARECIMENTO_FLIGHT_UTILITIES_PNR3"}
                #endregion
            };

            return tableTitles;
        }
    }
}

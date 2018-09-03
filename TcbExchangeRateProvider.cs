using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Xml;
using HtmlAgilityPack;
using Nop.Core;
using Nop.Core.Plugins;
using Nop.Plugin.ExchangeRate.TcbExchange;
using Nop.Services.Directory;
using Nop.Services.Localization;
using Nop.Services.Logging;

namespace Nop.Plugin.ExchangeRate.EcbExchange
{
    public class TcbExchangeRateProvider : BasePlugin, IExchangeRateProvider
    {
        #region Fields

        private readonly ILocalizationService _localizationService;
        private readonly ILogger _logger;

        #endregion

        #region Ctor

        public TcbExchangeRateProvider(ILocalizationService localizationService,
            ILogger logger)
        {
            this._localizationService = localizationService;
            this._logger = logger;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Gets currency live rates
        /// Rate Url is : https://www.tcb-bank.com.tw/finance_info/Pages/foreign_spot_rate.aspx
        /// </summary>
        /// <param name="exchangeRateCurrencyCode">Exchange rate currency code</param>
        /// <returns>Exchange rates</returns>
        public IList<Core.Domain.Directory.ExchangeRate> GetCurrencyLiveRates(string exchangeRateCurrencyCode)
        {
            if (exchangeRateCurrencyCode == null)
                throw new ArgumentNullException(nameof(exchangeRateCurrencyCode));

            //add euro with rate 1
            var ratesToTwd = new List<Core.Domain.Directory.ExchangeRate>
            {
                new Core.Domain.Directory.ExchangeRate
                {
                    CurrencyCode = "TWD",
                    Rate = 1,
                    UpdatedOn = DateTime.UtcNow
                }
            };

            var currencyList = new HtmlWeb().Load("https://www.tcb-bank.com.tw/finance_info/Pages/foreign_spot_rate.aspx")
                .DocumentNode.SelectNodes("//table[@id='ctl00_PlaceHolderEmptyMain_PlaceHolderMain_fecurrentid_gvResult']/tr");
            var rateResult = new List<TcbRateObject>();
            for (int i = 1; i < currencyList.Count; i+=2)
            {
                string currency = currencyList[i].SelectSingleNode("td[1]").InnerText.Trim().Replace("&nbsp;", String.Empty);
                string currencyCode = currencyList[i+1].SelectSingleNode("td[1]").InnerText.Trim().Replace("&nbsp;", String.Empty);
                string spotBuying = currencyList[i].SelectSingleNode("td[3]").InnerText.Trim();
                string cashBuying = currencyList[i].SelectSingleNode("td[4]").InnerText.Trim();
                string spotSelling = currencyList[i+1].SelectSingleNode("td[3]").InnerText.Trim();
                string cashSelling = currencyList[i+1].SelectSingleNode("td[4]").InnerText.Trim();

                if (String.IsNullOrEmpty(cashBuying) && String.IsNullOrEmpty(cashSelling))
                {
                    cashBuying = "-";
                    cashSelling = "-";
                }

                rateResult.Add(new TcbRateObject
                {
                    Currency = currency,
                    CurrencyCode = currencyCode,
                    CashBuying = (!cashBuying.Contains("-")) ? Convert.ToDecimal(cashBuying) : new decimal?(),
                    CashSelling = (!cashSelling.Contains("-")) ? Convert.ToDecimal(cashSelling) : new decimal?(),
                    SpotBuying = (!spotBuying.Contains("-")) ? Convert.ToDecimal(spotBuying) : new decimal?(),
                    SpotSelling = (!spotSelling.Contains("-")) ? Convert.ToDecimal(spotSelling) : new decimal?()
                });
            }

            //converr rate to nop currency rate object
            foreach(var rate in rateResult)
            { 
                var averageRate = rate.SpotBuying!=null
                    ? (1 / (( rate.SpotBuying.Value + rate.SpotSelling.Value ) / 2))
                    : (1 / (( rate.CashBuying.Value + rate.CashBuying.Value ) / 2));

                ratesToTwd.Add(new Core.Domain.Directory.ExchangeRate
                { 
                    CurrencyCode = rate.CurrencyCode,
                    Rate = averageRate,
                    UpdatedOn = DateTime.UtcNow,
                });
            }

            //return result for the euro
            if (exchangeRateCurrencyCode.Equals("twd", StringComparison.InvariantCultureIgnoreCase))
                return ratesToTwd;

            //use only currencies that are supported by TCB
            var exchangeRateCurrency = ratesToTwd.FirstOrDefault(rate => rate.CurrencyCode.Equals(exchangeRateCurrencyCode, StringComparison.InvariantCultureIgnoreCase));
            if (exchangeRateCurrency == null)
                throw new NopException(_localizationService.GetResource("Plugins.ExchangeRate.TcbExchange.Error"));

            //return result for the selected (not euro) currency
            return ratesToTwd.Select(rate => new Core.Domain.Directory.ExchangeRate
            {
                CurrencyCode = rate.CurrencyCode,
                Rate = Math.Round(rate.Rate / exchangeRateCurrency.Rate, 4),
                UpdatedOn = rate.UpdatedOn
            }).ToList();
        }

        /// <summary>
        /// Install the plugin
        /// </summary>
        public override void Install()
        {
            //locales
            _localizationService.AddOrUpdatePluginLocaleResource("Plugins.ExchangeRate.TcbExchange.Error", "You can use TCB (European central bank) exchange rate provider only when the primary exchange rate currency is supported by TCB");

            base.Install();
        }

        /// <summary>
        /// Uninstall the plugin
        /// </summary>
        public override void Uninstall()
        {
            //locales
            _localizationService.DeletePluginLocaleResource("Plugins.ExchangeRate.TcbExchange.Error");

            base.Uninstall();
        }

        #endregion

    }
}
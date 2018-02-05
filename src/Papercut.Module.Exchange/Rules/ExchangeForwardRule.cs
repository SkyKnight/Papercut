using Microsoft.Exchange.WebServices.Data;
using Papercut.Common.Extensions;
using Papercut.Core.Domain.Rules;
using Papercut.Rules;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Papercut.Module.Exchange.Rules
{
    [Serializable]
    public class ExchangeForwardRule : RuleBase
    {
        private string _autodiscoveryUrl;

        [Category("Settings")]
        [DisplayName("Autodiscovery URL")]
        public string AutodiscoveryUrl
        {
            get { return _autodiscoveryUrl; }
            set
            {
                _autodiscoveryUrl = value;
                OnPropertyChanged(nameof(AutodiscoveryUrl));
            }
        }

        private string _url;

        [Category("Settings")]
        [DisplayName("Exchange URL")]
        public string Url
        {
            get { return _url; }
            set
            {
                _url = value;
                OnPropertyChanged(nameof(Url));
            }
        }

        private string _userName;

        [Category("Settings")]
        [DisplayName("User name")]
        public string UserName
        {
            get { return _userName; }
            set
            {
                _userName = value;
                OnPropertyChanged(nameof(UserName));
            }
        }

        private string _password;

        [Category("Settings")]
        [DisplayName("Password")]
        [PasswordPropertyText]
        public string Password
        {
            get { return _password; }
            set
            {
                _password = value;
                OnPropertyChanged(nameof(Password));
            }
        }

        private string _domain;

        [Category("Settings")]
        [DisplayName("Domain name")]
        public string Domain
        {
            get { return _domain; }
            set
            {
                _domain = value;
                OnPropertyChanged(nameof(Domain));
            }
        }

        private ExchangeVersion _exchangeVersion;

        [Category("Settings")]
        [DisplayName("Exchange version")]
        [DefaultValue(ExchangeVersion.Exchange2007_SP1)]
        public ExchangeVersion ExchangeVersion
        {
            get { return _exchangeVersion; }
            set
            {
                _exchangeVersion = value;
                OnPropertyChanged(nameof(ExchangeVersion));
            }
        }

        protected override IEnumerable<KeyValuePair<string, Lazy<object>>> GetPropertiesForDescription()
        {
            return base.GetPropertiesForDescription().Concat(this.GetProperties());
        }
    }
}

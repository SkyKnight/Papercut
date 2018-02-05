using Papercut.Core.Domain.Rules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Papercut.Core.Domain.Message;
using Papercut.Message;
using System.Reactive.Linq;
using Papercut.Message.Helpers;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using Serilog;
using MimeKit;

namespace Papercut.Module.Exchange.Rules
{
    public class ExchangeForwardRuleDispatch : IRuleDispatcher<ExchangeForwardRule>
    {
        private readonly Lazy<MimeMessageLoader> _mimeMessageLoader;
        private readonly ILogger _logger;

        public ExchangeForwardRuleDispatch(Lazy<MimeMessageLoader> mimeMessageLoader, ILogger logger)
        {
            _mimeMessageLoader = mimeMessageLoader;
            _logger = logger;
        }

        public void Dispatch(ExchangeForwardRule rule, MessageEntry messageEntry)
        {
            if (rule == null) throw new ArgumentNullException(nameof(rule));
            if (messageEntry == null) throw new ArgumentNullException(nameof(messageEntry));

            _mimeMessageLoader.Value.Get(messageEntry)
                .Select(m => m.CloneMessage())
                .Subscribe(m =>
                {
                    try
                    {
                        var service = new ExchangeService(rule.ExchangeVersion);
                        if (!string.IsNullOrWhiteSpace(rule.UserName) && !string.IsNullOrWhiteSpace(rule.Password))
                            service.Credentials = new NetworkCredential(rule.UserName, rule.Password, rule.Domain);

                        if (!string.IsNullOrWhiteSpace(rule.Url))
                            service.Url = new Uri(rule.Url);
                        else if (!string.IsNullOrWhiteSpace(rule.AutodiscoveryUrl))
                            service.AutodiscoverUrl(rule.AutodiscoveryUrl);

                        var msg = new EmailMessage(service);
                        msg.Subject = m.Subject;
                        msg.Body = string.IsNullOrWhiteSpace(m.HtmlBody) ? new MessageBody(BodyType.Text, m.TextBody) : new MessageBody(BodyType.HTML, m.HtmlBody);
                        var from = m.From.OfType<MailboxAddress>().FirstOrDefault();
                        if(from != null)
                            msg.From = new EmailAddress(from.Name, from.Address);

                        foreach (var r in m.To)
                        {
                            var recipient = r as MailboxAddress;
                            if(recipient != null)
                                msg.ToRecipients.Add(new EmailAddress(recipient.Address));
                        }

                        msg.Send();
                    }
                    catch(Exception ex)
                    {
                        _logger.Error(ex, "Failure sending {@MessageEntry} for rule {@Rule}", messageEntry, rule);
                    }
                });
        }
    }
}

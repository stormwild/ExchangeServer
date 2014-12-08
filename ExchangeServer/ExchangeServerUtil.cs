using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeServer
{
	public class ExchangeServerUtil
	{
		private readonly ExchangeService _service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
		private EmailMessage _message;

		public ExchangeServerUtil(string email)
		{
			_service.UseDefaultCredentials = true;
			_service.AutodiscoverUrl(email, RedirectionUrlValidationCallback);
		}

		public void CreateMessage(IEnumerable<string> to, string subject, string body)
		{
			_message = new EmailMessage(_service);
			_message.ToRecipients.AddRange(to);

			_message.Subject = subject;
			_message.Body = body;
		}

		public void CreateMessage(IEnumerable<string> to, IEnumerable<string> cc, IEnumerable<string> bcc, string subject, string body,
			Importance priority, string attachment)
		{
			_message = new EmailMessage(_service);

			_message.ToRecipients.AddRange(to);
			_message.CcRecipients.AddRange(cc);
			_message.BccRecipients.AddRange(bcc);

			_message.Subject = subject;
			_message.Body = body;

			_message.Importance = priority;
			_message.Attachments.AddFileAttachment(attachment);
		}

		public void Send()
		{
			if (_message != null) { _message.Send(); }
		}

		private static bool RedirectionUrlValidationCallback(string redirectionUrl)
		{
			var result = false;

			var redirectionUri = new Uri(redirectionUrl);
			if (redirectionUri.Scheme == "https")
			{
				result = true;
			}
			return result;
		}
	}
}

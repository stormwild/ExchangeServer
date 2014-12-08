using System.Collections.Generic;
using ExchangeServer;

namespace ExchangeServerUtilTest
{
	class Program
	{
		static void Main(string[] args)
		{
			var util = new ExchangeServerUtil("user@contoso.com");
			util.CreateMessage(new List<string> { "user@example.com" }, "Test", "Test");
			util.Send();
		}
	}
}

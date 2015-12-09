using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Linq.Expressions;
using System.Dynamic;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using System.Reflection;
using System.Net;
using System.Xml.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.Security;

namespace OfficeDevPnP.Core.Tests.Utilities
{
    public class ClientContextWrapper
    {
        List<string> executedQueries;

        public ClientContextWrapper()
        {
            executedQueries = new List<string>();
        }
        public bool WrapExecuteQuery(Action action, ClientContext ctx)
        {
            action();

            var pendingRequest = ctx.PendingRequest;
            var pendingRequestQueryContent = GetQuery(pendingRequest);

            string result = string.Empty;

            XmlSerializer serializer = new XmlSerializer(typeof(Request));

            using (TextReader reader = new StringReader(pendingRequestQueryContent.ToString()))
            {
                Request xml = (Request)serializer.Deserialize(reader);


                result = JsonConvert.SerializeObject(xml);
            }

            if (executedQueries.Any(e => e == result))
            {
                Debug.WriteLine("Already called this");
                return true;
            }

            executedQueries.Add(result);

            ctx.ExecuteQuery();

            return false;
        }

        private StringBuilder GetQuery(ClientRequest req)
        {
            MethodInfo dynMethod = req.GetType().GetMethod("BuildQuery", BindingFlags.NonPublic | BindingFlags.Instance);
            var x = dynMethod.Invoke(req, new object[] { });

            FieldInfo field = x.GetType().GetField("m_sb", BindingFlags.NonPublic | BindingFlags.Instance);
            var value = field.GetValue(x);

            return value as StringBuilder;
        }
    }

    [TestClass]
    public class ClientContextWrapperTests
    {
        public static SecureString GetSecureString(string input)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("Input string is empty and cannot be made into a SecureString", "input");

            var secureString = new SecureString();
            foreach (char c in input.ToCharArray())
                secureString.AppendChar(c);

            return secureString;
        }
        [TestMethod]
        public void ItSkipsCachedClientContextCalls()
        {
            var devSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];

            var userName = ConfigurationManager.AppSettings["SPOUserName"];
            var password = ConfigurationManager.AppSettings["SPOPassword"];

            var secPassword = GetSecureString(password);
            var credentials = new SharePointOnlineCredentials(userName, secPassword);

            ClientContext context;
            context = new ClientContext(devSiteUrl);
            context.Credentials = credentials;

            using (var ctx = context)
            {

                var wrapper = new ClientContextWrapper();

                bool cached = wrapper.WrapExecuteQuery(() =>
                {
                    ctx.Load(ctx.Web, c => c.Title);
                }, ctx);

                Assert.AreEqual(false, cached);

                cached = wrapper.WrapExecuteQuery(() =>
                {
                    ctx.Load(ctx.Web, c => c.Title);
                }, ctx);

                Assert.AreEqual(false, cached);

                cached = wrapper.WrapExecuteQuery(() =>
                {
                    ctx.Load(ctx.Web, c => c.Title);
                }, ctx);

                Assert.AreEqual(true, cached);

                cached = wrapper.WrapExecuteQuery(() =>
                {
                    ctx.Load(ctx.Web, c => c.Title, c => c.UIVersion);
                }, ctx);

                Assert.AreEqual(false, cached);

                cached = wrapper.WrapExecuteQuery(() =>
                {
                    ctx.Load(ctx.Web, c => c.Title, c => c.Webs);
                }, ctx);

                Assert.AreEqual(false, cached);

                for (int i = 0; i < 1000; i++)
                {
                    cached = wrapper.WrapExecuteQuery(() =>
                    {
                        ctx.Load(ctx.Web, c => c.Title, c => c.Webs);
                    }, ctx);

                    if(cached)
                    {
                        //do nothing
                    }
                }
            }
        }
    }
}

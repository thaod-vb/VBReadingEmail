using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.FilterOperators;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VBDCS_SQLFactory.DataAccess;

namespace ReadingEmail
{
    class Program
    {
        static void Main(string[] args)
        {

            string constr = "Data Source=aphrodite;Initial Catalog=VideoBankMaster;User ID=sa;Password=vbpass12#";

            // Set up the config to load the user secrets
            //var config = new ConfigurationBuilder()
            //    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            //    .AddUserSecrets<Program>()
            //    .Build();

            // Define the credentials. 
            // Note: In your implementations of this code, please consider using managed identities, and avoid credentials in code or config.
            var credentials = new ClientSecretCredential(
                "a5ed794a-a4fe-4d02-87f8-3093b1f07ba8",//config["GraphMail:TenantId"],
                "e212c173-b9c7-45e1-b29e-efe9aa3d85f3",//config["GraphMail:ClientId"],
                "Nj~8Q~OWwQeuBkOz7aM5mbuJwt~Y-8OEC.YlkdhB",//config["GraphMail:ClientSecret"],
                new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud });

            // Define a new Microsoft Graph service client.            
            GraphServiceClient _graphServiceClient = new GraphServiceClient(credentials);

            using (ImageGenerationDBDataContext db = new ImageGenerationDBDataContext(constr))
            {
                //var outlookConfigs = db.OutlookConfigs.ToList();
                var currentDateTime = db.ExecuteQuery<DateTime>("Select Getdate()").First().AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                Console.WriteLine($"currentDateTime: {currentDateTime}");
                //foreach (var configItem in outlookConfigs)
                //{
                foreach (var ruleItem in db.OutlookRules.Where(r => r.IsActive.GetValueOrDefault() == true && r.UsernameId == 198))
                {
                    StringBuilder conditionFilterString = new StringBuilder();
                    StringBuilder conditionSearchString = new StringBuilder();
                    conditionFilterString.Append($"ReceivedDateTime le {currentDateTime}");
                    foreach (var conditionItem in ruleItem.OutlookConditions)
                    {
                        switch (conditionItem.ConditionType.GetValueOrDefault())
                        {
                            case 1:
                                conditionFilterString.Append($" and From/EmailAddress/Address eq '{conditionItem.ConditionValue}'");
                                break;
                                //case 2:
                                //    conditionFilterString.Append($" and ToRecipients/any(t:t/EmailAddress/Address eq '{conditionItem.ConditionValue}')");
                                //    //conditionSearchString.Append($"\"to:{conditionItem.ConditionValue}\"");
                                //    break;
                        }
                    }

                    Console.WriteLine($"conditionFilterString: {conditionFilterString.ToString()}");
                    //Console.WriteLine($"conditionSearchString: {conditionSearchString.ToString()}");
                    // Get the e-mails for a specific user.
                    var messages = _graphServiceClient.Users[ruleItem.Username.NotifyEmail].Messages.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter = conditionFilterString.ToString();
                            //requestConfiguration.QueryParameters.Search = conditionSearchString.ToString();
                            requestConfiguration.QueryParameters.Orderby = new string[] { "ReceivedDateTime desc" };
                        requestConfiguration.QueryParameters.Top = 10;
                        requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type='text'");
                            //requestConfiguration.QueryParameters.Count = true;
                            //requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                    }).Result.Value;

                    //foreach (var conditionItem in ruleItem.OutlookConditions)
                    //{
                    //    switch (conditionItem.ConditionType.GetValueOrDefault())
                    //    {
                    //        //case 1:
                    //        //    messages = messages.Where(m => m.From?.EmailAddress.Address == conditionItem.ConditionValue).ToList();
                    //        //    break;
                    //        case 2:
                    //            messages = (from a in messages
                    //                       where a.ToRecipients.Any(t => t.EmailAddress.Address == conditionItem.ConditionValue)
                    //                       select a).ToList();
                    //            //conditionSearchString.Append($"\"to:{conditionItem.ConditionValue}\"");
                    //            break;
                    //    }
                    //}

                    List<OutlookMessage> outlookMessages = new List<OutlookMessage>();
                    int count = 1;
                    foreach (var message in messages)
                    {
                        Console.WriteLine("=================================================================");
                        Console.WriteLine($"{message.ReceivedDateTime?.ToString("yyyy-MM-dd HH:mm:ss")} from {message.From?.EmailAddress.Address}");
                        Console.WriteLine($"To: {string.Join(";", message.ToRecipients.Select(o => o.EmailAddress.Address).ToList())}");
                        Console.WriteLine($"{message.InternetMessageId}");
                        Console.WriteLine($"{message.Subject}");
                        Console.WriteLine($"{message.Body.Content}");
                        Console.WriteLine($"{_graphServiceClient.Users[ruleItem.Username.NotifyEmail].Messages[message.Id].Content}");
                        Console.WriteLine($"{message.WebLink}");
                         var a = _graphServiceClient.Users[ruleItem.Username.NotifyEmail].Messages[message.Id].GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type='html'");
                        }).GetAwaiter().GetResult();
                        Console.WriteLine($"{a.Body.Content}");
                        //Console.WriteLine($"{Regex.Replace(message.Body.Content, "<.*?>", String.Empty)}");
                        Console.WriteLine("-----------------");
                        //Console.WriteLine($"{System.Net.WebUtility.HtmlDecode(message.Body.Content)}");
                        //Console.WriteLine(" ");
                        //Console.WriteLine($"{WebUtility.HtmlDecode(message.Body.Content)}");
                         var attachments = _graphServiceClient.Users[ruleItem.Username.NotifyEmail].Messages[message.Id].Attachments.GetAsync().GetAwaiter().GetResult();
                        if (attachments != null)
                        {
                            foreach (var attach in attachments.Value)
                            {
                                Console.WriteLine($"---------Attachments--------{attach.Name}");
                                string pathAttm = "C:\\temp\\email\\" + count + "_" + attach.Name;
                                Console.WriteLine($"pathFormat: {pathAttm}");
                                //var result = _graphServiceClient.Users[ruleItem.Username.NotifyEmail].Messages[message.Id].Attachments[attm.Id].GetAsync().GetAwaiter().GetResult();
                                File.WriteAllBytes(pathAttm, (attach as Microsoft.Graph.Models.FileAttachment).ContentBytes);
                                
                            }
                        }

                        //var newMessage = new VBDCS_SQLFactory.DataAccess.OutlookMessage();
                        //newMessage.OutlookRuleId = ruleItem.Id;
                        //newMessage.InternetMessageId = message.InternetMessageId;
                        //newMessage.Subject = message.Subject;
                        //newMessage.FromRecipient = message.From?.EmailAddress.Address;
                        //newMessage.ToRecipient = string.Join(";", message.ToRecipients.Select(o => o.EmailAddress.Address).ToList());
                        //newMessage.BccRecipient = string.Join(";", message.BccRecipients.Select(o => o.EmailAddress.Address).ToList());
                        //newMessage.CcRecipient = string.Join(";", message.CcRecipients.Select(o => o.EmailAddress.Address).ToList());
                        //newMessage.FormatDocumentText = new FormatDocumentText();
                        //newMessage.FormatDocumentText.DocumentTextHTML = message.Body.Content;
                        //newMessage.FormatDocumentText.DocumentText = message.BodyPreview;
                        //newMessage.FormatDocumentText.FormatIdKey = 240469;
                        //outlookMessages.Add(newMessage);
                        var mimeContentStream = _graphServiceClient.Users[ruleItem.Username.NotifyEmail].Messages[message.Id].Content.GetAsync().GetAwaiter().GetResult();
                        string pathFormat = "C:\\temp\\email\\" + count + ".msg";
                        Console.WriteLine($"pathFormat: {pathFormat}");
                        using (var fileStream = System.IO.File.Create(pathFormat))
                        {
                            // mimeContentStream.Seek(0, SeekOrigin.Begin);
                            mimeContentStream.CopyTo(fileStream);
                        };
                        count++;
                        Console.WriteLine("=================================================================");
                    }

                    if (outlookMessages.Count() > 0)
                    {
                        db.OutlookMessages.InsertAllOnSubmit(outlookMessages);
                        db.SubmitChanges();
                    }
                }
                //}
            }
            Console.WriteLine("***** DONE *****");
            Console.ReadLine();

            //try
            //{
            //    var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            //    var nameSpace = outlookApp.GetNamespace("MAPI");

            //    // Log in to Outlook (you can provide credentials if needed)
            //    nameSpace.Logon("", "", System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            //    // Access the desired folder (e.g., Inbox)
            //    var inboxFolder = nameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

            //    // Process unread mail items
            //    foreach (Microsoft.Office.Interop.Outlook.MailItem mailItem in inboxFolder.Items)
            //    {
            //        if (mailItem.UnRead)
            //        {
            //            Console.WriteLine($"Subject: {mailItem.Subject}");
            //            Console.WriteLine($"Body: {mailItem.Body}");
            //        }
            //    }

            //    // Log off from Outlook
            //    nameSpace.Logoff();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ex: " + ex);
            //    Console.ReadLine();

            //}
        }
    }

}

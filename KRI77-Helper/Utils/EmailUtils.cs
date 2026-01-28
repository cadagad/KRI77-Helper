using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;


namespace KRI77_Helper.Utils
{
    public static class EmailUtils
    {
        /* Hardcode paths for demo purposes */
        private const string inputPath = "https://mfc.sharepoint.com/:f:/r/sites/AICOE/Shared%20Documents/Project/Process%20Review/01%20Asset%20Management/Source%20Files/Demo/Input?csf=1&web=1&e=P90uI2";
        private const string outputPath = "https://mfc.sharepoint.com/:f:/r/sites/AICOE/Shared%20Documents/Project/Process%20Review/01%20Asset%20Management/Source%20Files/Demo/Output?csf=1&web=1&e=cSA3jn";

        public static int SendEmail(string processName, int rowsProcessed, int rowsDuplicate, string elapsedTime, bool isError)
        {
            try
            {
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                string? emailTo = configuration["FileSettings:EmailTo"];
                string? emailCc = configuration["FileSettings:EmailCc"];
                string? emailToError = configuration["FileSettings:EmailToError"];

                using var message = new MailMessage();
                message.From = new MailAddress("AI_COE_Automation@manulife.com", "AI-COE Automation");

                if (!isError) 
                {
                    if (!string.IsNullOrEmpty(emailTo))
                    {
                        foreach (string email in emailTo.Split(";"))
                        {
                            message.To.Add(new MailAddress(email));
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(emailToError))
                    {
                        foreach (string email in emailToError.Split(";"))
                        {
                            message.To.Add(new MailAddress(email));
                        }
                    }
                }

                if (!string.IsNullOrEmpty(emailCc))
                {
                    foreach (string email in emailCc.Split(";"))
                    {
                        message.CC.Add(new MailAddress(email));
                    }
                }

                if (message.To.Count() == 0 && message.CC.Count() == 0)
                {
                    Console.WriteLine("No email recipients specified. Email not sent.");
                    return 0;
                }

                if (!isError)
                {
                    message.Subject = $"KRI77 Handling : {processName} : Successfully processed - " + DateTime.UtcNow.ToString("MM/dd/yyyy");

                    // Construct the HTML body using StringBuilder for clarity
                    StringBuilder textBody = new StringBuilder();

                    textBody.Append($"<br /><h4>{processName} has been successfully processed.</h4><br />");

                    textBody.Append("<html><body>");
                    textBody.Append("<table border='1' cellpadding='5' cellspacing='0'>");
                    //textBody.Append("<tr bgcolor='#4da6ff' style='color: white;'>"); // Table header row with background color
                    //textBody.Append("<td><b>Column 1</b></td>");
                    //textBody.Append("<td><b>Column 2</b></td>");
                    //textBody.Append("</tr>");

                    // Add some data rows (can be dynamically generated from a database or data source)
                    textBody.Append($"<tr>");
                    textBody.Append($"<td>Input Path</td>");
                    textBody.Append($"<td><a href=\"{inputPath}\">AI COE - Documents - Input</a></td>");
                    textBody.Append($"</tr>");
                    textBody.Append($"<tr>");
                    textBody.Append($"<td>Output Path</td>");
                    textBody.Append($"<td><a href=\"{outputPath}\">AI COE - Documents - Output</a></td>");
                    textBody.Append($"</tr>");
                    textBody.Append($"<tr>");
                    textBody.Append($"<td>Rows Processed</td>");
                    textBody.Append($"<td>{rowsProcessed.ToString()}</td>");
                    textBody.Append($"</tr'>");
                    textBody.Append($"<tr>");
                    textBody.Append($"<td>Duplicate Rows</td>");
                    textBody.Append($"<td>{rowsDuplicate.ToString()}</td>");
                    textBody.Append($"</tr>");
                    textBody.Append($"<tr>");
                    textBody.Append($"<td>Processing Time</td>");
                    textBody.Append($"<td>{elapsedTime.ToString()}</td>");
                    textBody.Append($"</tr>");
                    textBody.Append("</table>");
                    textBody.Append("</body></html>");

                    textBody.Append($"<br /><br /><p><i>This is an automated email. Please do not reply to this message.</i></p>");



                    message.Body = textBody.ToString();

                }
                else
                {
                    message.Subject = $"{processName} : Error encountered during processing";
                    message.Body = "Please see log file for detailed error message.";
                }

                message.IsBodyHtml = true;

                using var smtpClient = new SmtpClient("mail.manulife.com");
                smtpClient.Port = 25; // Common port for internal SMTP without authentication
                smtpClient.EnableSsl = false; // Set to true if your server requires SSL/TLS
                smtpClient.UseDefaultCredentials = true; // Use current Windows credentials
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                smtpClient.Send(message);
            }
            catch (SmtpException ex)
            {
                Console.WriteLine($"SMTP Error: {ex.Message}");
                Console.WriteLine($"Status Code: {ex.StatusCode}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }

            /* Success - Email sent */
            return 0;
        }

        public static int SendErrorEmail(string in_file, string errorMessage, bool isError)
        {
            try
            {
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                string? emailTo = configuration["FileSettings:EmailTo"];
                string? emailCc = configuration["FileSettings:EmailCc"];
                string? emailToError = configuration["FileSettings:EmailToError"];

                using var message = new MailMessage();
                message.From = new MailAddress("AI_COE_Automation@manulife.com", "AI-COE Automation");

                if (!isError)
                {
                    if (!string.IsNullOrEmpty(emailTo))
                    {
                        foreach (string email in emailTo.Split(";"))
                        {
                            message.To.Add(new MailAddress(email));
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(emailToError))
                    {
                        foreach (string email in emailToError.Split(";"))
                        {
                            message.To.Add(new MailAddress(email));
                        }
                    }
                }

                if (!string.IsNullOrEmpty(emailCc))
                {
                    foreach (string email in emailCc.Split(";"))
                    {
                        message.CC.Add(new MailAddress(email));
                    }
                }

                if (message.To.Count() == 0 && message.CC.Count() == 0)
                {
                    Console.WriteLine("No email recipients specified. Email not sent.");
                    return 0;
                }

                if (isError)
                {
                    message.Subject = $"KRI77 Handling : Error encountered during processing";
                    //message.Body = "Please see log file for detailed error message.";
                    
                    //message.Subject = $"{processName} : Successfully processed - " + DateTime.UtcNow.ToString("MM/dd/yyyy");

                    // Construct the HTML body using StringBuilder for clarity
                    StringBuilder textBody = new StringBuilder();

                    textBody.Append($"<br /><h4>Error encountered when processing below file.</h4><br />");

                    textBody.Append("<html><body>");
                    textBody.Append("<table border='1' cellpadding='5' cellspacing='0'>");
                    //textBody.Append("<tr bgcolor='#4da6ff' style='color: white;'>"); // Table header row with background color
                    //textBody.Append("<td><b>Column 1</b></td>");
                    //textBody.Append("<td><b>Column 2</b></td>");
                    //textBody.Append("</tr>");

                    // Add some data rows (can be dynamically generated from a database or data source)
                    textBody.Append($"<tr>");
                    textBody.Append($"<td>Filename</td>");
                    textBody.Append($"<td>{in_file}</td>");
                    textBody.Append($"</tr>");
                    textBody.Append($"<td>Error Message</td>");
                    textBody.Append($"<td>{errorMessage}</td>");
                    textBody.Append($"</tr>");
                    textBody.Append("</table>");
                    textBody.Append("</body></html>");

                    textBody.Append($"<br /><br /><p><i>This is an automated email. Please do not reply to this message.</i></p>");

                    message.Body = textBody.ToString();

                }
                

                message.IsBodyHtml = true;

                using var smtpClient = new SmtpClient("mail.manulife.com");
                smtpClient.Port = 25; // Common port for internal SMTP without authentication
                smtpClient.EnableSsl = false; // Set to true if your server requires SSL/TLS
                smtpClient.UseDefaultCredentials = true; // Use current Windows credentials
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                smtpClient.Send(message);
            }
            catch (SmtpException ex)
            {
                Console.WriteLine($"SMTP Error: {ex.Message}");
                Console.WriteLine($"Status Code: {ex.StatusCode}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }

            /* Success - Email sent */
            return 0;
        }
    }
}

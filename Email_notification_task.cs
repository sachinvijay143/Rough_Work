using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.Runtime;
using Exception = System.Exception;

namespace Rough_Works
{
    internal class Email_notification_task
    {
        [CommandMethod("TESTNOTIFICATION")]
        public static async Task Main()
        {
            MessageBox.Show("Starting process...");

            // Simulate your task
            await DoSomeTask();

            // After task completes, send notification
            await SendNotification("vinesh.g@magnasoft.com", "Your process has completed successfully!");

            MessageBox.Show("Process finished. Notification sent.");
        }

        // Simulate a task that takes time
        public static async Task DoSomeTask()
        {
            MessageBox.Show("Task is running...");
            await Task.Delay(3000); // simulate 3 seconds of work
            MessageBox.Show("Task completed!");
        }

        // Send notification via Apps Script Web API
        public static async Task SendNotification(string email, string message)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var json = $@"{{
                        ""email"": ""{email}"",
                        ""message"": ""{message}""
                    }}";

                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    // Replace with your Apps Script Web App URL
                    string scriptUrl = "https://script.google.com/a/macros/magnasoft.com/s/AKfycby0-BqXR812SYfoRwf2Ts5It8T17J8Rinmvsfj6KQ2I0ecI1QQxoNkbhhgCbe9Zed3q/exec";

                    var response = await client.PostAsync(scriptUrl, content);

                    var result = await response.Content.ReadAsStringAsync();

                    Console.WriteLine("Apps Script response: " + result);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error sending notification: " + ex.Message);
            }
        }

        public static async Task SendNotifications(string email, string message)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var json = $@"{{
                ""email"": ""{email}"",
                ""message"": ""{message}""
            }}";

                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    string scriptUrl = "https://script.google.com/a/macros/magnasoft.com/s/AKfycbzGIP8Y7AT5SB5WM3FowoC5VbCJYD0KlY8zpIq7sth25kYw0e2zcAjhQoeULA4mKrTa/exec";

                    var response = await client.PostAsync(scriptUrl, content);

                    var result = await response.Content.ReadAsStringAsync();

                    MessageBox.Show($"Status: {response.StatusCode}\nResponse: {result}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        public static void SendEmail()
        {
            try
            {
                var fromAddress = new MailAddress("yourgmail@gmail.com", "Vinesh");
                var toAddress = new MailAddress("vinesh.g@magnasoft.com");

                const string fromPassword = "YOUR_APP_PASSWORD"; // ⚠️ Not Gmail password
                const string subject = "Notification";
                const string body = "Process completed successfully.";

                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                    Timeout = 20000
                };

                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body
                })
                {
                    smtp.Send(message);
                }

                MessageBox.Show("Email sent successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}

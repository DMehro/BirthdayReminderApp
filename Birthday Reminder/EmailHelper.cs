using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;

namespace Birthday_Reminder
{
    class EmailHelper
    {
        public void SendEmail(string empName,string Date)
        {

            MailMessage message = new System.Net.Mail.MailMessage();
            string fromEmail = "youremail";
            string fromPW = "password";
            string toEmail = "toemail";
            message.From = new MailAddress(fromEmail);
            message.To.Add(toEmail);
            message.Subject = "UpComing Birthday";
            message.Body = string.Format("Tommorrow is {0} Birthday", empName);
            message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.EnableSsl = true;
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential(fromEmail,
                                                        fromPW);

            smtpClient.Send(message.From.ToString(), message.To.ToString(),
            message.Subject, message.Body);

            
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using OpenPop.Mime;
using OpenPop.Pop3;



namespace AprilAppConsole_9
{
    class Program
    {
        static void Main(string[] args)
        {
            Pop3Client client = new Pop3Client();
            Console.Write("Пожалуйста, укажите почтовый сервер: ");
            string emailServer = Console.ReadLine();
            Console.Write("Пожалуйста, укажите порт почтового сервера: ");
            int portEmailServer = int.Parse(Console.ReadLine());
            
            Console.Write("Пожалуйста, укажите логин для авторизации: ");
            string emailLogin = Console.ReadLine();
            Console.Write("Пожалуйста, укажите пароль для авторизации: ");
            string emailPassword = Console.ReadLine();

            Console.WriteLine();
            Console.WriteLine("Подключение к почтовому серверу...");
            client.Connect(emailServer, portEmailServer, true); //("pop.mail.ru", 995, true);
            
            if (client.Connected) Console.WriteLine("Подключение установлено\nВыполняется авторизация...");
            client.Authenticate(emailLogin, emailPassword);
            Console.WriteLine("Авторизация успешна");
            Console.WriteLine();

            int count = client.GetMessageCount();

            for (int i = count - 13; i <= count; i++)
            {
                Message message = client.GetMessage(i);
                List<MessagePart> attachments = message.FindAllAttachments();

                int attachCNT = 0;
                if (attachments.Count > 0)
                {
                    foreach (MessagePart attachment in attachments)
                    {
                        if (attachment.ContentType.MediaType == "application/zip" || attachment.ContentType.MediaType == "application/x-rar")
                        {
                            string filePath = Path.Combine($@"{Environment.CurrentDirectory}\April_email\Attachment\{message.Headers.From.DisplayName}({message.Headers.From.Address})\{DateTime.Parse(message.Headers.Date.ToString()).ToShortDateString()}");
                            if (!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
                            using (FileStream fileStream = File.Create($@"{filePath}\{attachment.FileName}"))
                            {
                                using (BinaryWriter binaryWriter = new BinaryWriter(fileStream))
                                {
                                    binaryWriter.Write(attachment.Body);
                                }
                            }
                            attachCNT++;
                        }
                    }

                    Console.WriteLine($"Дата: {message.Headers.Date}\nОт кого: {message.Headers.From.DisplayName}\nТема: {message.Headers.Subject}\nСохранено волжений: {attachCNT}");
                    Console.WriteLine();
                }
            }        

            Console.ReadKey();
        }
    }
}

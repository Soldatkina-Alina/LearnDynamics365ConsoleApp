using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using log4net.Repository.Hierarchy;
using log4net;
using log4net.Config;
using System.Linq;

namespace LearnDynamics365ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            CrmServiceClient Service;
            log4net.Config.XmlConfigurator.Configure();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ILog _Log = LogManager.GetLogger("LOGGER");

            var connectionString = $@"AuthType=OAuth;Url=https://homertrialtest.api.crm4.dynamics.com/;"+ MyOptions.login + MyOptions.password +"AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=app://58145B91-0C36-4500-8554-080854F2AC97;"; 

            Service = new CrmServiceClient(connectionString);

            if (Service == null)
                _Log.Info("Ошибка соединения");

            RetriveContact(Service, _Log);
            Console.ReadKey();
        }

        public static void RetriveContact(IOrganizationService _service, ILog _Log)
        {
            // Выбираем тех у кого есть телефон1 и email1 и нет связи со средствами связи
            string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                    <entity name='contact'>
                                    <attribute name='fullname'/>
                                    <attribute name='telephone1'/>
                                    <attribute name='contactid'/>
                                    <order attribute='fullname' descending='false'/>
                                    <filter type='and'>
                                        <filter type='or'>
                                        <condition attribute='telephone1' operator='not-null'/>
                                        <condition attribute='emailaddress2' operator='not-null'/>
                                        </filter>
                                    </filter>
                                    <link-entity name='ptest_communication' from='ptest_contactid' to='contactid' link-type='outer' alias='am' />
                                        <filter type = 'and' >
                                        <condition entityname='am' attribute='ptest_contactid' operator='null' />
                                        </filter > 
                                     </entity>
                                </fetch>";

            
            EntityCollection result = _service.RetrieveMultiple(new FetchExpression(fetchXml));

            var listContacts = result.Entities;

            _Log.Info($"Выполнен fetchXml запрос. Количество записей {listContacts.Count}");

            foreach(var contact in listContacts)
            {
                string email, phone;
                contact.TryGetAttributeValue("emailaddress1", out email);
                contact.TryGetAttributeValue("telephone1", out phone);

                _Log.Info($"Начата запись новых элементов \"Средства связи\"");
                //объевление нового элемента для сохранения
                Entity communication = new Entity("ptest_communication");
                //Создание связи с контактом
                EntityReference conctactRef = new EntityReference("contact", contact.GetAttributeValue<Guid>("contactid"));

                //Заполнение полей
                communication["ptest_contactid"] = conctactRef;
                communication["ptest_name"] = contact["fullname"];

                if (phone.Length > 0)
                {
                    communication["ptest_phone"] = phone;
                    communication["ptest_type"] = new OptionSetValue(234260000);
                    communication["ptest_main"] = true;
                    Guid communicationId = _service.Create(communication);
                    _Log.Info($"Создано новое средство связи с типом телефон. Id: {communicationId}");
                }
                if (email.Length > 0)
                {
                    communication["ptest_email"] = email;
                    communication["ptest_type"] = new OptionSetValue(234260001);
                    communication["ptest_main"] = false;
                    Guid communicationId = _service.Create(communication);
                    _Log.Info($"Создано новое средство связи с типом email. Id: {communicationId}");
                }
            }
        }
    }
}

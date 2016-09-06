using ExchnagePullController.Data;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Instrumentation;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml.Serialization;

namespace ExchnagePullController.Util
{
    class SerializeHelper
    {

        public static SerializeHelper Instanse { get; } = new SerializeHelper();

        private SerializeHelper(){}


        public string ToJson(ExchangeUser[] users)
        {
            var serializer = new JavaScriptSerializer();
            var serializedResult = serializer.Serialize(users);

            return serializedResult;
        }

        public bool WriteJsonFile(string fileName, ExchangeUser[] users)
        {
            string content = ToJson(users);
            if (!string.IsNullOrEmpty(content))
            {
                if (File.Exists(fileName))
                {
                    // TODO - what to do?
                    File.Delete(fileName);
                }

                using (StreamWriter writer = File.CreateText(fileName))
                {
                    var task = writer.WriteLineAsync(content);
                    task.Wait();

                    return true;
                }
            }

            return false;
        }

        //public string ToXml(ExchangeUser[] users)
        public string ToXml(ExchangeUser[] toSerialize)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(toSerialize.GetType());

            using (StringWriter textWriter = new StringWriter())
            {
                xmlSerializer.Serialize(textWriter, toSerialize);
                return textWriter.ToString();
            }
        }

    }
}

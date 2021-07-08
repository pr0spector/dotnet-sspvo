using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Xml;
using System.Threading;

namespace ConsoleApp2
{
    public class RepsonseSS
    {
        public string json;
        public string xml;
    }

    public class SSClient
    {
        private readonly Crypto _csp;
        private readonly string _apiHost;
        private readonly string _ogrn;
        private readonly string _kpp;
        private readonly string _path;
        private readonly int _rertryCount;
        private readonly int _msRetryDelay;

        public bool debugMode
            { get; set; }

        public SSClient(string ogrn, string kpp, string apiHost, Crypto csp, string path)
        {
            if (string.IsNullOrWhiteSpace(ogrn)) throw new ArgumentNullException(nameof(ogrn));
            if (string.IsNullOrWhiteSpace(kpp)) throw new ArgumentNullException(nameof(kpp));
            if (string.IsNullOrWhiteSpace(apiHost)) throw new ArgumentNullException(nameof(apiHost));

            _csp = csp ?? throw new ArgumentNullException(nameof(csp));
            _ogrn = ogrn;
            _kpp = kpp;
            _apiHost = apiHost;
            _path = path;
            _rertryCount = 10;
            _msRetryDelay = 5000;
            debugMode = false;
        }

        public async Task<bool> saveXmlBySnils(string snils, string[] appUids = null)
        {
            Console.WriteLine("Обработка СНИЛС:" + snils);
            uint idJwt = await getSnilsIdJwt(snils);
            //Console.WriteLine("IdJwt:" + idJwt);
            RepsonseSS resp = await getServiceQueueMsg(idJwt);
            /*
            Console.WriteLine(resp.json);
            Console.WriteLine("\n\n");
            Console.WriteLine(resp.xml);
            */
            string fname = _path + "\\" + snils;
            writeFile(fname + ".xml", resp.xml);
            writeFile(fname + ".json", resp.json);
            bool confirmed = await confirmMessage(idJwt);

            // Обработка документов
            bool parsed = await parseAbiturXml(snils, resp.xml, appUids);

            // Обработка заявлений
            if (appUids != null)
            {
                foreach(string appUid in appUids)
                {
                    uint appIdJwt = await getApplicationIdJwt(appUid);
                    RepsonseSS appResp = await getServiceQueueMsg(appIdJwt);
                    if (appResp!=null)
                    {
                        string appFileName = _path + "\\" + snils + "-application-" + appUid;
                        writeFile(appFileName + ".xml", appResp.xml);
                        writeFile(appFileName + ".json", appResp.json);
                    }
                }
            }
            return true;
        }

        private async Task<bool> parseAbiturXml(string snils, string xml, string[] appUids)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            /*
            XmlNode guid_node = xmlDoc.SelectSingleNode($"//ServiceEntrant//GUID");
            if (guid_node == null)
            {
                Console.WriteLine("tag GUID not found in xml");
                return false;
            }
            string person_guid = guid_node.InnerText;
            */

            var tags = new string[] { "Identifications", "Documents", "EpguEntrantAchievement", "AppAchievement", "Contract" };

            foreach (string tag in tags)
            {
                XmlNodeList nodes = xmlDoc.SelectNodes($"//{tag}//UIDEpgu");
                if (debugMode && nodes.Count > 0)
                {
                    Console.WriteLine($"SNILS:{snils} tag:{tag}");
                }

                foreach (XmlNode node in nodes)
                {
                    string uid = node.InnerText;
                    uint idJwt = 0;
                    if (tag == "Identifications")
                    {
                        idJwt = await getIdentificationIdJwt(snils, uid);
                    }
                    else if (tag == "Documents")
                    {
                        idJwt = await getDocumentIdJwt(snils, uid);
                    }
                    // При получении заявления из очереди epgu tag AppAchievement,  при запросе ServiceEntrant - EpguEntrantAchievement
                    else if (tag == "EpguEntrantAchievement" || tag == "AppAchievement")
                    {
                        if (appUids == null || !appUids.Any()) continue;
                        // для тестов возьмём хотя бы одно заявление
                        var appUid = appUids[0];
                        idJwt = await getAchievementIdJwt(appUid, uid);
                    }
                    else if (tag == "Contract")
                    {
                        idJwt = await getContractIdJwt(snils, uid);
                    }

                    //Console.WriteLine(tag + " idJwt:" + tagIdJwt);

                    if (idJwt > 0)
                    {
                        RepsonseSS response = await getServiceQueueMsg(idJwt);
                        string fname = $"{_path}\\{snils}-";
                        if (tag == "Documents")
                        {
                            string docType = node.ParentNode.ParentNode.Name;
                            fname += $"{docType}-{uid}";
                        }
                        else
                        {
                            fname += $"{tag}-{uid}";
                        }
                        writeFile(fname + ".xml", response.xml);
                        writeFile(fname + ".json", response.json);
                        bool confirmed = await confirmMessage(idJwt);
                        if (response.xml.Length > 0)
                        {
                            parseDocumentResponse(response.xml, fname);
                        }
                        else
                        {
                            Console.WriteLine("Пустой xml " + fname);
                        }
                    } else
                    {
                        Console.WriteLine("idJwt <= 0" + idJwt);
                    }
                }
            }
            return true;
        }

        private void parseDocumentResponse(string str_xml, string fname) {

            var xml = new XmlDocument();
            xml.LoadXml(str_xml);
            XmlNode nodeFile = xml.SelectSingleNode("//Base64File");
            if (nodeFile == null)
            {
                // Achievement
                nodeFile = xml.SelectSingleNode(".//File//Base64");
            }
            if (nodeFile != null)
            {
                var fileExt = nodeFile.ParentNode.SelectSingleNode(".//FileType");
                string extension = fileExt != null ? fileExt.InnerText : ".unknown";
                if (extension[0] != '.') // в Achievement расширение без точки
                {
                    extension = "." + extension;
                }
                byte[] bytes = Convert.FromBase64String(nodeFile.InnerText);
                writeFile(fname + extension, bytes);
            }
            return;
        }

        public async Task<uint> getApplicationIdJwt(string uidEpgu)
        {
            String header = new JObject
            {
                {"action", "get"},
                {"entityType", "serviceApplication"},
                {"ogrn", _ogrn},
                {"kpp", _kpp}
            }.ToString();
            var payload = $"<PackageData><ServiceApplication><IDApplicationChoice><UIDEpgu>{uidEpgu}</UIDEpgu></IDApplicationChoice></ServiceApplication></PackageData>";
            var response = await sendMessage("/api/token/new", header, payload);
            if (response == null)
            {
                Console.WriteLine($"No reponse for serviceApplication uidEpgu: {uidEpgu}");
                return 0;
            }
            var json = response.Content;
            uint idJwt = getIdJwtFromReponse(json);
            return idJwt;
        }

        public async Task<uint> getSnilsIdJwt(string snils)
        {
            String header = new JObject
            {
                {"action", "get"},
                {"entityType", "serviceEntrant"},
                {"ogrn", _ogrn},
                {"kpp", _kpp}
            }.ToString();
            var payload = $"<PackageData><ServiceEntrant><IDEntrantChoice><SNILS>{snils}</SNILS></IDEntrantChoice></ServiceEntrant></PackageData>";
            var response = await sendMessage("/api/token/new", header, payload);
            if (response == null)
            {
                Console.WriteLine($"No response for  serviceEntrant snils {snils}");
                return 0;
            }
            var json = response.Content;
            uint idJwt = getIdJwtFromReponse(json);
            return idJwt;
        }

        public async Task<uint> getDocumentIdJwt(string snils, string documentGuid)
        {
            return await getEntityIdJwt("document", "IDDocChoice", snils, documentGuid);
        }

        public async Task<uint> getIdentificationIdJwt(string snils, string documentGuid)
        {
            return await getEntityIdJwt("identification", "IDChoice", snils, documentGuid);
        }

        public async Task<uint> getAchievementIdJwt(string appUid, string documentGuid)
        {
            /*
             Я знаю, Вы так-таки будете смеяться, но этот пакет отличается
             ApplicationIDChoice вместо IDEntrantChoice  :)
             */
            string header = new JObject
            {
                {"action", "get"},
                {"entityType", "appAchievement"},
                {"ogrn", _ogrn},
                {"kpp", _kpp}
            }.ToString();
            var payload = $"<PackageData><AppAchievement><ApplicationIDChoice><UIDEpgu>{appUid}</UIDEpgu></ApplicationIDChoice><AchievementIDChoice><UIDEpgu>{documentGuid}</UIDEpgu></AchievementIDChoice></AppAchievement></PackageData>";
            var response = await sendMessage("/api/token/new", header, payload);
            var json = response.Content;
            return getIdJwtFromReponse(json);
        }

        public async Task<uint> getContractIdJwt(string snils, string documentGuid)
        {
            return await getEntityIdJwt("contract", "IDChoice", snils, documentGuid);
        }

        public async Task<uint> getEntityIdJwt(string entity, string choiceName, string snils, string documentGuid)
        {
            String header = new JObject
            {
                {"action", "get"},
                {"entityType", entity},
                {"ogrn", _ogrn},
                {"kpp", _kpp}
            }.ToString();
            string xmlEmtityTag = capitalizeFirstLetter(entity);
            var payload = $"<PackageData><{xmlEmtityTag}><IDEntrantChoice><SNILS>{snils}</SNILS></IDEntrantChoice><{choiceName}><UIDEpgu>{documentGuid}</UIDEpgu></{choiceName}></{xmlEmtityTag}></PackageData>";
            var response = await sendMessage("/api/token/new", header, payload);
            if (response == null)
            {
                Console.WriteLine($"No repsonse for getEntity entity {entity} snils {snils} docGuid {documentGuid}");
                return 0;
            }
            var json = response.Content;
            //Console.WriteLine($"entity {entity} Response:" + json);
            return getIdJwtFromReponse(json);
        }

        public async Task<RepsonseSS> getServiceQueueMsg(uint idJwt, bool debugMode = false)
        {
            String header = new JObject
            {
                {"action", "getMessage"},
                {"idJwt", idJwt},
                {"ogrn", _ogrn},
                {"kpp", _kpp}
            }.ToString();

            var response = await sendMessage("/api/token/service/info", header, "");
            if (response == null)
            {
                Console.WriteLine($"Cannot get msg idJwt:{idJwt}");
                return null;
            }
            var json = response.Content;
            if (debugMode)
            {
                Console.WriteLine(response.Content);
            }
            return parseReponseToken(json);
        }

        private uint getIdJwtFromReponse(String response)
        {
            Dictionary<String, String> dict = JsonConvert.DeserializeObject<Dictionary<String, String>>(response);
            if (!dict.ContainsKey("idJwt"))
            {
                Console.WriteLine("No idJwt in reponse!");
            }
            uint idJwt = uint.Parse(dict["idJwt"]);
            //Console.WriteLine("idJwt " + idJwt);
            return idJwt;
        }

        public async Task<IRestResponse> sendMessage(String url, String header, String payload)
        {

            IRestResponse response = null;
            JObject json = getSignedObject(_csp, header, payload);
            //Console.WriteLine(json);
            for (int i = 0; i < _rertryCount; i++)
            {
                var request = new RestRequest(url, Method.POST);
                request.AddParameter("application/json", json.ToString(), ParameterType.RequestBody);
                var client = new RestClient(_apiHost);

                response = await client.ExecutePostAsync(request);
                if (response.StatusCode != System.Net.HttpStatusCode.OK)
                {
                    Console.WriteLine("HTTP StatusCode:" + response.StatusCode);
                    Console.WriteLine($"Sleep for {_msRetryDelay / 1000} secs");
                    Thread.Sleep(_msRetryDelay);
                    continue;
                }
                Dictionary<String, String> result = JsonConvert.DeserializeObject<Dictionary<String, String>>(response.Content);
                if (result.ContainsKey("error"))
                {
                    Console.WriteLine("Error:" + result["error"]);
                }
                break;
            }
            return response;
        }

        public async Task<Boolean> confirmMessage(uint idJwt)
        {
            // Один и тот же пакет для confirm обоих очередей (service и epgu)

            string header = new JObject
            {
                {"action", "messageConfirm"},
                {"idJwt", idJwt},
                {"ogrn", _ogrn},
                {"kpp", _kpp}
            }.ToString();
            var response = await sendMessage("/api/token/confirm", header, "");
            if (response == null)
            {
                Console.WriteLine($"No reponse for confirm idJwt: {idJwt}");
                return false;
            }
            string json = response.Content;
            Dictionary<String, String> result = JsonConvert.DeserializeObject<Dictionary<String, String>>(json);
            if (!result.ContainsKey("result"))
            {
                Console.WriteLine("Error no Result on Confirm:" + json);
            }
            else if (result["result"] != "true")
            { // :)
                Console.WriteLine("Bad confirm result:" + json);
                return false;
            }
            return true;
        }


        public static JObject getSignedObject(Crypto cryptoService, String header, String payload)
        {
            //string joHeader = msgJwt.JHeader;
            string b64Header = ToBase64String(header);
            string b64Payload = ToBase64String(payload);
            string stringToSign = $"{b64Header}.{b64Payload}";

            try
            {
                byte[] signed = cryptoService.SignData(System.Text.Encoding.UTF8.GetBytes(stringToSign));
                string signature = Convert.ToBase64String(signed);
                return new JObject
                {
                    { "token", $"{b64Header}.{b64Payload}.{signature}" }
                };
            }
            catch (Exception e)
            {
                throw;
            }
        }

        private static RepsonseSS parseReponseToken(String response)
        {
            var result = JsonConvert.DeserializeObject<Dictionary<String, String>>(response);

            if (result.ContainsKey("responseToken"))
            {
                string responseToken = result["responseToken"];
                string[] parts = responseToken.Split('.');
                RepsonseSS resp = new RepsonseSS
                {
                    json = FromBase64String(parts[0]),
                    xml = FromBase64String(parts[1])
                };
                return resp;
            }
            else
            {
                Console.WriteLine("No jsOnToken in Reponse:" + response);
                throw new NullReferenceException("No jsOnToken in Reponse!");
            }
        }

        private static string ToBase64String(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string FromBase64String(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        private static string capitalizeFirstLetter(string text)
        {
            return text?.First().ToString().ToUpper() + text?.Substring(1);
        }

        private static void writeFile(string fname, string text)
        {
            EnsureDirectoryExists(fname);

            FileStream file = new FileStream(fname, FileMode.Create);
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            file.Write(bytes, 0, bytes.Length);
            file.Close();
        }

        private static void writeFile(string fname, byte[] bytes)
        {
            EnsureDirectoryExists(fname);

            FileStream file = new FileStream(fname, FileMode.Create);
            file.Write(bytes, 0, bytes.Length);
            file.Close();
        }

        private static void EnsureDirectoryExists(string path)
        {
            string dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        }
    }
}
using Google.Apis.Sheets.v4;
using System.IO;
using System;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using Entities;
using GoogleAPI;
using Newtonsoft.Json;
using Google.Apis.Sheets.v4.Data;

namespace ConsumindoAPIGoogle
{
    class Program
    {
        public static string SpreadsheetId = "1lxrass8OfuM9W82rcGRbFCE_lH-U6rXyF4Wcc5CEy40";
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string GoogleCredencialFileName = "credential.json";
        private const string ReadRange = "A4:H99";
        private const string valueInputOption = "USER_ENTERED";

        static void Main(string[] args)
        {
            List<IList<object>> obj = new List<IList<object>>();
            UserCredential credential;

            using (var stream = new FileStream(GoogleCredencialFileName, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets,
                                                                        Scopes, "user", CancellationToken.None,
                                                                        new FileDataStore("token.json", true)).Result;
            }
            var service = SheetsServices.GetSheetsService(credential);
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SpreadsheetId, ReadRange);
            ValueRange response = request.Execute();

            List<Student> students = new List<Student>();
            IList<IList<Object>> result = response.Values;
            if (result != null)
            {
                foreach (var row in result)
                {
                    students.Add(new Student()
                    {
                        Registration = int.Parse(row[0].ToString()),
                        Name = row[1].ToString(),
                        Absences = int.Parse(row[2].ToString()),
                        P1 = int.Parse(row[3].ToString()),
                        P2 = int.Parse(row[4].ToString()),
                        P3 = int.Parse(row[5].ToString()),
                    });
                    foreach (var student in students)
                    {
                        student.Average = (student.P1 + student.P2 + student.P3) / 3;
                        if (student.Absences > 15)
                        {
                            student.Situation = "Reprovado por Falta";
                            student.FinalP = 0;
                        }
                        else if (student.Average < 50)
                        {
                            student.Situation = "Reprovado por Nota";
                            student.FinalP = 0;
                        }
                        else if (student.Average >= 50 && student.Average < 70)
                        {
                            student.Situation = "Exame final";
                            student.FinalP = (student.Average + 70) / 2;
                        }
                        else if (student.Average >= 70)
                        {
                            student.Situation = "Aprovado";
                            student.FinalP = 0;
                        }
                    }
                }
                foreach (var studentObj in students)
                {
                    obj.Add(new List<object>() {
                            studentObj.Registration,
                            studentObj.Name,
                            studentObj.Absences,
                            studentObj.P1,
                            studentObj.P2,
                            studentObj.P3,
                            studentObj.Situation,
                            studentObj.FinalP});
                }
                List<ValueRange> updateData = new List<ValueRange>();
                var dataValueRange = new ValueRange
                {
                    Range = ReadRange,
                    Values = obj
                };
                updateData.Add(dataValueRange);
                BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest
                {
                    ValueInputOption = valueInputOption,
                    Data = updateData
                };

                var newRequest = service.Spreadsheets.Values.BatchUpdate(requestBody, SpreadsheetId);

                BatchUpdateValuesResponse newResponse = newRequest.Execute();

                JsonConvert.SerializeObject(newResponse); 
            }
        }
    }
}


using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Web.Mvc;
using HotSpotUserInfo.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace HotSpotUserInfo.Controllers
{
    public class HomeController : Controller
    {
        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private const string Hapikey = "8b03825a-4f4e-4659-9348-c5b1d89fb7e0";

        public ActionResult Index()
        {
            return RedirectToAction("A", new { modifiedOnOrAfter = DateTime.Now.AddDays(-90) });
        }

        public List<Contact> A(DateTime modifiedOnOrAfter)
        {
            var startDateInMilliseconds = modifiedOnOrAfter.ToUniversalTime().Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;

            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var response = client.GetAsync("https://api.hubapi.com/contacts/v1/lists/recently_updated/contacts/recent?hapikey=" + Hapikey).Result;
            var responseString = response.Content.ReadAsStringAsync().Result;

            var contactsList = new List<Contact>();

            var jContacts = JObject.Parse(responseString);
            var length = ((JArray)jContacts["contacts"]).Count;
            for (var i = 0; i < length; i++)
            {
                var lastModifiedDateMilliseconds = (double)jContacts.SelectToken("contacts[" + i + "].properties.lastmodifieddate.value");
                if (!(lastModifiedDateMilliseconds >= startDateInMilliseconds)) continue;
                var contact = GetAndFillContact((int)jContacts.SelectToken("contacts[" + i + "].vid"));
                contactsList.Add(contact);
            }

            DisplayInExcel(contactsList);
            return contactsList;
        }

        public Contact GetAndFillContact(int vid)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var contactResponse = client.GetAsync("https://api.hubapi.com/contacts/v1/contact/vid/"+ vid +"/profile?hapikey=" + Hapikey).Result;
            var contactString = contactResponse.Content.ReadAsStringAsync().Result;

            var contact = JsonConvert.DeserializeObject<Contact>(contactString);

            var jContact = JObject.Parse(contactString);
            contact.FirstName = jContact.SelectToken("properties.firstname.value").ToString();
            contact.LastName = jContact.SelectToken("properties.lastname.value").ToString();

            var lifeCycleStageMilleseconds = jContact.SelectToken("properties.hs_lifecyclestage_lead_date.value").ToString();
            contact.LifeCycleStage = (new DateTime(1970, 1, 1)).AddMilliseconds(double.Parse(lifeCycleStageMilleseconds)).Date.ToShortDateString();

            contact.Company  = GetAndFillCompany((int)jContact.SelectToken("properties.associatedcompanyid.value"));

            return contact;
        }

        public Company GetAndFillCompany(int companyId)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var companyResponse = client.GetAsync("https://api.hubapi.com/companies/v2/companies/" + companyId + "?hapikey=" + Hapikey).Result;
            var companyString = companyResponse.Content.ReadAsStringAsync().Result;

            var company = JsonConvert.DeserializeObject<Company>(companyString);

            var jContact = JObject.Parse(companyString);
            company.Company_Name = jContact.SelectToken("properties.name.value").ToString();
            company.Company_WebSite = jContact.SelectToken("properties.website.value").ToString();
            company.Company_ZipCode = (int)jContact.SelectToken("properties.zip.value");

            var jCitySate = GetCityAndStateFromZip(company.Company_ZipCode);
            dynamic data = JObject.Parse(jCitySate);
            company.Company_City = data.city;
            company.Company_State = data.state;

            company.Company_Phone = jContact.SelectToken("properties.phone.value").ToString();

            return company;
        }

        public string GetCityAndStateFromZip(int zip)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var cityAndState = client.GetAsync("http://ZiptasticAPI.com/" + zip).Result;
            var cityAndStateString = cityAndState.Content.ReadAsStringAsync().Result;

            return cityAndStateString;
        }

        public void DisplayInExcel(List<Contact> contacts)
        {
            var excelApp = new Excel.Application {Visible = true};
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            var type = typeof(Contact);
            var properties = type.GetProperties();
            for (var j =0; j < properties.Length; j++)
            {
                var curLetter = j;
                if (properties[j].Name == "Company")
                {
                    Type subType = typeof(Company);
                    PropertyInfo[] subProperties = subType.GetProperties();
                    var k = curLetter;
                    foreach (var pInfo in subProperties)
                    {
                        if (pInfo.Name == "Id") continue;
                        workSheet.Cells[1, Alphabet[k].ToString()] = pInfo.Name;
                        var rng = (Excel.Range)workSheet.Cells[1, Alphabet[k].ToString()];
                        rng.Font.Bold = true;
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        rng.Borders.Weight = 3d;
                        k++;
                    }
                }
                else
                {
                    workSheet.Cells[1, Alphabet[j].ToString()] = properties[j].Name;
                    var rng = (Excel.Range)workSheet.Cells[1, Alphabet[j].ToString()];
                    rng.Font.Bold = true;
                    rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng.Borders.Weight = 3d;
                }
            }
            var row = 2;
            var column = 0;
            foreach (var contact in contacts)
            {
                workSheet.Cells[row, column + 1].Value = contact.Id;
                workSheet.Cells[row, column + 2].Value = contact.FirstName;
                workSheet.Cells[row, column + 3].Value = contact.LastName;
                workSheet.Cells[row, column + 4].Value = contact.LifeCycleStage;
                workSheet.Cells[row, column + 5].Value = contact.Company.Company_Name;
                workSheet.Cells[row, column + 6].Value = contact.Company.Company_WebSite;
                workSheet.Cells[row, column + 7].Value = contact.Company.Company_City;
                workSheet.Cells[row, column + 8].Value = contact.Company.Company_State;
                workSheet.Cells[row, column + 9].Value = contact.Company.Company_ZipCode;
                workSheet.Cells[row, column + 10].Value = contact.Company.Company_Phone;

                workSheet.Columns.AutoFit();
                column = 0;
                row++;
            }
        }
    }
}
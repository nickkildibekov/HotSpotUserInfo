using Newtonsoft.Json;

namespace HotSpotUserInfo.Models
{
    public class Company
    {
        [JsonProperty(PropertyName = "companyId")]
        public int Id { get; set; }
        [JsonProperty(PropertyName = "name")]
        public string Company_Name { get; set; }
        [JsonProperty(PropertyName = "website")]
        public string Company_WebSite { get; set; }
        [JsonProperty(PropertyName = "city")]
        public string Company_City { get; set; }
        [JsonProperty(PropertyName = "state")]
        public string Company_State { get; set; }
        [JsonProperty(PropertyName = "zip")]
        public int Company_ZipCode { get; set; }
        [JsonProperty(PropertyName = "phone")]
        public string Company_Phone { get; set; }
    }
}
using Newtonsoft.Json;

namespace HotSpotUserInfo.Models
{
    public class Contact
    {
        [JsonProperty(PropertyName = "vid")]
        public int Id { get; set; }
        [JsonProperty(PropertyName = "firstname")]
        public string FirstName { get; set; }
        [JsonProperty(PropertyName = "lastname")]
        public string LastName { get; set; }
        [JsonProperty(PropertyName = "hs_lifecyclestage_lead_date")]
        public string LifeCycleStage { get; set; }
        [JsonProperty(PropertyName = "associated_company")]
        public Company Company { get; set; }
        public Contact()
        {
            Company = new Company();
        }
    }
}
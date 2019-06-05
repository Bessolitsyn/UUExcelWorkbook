using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace UUExcelWorkbook
{
    [DataContract]
    class FlowMeter
    {
        [DataMember]
        public int Id;
        [DataMember]
        public string Code;
        [DataMember]
        public string Type;
        [DataMember]
        public string Reducer;
        [DataMember]
        public string DIA1;
        [DataMember]
        public string DIA0;
        [DataMember]
        public string CostPrice;
        [DataMember]
        public string Price;
        [DataMember]
        public string ConnectionType;
        [DataMember]
        public string FlowRateMin;
        [DataMember]
        public string FlowRateMax;
        [DataMember]
        public string Name;

    }

    [DataContract]
    class FlowMeterProperty
    {
        [DataMember]
        int Id;
        [DataMember]
        public int FlowMeterId;
        [DataMember]
        public string Column;
        [DataMember]
        public string Name;
        [DataMember]
        public string Value;      

    }

    [DataContract]
    class FullFLowMeter
    {
        [DataMember]
        public FlowMeter FlowMeter;
        [DataMember]
        public List<FlowMeterProperty> FlowMeterProperties;
    }
}

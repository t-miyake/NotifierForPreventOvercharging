namespace NotifierForPreventOvercharging
{
    public class BatteryInformation
    {
        public uint CurrentCapacity { get; set; }
        public int DesignedMaxCapacity { get; set; }
        public int FullChargeCapacity { get; set; }
        public uint Voltage { get; set; }
        public int DischargeRate { get; set; }
    }
}
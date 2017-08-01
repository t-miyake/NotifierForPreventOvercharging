using System.Windows;

namespace NotifierForPreventOvercharging
{
    /// <summary>
    /// BatteryInfomationWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class BatteryInformationWindow : Window
    {
        public BatteryInformationWindow()
        {
            InitializeComponent();
            UpdateInfomation();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            UpdateInfomation();
        }

        private void UpdateInfomation()
        {
            var info = BatteryInfo.GetBatteryInformation();
            DesignCapacity.Text = info.DesignedMaxCapacity.ToString("#,0") + " mWh";
            FullChargeCapacity.Text = info.FullChargeCapacity.ToString("#,0") + " mWh";
            CurrentCapacity.Text = info.CurrentCapacity.ToString("#,0") + " mWh";
        }
    }
}
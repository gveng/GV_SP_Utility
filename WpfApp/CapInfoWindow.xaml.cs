using System.Collections.Generic;
using System.Windows;

namespace WpfApp
{
    public partial class CapInfoWindow : Window
    {
        public CapInfoWindow()
        {
            InitializeComponent();
        }

        public void UpdateCapacitors(List<CapacitorAssignment> capacitors)
        {
            CapList.Items.Clear();
            if (capacitors.Count == 0)
            {
                CapList.Items.Add("No capacitors placed yet.");
                return;
            }

            for (int i = 0; i < capacitors.Count; i++)
            {
                var cap = capacitors[i];
                CapList.Items.Add($"C{i + 1}: {cap.FileName}");
                CapList.Items.Add($"     {cap.Coordinates}");
                CapList.Items.Add($"     {cap.RlcInfo}");
            }
        }
    }
}

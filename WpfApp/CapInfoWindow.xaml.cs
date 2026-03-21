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

            // Group by domain (empty domain → "Unassigned")
            var groups = new Dictionary<string, List<(int Index, CapacitorAssignment Cap)>>(System.StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < capacitors.Count; i++)
            {
                string dom = string.IsNullOrWhiteSpace(capacitors[i].DomainName) ? "Unassigned" : capacitors[i].DomainName;
                if (!groups.ContainsKey(dom))
                    groups[dom] = new List<(int, CapacitorAssignment)>();
                groups[dom].Add((i, capacitors[i]));
            }

            foreach (var kv in groups)
            {
                CapList.Items.Add($"─── {kv.Key} ({kv.Value.Count} caps) ───");
                foreach (var (idx, cap) in kv.Value)
                {
                    CapList.Items.Add($"  C{idx + 1}: {cap.FileName}");
                    CapList.Items.Add($"       {cap.Coordinates}");
                    if (!string.IsNullOrEmpty(cap.RlcInfo))
                        CapList.Items.Add($"       {cap.RlcInfo}");
                }
            }
        }
    }
}

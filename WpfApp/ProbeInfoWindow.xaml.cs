using System.Collections.Generic;
using System.Windows;

namespace WpfApp
{
    public partial class ProbeInfoWindow : Window
    {
        public ProbeInfoWindow()
        {
            InitializeComponent();
        }

        public void UpdateProbes(List<Point> probePoints, List<string> probeDomains)
        {
            ProbeList.Items.Clear();
            if (probePoints.Count == 0)
            {
                ProbeList.Items.Add("No probes placed yet.");
                return;
            }

            // Group by domain
            var groups = new Dictionary<string, List<int>>(System.StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < probePoints.Count; i++)
            {
                string dom = i < probeDomains.Count ? probeDomains[i] : "Default";
                if (!groups.ContainsKey(dom)) groups[dom] = new List<int>();
                groups[dom].Add(i);
            }

            foreach (var kv in groups)
            {
                ProbeList.Items.Add($"─── {kv.Key} ({kv.Value.Count} probes) ───");
                foreach (int i in kv.Value)
                    ProbeList.Items.Add($"  P{i + 1}:  X={probePoints[i].X:F2}  Y={probePoints[i].Y:F2}");
            }
        }
    }
}

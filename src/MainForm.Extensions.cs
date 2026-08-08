using System.Windows.Forms;

namespace SnakeTail
{
    // Extension hook to launch Monitor Rules UI from MainForm
    public partial class MainForm
    {
        public void LaunchMonitorRulesForm()
        {
            var f = new MonitorRulesForm();
            f.Show(this);
        }
    }
}

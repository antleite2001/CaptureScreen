using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ni
{
    public partial class frmSecondsBetweenCapture : Form
    {
        public frmSecondsBetweenCapture(UInt32 SecondsBetweenCapture)
        {
            InitializeComponent();
            tbSecondsBetweenCapture.Text = SecondsBetweenCapture.ToString();
        }

        

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Visible = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Visible = false;
        }

        private void tbSecondsBetweenCapture_TextChanged(object sender, EventArgs e)
        {
           
            if (uint.TryParse(tbSecondsBetweenCapture.Text, out uint result))
            {
                if ((result >= 1) && (result <= 86400))
                {
                    lblMinutesHours.Text =  (Convert.ToDouble( result) / 60.0F).ToString("0.00") + " minutes";
                    btnOk.Enabled = true;
                    return;
                }
            }
            lblMinutesHours.Text = "";
            btnOk.Enabled = false;
        }
    }
}

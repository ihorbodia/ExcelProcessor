using System.Windows.Forms;

namespace ExcelProcessor
{
	public partial class ExcelProcessorMainGui : Form
	{
		public ExcelProcessorMainGui()
		{
			InitializeComponent();
			StatusLabelText.Text = "Choose folder";
		}
	}
}

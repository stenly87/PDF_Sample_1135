using Spire.Doc;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp12
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string Text1 { get; set; }
        public string Text2{ get; set; }
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            try
            {
                Random random = new Random();
                DateTime date = DateTime.Now;
                int orderNumber = 1199;
                double sum = 5000;
                double скидка = 0.1;
                string пунктВыдачи = "Бородинская, 16, каб 4";
                int code = random.Next(100, 1000);

                List<string> content = new List<string>();
                content.Add("Кофе со сливками");
                content.Add("Маринованные помидоры");
                content.Add("Оливье");
                content.Add("Мандаринка");

                Document document = new Document();
                var section = document.AddSection();
                section.PageSetup.PageSize = new System.Drawing.SizeF(300, 300);

                var p = document.Sections[0].AddParagraph();
                p.AppendText($"Дата заказа: {date.ToShortDateString()}");

                p = document.Sections[0].AddParagraph();
                p.AppendText($"Номер заказа: {orderNumber}");
                document.Sections[0].AddParagraph();

                p = document.Sections[0].AddParagraph();
                var range = p.AppendText($"{code}");
                range.CharacterFormat.FontSize = 30;
                range.CharacterFormat.Bold = true;
                document.Sections[0].AddParagraph();

                p = document.Sections[0].AddParagraph();
                p.AppendText($"Состав заказа:");
                foreach (var item in content)
                {
                    p = document.Sections[0].AddParagraph();
                    p.AppendText(item);
                }
                document.Sections[0].AddParagraph();

                p = document.Sections[0].AddParagraph();
                p.AppendText($"Сумма заказа: {sum}р. Размер скидки {скидка}р.");

                p = document.Sections[0].AddParagraph();
                p.AppendText($"Ваш заказ лежит тут: {пунктВыдачи}");

                document.SaveToFile("newFile.pdf", FileFormat.PDF);
                var info = new ProcessStartInfo("explorer.exe");
                info.Arguments = Environment.CurrentDirectory + "\\newFile.pdf";
                Process.Start(info);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
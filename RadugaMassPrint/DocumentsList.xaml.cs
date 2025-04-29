using RadugaMassPrint.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace RadugaMassPrint
{
    /// <summary>
    /// Логика взаимодействия для DocumentsList.xaml
    /// </summary>
    public partial class DocumentsList : Window
    {
        internal bool IsConfirm { get; private set; }
        public DocumentsList(IEnumerable<DocumentData> documentDatas)
        {
            InitializeComponent();

            foreach (var documentData in documentDatas.Where(dd => dd.DocumentName.ToLower().Contains("акт")))
            {
                var invoiceDocument = documentDatas.FirstOrDefault(dd => dd.DocID is 72 or 80 && dd.AgreementNumber == documentData.AgreementNumber);
                documentData.DifferentSum = invoiceDocument != null && invoiceDocument?.Sum != documentData.Sum;
            }


            DocumentsListDataGrid.ItemsSource = documentDatas;
        }


        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            IsConfirm = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsConfirm = false;
            this.Close();
        }
    }
}

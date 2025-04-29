using MassPrint.Services;
using Microsoft.Extensions.Configuration;
using RadugaMassPrint.Models;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using WPF = System.Windows.Controls;

namespace RadugaMassPrint
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly IConfigurationRoot _configure = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
        private static readonly Dictionary<string, string> operators = _configure.GetSection("Operators")
                                                                   .GetChildren()
                                                                   .ToDictionary(section => section.Key, section => section.Value);
        public MainWindow()
        {
            InitializeComponent();
            OperatorsComboBox.ItemsSource = operators.OrderBy(op => int.Parse(op.Key));
            OperatorsComboBox.SelectedIndex = 0;

            var documentTypes = _configure
                                        .GetSection("DocumentTypes")
                                        .GetChildren()
                                        .ToDictionary(section => section.Key, section => section.Value)
                                        .OrderBy(dt => int.Parse(dt.Key));

            foreach (var documentType in documentTypes)
            {
                documentTypesUniformGrid.Children.Add( new WPF.CheckBox()
                {
                    Content = documentType.Value,
                    Tag = documentType.Key,
                    Margin = new Thickness(5)
                });
            }
        }

        private void MonthPicker_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is WPF.DatePicker dp)
            {
                DateTime today = DateTime.Now;
                DateTime minDate = new DateTime(today.Year - 2, 1, 1);
                DateTime maxDate = new DateTime(today.Year, today.Month, 1);

                dp.SelectedDate = new DateTime(today.Year, today.Month, 1);
                dp.DisplayDateStart = minDate;
                dp.DisplayDateEnd = maxDate;

                for (var date = minDate; date < maxDate; date = date.AddDays(1))
                {
                    if (date.Day == 1)
                    {
                        continue;
                    }
                    dp.BlackoutDates.Add(new WPF.CalendarDateRange(date));
                }
            }
        }

        private async void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is WPF.Button bt)
                {
                    bt.IsEnabled = false;
                }

                var docIds = documentTypesUniformGrid.Children.OfType<WPF.CheckBox>().Where(cb => cb.IsChecked == true).Select(cb => int.Parse(cb.Tag.ToString()!)).ToList();
                var dateFrom = MonthPicker.SelectedDate;
                IEnumerable<int> operators = OperatorsComboBox.SelectedValue.ToString() == "0" ? MainWindow.operators.Where(o => o.Key != "0").Select(o => int.Parse(o.Key)) : new List<int>() { int.Parse(OperatorsComboBox.SelectedValue.ToString()) };
                string? accType = accTypeContainer.Children.OfType<WPF.RadioButton>().FirstOrDefault(rb => rb.IsChecked == true)?.Tag.ToString();

                if (docIds.Count() == 0)
                {
                    System.Windows.MessageBox.Show("Необходимо выбрать типы документов", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (accType == null)
                {
                    System.Windows.MessageBox.Show("Необходимо выбрать тип пользователя", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (string.IsNullOrEmpty(FoldersName.Text))
                {
                    System.Windows.MessageBox.Show("Необходимо выбрать папку для формирования документов", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (var mySqlManager = new MySqlManager(_configure.GetConnectionString("Billing")!))
                {
                    IEnumerable<DocumentData> documentDatas = await mySqlManager.GetFilesFolder(docIds, dateFrom!.Value, operators, int.Parse(accType));

                    if (int.Parse(accType) == 2)
                    {
                        documentDatas = documentDatas.OrderBy(dd =>
                        {
                            var addressParts = (dd.Address ?? "").Split(',');
                            return string.Join(",", addressParts.Where((part, index) => index != 4));
                        })
                                                      .ThenBy(dd =>
                                                      {
                                                          var match = Regex.Match(dd.BuildingName ?? "", @"\d+", RegexOptions.IgnoreCase);
                                                          return match.Success ? int.Parse(match.Value) : 0;
                                                      })
                                                     .ThenBy(dd => dd.AccountName)
                                                     .ThenBy(dd => dd.AgrmID)
                                                     .ToList();
                    }
                    else
                    {
                        documentDatas = documentDatas
                            .OrderBy(dd =>
                            {
                                var match = Regex.Match(dd.AgreementNumber ?? "", @"\d+", RegexOptions.IgnoreCase);
                                return match.Success ? int.Parse(match.Value) : 0;
                            })
                            .ThenBy(dd => dd.AccountName)
                            .ThenBy(dd => dd.AgrmID)
                            .ThenBy(dd =>
                            {
                                var math = Regex.Match(dd.BuildingName ?? "", @"\d+", RegexOptions.IgnoreCase);
                                return math.Success ? int.Parse(math.Value) : 0;
                            }).ToList();
                    }

                    DocumentsList documentsListView = new DocumentsList(documentDatas);
                    documentsListView.ShowDialog();

                    if (!documentsListView.IsConfirm)
                    {
                        return;
                    }

                    DocumentsProgressBar.Value = 0;
                    DocumentsProgressBar.Maximum = documentDatas.Count();

                    var progress = new Progress<(int level,int count)>( data =>
                    {
                        ProgressTextBlock.Text = $"Этап {data.level} - {(data.level == 1 ? $"Скачено файлов {data.count} из {documentDatas.Count()}" : $"Сформировано/Печать {data.count} из {documentDatas.Count()}")}";
                        DocumentsProgressBar.Value = data.count;
                    });

                    await Task.Run(async () =>
                    {
                        using (var sftp = new SftpClient(_configure["SSH:host"]!, _configure["SSH:username"]!, _configure["SSH:password"]!))
                        {
                            sftp.Connect();

                            int completed = 0;


                            foreach (var documentData in documentDatas)
                            {
                                string filePath = Regex.Match(documentData.FileName, @"[^/\\]+$", RegexOptions.None).Value;
                                using (var fileStream = new FileStream(filePath, FileMode.Create))
                                {
                                    sftp.DownloadFile(documentData.FileName, fileStream);
                                }

                                (progress as IProgress<(int, int)>)?.Report((1, ++completed));

                            }


                            if (docIds.Count == 1 && docIds[0] == 85)
                            {
                                string folderName = "";
                                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                {
                                    folderName = FoldersName.Text;
                                });
                                WordService.JoinDocuments(documentDatas.Select(dd => Regex.Match(dd.FileName, @"[^/\\]+$", RegexOptions.None).Value),folderName, progress);
                            }
                            else
                            {
                                await WordService.PrintDocument(documentDatas.Select(dd => Regex.Match(dd.FileName, @"[^/\\]+$", RegexOptions.None).Value), progress);
                            }
                        }
                    });

                }
            }
            catch(Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (sender is WPF.Button bt)
                {
                    bt.IsEnabled = true;
                }
            }
        }

        private void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using var saveFileDialog = new OpenFileDialog()
            {
                Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx",
                FileName = "NewDocument.doc"
            };

            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (!File.Exists(saveFileDialog.FileName))
                {
                    File.Create(saveFileDialog.FileName);
                }

                FoldersName.Text = saveFileDialog.FileName;
            }
        }
    }
}

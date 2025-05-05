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
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
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

        private CancellationTokenSource CancellationTokenSource;
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
                var cb = new WPF.CheckBox()
                {

                    Content = documentType.Value,
                    Tag = documentType.Key,
                    Margin = new Thickness(5)
                };

                cb.Checked += CheckBox_Checked;
                cb.Unchecked += CheckBox_Unchecked;

                documentTypesUniformGrid.Children.Add(cb);
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs args)
        {
            if (sender is WPF.CheckBox cb && cb != null)
            {
                if (int.Parse(cb.Tag.ToString()) == 63)
                {
                    var kTVCb = documentTypesUniformGrid.Children.OfType<WPF.CheckBox>().First(cb => int.Parse(cb.Tag.ToString()) == 71);
                    kTVCb.IsEnabled = false;
                    kTVCb.IsChecked = false;
                }
                else if (int.Parse(cb.Tag.ToString()) == 71)
                {
                    var domofonCB = documentTypesUniformGrid.Children.OfType<WPF.CheckBox>().First(cb => int.Parse(cb.Tag.ToString()) == 63);
                    domofonCB.IsChecked = false;
                    domofonCB.IsEnabled = false;
                }
            }
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs args)
        {
            if (sender is WPF.CheckBox cb && cb != null)
            {
                if (int.Parse(cb.Tag.ToString()) == 63)
                {
                    var kTVCb = documentTypesUniformGrid.Children.OfType<WPF.CheckBox>().First(cb => int.Parse(cb.Tag.ToString()) == 71);
                    kTVCb.IsEnabled = true;
                }
                else if (int.Parse(cb.Tag.ToString()) == 71)
                {
                    var domofonCB = documentTypesUniformGrid.Children.OfType<WPF.CheckBox>().First(cb => int.Parse(cb.Tag.ToString()) == 63);
                    domofonCB.IsEnabled = true;
                }
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
            CancellationTokenSource = new CancellationTokenSource();
            try
            {
                if (sender is WPF.Button bt)
                {
                    bt.IsEnabled = false;
                    CancelButton.IsEnabled = true;
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

                //if (string.IsNullOrEmpty(FoldersName.Text))
                //{
                //    System.Windows.MessageBox.Show("Необходимо выбрать папку для формирования документов", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                //    return;
                //}

                using (var mySqlManager = new MySqlManager(_configure.GetConnectionString("Billing")!))
                {
                    IEnumerable<DocumentData> documentDatas = await mySqlManager.GetFilesFolder(docIds, dateFrom!.Value, operators, int.Parse(accType));

                    var packDocs = documentDatas.Where(dd => dd.AgreementNumber.Contains('П')).Select(dd => dd.AgrmID).ToList();
                    documentDatas = documentDatas.Where(dd => !packDocs.Contains(dd.AgrmID)).ToList();

                    if (docIds.Contains(63))
                    {
                        var docs = documentDatas.Where(dd => dd.DocID == 63 && !dd.AgreementNumber.Contains('Д')).Select(dd => dd.AgrmID).ToList();
                        documentDatas = documentDatas.Where(dd => !docs.Contains(dd.AgrmID)).ToList();
                    }

                    if (docIds.Contains(71))
                    {
                        var docs = documentDatas.Where(dd => dd.DocID == 71 && !dd.AgreementNumber.Contains("ТВ")).Select(dd => dd.AgrmID).ToList();
                        documentDatas = documentDatas.Where(dd => !docs.Contains(dd.AgrmID)).ToList();
                    }

                    if (int.Parse(accType) == 2)
                    {
                        documentDatas = documentDatas.OrderBy(dd =>
                        {
                            var addressParts = (dd.Address ?? "").Split(',');

                            return string.Join(",", addressParts.Take(6));
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

                    var progress = new Progress<(int level, int count)>(data =>
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

                            try
                            {

                                foreach (var documentData in documentDatas)
                                {
                                    string filePath = Regex.Match(documentData.FileName, @"[^/\\]+$", RegexOptions.None).Value;
                                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                                    {
                                        // Создаем задачу для скачивания файла
                                        var downloadTask = Task.Run(() =>
                                        {
                                            // Проверка на отмену перед скачиванием
                                            CancellationTokenSource.Token.ThrowIfCancellationRequested();
                                            sftp.DownloadFile(documentData.FileName, fileStream);
                                        }, CancellationTokenSource.Token);

                                        try
                                        {
                                            // Ожидаем завершения задачи
                                            await downloadTask;
                                        }
                                        catch (OperationCanceledException)
                                        {
                                            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                            {
                                                ProgressTextBlock.Text = "Отмена пользователем";
                                            });
                                            return;
                                        }
                                    }

                                    (progress as IProgress<(int, int)>)?.Report((1, ++completed));

                                }


                                var downLoadTask = Task.Run(async () =>
                                {
                                    if (docIds.Count == 1 && (docIds[0] == 85 || docIds[0] == 63 || docIds[0] == 71))
                                    {

                                        string folderName = "";
                                        //await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                        //{
                                        //    folderName = FoldersName.Text;
                                        //});
                                        await WordService.JoinDocumentsAndPrint(documentDatas.Select(dd => Regex.Match(dd.FileName, @"[^/\\]+$", RegexOptions.None).Value).ToList(), folderName, progress, CancellationTokenSource.Token);
                                    }
                                    else
                                    {
                                        await WordService.JoinDocumentsWithPageBreakAndPrint(documentDatas.Select(dd => Regex.Match(dd.FileName, @"[^/\\]+$", RegexOptions.None).Value).ToList(), "", progress, CancellationTokenSource.Token);
                                    }
                                }, CancellationTokenSource.Token);

                                try
                                {
                                    // Ожидаем завершения задачи
                                    await downLoadTask; // КТВ, домофон 
                                }
                                catch (OperationCanceledException)
                                {
                                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                    {
                                        ProgressTextBlock.Text = "Отмена пользователем";
                                    });
                                    return;
                                }
                            }
                            catch (OperationCanceledException)
                            {
                                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                {
                                    ProgressTextBlock.Text = "Отмена пользователем";
                                });
                            }

                            catch (Exception e)
                            {
                                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                {
                                    System.Windows.MessageBox.Show(e.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                    ProgressTextBlock.Text = $"Ошибка - {e.Message}";
                                });

                            }
                            finally
                            {
                                foreach (var documentData in documentDatas)
                                {
                                    string filePath = Regex.Match(documentData.FileName, @"[^/\\]+$", RegexOptions.None).Value;
                                    if (File.Exists(System.IO.Path.Combine(Environment.CurrentDirectory, filePath)))
                                    {
                                        File.Delete(System.IO.Path.Combine(Environment.CurrentDirectory, filePath));
                                    }
                                }
                            }
                        }
                    }, CancellationTokenSource.Token);

                    ProgressTextBlock.Text = "Готово";
                    DocumentsProgressBar.Value = 0;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (sender is WPF.Button bt)
                {
                    bt.IsEnabled = true;
                    CancelButton.IsEnabled = false;
                }
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CancellationTokenSource?.Cancel();
            if (sender is WPF.Button bt)
            {
                bt.IsEnabled = false;
                printButton.IsEnabled = true;
            }
        }

        //private void SelectFolder_Click(object sender, RoutedEventArgs e)
        //{
        //    using var saveFileDialog = new OpenFileDialog()
        //    {
        //        Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx",
        //        FileName = "NewDocument.doc"
        //    };

        //    if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //    {
        //        if (!File.Exists(saveFileDialog.FileName))
        //        {
        //            File.Create(saveFileDialog.FileName);
        //        }

        //        FoldersName.Text = saveFileDialog.FileName;
        //    }
        //}
    }
}

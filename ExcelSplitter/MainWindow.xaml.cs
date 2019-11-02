using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelSplitter.Model;
using win32 = Microsoft.Win32;
using System.Windows.Forms;
using ExcelDataReader;

namespace ExcelSplitter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            _excelSpreadSheets = new ObservableCollection<TableComboBox>();
            _headerRowList = new ObservableCollection<TableComboBox>();
            _blankRowList = new ObservableCollection<TableComboBox>();
        }

        private DataTableCollection _tables;
        private DataTable _table;
        private int headerRowLocation;
        private int blankRowLocation;
        private ObservableCollection<TableComboBox> _excelSpreadSheets;
        private ObservableCollection<TableComboBox> _headerRowList;
        private ObservableCollection<TableComboBox> _blankRowList;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new win32.OpenFileDialog();
            var dialogResult = fileDialog.ShowDialog();
            if (!dialogResult.HasValue) return;
            var fileName = fileDialog.FileName;
            FileNameTxtBox.Text = fileName;
            //Clear all textboxes and combo boxes

        }

        private void FileNameTxtBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var txtBox = FileNameTxtBox.Text;
            if (String.IsNullOrWhiteSpace(txtBox)) return;
            var fileName = Path.GetFileName(txtBox);
            if (ValidateFileExtension(fileName))
            {
                ExcelColumnsAndRows(FileNameTxtBox.Text);    
            }
            

        }

        private bool ValidateFileExtension(string fileName)
        {
            if (!Path.HasExtension(fileName))
            {
                MessageBoxDialog();
                return false;
            }
            if (!Path.HasExtension(fileName)) return false;
            
            var fileExtension = Path.GetExtension(fileName);
            
            if (fileExtension == null) return false;
            
            if (fileExtension.Equals(".xlsx")) return true;
            
            MessageBoxDialog();
            
            return false;
        }

        private void MessageBoxDialog()
        {
            System.Windows.MessageBox.Show("Please choose an Excel file.", "File Error", MessageBoxButton.OK, MessageBoxImage.Question);
            FileNameTxtBox.Text = "";
        }

        private void MessageBoxDialog(string messageboxtext, string captiontext, System.Windows.Controls.TextBox[] textboxes)
        {
            System.Windows.MessageBox.Show(messageboxtext, captiontext, MessageBoxButton.OK, MessageBoxImage.Question);
            foreach(var textbox in textboxes)
            {
                textbox.Text = "";
            }
            
        }
        private void MessageBoxDialog(string messageboxtext, string captiontext)
        {
            System.Windows.MessageBox.Show(messageboxtext, captiontext, MessageBoxButton.OK, MessageBoxImage.Question);
        }

        private void ExcelColumnsAndRows(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //4. DataSet - Create column names from first row
            //excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();
            _tables = result.Tables;
            for (var i = 0; i < _tables.Count; i++)
            {
                var tableComboBox = new TableComboBox
                {
                    TableName = _tables[i].TableName,
                    TablePosition = i
                };
                _excelSpreadSheets.Add(tableComboBox);
            }
            TableSelectComboBox.ItemsSource = _excelSpreadSheets;
            TableSelectComboBox.DisplayMemberPath = "TableName";
            TableSelectComboBox.SelectedValuePath = "TablePosition";
            excelReader.Close();
        }
        //TableSelectComboBoxHeaderRow
        private void TableSelectComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var items = e.AddedItems;
            if(items != null && items.Count != 0)
            {
                var item = (TableComboBox)items[0];
                _table = _tables[item.TablePosition];
                var rowCount = _table.Rows.Count;
                var rows = _table.Rows;
                var columnCount = rows[0].ItemArray.Count();
                RowsTxtBox.Text = rowCount.ToString(CultureInfo.InvariantCulture);
                ColumnsTxtBox.Text = columnCount.ToString(CultureInfo.InvariantCulture);
                int rowCountDropDown = rowCount;
                for (var i = 1; i <= rowCount; i++)
                {
                    var tableComboBox = new TableComboBox
                    {
                        TableName = i.ToString(),
                        TablePosition = i
                    };
                    _blankRowList.Add(tableComboBox);
                    _headerRowList.Add(tableComboBox);
                }
                TableSelectComboBoxHeaderRow.ItemsSource = _headerRowList;
                TableSelectComboBoxHeaderRow.DisplayMemberPath = "TableName";
                TableSelectComboBoxHeaderRow.SelectedValuePath = "TablePosition";
                BlankRowSelectionComboBox.ItemsSource = _blankRowList;
                BlankRowSelectionComboBox.DisplayMemberPath = "TableName";
                BlankRowSelectionComboBox.SelectedValuePath = "TablePosition";
            }


            //var tableName = tables[0].TableName;
        }

        private void TableSelectComboBoxHeaderRow_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = sender as System.Windows.Controls.ComboBox;
            if(comboBox.SelectedIndex != -1)
            {
                var stringvalue = comboBox.SelectedValue.ToString();
                var value = int.Parse(stringvalue);
                //string selectedItem = (sender as System.Windows.Controls.ComboBox).SelectedItem as string;
                headerRowLocation = value;
            }
            

        }

        private void BlankRowSelectionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = sender as System.Windows.Controls.ComboBox;
            if(comboBox.SelectedIndex != -1)
            {
                var stringvalue = comboBox.SelectedValue.ToString();
                var value = int.Parse(stringvalue);
                //string selectedItem = (sender as System.Windows.Controls.ComboBox).SelectedItem as string;
                blankRowLocation = value;
            }
            
        }



        private void Button_Click_1(object sender, RoutedEventArgs e)
        {   
            if(string.IsNullOrEmpty(OutDirTxtBox.Text) || string.IsNullOrEmpty(OutFileNameTxtBox.Text))
            {
                MessageBoxDialog("Please fill out all text boxes and drop down lists", "Error", new[] { OutDirTxtBox, OutFileNameTxtBox });
                return;
            }
            bool anyBlankRows;

            if (BlankRowSelectionComboBox.SelectedIndex == -1)
            {
                anyBlankRows = false;
            }
            else
            {
                anyBlankRows = true;
            }
            int headerRowNum = headerRowLocation - 1;
            int blankRowNum = blankRowLocation - 1;
            DataRowCollection rows = _table.Rows;
            DataColumnCollection columns = _table.Columns;
            DataRow headerRow = rows[headerRowNum];
            object[] itemArray = headerRow.ItemArray;
            
            string output = OutDirTxtBox.Text + "\\" + OutFileNameTxtBox.Text + ".xlsx";
            int nupFactor = int.Parse(DivideFactor.Text);
            
            

            //Rename columns
            for (var i = 0; i < columns.Count; i ++ )
            {
                var caption = columns[i].Caption = string.Format("[1]{0}", itemArray[i]);
                var name = columns[i].ColumnName = string.Format("[1]{0}", itemArray[i]);
            }
            //Add New Columns
            var j = 2;
            while(j <= nupFactor)
            {
                foreach(var item in itemArray)
                {
                    columns.Add(new DataColumn(string.Format("[{0}]{1}", j, item)));
                }
                j++;
            }

            //Delete all rows before header row
            for (int k = headerRowNum; k >= 0; k-- )
            {
                rows[k].Delete();
            }
            _table.AcceptChanges();
            //Delete all empty rows
            if(anyBlankRows)
            {
                for (int l = blankRowNum; l < rows.Count; l++)
                {
                    rows[l].Delete();

                }
                _table.AcceptChanges();
            }
            
            int columnCount = columns.Count;
            int columnRemainder = columnCount / nupFactor;

            int rowModulo = rows.Count % nupFactor;
            int rowRemainder = rows.Count / nupFactor;
            if(rowModulo != 0)
            {
                rowRemainder++;
            }
            
            int rowsToFill = nupFactor;
            int startingRowCount = 0;

            

            if (sortedCheckBox.IsChecked ?? false)
            {
                int filledRowCount = 1;
                int currentSplitRowCount = rowRemainder;
                int currentColumnCount = columnRemainder;
                while (currentSplitRowCount < rows.Count)
                {
                    if (filledRowCount < rowsToFill)
                    {
                        var currentSplitRow = rows[currentSplitRowCount];
                        var currentSplitRowValues = currentSplitRow.ItemArray;
                        foreach (var val in currentSplitRowValues)
                        {
                            var fillToCount = (filledRowCount + 1) * columnRemainder;
                            if (currentColumnCount == fillToCount || currentColumnCount == columnCount)
                            {
                                break;
                            }
                            rows[startingRowCount][currentColumnCount] = val;
                            currentColumnCount++;
                        }
                        filledRowCount++;
                        currentSplitRowCount++;

                    }
                    else
                    {
                        startingRowCount++;
                        currentColumnCount = columnRemainder;
                        filledRowCount = 1;
                    }
                }
            } 
            else
            {

                int filledRowCount = 0;               
                int currentSplitRowCount = 0;
                int currentColumnCount = 0;
                while (currentSplitRowCount < rows.Count)
                {
                    if (filledRowCount < rowsToFill)
                    {
                        var currentSplitRow = rows[currentSplitRowCount];
                        var currentSplitRowValues = currentSplitRow.ItemArray;
                        foreach (var val in currentSplitRowValues)
                        {
                            var fillToCount = (filledRowCount + 1) * columnRemainder;
                            if (currentColumnCount == fillToCount || currentColumnCount == columnCount)
                            {
                                break;
                            }
                            rows[startingRowCount][currentColumnCount] = val;
                            currentColumnCount++;
                        }
                        filledRowCount++;
                        currentSplitRowCount++;

                    }
                    else
                    {
                        startingRowCount++;
                        currentColumnCount = 0;
                        filledRowCount = 0;
                    }
                }
            }


            //Delete all extra rows
            //rowRemainder
            for (int l = rows.Count - 1; l >= rowRemainder; l--)
            {
                var row = rows[l];
                row.Delete();
                //_table.AcceptChanges();
            }
            _table.AcceptChanges();
            _table.ExportToExcel(output);
            if (TableSelectComboBox.HasItems)
            {
                TableSelectComboBox.SelectedIndex = -1;
                while(_excelSpreadSheets.Count > 0)
                {
                    _excelSpreadSheets.RemoveAt(0);
                }
                //Additional information: Operation is not valid while ItemsSource is in use. Access and modify elements with ItemsControl.ItemsSource instead.
            }
            if (BlankRowSelectionComboBox.HasItems)
            {
                BlankRowSelectionComboBox.SelectedIndex = -1;

                while (_blankRowList.Count > 0)
                {
                    _blankRowList.RemoveAt(0);
                }
            }
                
            if (TableSelectComboBoxHeaderRow.HasItems)
            {
                TableSelectComboBoxHeaderRow.SelectedIndex = -1;
                while(_headerRowList.Count > 0)
                {
                    _headerRowList.RemoveAt(0);
                }
            }
                
            if (!string.IsNullOrEmpty(RowsTxtBox.Text))
            {
                RowsTxtBox.Clear();
            }
                
            if (!string.IsNullOrEmpty(OutDirTxtBox.Text))
            {
                OutDirTxtBox.Clear();
            }
                
            if (!string.IsNullOrEmpty(OutFileNameTxtBox.Text))
            {
                OutFileNameTxtBox.Clear();
            }
                
            if (!string.IsNullOrEmpty(DivideFactor.Text))
            {
                DivideFactor.Clear();
            }
                
            if (!string.IsNullOrEmpty(ColumnsTxtBox.Text))
            {
                ColumnsTxtBox.Clear();
            }
            if (!string.IsNullOrEmpty(FileNameTxtBox.Text))
            {
                FileNameTxtBox.Clear();
            }
                
            MessageBoxDialog("Excel File successfully generated!", "Success");
            //_table.ExportToExcel
        }

        private void Button_Click_Outdir(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            OutDirTxtBox.Text = dialog.SelectedPath;
            
            
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }




    }
}

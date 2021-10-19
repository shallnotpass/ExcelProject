using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelProject1.Models;

namespace ExcelProject1
{
    public partial class Form1 : Form
    {
        private DataProvider DBDataProvider;
        private ExcelDataProvider ExcelProvider;
        private List<TimedValueDBO> TimedValues;
        private List<ValueDBO> Values;
        public Form1()
        {
            InitializeComponent();
            this.ExcelProvider = new ExcelDataProvider();
            this.DBDataProvider = new DataProvider();
        }

        private async void SavingIntoFileButton(object sender, EventArgs e)
        {
            Task[] tasks = new Task[2]
            {
                Task.Factory.StartNew(() => { ExcelProvider.SaveDataCsv(this.Values); }),
                Task.Factory.StartNew(() => { ExcelProvider.SaveDataXlsx(this.TimedValues); })
            };
            await Task.WhenAll(tasks);
            this.ShowMessage("Запись в файлы произведена");
        }

        private async void GettingFromFileButton(object sender, EventArgs e)
        {
            Task<List<ValueDBO>> readingCsv = ExcelProvider.GetCsvData("Excel2.csv");
            Task<List<TimedValueDBO>> readingXlsx = ExcelProvider.GetXlsxData("Excel1.xlsx");
            await Task.WhenAll(readingCsv, readingXlsx);
            this.Values = readingCsv.Result;
            this.TimedValues = readingXlsx.Result;
            this.Values.OrderBy(x => x.TagName);
            this.TimedValues.OrderBy(x => x.TagName);
            this.ShowMessage("Данные из файлов получены");
            this.button1.Visible = true;
            this.button3.Visible = true;
        }

        private void ShowMessage(string message)
        {
            this.textBox1.Text = message;
        }

        private async void SavingIntoBDButton(object sender, EventArgs e)
        {
            var savingTask = Task.Factory.StartNew(() => { this.DBDataProvider.SaveDataToDatabase(this.Values, this.TimedValues); });
            await Task.WhenAll(savingTask);
            this.ShowMessage("Данные cохранены в БД");
        }

        private async void GettingFromDBButton(object sender, EventArgs e)
        {
            Task<List<ValueDBO>> readingCsv = this.DBDataProvider.ReadValuesFromDatabase();
            Task<List<TimedValueDBO>> readingXlsx = this.DBDataProvider.ReadTimedValuesFromDatabase();
            await Task.WhenAll(readingCsv, readingXlsx);
            this.Values = readingCsv.Result;
            this.TimedValues = readingXlsx.Result;
            this.Values.OrderBy(x => x.TagName);
            this.TimedValues.OrderBy(x => x.TagName);
            this.ShowMessage("Данные из БД получены");
            this.button1.Visible = true;
            this.button3.Visible = true;
        }
    }
}

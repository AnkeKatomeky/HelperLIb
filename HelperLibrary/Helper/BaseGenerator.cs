using HelperLibrary;
using HelperLibrary.ExcelOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace HelperLibrary
{
    public abstract class BaseGenerator
    {
        protected List<string> _errors;
        protected StringBuilder _errorMessageBuilder;
        protected System.ComponentModel.BackgroundWorker backWorker;

        public delegate void ReportProgressProc(int percentProgress);
        public delegate void ReportProgressProcFull(int percentProgress, object userState);
        public event EventHandler<ProgressChangedEventArgs> ProgressChanged;
        public event EventHandler<OperationCompletedEventArgs> OperationCompleted;

        /// <summary>
        /// Запускающий работника метод
        /// </summary>
        public virtual void Create()
        {
            _errors.Clear();
            _errorMessageBuilder.Clear();
            backWorker = new System.ComponentModel.BackgroundWorker();
            backWorker.WorkerReportsProgress = true;
            backWorker.ProgressChanged += _backWorker_ProgressChanged;
            backWorker.DoWork += _backWorker_DoWork;
            backWorker.RunWorkerCompleted += _backWorker_RunWorkerCompleted;
            backWorker.RunWorkerAsync();
        }

        protected void _backWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            OperationCompleted?.Invoke(this, new OperationCompletedEventArgs(_errorMessageBuilder.ToString()));
        }

        protected void _backWorker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            ProgressChanged?.Invoke(this, new ProgressChangedEventArgs(e.ProgressPercentage, (string)e.UserState));
        }

        protected void _backWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            int fromPercent;
            int toPercent;

            try
            {
                System.ComponentModel.BackgroundWorker worker = sender as System.ComponentModel.BackgroundWorker;

                fromPercent = 0;
                toPercent = 50;
                worker.ReportProgress(fromPercent, "Загрузка отчетов ...");
                if (!LoadData(worker.ReportProgress, fromPercent, toPercent))
                {
                    _errorMessageBuilder.AppendLine("Не удалось загрузить отчеты!");
                    return;
                }

                fromPercent = toPercent;
                toPercent = 80;
                worker.ReportProgress(fromPercent, "Формирование отчета...");
                if (!Modeling(worker.ReportProgress, fromPercent, toPercent))
                {
                    _errorMessageBuilder.AppendLine("Не удалось сформировать отчет!");
                    return;
                }

                fromPercent = toPercent;
                toPercent = 100;
                worker.ReportProgress(fromPercent, "Сохранение...");
                if (!Save(worker.ReportProgress, fromPercent, toPercent))
                {
                    _errorMessageBuilder.AppendLine("Не удалось сохранить!");
                    return;
                }
            }
            catch (Exception ex)
            {
                _errorMessageBuilder.AppendLine("Неотложеная ошибка:\r\n\n" + ex.Message + "#" + ex.StackTrace);
                return;
            }
            string errorsFileName = SupportApplication.StartupPath + @"\Input\Errors.txt";
            try
            {
                if (_errors.Count > 0)
                {
                    System.IO.File.WriteAllLines(errorsFileName, _errors.Distinct());
                    System.Diagnostics.Process.Start(errorsFileName);
                }
            }
            catch
            {
                _errorMessageBuilder.AppendLine("Файл" + errorsFileName + " уже открыт:\r\n\n");
            }

            Open();
        }

        /// <summary>
        /// Метод загрузки исходника. Рекомендуется использовать совместно с методом Map для нормальной разметки процента выполнения
        /// </summary>
        /// <param name="reportProgressProc">Делегат(метод как значение) через который можно отображать статус исполнения загрузки</param>
        /// <param name="currentPercentage">Нижний порог прцента</param>
        /// <param name="maxPercentage">Верхний порог процента</param>
        /// <returns>В возврате верно ли все загрузилось</returns>
        protected abstract bool LoadData(ReportProgressProcFull reportProgressProc, int currentPercentage, int maxPercentage);

        /// <summary>
        /// Метод обработки данных из исходников. Рекомендуется использовать совместно с методом Map для нормальной разметки процента выполнения
        /// </summary>
        /// <param name="reportProgressProc">Делегат(метод как значение) через который можно отображать статус исполнения загрузки</param>
        /// <param name="currentPercentage">Нижний порог прцента</param>
        /// <param name="maxPercentage">Верхний порог процента</param>
        /// <returns>В возврате верно ли все обработалось</returns>
        protected abstract bool Modeling(ReportProgressProc reportProgressProc, int currentPercentage, int maxPercentage);

        /// <summary>
        /// Метод сохранения обработаных данных из исходников. Рекомендуется использовать совместно с методом Map для нормальной разметки процента выполнения
        /// </summary>
        /// <param name="reportProgressProc">Делегат(метод как значение) через который можно отображать статус исполнения загрузки</param>
        /// <param name="currentPercentage">Нижний порог прцента</param>
        /// <param name="maxPercentage">Верхний порог процента</param>
        /// <returns>В возврате верно ли все загрузилось</returns>
        protected abstract bool Save(ReportProgressProc reportProgressProc, int currentPercentage, int maxPercentage);

        /// <summary>
        /// Открытие собраных отчетов
        /// </summary>
        protected abstract void Open();

        protected virtual double Map(double value, double fromLower, double fromUpper, double toLower, double toUpper)
        {
            return (toLower + (value - fromLower) / (fromUpper - fromLower) * (toUpper - toLower));
        }
        protected virtual ExcelCell DateOrEmpty(DateTime? p)
        {
            if (!p.HasValue)
            {
                return new ExcelCellString(null);
            }
            else if (p.Value.Year < 1905)
            {
                return new ExcelCellString(null);
            }
            else
            {
                return new ExcelCellDate(p);
            }
        }
        protected virtual ExcelCell DecimalOrEmpty(decimal d)
        {
            if (d < 0)
            {
                return new ExcelCellString("");
            }
            else
            {
                return new ExcelCellNumberFractional(d);
            }
        }
        protected virtual ExcelCell DecimalOrInt(decimal? d)
        {
            if (d % 1 != 0)
            {
                return new ExcelCellNumberFractional(d);
            }
            else
            {
                return new ExcelCellNumberIntegral((int?)d);
            }
        }
        protected virtual ExcelCell ObjectToCell(object obj)
        {
            if (obj is string)
            {
                return new ExcelCellString((string)obj);
            }
            else if (obj is DateTime)
            {
                return new ExcelCellDate((DateTime)obj);
            }
            else
            {
                return new ExcelCellString(null);
            }
        }
        protected virtual ExcelCell ObjectToCell(object obj, List<IfFormat> list)
        {
            if (obj is string)
            {
                return new ExcelCellString((string)obj, list);
            }
            else if (obj is DateTime)
            {
                return new ExcelCellDate((DateTime)obj, list);
            }
            else
            {
                return new ExcelCellString(null, list);
            }
        }
        protected virtual string ReportFileName(string fileName)
        {
            return string.Format(SupportApplication.StartupPath + fileName, DateTime.Today.ToString("yyyy.MM.dd"));
        }
        protected virtual double DoubleFromStrin(string input)
        {
            double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double output);
            return output;
        }
        protected virtual void ValidateFile(string excelfile)
        {
            if (!ExcelOpenXmlSaxWriter.ValidateFile(excelfile))
            {
                _errors.Add(string.Format("Файл {0} не прошел валидациию. Данные в таблице могут отсутствовать или быть некорректными.", excelfile));
            }
        }
    }
}

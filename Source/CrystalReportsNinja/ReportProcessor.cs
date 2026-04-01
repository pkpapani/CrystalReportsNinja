using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Data.SqlClient;
using System.Globalization;
using CeLocale = CrystalDecisions.ReportAppServer.DataDefModel.CeLocale;

namespace CrystalReportsNinja
{
    public class ReportProcessor
    {
        private string _sourceFilename;
        private string _outputFilename;
        private string _outputFormat;
        private bool _printToPrinter;
        private string _logfilename;

        private ReportDocument _reportDoc;
        private LogWriter _logger;

        public ArgumentContainer ReportArguments { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="logfilename"></param>
        public ReportProcessor(string logfilename)
        {
            _reportDoc = new ReportDocument();
            _logfilename = logfilename;
            _logger = new LogWriter(_logfilename);
        }

        /// <summary>
        /// Load external Crystal Report file into Report Document
        /// </summary>
        private void LoadReport()
        {
            _sourceFilename = ReportArguments.ReportPath.Trim();
            if (_sourceFilename == null || _sourceFilename == string.Empty)
            {
                throw new Exception("Invalid Crystal Reports file");
            }

            if (_sourceFilename.LastIndexOf(".rpt") == -1)
                throw new Exception("Invalid Crystal Reports file");

            if (ReportArguments.LCID.HasValue)
            {
                CultureInfo locale = CultureInfo.GetCultureInfo(ReportArguments.LCID.Value);
                Console.WriteLine("Using locale {0} {1}", locale.LCID, locale.EnglishName);
                _reportDoc.ReportClientDocument.LocaleID = (CeLocale)locale.LCID;
                _reportDoc.ReportClientDocument.PreferredViewingLocaleID = (CeLocale)locale.LCID;
                _reportDoc.ReportClientDocument.ProductLocaleID = (CeLocale)locale.LCID;
            }

            _reportDoc.Load(_sourceFilename, OpenReportMethod.OpenReportByDefault);
            _logger.Write(string.Format("Report loaded successfully"));
            Console.WriteLine("Report loaded successfully");
        }

        /// <summary>
        /// Match User input parameter values with Report parameters
        /// </summary>
        private void ProcessParameters()
        {
            ParameterCore paraCore = new ParameterCore(_logfilename, ReportArguments.ParameterCollection);
            paraCore.ProcessRawParameters();

            foreach (ParameterField pf in _reportDoc.ParameterFields)
            {
                if (pf.ReportParameterType != ParameterType.QueryParameter &&
                    pf.ReportParameterType != ParameterType.StoreProcedureParameter &&
                    !string.IsNullOrEmpty(pf.ReportName)/* CR do not like setting value to a subreport parameter */) continue;

                object value = paraCore.GetParameterValue(pf.Name);
                _reportDoc.SetParameterValue(pf.Name, value);
            }
        }

        /// <summary>
        /// Validate configurations related to program output.
        /// </summary>
        /// <remarks>
        /// Program output can be TWO forms
        /// 1. Export as a file
        /// 2. Print to printer
        /// </remarks>
        private void ValidateOutputConfigurations()
        {
            _outputFilename = ReportArguments.OutputPath;
            _outputFormat = ReportArguments.OutputFormat;
            _printToPrinter = ReportArguments.PrintOutput;

            bool specifiedFileName = _outputFilename != null ? true : false;
            bool specifiedFormat = _outputFormat != null ? true : false;

            if (!_printToPrinter)
            {
                string fileExt = "";

                //default set to text file
                if (!specifiedFileName && !specifiedFormat)
                    _outputFormat = "txt";

                // Use output format to set output file name extension
                if (specifiedFormat)
                {
                    if (_outputFormat.ToUpper() == "XLSDATA")
                        fileExt = "xls";
                    else if (_outputFormat.ToUpper() == "TAB")
                        fileExt = "txt";
                    else if (_outputFormat.ToUpper() == "ERTF")
                        fileExt = "rtf";
                    else
                        fileExt = _outputFormat;
                }

                // Use output file name extension to set output format
                if (specifiedFileName && !specifiedFormat)
                {
                    int lastIndexDot = _outputFilename.LastIndexOf(".");
                    fileExt = _outputFilename.Substring(lastIndexDot + 1, 3); //what if file ext has 4 char

                    //ensure filename extension has 3 char after the dot (.)
                    if ((_outputFilename.Length == lastIndexDot + 4) && (fileExt.ToUpper() == "RTF" || fileExt.ToUpper() == "TXT" || fileExt.ToUpper() == "CSV" || fileExt.ToUpper() == "PDF" || fileExt.ToUpper() == "RPT" || fileExt.ToUpper() == "DOC" || fileExt.ToUpper() == "XLS" || fileExt.ToUpper() == "XML" || fileExt.ToUpper() == "HTM"))
                        _outputFormat = _outputFilename.Substring(lastIndexDot + 1, 3);
                }

                if (specifiedFileName && specifiedFormat)
                {
                    int lastIndexDot = _outputFilename.LastIndexOf(".");
                    if (fileExt != _outputFilename.Substring(lastIndexDot + 1, 3)) //what if file ext has 4 char
                    {
                        _outputFilename = string.Format("{0}.{1}", _outputFilename, fileExt);
                    }
                }

                if (!specifiedFileName)
                    _outputFilename = String.Format("{0}-{1}.{2}", _sourceFilename.Substring(0, _sourceFilename.LastIndexOf(".rpt")), DateTime.Now.ToString("yyyyMMddHHmmss"), fileExt);

                _logger.Write(string.Format("Output Filename : {0}", _outputFilename));
                _logger.Write(string.Format("Output format : {0}", _outputFormat));
            }
        }

        /// <summary>
        /// Perform Login to database tables
        /// </summary>
        private void PerformDBLogin()
        {
            bool toRefresh = ReportArguments.Refresh;

            var server = ReportArguments.ServerName;
            var database = ReportArguments.DatabaseName;
            var username = ReportArguments.UserName;
            var password = ReportArguments.Password;

            if (toRefresh)
            {
                ApplyLogonToTables(_reportDoc.Database.Tables, server, database, username, password);

                foreach (ReportDocument subreport in GetSubreports(_reportDoc))
                    ApplyLogonToTables(subreport.Database.Tables, server, database, username, password);

                Console.WriteLine("Database Login done");
            }
        }

        private void ApplyLogonToTables(Tables tables, string server, string database, string username, string password)
        {
            foreach (Table table in tables)
            {
                var logonInfo = new TableLogOnInfo();

                if (server != null)
                    logonInfo.ConnectionInfo.ServerName = server;

                if (database != null)
                    logonInfo.ConnectionInfo.DatabaseName = database;

                if (username == null && password == null)
                    logonInfo.ConnectionInfo.IntegratedSecurity = true;
                else
                {
                    if (username != null && username.Length > 0)
                        logonInfo.ConnectionInfo.UserID = username;

                    if (password == null) //to support blank password
                        logonInfo.ConnectionInfo.Password = "";
                    else
                        logonInfo.ConnectionInfo.Password = password;
                }
                TestSqlConnection(logonInfo, table.Name);
                table.ApplyLogOnInfo(logonInfo);
            }
        }

        private void TestSqlConnection(TableLogOnInfo logonInfo, string tableName)
        {
            var info = logonInfo.ConnectionInfo;
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = info.ServerName,
                InitialCatalog = info.DatabaseName,
                IntegratedSecurity = info.IntegratedSecurity
            };

            if (!info.IntegratedSecurity)
            {
                builder.UserID = info.UserID;
                builder.Password = info.Password;
            }

            _logger.Write(string.Format("Testing SQL connection for table '{0}' on server '{1}', database '{2}'.", tableName, info.ServerName, info.DatabaseName));

            try
            {
                using (var conn = new SqlConnection(builder.ConnectionString))
                    conn.Open();
            }
            catch (Exception ex)
            {
                _logger.Write(string.Format("SQL connection test failed for table '{0}': {1}", tableName, ex.Message));
            }
        }

        private System.Collections.Generic.IEnumerable<ReportDocument> GetSubreports(ReportDocument reportDoc)
        {
            foreach (Section section in reportDoc.ReportDefinition.Sections)
            {
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject is SubreportObject subreportObject)
                        yield return subreportObject.OpenSubreport(subreportObject.SubreportName);
                }
            }
        }

        /// <summary>
        /// Set export file type or printer to Report Document object.
        /// </summary>
        private void ApplyReportOutput()
        {
            if (_printToPrinter)
            {
                var printerName = ReportArguments.PrinterName != null ? ReportArguments.PrinterName.Trim() : "";

                if (printerName.Length > 0)
                {
                    _reportDoc.PrintOptions.PrinterName = printerName;
                }
                else
                {
                    System.Drawing.Printing.PrinterSettings prinSet = new System.Drawing.Printing.PrinterSettings();

                    if (prinSet.PrinterName.Trim().Length > 0)
                        _reportDoc.PrintOptions.PrinterName = prinSet.PrinterName;
                    else
                        throw new Exception("No printer name is specified");
                }
            }
            else
            {
                if (_outputFormat.ToUpper() == "RTF")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.RichText;
                else if (_outputFormat.ToUpper() == "TXT")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Text;
                else if (_outputFormat.ToUpper() == "TAB")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.TabSeperatedText;
                else if (_outputFormat.ToUpper() == "CSV")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                else if (_outputFormat.ToUpper() == "PDF")
                {
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;

                    var grpCnt = _reportDoc.DataDefinition.Groups.Count;
                    if (grpCnt > 0)
                        _reportDoc.ExportOptions.ExportFormatOptions = new PdfFormatOptions { CreateBookmarksFromGroupTree = true };
                }
                else if (_outputFormat.ToUpper() == "RPT")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.CrystalReport;
                else if (_outputFormat.ToUpper() == "DOC")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows;
                else if (_outputFormat.ToUpper() == "XLS")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Excel;
                else if (_outputFormat.ToUpper() == "XLSDATA")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                else if (_outputFormat.ToUpper() == "ERTF")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.EditableRTF;
                else if (_outputFormat.ToUpper() == "XML")
                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.Xml;
                else if (_outputFormat.ToUpper() == "HTM")
                {
                    HTMLFormatOptions htmlFormatOptions = new HTMLFormatOptions();

                    if (_outputFilename.LastIndexOf("\\") > 0) //if absolute output path is specified
                        htmlFormatOptions.HTMLBaseFolderName = _outputFilename.Substring(0, _outputFilename.LastIndexOf("\\"));

                    htmlFormatOptions.HTMLFileName = _outputFilename;
                    htmlFormatOptions.HTMLEnableSeparatedPages = false;
                    htmlFormatOptions.HTMLHasPageNavigator = true;
                    htmlFormatOptions.FirstPageNumber = 1;

                    _reportDoc.ExportOptions.ExportFormatType = ExportFormatType.HTML40;
                    _reportDoc.ExportOptions.FormatOptions = htmlFormatOptions;
                }
            }
        }

        /// <summary>
        /// Refresh Crystal Report if no input of parameters
        /// </summary>
        private void PerformRefresh()
        {
            bool toRefresh = ReportArguments.Refresh;
            bool noParameter = (_reportDoc.ParameterFields.Count == 0) ? true : false;

            if (toRefresh && noParameter)
                _reportDoc.Refresh();
        }

        /// <summary>
        /// Print or Export Crystal Report
        /// </summary>
        private void PerformOutput()
        {
            if (_printToPrinter)
            {
                var copy = ReportArguments.PrintCopy;
                _reportDoc.PrintToPrinter(copy, true, 0, 0);
                _logger.Write(string.Format("Report printed to : {0}", _reportDoc.PrintOptions.PrinterName));
            }
            else
            {
                _reportDoc.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                DiskFileDestinationOptions diskOptions = new DiskFileDestinationOptions
                {
                    DiskFileName = _outputFilename
                };

                _reportDoc.ExportOptions.DestinationOptions = diskOptions;
                _reportDoc.Export();
                _logger.Write(string.Format("Report exported to : {0}", _outputFilename));
            }
            Console.WriteLine("Completed");
        }

        /// <summary>
        /// Run the Crystal Reports Exporting or Printing process.
        /// </summary>
        public void Run()
        {
            try
            {
                LoadReport();
                ValidateOutputConfigurations();

                PerformDBLogin();
                ApplyReportOutput();
                ProcessParameters();

                PerformRefresh();
                PerformOutput();
            }
            catch (Exception ex)
            {
                _logger.Write(ex.ToString());
                throw ex;
            }
            finally
            {
                _reportDoc.Close();
            }
        }
    }
}

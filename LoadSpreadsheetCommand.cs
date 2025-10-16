namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Command to load the spreadsheet.
    /// </summary>
    public class LoadSpreadsheetCommand : ICommand
    {
        private Spreadsheet spreadsheet; // the spredsheet we want to load to
        private FileStream fs;

        /// <summary>
        /// Initializes a new instance of the <see cref="LoadSpreadsheetCommand"/> class.
        /// </summary>
        /// <param name="spreadsheet">spreadsheet.</param>
        /// <param name="fs">filestream.</param>
        public LoadSpreadsheetCommand(Spreadsheet spreadsheet, FileStream fs)
        {
            this.spreadsheet = spreadsheet;
            this.fs = fs;
        }

        public string Title
        {
            get
            {
                return "load";
            }
        }

        /// <summary>
        /// Executes the command.
        /// </summary>
        public void Execute()
        {
            this.spreadsheet.Load(this.fs);
        }

        /// <summary>
        /// Un-executes the command.
        /// </summary>
        public void UnExecute()
        {
        }
    }
}

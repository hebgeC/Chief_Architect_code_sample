namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface to ensure all commands implement execute method.
    /// </summary>
    public interface ICommand
    {
        /// <summary>
        /// Gets title of command.
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// Method to execute command.
        /// </summary>
        public void Execute();

        /// <summary>
        /// Method to un-execute a command.
        /// </summary>
        public void UnExecute();
    }
}

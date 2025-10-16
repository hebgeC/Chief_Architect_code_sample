namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Invokes commands on the spreadsheet.
    /// </summary>
    public class SpreadSheetInvoker : INotifyPropertyChanged
    {
        private Stack<ICommand> undo = new Stack<ICommand>();
        private Stack<ICommand> redo = new Stack<ICommand>();
        private ICommand staticCommand; // command that shouldn't undo/ redo.
        private bool isUndoEmpty = true;
        private bool isRedoEmpty = true;
        private string undoCommandTitle;
        private string redoCommandTitle;

        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadSheetInvoker"/> class.
        /// </summary>
        public SpreadSheetInvoker()
        {
        }

        /// <summary>
        /// Event that is triggered when a property in the class changes.
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged = (sender, e) => { };

        /// <summary>
        /// Gets the command title of the top command from undo stack.
        /// </summary>
        public string UndoCommandTitle
        {
            get
            {
                return this.undoCommandTitle;
            }

            private set
            {
                if (value != this.undoCommandTitle)
                {
                    this.undoCommandTitle = value;
                    this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.UndoCommandTitle)));
                }
            }
        }

        /// <summary>
        /// Gets the command title of the top command from redo stack.
        /// </summary>
        public string RedoCommandTitle
        {
            get
            {
                return this.redoCommandTitle;
            }

            private set
            {
                if (value != this.redoCommandTitle)
                {
                    this.redoCommandTitle = value;
                    this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.RedoCommandTitle)));
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the undo stack is empty.
        /// </summary>
        public bool IsUndoEmpty
        {
            get
            {
                return this.isUndoEmpty;
            }

            private set
            {
                if (value != this.isUndoEmpty)
                {
                    this.isUndoEmpty = value;
                    this.PropertyChanged(this, new PropertyChangedEventArgs(nameof(this.IsUndoEmpty)));
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the redo stack is empty.
        /// </summary>
        public bool IsRedoEmpty
        {
            get
            {
                return this.isRedoEmpty;
            }

            private set
            {
                if (value != this.isRedoEmpty)
                {
                    this.isRedoEmpty = value;
                    this.PropertyChanged(this, new PropertyChangedEventArgs(nameof(this.IsRedoEmpty)));
                }
            }
        }

        /// <summary>
        /// Enqueue a command.
        /// </summary>
        /// <param name="command">command.</param>
        public void SetCommand(ICommand command)
        {
            this.staticCommand = command;
        }

        /// <summary>
        /// pops and executes command from command queue.
        /// </summary>
        /// <param name="spreadsheet">reference to spreadsheet.</param>
        public void ExecuteCommand()
        {
            if (this.staticCommand != null)
            {
                if (this.staticCommand.Title == "load")
                {
                    this.undo.Clear();
                    this.redo.Clear();
                    this.UpdateProps();
                }

                this.staticCommand.Execute();
                this.staticCommand = null;
                return;
            }

            this.undo.Peek().Execute();
            this.UpdateProps();
        }

        /// <summary>
        /// Undo function.
        /// </summary>
        public void Undo()
        {
            if (this.undo.Count > 0)
            {
                ICommand command = this.undo.Pop();
                this.redo.Push(command);
                command.UnExecute();
            }

            this.UpdateProps();
        }

        /// <summary>
        /// redo function.
        /// </summary>
        public void Redo()
        {
            if (this.redo.Count > 0)
            {
                ICommand command = this.redo.Pop();
                this.undo.Push(command);
                command.Execute();
            }

            this.UpdateProps();
        }

        /// <summary>
        /// Add an undoable command to the stack.
        /// </summary>
        /// <param name="command">command.</param>
        public void AddUndo(ICommand command)
        {
            this.undo.Push(command);
            this.redo.Clear();
            this.UpdateProps();
        }

        /// <summary>
        /// Update the properties of the undo and redo stacks.
        /// </summary>
        private void UpdateProps()
        {
            if (this.undo.Count > 0)
            {
                this.IsUndoEmpty = false;
                this.UndoCommandTitle = this.undo.Peek().Title;
            }
            else
            {
                this.IsUndoEmpty = true;
                this.UndoCommandTitle = string.Empty;
            }

            if (this.redo.Count > 0)
            {
                this.IsRedoEmpty = false;
                this.RedoCommandTitle = this.redo.Peek().Title;
            }
            else
            {
                this.IsRedoEmpty = true;
                this.RedoCommandTitle = string.Empty;
            }
        }
    }
}

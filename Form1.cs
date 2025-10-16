// <copyright file="Form1.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>
// Connor Chase.
// 11796917.

namespace Spreadsheet_Connor_Chase
{
    using System.ComponentModel;
    using System.Security.Cryptography;
    using System.Windows.Forms;
    using SpreadsheetEngine;
    using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

    /// <summary>
    /// The UI class for this Form.
    /// </summary>
    public partial class Form1 : Form
    {
        private Spreadsheet spreadsheet;
        private SpreadSheetInvoker invoker;

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            this.InitializeComponent();
            this.InitializeDataGrid();
            this.spreadsheet = new Spreadsheet(50, 26);
            this.invoker = new SpreadSheetInvoker();
            this.invoker.PropertyChanged += this.SpreadsheetInvoker_PropertyChanged;
            this.spreadsheet.CellPropertyChanged += this.Spreadsheet_CellPropertyChanged;

            this.undoToolStripMenuItem.Enabled = false;
            this.redoToolStripMenuItem.Enabled = false;
        }

        private void SpreadsheetInvoker_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            // handle disabling/ enabling undo/redo buttons. handle setting button text with title command
            if (sender is SpreadSheetInvoker inv)
            {
                if (e.PropertyName == "IsUndoEmpty")
                {
                    if (inv.IsUndoEmpty)
                    {
                        this.undoToolStripMenuItem.Enabled = false;
                    }
                    else
                    {
                        this.undoToolStripMenuItem.Enabled = true;
                    }
                }

                if (e.PropertyName == "IsRedoEmpty")
                {
                    if (inv.IsRedoEmpty)
                    {
                        this.redoToolStripMenuItem.Enabled = false;
                    }
                    else
                    {
                        this.redoToolStripMenuItem.Enabled = true;
                    }
                }

                if (e.PropertyName == "UndoCommandTitle")
                {
                    this.undoToolStripMenuItem.Text = $"Undo {inv.UndoCommandTitle}";
                }

                if (e.PropertyName == "RedoCommandTitle")
                {
                    this.redoToolStripMenuItem.Text = $"Redo {inv.RedoCommandTitle}";
                }
            }
        }

        private void Spreadsheet_CellPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            // when a cells Text changes, it gets updated in the datagridview
            if (sender is Cell cell && e.PropertyName == "Text")
            {
                DataGridViewRow row = this.dataGridView1.Rows[cell.RowIndex];
                row.Cells[cell.ColumnIndex].Value = cell.Value;
            }

            // when a cells background color changed, update it in the datagridview
            else if (sender is Cell && e.PropertyName == "BGColor")
            {
                cell = (Cell)sender;

                DataGridViewRow row = this.dataGridView1.Rows[cell.RowIndex];
                row.Cells[cell.ColumnIndex].Style.BackColor = Color.FromArgb((int)cell.BGColor);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        /// <summary>
        /// displpays evaluation of cell upon end editing.
        /// </summary>
        /// <param name="sender">the object sender.</param>
        /// <param name="e">event arguments.</param>
        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // get value from dataGridView at specified index
            string inputValue = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty;
            int rowIndex = e.RowIndex;
            int columnIndex = e.ColumnIndex;

            // set the text of the cell in the spreadsheet equal to dateaGridView value
            Cell internalCell = this.spreadsheet.GetCell(rowIndex, columnIndex);
            ChangeCellTextCommand command = new ChangeCellTextCommand(internalCell, inputValue);
            this.invoker.SetCommand(command);
            this.invoker.AddUndo(command);
            this.invoker.ExecuteCommand();
            var internalCellValue = this.spreadsheet.GetCell(rowIndex, columnIndex).Value;
            this.dataGridView1.Rows[rowIndex].Cells[columnIndex].Value = internalCellValue;
        }

        /// <summary>
        /// displays test property of cell upon editing.
        /// </summary>
        /// <param name="sender">the object sender.</param>
        /// <param name="e">event arguments.</param>
        private void DataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            string textProp = this.spreadsheet.GetCell(e.RowIndex, e.ColumnIndex).Text;

            if (textProp != string.Empty)
            {
                this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = textProp;
            }
        }

        /// <summary>
        /// Initialized grid with columns A-Z and rows 1-50.
        /// </summary>
        private void InitializeDataGrid()
        {
            this.InitializeColumns();
            this.InitializeRows();
        }

        /// <summary>
        /// Initializes columns 1-50.
        /// </summary>
        private void InitializeColumns()
        {
            string str = new string(string.Empty); // so that c.ToString() is not called multipte times
            for (char c = 'A'; c <= 'Z'; c++)
            {
                str = c.ToString();
                DataGridViewColumn newCol = new DataGridViewColumn();
                newCol.HeaderText = str;
                newCol.Name = str;
                newCol.CellTemplate = new DataGridViewTextBoxCell(); // this part is very important. otherwise you spend an hour wondering why you cant add rows
                this.dataGridView1.Columns.Add(newCol);
            }
        }

        /// <summary>
        /// Initializes rows 1-50.
        /// </summary>
        private void InitializeRows()
        {
            for (int i = 1; i <= 50; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.HeaderCell = new DataGridViewRowHeaderCell();
                row.HeaderCell.Value = i.ToString();

                this.dataGridView1.Rows.Add(row);
            }
        }

        /// <summary>
        /// populates columns A and B, as well as random 50 cells.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">e.</param>
        private void Demo_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 50; i++)
            {
                Random rnd = new Random();
                int ranRowIndex = rnd.Next(0, 50);
                int ranColumnIndex = rnd.Next(0, 26);

                // populate random cells
                this.spreadsheet.GetCell(ranRowIndex, ranColumnIndex).Text = "It Worked!!";
            }

            for (int i = 0; i < 50; i++)
            {
                // populate all "B" cells
                this.spreadsheet.GetCell(i, 1).Text = "this is cell B" + (i + 1);

                // populate all "A" cells
                this.spreadsheet.GetCell(i, 0).Text = "=B" + (i + 1);
            }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.invoker.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.invoker.Redo();
        }

        private void changeBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog MyDialog = new ColorDialog();
            MyDialog.AllowFullOpen = false;

            DataGridViewSelectedCellCollection selectedCells = this.dataGridView1.SelectedCells;

            // sets the initial color select to the default color
            MyDialog.Color = Color.FromArgb(unchecked((int)0xFFFFFFFF));

            if (MyDialog.ShowDialog() == DialogResult.OK)
            {
                List<Cell> internalCells = new List<Cell>();

                foreach (DataGridViewCell cell in selectedCells)
                {
                    int rowIndex = cell.RowIndex;
                    int columnIndex = cell.ColumnIndex;
                    internalCells.Add(this.spreadsheet.GetCell(rowIndex, columnIndex));
                }

                ChangeCellBGColorCommand command = new ChangeCellBGColorCommand(internalCells, (uint)MyDialog.Color.ToArgb());
                this.invoker.SetCommand(command);
                this.invoker.AddUndo(command);
                this.invoker.ExecuteCommand();
            }
        }

        /// <summary>
        /// button for loading a spreadsheet.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">event args.</param>
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XML Files (*.xml)|*.xml";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(openFileDialog.FileName, FileMode.Open);
                LoadSpreadsheetCommand command = new LoadSpreadsheetCommand(this.spreadsheet, fs);
                this.invoker.SetCommand(command);
                this.invoker.ExecuteCommand();
                fs.Dispose();
            }
        }

        /// <summary>
        /// button for saving a spreadsheet.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">event args.</param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML Files (*.xml)|*.xml";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.OpenOrCreate);
                SaveSpreadsheetCommand command = new SaveSpreadsheetCommand(this.spreadsheet, fs);
                this.invoker.SetCommand(command);
                this.invoker.ExecuteCommand();
                fs.Dispose();
            }
        }
    }
}

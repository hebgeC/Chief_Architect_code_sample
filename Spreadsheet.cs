// <copyright file="Spreadsheet.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>
// Connor Chase.
// 11796917.
namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Runtime.CompilerServices;
    using System.Runtime.Intrinsics.Arm;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;

    /// <summary>
    /// Class to instantiate spreadsheet.
    /// </summary>
    public class Spreadsheet
    {
        private const string BADREF = "BAD_REF";
        private const string SELFREF = "SELF_REF";
        private const string CIRCREF = "CIRC_REF";
        private const string DIVZERO = "DIV_ZERO";

        private Cell[,] grid;
        private Dictionary<Cell, HashSet<Cell>> dependents; // each key (cell) is paired with a list of cells that depend upon the key
        private Dictionary<Cell, HashSet<Cell>> dependencies; // each key (cell) is paired with a list of cells that the key depends on

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="rows">rows.</param>
        /// <param name="columns">columns.</param>
        public Spreadsheet(int rows, int columns)
        {
            if (rows <= 0 || columns <= 0)
            {
                throw new ArgumentException("Rows and columns must be greater than zero.");
            }

            this.grid = new Cell[rows, columns];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    this.grid[i, j] = new TextCell(i, j);
                    this.grid[i, j].rowIndex = i;
                    this.grid[i, j].columnIndex = j;
                    this.grid[i, j].PropertyChanged += this.Cell_PropertyChangedHandler; // each cell in the grid subscribes to the cell event handler
                }
            }

            this.dependents = new Dictionary<Cell, HashSet<Cell>>();
            this.dependencies = new Dictionary<Cell, HashSet<Cell>>();

            foreach (Cell cell in this.grid)
            {
                this.dependents[cell] = new HashSet<Cell>();
                this.dependencies[cell] = new HashSet<Cell>();
            }
        }

        /// <summary>
        /// notifies when a property of any cell changes.
        /// </summary>
        public event PropertyChangedEventHandler? CellPropertyChanged = (sender, e) => { };

        /// <summary>
        /// Gets the specified cell.
        /// </summary>
        /// <param name="rowIndex"> row index.</param>
        /// <param name="columnIndex">column index.</param>
        /// <returns>Returns the specified cell.</returns>
        public Cell GetCell(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex >= this.grid.GetLength(0))
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "rowIndex is out of bounds");
            }

            if (columnIndex < 0 || columnIndex >= this.grid.GetLength(1))
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "columnIndex is out of bounds");
            }

            return this.grid[rowIndex, columnIndex];
        }

        /// <summary>
        /// Save spreadsheet to XML file.
        /// </summary>
        /// <param name="fs">filestream.</param>
        public void Save(FileStream fs)
        {
            XmlSaver.Save(this.grid, fs);
        }

        /// <summary>
        /// Load spreadsheet from XML file.
        /// </summary>
        /// <param name="fs">filestream.</param>
        public void Load(FileStream fs)
        {
            XmlLoader.Load(this.grid, fs);
        }

        /// <summary>
        /// defines what to do in the event an expression reference changes.
        /// </summary>
        /// <param name="sender">the object that fired the event.</param>
        /// <param name="e">event arguments.</param>
        /// <exception cref="NotFiniteNumberException">Not A Number.</exception>
        private void OnExpressionRef_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (sender is Cell cell)
            {
                if (e.PropertyName == "Text")
                {
                    string refCellContent;

                    foreach (Cell dependent in this.dependents[cell])
                    {
                        ExpressionTree expression = new ExpressionTree("empty");

                        if (dependent.Text.Length > 0)
                        {
                            expression = new ExpressionTree(dependent.Text.Substring(1));
                        }
                        else
                        {
                            expression = new ExpressionTree("00");
                        }

                        List<string> variableList = expression.GetVariableList();

                        foreach (string variable in variableList)
                        {
                            // get value of cell. (e.g. 'B2')
                            if (this.GetCellFromReference(variable, out Cell? refCell))
                            {
                                if (this.HandleSelfRef(dependent, refCell))
                                {
                                    return;
                                }

                                if (this.HasCircularReference(cell))
                                {
                                    cell.Value = this.SetErrorMessage(CIRCREF);
                                    return;
                                }

                                refCellContent = refCell.Value;


                                if (double.TryParse(refCellContent, out double number))
                                {
                                    // set the value of the varialbe
                                    expression.SetVariable(variable, number);
                                }
                                else
                                {
                                    expression.SetVariable(variable, 0); // NAN  
                                }
                            }
                            else
                            {
                                dependent.Value = this.SetErrorMessage(BADREF);
                            }
                        }

                        // set the value of the cell to the evaluated expression
                        try
                        {
                            double value = expression.Evaluate();
                            dependent.Value = value.ToString();
                        }
                        catch (DivideByZeroException)
                        {
                            dependent.Value = this.SetErrorMessage(DIVZERO);
                        }

                        this.CellPropertyChanged?.Invoke(dependent, e);
                    }
                }
                else if (e.PropertyName == "Value")
                {
                    this.UpdateDependents(cell);
                }
            }
        }

        /// <summary>
        /// defines what to do in the event an assignment reference changes.
        /// </summary>
        /// <param name="sender">object who fired event.</param>
        /// <param name="e">event arguments.</param>
        private void OnAssignmentRef_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (sender is Cell cell)
            {
                // sender is the referenced cell
                if (e.PropertyName == "Text")
                {
                    foreach (Cell dependant in this.dependents[cell]) // get cells that depend on sender
                    {
                        // we unsubscribe from the event handler because when the new Text property is changed, the necessary references will be made again
                        dependant.PropertyChanged -= this.OnAssignmentRef_PropertyChanged;
                        dependant.Text = cell.Text;
                    }
                }
                else if (e.PropertyName == "Value") // sender is the referenced cell
                {
                    this.UpdateDependents(cell);
                }
            }
        }

        /// <summary>
        /// Defines what to do when a cell property has been changed.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">e.</param>
        private void Cell_PropertyChangedHandler(object? sender, PropertyChangedEventArgs e)
        {
            if (sender is Cell cell)
            {
                if (e.PropertyName == "Text")
                {
                    this.ClearDependencies(cell);

                    // formula is a simple cell assignment
                    if (Regex.IsMatch(cell.Text, @"^=[A-Z]+[1-9]*$") || Regex.IsMatch(cell.Text, @"^=[a-z]+[1-9]*$"))
                    {
                        this.HandleAssignment(cell);
                    }

                    // formula is expression
                    if (cell.Text.StartsWith("="))
                    {
                        this.HandleExpression(cell);
                    }
                    else
                    {
                        cell.Value = cell.Text;
                    }

                    this.grid[cell.RowIndex, cell.ColumnIndex].Value = cell.Value;
                }
                else if (e.PropertyName == "BGColor")
                {
                    this.grid[cell.RowIndex, cell.ColumnIndex].BGColor = cell.BGColor;
                }

                this.CellPropertyChanged?.Invoke(sender, e);
            }
        }

        private void UpdateDependents(Cell cell)
        {
            // update any cells that depend on cell
            foreach (Cell dependant in this.dependents[cell])
            {
                dependant.Value = cell.Value;
            }
        }

        /// <summary>
        /// handles expression formula.
        /// </summary>
        /// <param name="cell">cell containing expression.</param>
        /// <exception cref="NotFiniteNumberException">not a number.</exception>
        private void HandleExpression(Cell cell)
        {
            string refCellContent;

            // STEP 1: build expression tree from expression
            ExpressionTree expression = new ExpressionTree(cell.Text.Substring(1));

            // STEP 2: get list of variables from expression tree
            List<string> variableList = expression.GetVariableList();

            // STEP 3: clear the cell's dependacies
            foreach (var item in this.dependencies[cell])
            {
                this.dependents[item].Remove(cell);
                item.PropertyChanged -= this.OnExpressionRef_PropertyChanged;
            }

            this.dependencies[cell].Clear();

            // STEP 4: set the value of each variable in tree
            foreach (string variable in variableList)
            {
                // get value of cell. (e.g. 'B2')
                if (this.GetCellFromReference(variable, out Cell? refCell))
                {
                    if (this.HandleSelfRef(cell, refCell))
                    {
                        return;
                    }

                    if (this.HasCircularReference(cell))
                    {
                        cell.Value = this.SetErrorMessage(CIRCREF);
                        return;
                    }

                    refCellContent = refCell.Value;

                    if (double.TryParse(refCellContent, out double number))
                    {
                        // set the value of the varialbe
                        expression.SetVariable(variable, number);

                        // update cell dependancies
                        this.dependents[refCell].Add(cell);
                        this.dependencies[cell].Add(refCell);
                        refCell.PropertyChanged += this.OnExpressionRef_PropertyChanged;
                    }
                    else
                    {
                        this.dependents[refCell].Add(cell);
                        this.dependencies[cell].Add(refCell);
                        expression.SetVariable(variable, 0);
                        refCell.PropertyChanged += this.OnExpressionRef_PropertyChanged;
                    }
                }
                else
                {
                    cell.Value = this.SetErrorMessage(BADREF);
                    return;
                }
            }

            // set the value of the cell to the evaluated expression
            try
            {
                double value = expression.Evaluate();
                cell.Value = value.ToString();
            }
            catch (DivideByZeroException)
            {
                cell.Value = this.SetErrorMessage(DIVZERO);
            }
        }

        private bool HandleSelfRef(Cell cell, Cell? refCell)
        {
            if (refCell == cell)
            {
                cell.Value = this.SetErrorMessage(SELFREF);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Gets the cell from the reference.
        /// </summary>
        /// <param name="variable">variable containing reference.</param>
        private bool GetCellFromReference(string variable, out Cell? cell)
        {
            int specifiedColumn = char.ToUpper(variable[0]) - 'A';
            int specifiedRow = Convert.ToInt32(variable.Substring(1)) - 1;

            // check if cell is out of bounds
            if (specifiedRow < 0 || specifiedRow >= this.grid.GetLength(0) || specifiedColumn < 0 || specifiedColumn >= this.grid.GetLength(1))
            {
                cell = null;
                return false;
            }

            cell = this.GetCell(specifiedRow, specifiedColumn);
            return true;
        }

        /// <summary>
        /// handles assignment formula.
        /// </summary>
        /// <param name="cell">cell containing assignment formula.</param>
        private void HandleAssignment(Cell cell)
        {
            // pull value from specified cell and put in current cell
            if (this.GetCellFromReference(cell.Text.Substring(1), out Cell? refCell))
            {
                if (this.HandleSelfRef(cell, refCell))
                {
                    return;
                }

                if (this.HasCircularReference(cell))
                {
                    cell.Value = this.SetErrorMessage(CIRCREF);
                    return;
                }

                refCell.PropertyChanged += this.OnAssignmentRef_PropertyChanged;
                cell.Value = refCell.Value;

                // cell depends on refCell
                // cell is added to list of cells that depend on refcell
                this.dependents[refCell].Add(cell);
            }
            else
            {
                cell.Value = this.SetErrorMessage(BADREF);
            }
        }

        /// <summary>
        /// Clears the dependencies of a cell.
        /// </summary>
        /// <param name="cell">key.</param>
        private void ClearDependencies(Cell cell)
        {
            foreach (var dependency in this.dependencies[cell])
            {
                this.dependents[dependency].Remove(cell);
                dependency.PropertyChanged -= this.OnExpressionRef_PropertyChanged;
                dependency.PropertyChanged -= this.OnAssignmentRef_PropertyChanged;
            }

            this.dependencies[cell].Clear();
        }

        /// <summary>
        /// Sets the error message.
        /// </summary>
        /// <param name="type">error type.</param>
        /// <returns>Error message.</returns>
        private string SetErrorMessage(string type)
        {
            return $"!({type})";
        }

        private bool HasCircularReference(Cell start)
        {
            HashSet<Cell> visited = new HashSet<Cell>();
            return this.DetectCycle(start, start, visited);
        }

        private bool DetectCycle(Cell start, Cell current, HashSet<Cell> visited)
        {
            if (!this.dependencies.ContainsKey(current))
            {
                return false;
            }

            foreach (Cell neighbor in this.dependencies[current])
            {
                if (neighbor == start)
                {
                    return true;
                }

                if (!visited.Contains(neighbor))
                {
                    visited.Add(neighbor);
                    if (this.DetectCycle(start, neighbor, visited))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Wraps Cell in TextCell.
        /// </summary>
        internal class TextCell : Cell
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="TextCell"/> class.
            /// </summary>
            /// <param name="rowIndex">row index.</param>
            /// <param name="columnIndex">column index.</param>
            public TextCell(int rowIndex, int columnIndex)
            : base(rowIndex, columnIndex)
            {
            }
        }
    }
}

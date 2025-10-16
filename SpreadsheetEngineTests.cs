// <copyright file="SpreadsheetEngineTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>
namespace SpreadSheetEngineTests
{
    using System.Reflection;
    using NuGet.Frameworks;
    using SpreadsheetEngine;

    /// <summary>
    /// tests for spreadsheet engine.
    /// </summary>
    public class SpreadsheetEngineTests
    {
        /// <summary>
        /// set up for tests.
        /// </summary>
        [SetUp]
        public void Setup()
        {
        }

        /// <summary>
        /// tests the GetCell() function from Spreadsheet.
        /// </summary>
        [Test]
        public void TestGetCell()
        {
            Spreadsheet spreadsheet = new Spreadsheet(4, 4); // '4, 4' is arbitrary. just has to be at least 1
            Cell testCell = spreadsheet.GetCell(1, 1);
            Assert.That(testCell.RowIndex, Is.EqualTo(1));
            Assert.That(testCell.ColumnIndex, Is.EqualTo(1));
        }

        /// <summary>
        /// tests GetCell() function to handle arguments greater than array bounds.
        /// </summary>
        [Test]
        public void TestGetCell_outOfBounds()
        {
            Spreadsheet spreadsheet = new Spreadsheet(4, 4); // '4, 4' is arbitrary. just has to be at least 1
            Assert.Throws<ArgumentOutOfRangeException>(() => spreadsheet.GetCell(5, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => spreadsheet.GetCell(0, 5));
        }

        /// <summary>
        /// tests GetCell() function to handle negative arguments for array access.
        /// </summary>
        [Test]
        public void TestGetCell_negativeIndices()
        {
            Spreadsheet spreadsheet = new Spreadsheet(4, 4); // '4, 4' is arbitrary. just has to be at least 1
            Assert.Throws<ArgumentOutOfRangeException>(() => spreadsheet.GetCell(-5, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => spreadsheet.GetCell(0, -5));
        }

        /// <summary>
        /// Tests SetVairable.
        /// </summary>
        [TestCase]
        public void TestSetVariable()
        {
            string input = "A";
            ExpressionTree tree = new ExpressionTree(input);

            double expected = 3;

            tree.SetVariable("A", expected);

            double result = tree.Evaluate();

            Assert.AreEqual(expected, result);
        }

        /// <summary>
        /// tests Evaluate with division by Zero.
        /// </summary>
        [TestCase]
        public void TestSetVariable_doesntExist()
        {
            string input = "A";
            ExpressionTree tree = new ExpressionTree(input);

            Exception exception = Assert.Throws<Exception>(() => tree.SetVariable("B", 3));

            Assert.AreEqual("Variable doesnt exist", exception.Message);
        }

        /// <summary>
        /// tests Evaluate with '+' operator.
        /// </summary>
        [TestCase]
        public void TestEvaluate_addition()
        {
            string input = "1+1+1+1";
            ExpressionTree tree = new ExpressionTree(input);

            double expected = 4;

            double result = tree.Evaluate();

            Assert.AreEqual(expected, result);
        }

        /// <summary>
        /// tests Evaluate with subtration and negative numbers.
        /// </summary>
        [TestCase]
        public void TestEvaluate_subtraction()
        {
            string input = "6-3-2-1";
            ExpressionTree tree = new ExpressionTree(input);

            double expected = 0;

            double result = tree.Evaluate();

            Assert.AreEqual(expected, result);
        }

        /// <summary>
        /// tests Evaluate with multiplication.
        /// </summary>
        [TestCase]
        public void TestEvaluate_mult()
        {
            string input = "2*3";
            ExpressionTree tree = new ExpressionTree(input);

            double expected = 6;

            double result = tree.Evaluate();

            Assert.AreEqual(expected, result);
        }

        /// <summary>
        /// tests Evaluate with division.
        /// </summary>
        [TestCase]
        public void TestEvaluate_div()
        {
            string input = "8/4/2";
            ExpressionTree tree = new ExpressionTree(input);

            double expected = 1;

            double result = tree.Evaluate();

            Assert.AreEqual(expected, result);
        }

        /// <summary>
        /// tests Evaluate with division by Zero.
        /// </summary>
        [TestCase]
        public void TestEvaluate_divByZero()
        {
            string input = "8/0";
            ExpressionTree tree = new ExpressionTree(input);

            Exception exception = Assert.Throws<Exception>(() => tree.Evaluate());

            Assert.AreEqual("Cant divide by Zero", exception.Message);
        }

        /// <summary>
        /// tests GetVariableList.
        /// </summary>
        [Test]
        public void TestGetVariableList()
        {
            ExpressionTree tree = new ExpressionTree("a1+a2");

            List<string> expectedVariableList = new List<string>();
            expectedVariableList.Add("a1");
            expectedVariableList.Add("a2");

            List<string> actualVariableList = tree.GetVariableList();

            Assert.AreEqual(expectedVariableList, actualVariableList);
        }

        /// <summary>
        /// Tests the undo functionality of SpreadsheetInvoker.
        /// </summary>
        [Test]
        public void TestUndo()
        {
            Spreadsheet spreadsheet = new Spreadsheet(5, 5);
            SpreadSheetInvoker invoker = new SpreadSheetInvoker();
            Cell testCell = spreadsheet.GetCell(1, 1);
            ChangeCellTextCommand command = new ChangeCellTextCommand(testCell, "test!");

            invoker.SetCommand(command);
            invoker.AddUndo(command);
            invoker.ExecuteCommand();

            invoker.Undo();

            string resultCommandTitle = invoker.RedoCommandTitle;
            bool resultUndoStackIsEmpty = invoker.IsUndoEmpty;
            bool resultRedoStackIsEmpty = invoker.IsRedoEmpty;

            string expectedCommandTitle = "cell text change";
            bool expectedUndoStackIsEmpty = true;
            bool expectedRedoStackIsEmpty = false;

            Assert.AreEqual (expectedCommandTitle, resultCommandTitle);
            Assert.AreEqual(resultUndoStackIsEmpty, expectedUndoStackIsEmpty);
            Assert.AreEqual(resultRedoStackIsEmpty, expectedRedoStackIsEmpty);
        }

        /// <summary>
        /// Tests the redo functionality of SpreadsheetInvoker.
        /// </summary>
        [Test]
        public void TestRedo()
        {
            Spreadsheet spreadsheet = new Spreadsheet(5, 5);
            SpreadSheetInvoker invoker = new SpreadSheetInvoker();
            Cell testCell = spreadsheet.GetCell(1, 1);
            ChangeCellTextCommand command = new ChangeCellTextCommand(testCell, "test!");

            invoker.SetCommand(command);
            invoker.AddUndo(command);
            invoker.ExecuteCommand();

            invoker.Redo();

            string resultCommandTitle = invoker.UndoCommandTitle;
            bool resultUndoStackIsEmpty = invoker.IsUndoEmpty;
            bool resultRedoStackIsEmpty = invoker.IsRedoEmpty;

            string expectedCommandTitle = "cell text change";
            bool expectedUndoStackIsEmpty = false;
            bool expectedRedoStackIsEmpty = true;

            Assert.AreEqual(expectedCommandTitle, resultCommandTitle);
            Assert.AreEqual(resultUndoStackIsEmpty, expectedUndoStackIsEmpty);
            Assert.AreEqual(resultRedoStackIsEmpty, expectedRedoStackIsEmpty);
        }
    }
}
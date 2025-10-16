// <copyright file="ExpressionTree.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using static SpreadsheetEngine.ExpressionTree;

    /// <summary>
    /// Constructs expression tree to evaluate given expression.
    /// </summary>
    public class ExpressionTree
    {
        private Dictionary<string, double> variables = []; // this notation '[]' is the same as 'new Dictionary<string, double>()'

        private Node root;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpressionTree"/> class.
        /// Build expression tree from given expression.
        /// </summary>
        /// <param name="expression">expression to build tree from.</param>
        public ExpressionTree(string expression)
        {
            this.root = this.Construct(expression);
        }

        /// <summary>
        /// Sets the value of the specified variable within the ExpressionTree variables dictionary.
        /// </summary>
        /// <param name="variableName">variable name to set.</param>
        /// <param name="variableValue">variable value to set.</param>
        public void SetVariable(string variableName, double variableValue)
        {
            if (this.variables.ContainsKey(variableName))
            {
                this.variables[variableName] = variableValue;
            }
            else
            {
                throw new Exception("Variable doesnt exist");
            }
        }

        /// <summary>
        /// getsList of strings of all variables used in expression.
        /// </summary>
        /// <returns>List of strings of all variables used in expression.</returns>
        public List<string> GetVariableList()
        {
            List<string> variableNames = new List<string>();

            foreach (var variable in this.variables)
            {
                variableNames.Add(variable.Key);
            }

            return variableNames;
        }

        /// <summary>
        /// evaluates the expression to a double value.
        /// </summary>
        /// <returns>evaluated value of expression tree.</returns>
        public double Evaluate()
        {
            return this.root.Evaluate(this.variables);  
        }

        /// <summary>
        /// Parses given expression into expression tree.
        /// </summary>
        /// <returns>root node of expression tree.</returns>
        private Node Construct(string expression)
        {
            List<Node> expressionList = this.ToList(expression);
            List<Node> postfixExpression = new List<Node>(this.ConvertToPostfix(expressionList));

            Stack<Node> stack = new Stack<Node>();

            foreach (Node item in postfixExpression)
            {
                if (item is VariableNode)
                {
                    VariableNode tempVar = item as VariableNode;

                    this.variables.Add(tempVar.Name, 0);
                    stack.Push(tempVar);
                }
                else if (item is ConstantNode)
                {
                    ConstantNode tempConst = item as ConstantNode;
                    stack.Push(tempConst);
                }
                else
                {
                    OperatorNode tempOp = item as OperatorNode;

                    Node rightOperand = stack.Pop();
                    Node leftOperand = stack.Pop();

                    tempOp.Left = leftOperand;
                    tempOp.Right = rightOperand;

                    stack.Push(tempOp);
                }
            }

            return stack.Pop();
        }

        /// <summary>
        /// Converts a given expression to postfix notation.
        /// </summary>
        /// <param name="expression">given infix expression.</param>
        /// <returns>list of each expression term and operator in postfix order.</returns>
        private List<Node> ConvertToPostfix(List<Node> expression)
        {
            Stack<OperatorNode> stack = new Stack<OperatorNode>();
            List<Node> postfixExpression = new List<Node>();

            foreach (Node item in expression)
            {
                if (item is VariableNode || item is ConstantNode)
                {
                    postfixExpression.Add(item);
                }
                else if (item is OperatorNode temp)
                {
                    if (temp.Operator == '(')
                    {
                        stack.Push(temp);
                    }
                    else if (temp.Operator == ')')
                    {
                        while (stack.Count > 0 && stack.Peek().Operator != '(')
                        {
                            postfixExpression.Add(stack.Pop());
                        }

                        if (stack.Count > 0 && stack.Peek().Operator == '(')
                        {
                            stack.Pop();
                        }
                    }
                    else
                    {
                        while (stack.Count > 0 && stack.Peek().Priority >= temp.Priority)
                        {
                            postfixExpression.Add(stack.Pop());
                        }

                        stack.Push(temp);
                    }
                }
            }

            while (stack.Count > 0)
            {
                postfixExpression.Add(stack.Pop());
            }

            return postfixExpression;
        }

        /// <summary>
        /// Converts the given expression to a list of nodes. maintains the infix notation.
        /// </summary>
        /// <param name="expression">given expression.</param>
        /// <returns>expression as List of nodes.</Node></returns>
        private List<Node> ToList(string expression)
        {
            int asciiValue;
            StringBuilder currentTerm = new StringBuilder();
            List<Node> result = new List<Node>();
            Node operatorNode;
            Node currentOperand;

            foreach (char character in expression)
            {
                asciiValue = (int)character;

                // ignore whitespace
                if (asciiValue != 32)
                {
                    // need to check if its an operator using ascii values. that way we dont need to add more hard coded operators in the future.
                    // checking of the character is a letter or number
                    if ((asciiValue < 48 || asciiValue > 57) && (asciiValue < 65 || asciiValue > 90) && (asciiValue < 97 || asciiValue > 122))
                    {
                        operatorNode = NodeFactory.CreateNode(character.ToString());
                        currentOperand = NodeFactory.CreateNode(currentTerm.ToString());

                        if (currentTerm.Length != 0)
                        {
                            result.Add(currentOperand);
                        }

                        result.Add(operatorNode);

                        currentTerm.Clear();
                    }
                    else
                    {
                        // no operator was encountered. build the current term with the incoming character
                        currentTerm.Append(character);
                    }
                }
            }

            // we have gone through the entire expression.
            // create node for current term and add to result
            if (currentTerm.Length != 0)
            {
                currentOperand = NodeFactory.CreateNode(currentTerm.ToString());

                result.Add(currentOperand);
            }

            return result;
        }
    }
}

namespace SpreadsheetEngine
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.Metadata;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Factory class for creating Nodes.
    /// </summary>
    internal class NodeFactory
    {
        /// <summary>
        /// Creates a node of a specific type.
        /// </summary>
        /// <param name="term">term can be operaotr or operand.</param>
        /// <returns>node of specific type.</returns>
        public static Node CreateNode(string term)
        {
            string[] operators = { "+", "-", "*", "/", "(", ")" };

            bool isConstant = double.TryParse(term, out double number);

            // if term isnt an operator and isnt a constant
            if (!operators.Contains(term) && !isConstant)
            {
                // return new variable node
                return new VariableNode(term);
            }

            // not an operator, and we know its not a variable if it reaches this point
            else if (isConstant)
            {
                // return new constant node
                return new ConstantNode(number);
            }

            // its an operator
            else
            {
                // return new operator node
                return OperatorNodeFactory.CreateNode(term);
            }

            FindDerivedClasses<OperatorNode>(Assembly.GetExecutingAssembly());
        }

        private static void FindDerivedClasses<TBase>(Assembly assembly)
        {
            var derivedTypes = assembly.GetTypes()
                .Where(type => type.IsClass && !type.IsAbstract && type.IsSubclassOf(typeof(TBase)));

            foreach (var type in derivedTypes)
            {
                Console.WriteLine($"Found: {type.FullName}");
            }
        }
    }
}

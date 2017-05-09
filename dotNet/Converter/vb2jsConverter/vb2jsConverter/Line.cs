/*
 * Copyright 2010 Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace com.google.vb2js {



    /**
     * This localizes most of the processing for tokenizing a single line of input. Constructor
     * does strings and comments.
     * Lots of ad hocery here: ! is a component separator in VB, just blindly included in IDs here.
     * $ is valid at end of VB name.
     *
     * @author Brian Kernighan
     * @author Nikhil Singhal
     */
     public class Line {

        // precedence order, high to low, from VB manual:
        // Arithmetic and Concatenation Operators
        // Exponentiation (^)
        // Unary identity and negation (+, -)
        // Multiplication and floating-point division (*, /)
        // Integer division (\)
        // Modulus arithmetic (Mod)
        // Addition and subtraction (+, -), string concatenation (+)
        // String concatenation (&)
        // Arithmetic bit shift (<<, >>)
        // Comparison Operators
        // =, <>, <, <=, >, >=, Is, IsNot, Like, TypeOf...Is
        // Logical and Bitwise Operators
        // Negation (Not)
        // Conjunction (And, AndAlso)
        // Inclusive disjunction (Or, OrElse)
        // Exclusive disjunction (Xor)
        // all operators are evaluated left to right
        // other complications:
        // And, Or, Xor are both bitwise if operands are integer
        // and logical if they are bool s (e.g., relational tests).
        // AndAlso, OrElse are short-circuit (really && and ||)

        // We're using an ImmutableMap here since the order in which the elements are
        // accessed matters.
        private static RegexOptions op = RegexOptions.Singleline;

        private static Dictionary<Regex, String> TYPES =
              new Dictionary<Regex, string>() {
            {new Regex("^(?i)\\b(Mod|Is|Not|AndAlso|And|OrElse|Or|Xor|Eqv|Like|New)\\b",op),
                "OP"},
            {new Regex("^(?i)\\b(End +(If|Sub|Function|While|With|Select))\\b",op), "ENDXX"},
            {new Regex("^(?i)\\b(Exit)\\b",op), "EXIT"},
            {new Regex("^(?i)\\b(Private|Public|Static|Let|Set)\\b",op), "TOSS"},
            {new Regex("^(?i)\\b(Attribute|Option|Declare)\\b",op), "PUNT"},
            {new Regex("^(?i)\\b(Open .* For |Close #\\w+)\\b",op), "PUNT"},
            {new Regex("^(?i)\\b(Print #|Line Input #)\\b",op), "PUNT"},
            {new Regex("^(?i)\\b(On Error (Resume Next|GoTo 0)|Resume|GoTo)\\b",op), "PUNT"},
            {new Regex("^(?i)\\b(On Error)\\b",op), "ONERROR"},
            {new Regex("^(?i)\\b(Then|Else|To|Downto|Step|As|ByVal|ByRef)\\b",op), "KEY"},
            {new Regex("^(?i)\\b(Type|End Type)\\b",op), "TYPE"},
            {new Regex("^[a-zA-Z](\\w)*\\$?",op), "ID"},
            {new Regex("^#\\d+/\\d+/\\d+#",op), "DATE"},
            {new Regex("^((\\d+\\.?\\d*)|(\\.\\d+))([eE][-+]?\\d+)?[&#]?",op), "NUM"},
            {new Regex("^&H[a-fA-F0-9]+",op), "HEX"},
            {new Regex("^<>|<=|>=|:=",op), "OP"},
            {new Regex("^[*^/\\\\+\\-&=><]",op), "OP"},
            {new Regex("^\"[^\"]*\"",op), "STR"},
            {new Regex("^\".*",op), "COMMENT"},
            {new Regex("^.",op), "CHR"},
            {new Regex("^$",op), "END"}
              };

        private static Dictionary<String, String> KEYWORDS =
              new Dictionary<String, String>() {
            {"and", "And"},
            {"as", "As"},
            {"byref", "ByRef"},
            {"byval", "ByVal"},
            {"case", "Case"},
            {"const", "Const"},
            {"dim", "Dim"},
            {"do", "Do"},
            {"double", "Double"},
            {"downto", "Downto"},
            {"each", "Each"},
            {"else", "Else"},
            {"elseif", "ElseIf"},
            {"end", "End"},
            {"end function", "End Function"},
            {"end if", "End If"},
            {"end sub", "End Sub"},
            {"end select", "End Select"},
            {"end while", "End While"},
            {"end with", "End With"},
            {"error", "Error"},
            {"exit", "Exit"},
            {"false", "False"},
            {"for", "For"},
            {"function", "Function"},
            {"global", "Global"},
            {"goto", "GoTo"},
            {"if", "If"},
            {"integer", "Integer"},
            {"is", "Is"},
            {"like", "Like"},
            {"loop", "Loop"},
            {"mod", "Mod"},
            {"new", "New"},
            {"next", "Next"},
            {"not", "Not"},
            {"nothing", "Nothing"},
            {"null", "Null"},
            {"on", "On"},
            {"or", "Or"},
            {"private", "Private"},
            {"public", "Public"},
            {"resume", "Resume"},
            {"select", "Select"},
            {"single", "Single"},
            {"static", "Static"},
            {"step", "Step"},
            {"sub", "Sub"},
            {"then", "Then"},
            {"to", "To"},
            {"true", "True"},
            {"type", "Type"},
            {"until", "Until"},
            {"while", "While"},
            {"with", "With"},
            {"xor", "Xor"}
            };

        private static HashSet<string> LOGICAL_OPS = new HashSet<string>(new string[] {
            "And",
            "Or",
            "Xor" });

        private static HashSet<String> RELATIONAL_OPS = new HashSet<string>(new string[] {
            "<",
            ">",
            "=",
            "<=",
            ">=",
            "<>",
            "Is",
            "IsNot",
            "Like" });

        private static HashSet<String> ARITHMETIC_OPS = new HashSet<string>(new string[] {
            "+",
            "-",
            "*",
            "/",
            "\\",
            "Mod",
            "&",
            ">>",
            "<<" });

        private static int MAX_PEEK_LIMIT = 1000;

        private  GlobalState globalState;

        private String original;
        private String converted;
        private int peekCount;
        private String comment;

        private String tokenType;
        private String token;

        public Line(GlobalState globalState) {
            this.globalState = globalState;
        }

        public Line() :this(null) {
            
        }

        /**
         * Step over expected token.
         */
        public void eat(String expected) {
            String token = getToken(true);
        }

        /**
         * Returns a balanced-paren sequence of tokens. Called with ( as peek token.
         * Includes the parens in result. Tries to convert array(i) to array[i].
         */
        // Why is this different from exprlist?
        // These items don't have to be expressions. Not clear the separation is necessary, however.
        public String getBalancedParentheses() {
            StringBuilder balanced = new StringBuilder(getToken(true));
            while (!peek().Equals(")") && !peek().IsEmpty()) {
                if (peek().Equals("(")) {
                    balanced.Append(getBalancedParentheses());
                } else if (peek().Equals(".")) {
                    balanced.Append(globalState.getWithName()).Append(getToken(true)).Append(getName());
                } else if (tokenType.Equals("ID")) {
                    String name = getName();
                    balanced.Append(name);
                    if (globalState.isArrayName(name) && peek().Equals("(")) {
                        balanced.Append(setBrackets(getBalancedParentheses()));
                    }
                } else {
                    balanced.Append(ConverterUtil.fixOperators(getToken(true)));
                }
            }
            balanced.Append(getToken(true)); // adds terminating )
            return balanced.ToString();
        }

        /**
         * Tacked onto output lines in gen().
         */
        public String getComment() {
            return comment;
        }

        /**
         * Returns current token.
         */
        public String getCurrentToken() {
            return token;
        }

        /**
         * Returns next expression from input. Handles .names, nested constructs etc.
         * Drops whitespace after , This has 6-7 levels of precedence. From the
         * bottom:
         * :=, logical ops, negation, relational ops, arith, unary, expon.
         * This isn't complete but it's simpler; assumes that the input is already
         * sensibly parenthesized so it doesn't generate spurious parens.
         */
        public String getExpression() {
            String expression = getArg();
            if (peek().Equals(":=")) { // named argument
                getToken(true);
                expression = "\"" + expression + " :=\", " + getLogic();
            }
            return expression;
        }

        /**
         * Returns whatever is left of the line.
         */
        public String getLine() {
            return converted.Trim();
        }

        /**
         * Returns next name from input, with . expanded, () => [], etc.
         */
        public String getName() {
            if (peek().Equals(".")) {
                return globalState.getWithName() + getToken(true) + getName();
            }
            if (!tokenType.Equals("ID")) {
                return "";
            }
            StringBuilder name = new StringBuilder(getToken(true));
            if (peek().Equals("(")) { // e.g., Range("A3")
                String expressions = getExpressionList();
                if (globalState.isArrayName(name.ToString())) {
                    expressions = setBrackets(expressions);
                }
                name.Append(expressions);
            }
            if (peek().Equals("(")) { // e.g., Range("A1")(cnt)...
                name.Append(getExpressionList());
            }
            while (peek().Equals(".")) { // e.g., Range("A3").Selection.Cells(1,j)
                name.Append(getToken(true));
                name.Append(getName());
            }
            return name.ToString();
        }

        /**
         * Returns the trimmed original input.
         */
        public String getOriginal() {
            return original.Trim();
        }

        /**
         * Returns whatever remains of the current input line.
         */
        // Perhaps should do get_expr or the like to handle array subscripting?
        public String getRest() {
            StringBuilder rest = new StringBuilder();
            while (!peek().IsEmpty() && !peek().Equals(ConverterUtil.EOF)) {
                rest.Append(ConverterUtil.fixOperators(getToken(true)));
            }
            return rest.ToString();
        }

        /**
         * Returns next token.
         */
        // Not sure that string process is right yet, since a quoted string is
        // canonicalized by constructor but found here by a simple RE that doesn't
        // handle \" within a string.
        public String getToken(bool  advance) {
            if (original.Trim().Equals(ConverterUtil.EOF)) {
                return ConverterUtil.EOF;
            }

            converted = converted.Trim();
            foreach (var type in TYPES.Keys) {
                var m = type.Match(converted);
                if (m.Success) {
                    tokenType = TYPES[type];
                    token = converted.Substring(m.Index,m.Length); // the matching part

                    if (tokenType.Equals("TOSS")) {
                        converted = converted.Substring(token.Length); // left for next time
                        continue;
                    }
                    if (tokenType.Equals("STR")) { // re for strings isn't right so clean up
                        token = getStr(converted);
                    }
                    if (tokenType.Equals("DATE")) { // replace # by "
                        token = "\"" + token.Substring(1, token.Length) + "\"";
                    }
                    if (tokenType.Equals("HEX")) {
                        token = token.ReplaceFirst("&H", "0x");
                    }
                    if (token.Equals("!")) { // maybe too exuberant?
                        token = ".";
                    }
                    if (advance) {
                        converted = converted.Substring(token.Length); // left for next time
                    }
                    if (tokenType.Equals("NUM")) { // get rid of vb type indicator
                        token = token.ReplaceFirst("[&#]$", "");
                    }

                    return toUpperCase(token);
                }
            }

            throw new ParseException("Unknown token, can't parse: " + converted);
        }

        public bool  hasComment() {
            return !getComment().IsEmpty();
        }

        public bool hasToken() {
            return !getCurrentToken().IsEmpty();
        }

        /**
         * Tries to isolate a comment if any, while partially coping with horrors like
         * single quotes inside double, quotes in comments, etc.
         */
        public Line parseLine(String line) {
            this.original = line;
            this.peekCount = 0;
            this.comment = "";
            this.converted = "";
            this.tokenType = "";

            while (!line.IsEmpty()) {
                char first = line[0];
                if (first == '\'') {
                    comment = line.Substring(1);
                    break;
                } else if (first == '"') {
                    // getstring returns the quoted string and the residue as a tuple
                    // in java, getstring can append to _str and return residue
                    line = getString(line);
                } else if (first == '[') {
                    // getbrack returns the quoted string and the residue as a tuple
                    // in java, getbrack can append to _str and return residue
                    line = getBracketed(line);
                } else {
                    converted += first;
                    line = line.Substring(1);
                }
            }
            converted = canonicalize(converted.Trim());

            return this;
        }

        /**
         * Returns next token without consuming it.
         */
        public String peek() {
            if (original.Trim().Equals(ConverterUtil.EOF)) {
                return ConverterUtil.EOF;
            }

            // kludge to detect potential bad input
            // alternative is to decorate all calls with "and cur.peek() != EOF"
            ++peekCount;
            if (peekCount > MAX_PEEK_LIMIT) {
                throw new ParseException("Looping because of illegal input: " + original);
            }

            return getToken(false);
        }

        /**
         * Returns the type of the next token. (Assumes that peek() has just been
         * called).
         */
        public String peekTokenType() {
            return tokenType;
        }

        /**
         * Add outer parens if !s appears to need them.
         */
        private String addParen(String str) {
            if (str.matches(".*[-+*/%^<>=!&|].*")) { // watch out: needs unanchored
                return "(" + str + ")";
            } else {
                return str;
            }
        }

        /**
         * Canonicalize some lexical stuff, like Public, that will simplify
         * subsequent processing.
         */
        private String canonicalize(String str) {
            str = str.Replace("Property Get ", "Function Get")
                      .Replace("Property Let ", "Function Let")
                      .Replace("Property Set ", "Function Set")
                      .Replace("End Property", "End Function");
            str = Regex.Replace(str, "(Public|Private|Friend) +Sub", "Sub");
            str = Regex.Replace(str, "(Public|Private|Friend) +Function", "Function");
            str = Regex.Replace(str, "(Public|Private|Friend) +Dim", "Dim");
            str = Regex.Replace(str, "(Public|Private|Friend) +Global", "Global");
            str = Regex.Replace(str, "(Public|Private|Friend|Global) +Const", "Const");
            str = Regex.Replace(str, "(Public|Private|Friend) +Declare", "Declare");
            str = Regex.Replace(str, "(Public|Private|Static)", "Dim"); 
            return str;
        }

        /**
         * Returns next expression from input. Handles .names, nested constructs etc.
         * Drops whitespace after ,. This has about 6 levels of precedence, from the
         * bottom:
         * logical operators, negation, comparison, arith, unary, expon.
         * This isn't complete but it's simpler; assumes that the input is already
         * sensibly parenthesized.
         */
        private String getArg() {
            StringBuilder arg = new StringBuilder(getLogic());
            while (LOGICAL_OPS.Contains(peek())) {
                String op = ConverterUtil.fixOperators(getToken(true));
                arg.Append(op).Append(getLogic());
            }
            return arg.ToString();
        }

        private String getArithmeticOp() {
            String op = getFactor();
            while ("^".Equals(peek())) {
                getToken(true);
                op = "exp(" + op + ", " + getArithmeticOp() + ")";
            }
            return op;
        }

        /**
         * Collect [...]
         */
        private String getBracketed(String str) {
            StringBuilder inside = new StringBuilder();
            String bracketed = str.Substring(1);
            while (true) {
                char first = bracketed[0];
                if (first == ']') {  // the end
                    bracketed = bracketed.Substring(1);
                    break;
                } else if (first == '!') {
                    inside.Append(".");
                    bracketed = bracketed.Substring(1);
                } else {
                    inside.Append(first);
                    bracketed = bracketed.Substring(1);
                }
            }
            converted += "Range(\"" + inside.ToString() + "\")";
            return bracketed;
        }

        private String getCompare() {
            StringBuilder expr = new StringBuilder(getUnary());
            while (ARITHMETIC_OPS.Contains(peek())) {
                String op = ConverterUtil.fixOperators(getToken(true));
                expr.Append(op).Append(getUnary());
            }
            return expr.ToString();
        }

        /**
         * Returns a list of expressions. Called with ( as peek token. Includes the
         * parens in result. Tries to convert array(i) to array[i]. Flags empty
         * exprs.
         */
        // Note: Logic is too convoluted: getFactor() should be smarter.
        private String getExpressionList() {
            StringBuilder expressions = new StringBuilder(getToken(true)); // "("
            while (!peek().Equals(")") && !peek().IsEmpty()) {
                if (peek().Equals(",")) { // empty expr
                    expressions.Append(ConverterUtil.EMPTY).Append(getToken(true)).Append(" ");
                    if (peek().Equals(")")) { // empty expr
                        expressions.Append(ConverterUtil.EMPTY);
                    }
                    continue;
                }
                expressions.Append(getExpression());
                if (peek().Equals(",")) {
                    expressions.Append(getToken(true)).Append(" ");
                    if (peek().Equals(")")) { // empty expr
                        expressions.Append(ConverterUtil.EMPTY);
                    }
                }
            }
            expressions.Append(getToken(true)); // adds terminating )
            return expressions.ToString();
        }

        /**
         * Returns single entity -- number, name, or (expr). This also returns things
         * like comma, which is a botch.
         */
        private String getFactor() {
            StringBuilder expr = new StringBuilder();
            String peeks = peek();
            if (tokenType.Equals("ID")) {
                String name = getName();
                expr.Append(name);
                if (globalState.isArrayName(name) && peek().Equals("(")) {
                    String bp = getBalancedParentheses();
                    expr.Append(setBrackets(bp));
                }
            } else if (tokenType.Equals("NUM")) {
                expr.Append(getToken(true));
            } else if (tokenType.Equals("STR")) {
                expr.Append(getToken(true));
            } else if (peeks.Equals(".")) { // .name
                expr.Append(globalState.getWithName()).Append(getToken(true)).Append(getName());
            } else if (peeks.Equals("Not")) { // BUG?
                expr.Append(getLogic());
            } else if (peeks.Equals("(")) {
                expr.Append(getToken(true)).Append(getExpression()).Append(getToken(true));
            } else {
                expr.Append(getToken(true));
            }
            return expr.ToString();
        }

        private String getLogic() {
            StringBuilder expr;
            if (!peek().Equals("Not")) {
                expr = new StringBuilder(getNotOp());
            } else {
                expr = new StringBuilder();
            }
            while (peek().Equals("Not")) {
                String op = ConverterUtil.fixOperators(getToken(true));
                expr.Append(op).Append(addParen(getLogic()));
            }
            return expr.ToString();
        }

        private String getNotOp() {
            String expr = getCompare();
            while (RELATIONAL_OPS.Contains(peek())) {
                String op = ConverterUtil.fixOperators(getToken(true));
                if (op.Equals("Like")) {
                    expr = "Like(" + expr + "," + getCompare() + ")";
                } else {
                    expr += op + getCompare();
                }
            }
            return expr;
        }

        // TODO(nikhil): Rename getStr() and getString()
        /**
         * Returns the real string, skipping embedded \"'s.
         */
        private String getStr(String str) {
            int i = 1;
            while (i < str.Length) {
                if (str.Substring(i, 1).Equals("\"")) {
                    break;
                }
                if (str.Substring(i,1).Equals("\\")) {
                    ++i;
                }
                ++i;
            }
            return str.Substring(0, 1);
        }

        /**
         * Collect quoted string, handle "" and \.
         */
        private String getString(String str) {
            StringBuilder parsed = new StringBuilder(str.Substring(0, 1)); // the " at the front
            String input = str.Substring(1);
            while (true) {
                if (input.Substring(0, 1).Equals("\\")) {
                    parsed.Append("\\").Append(input.Substring(0, 1));
                    input = input.Substring(1);
                } else if (input.Substring(0, 1).Equals("\"") &&
                    (input.Length > 1) &&
                    input.Substring(1, 2).Equals("\"")) {
                    parsed.Append("\\\"");
                    input = input.Substring(2);
                } else if (input.Substring(0, 1).Equals("\"")) {
                    parsed.Append("\"");
                    input = input.Substring(1);
                    break;
                } else {
                    parsed.Append(input.Substring(0, 1));
                    input = input.Substring(1);
                }
            }
            converted += parsed.ToString();
            return input;
        }

        private String getUnary() {
            String op = "";
            while (peek().Equals("+") || peek().Equals("-")) {
                op += getToken(true);
            }
            String expr = getArithmeticOp();
            expr = op + expr;
            return expr;
        }

        /**
         * Set brackets in s to convert from (...) to [...]""".
         */
        // Note: This is probably too aggressive: won't work if there are nested
        // commas. e.g., in function calls in subscripts, or in strings.
        private String setBrackets(String str) {
            String input = str.Substring(1, str.Length);
            if (input.IndexOf('(') == -1) {
                input = input.Replace(", *", "][");
            }
            return "[" + input + "]";
        }

        /**
         * Canonicalizes the case of a likely keyword.
         */
        private String toUpperCase(String str) {
            String lowerCase = str.ToLower();
            if (KEYWORDS.ContainsKey(lowerCase)) {
                return KEYWORDS[lowerCase];
            } else {
                return str;
            }
        }
    }
}
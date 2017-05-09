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
using System.Text.RegularExpressions;

namespace com.google.vb2js {



    /**
     * Helper/utility class for the converter. Contains some constants, data structures and
     * static functions.
     *
     * @author Brian Kernighan
     * @author Nikhil Singhal
     */
    static class ConverterUtil {

        // Note: The order is relevant here.
        /**
         * Map to convert between VB and JS operators.
         */
        private static Dictionary<String, String> POSSIBLE_FIXES =
              new Dictionary<string, string> {
            {"=", " == "},
            {"<>", " != "},
            {"<=", " <= "},
            {">=", " >= "},
            {"<", " < "},
            {">", " > "},
            {"&", " + "},
            {"\\+", " + "},
            {"-", " - "},
            {"\\*", " * "},
            {"/", " / "},
            {"\\\\", " / "},
            {"\\^", " BUG exp(), "},
            {"\\bXor\\b", " ^ "},
            {"\\bAnd\\b", " && "},
            {"\\bOr\\b", " || "},
            {"\\bIs\\b", " == "},
            {"\\bIsNot\\b", " != "},
            {"\\bMod\\b", " % "},
            {"\\bNew\\b", "new "},
            {"\\bNot\\b", "!"}
            };

        private static  String ONE_LINE_IF_THEN_ELSE = "(?i).*Then .+ Else .*";
  private static  String ONE_LINE_IF_THEN = "(?i).*Then .+";

        /** Marker for after last line */
        public static String EOF = "(EOF)";

        /** Cross-platform line separator */
        public static String LINE_SEPARATOR = Environment.NewLine;

  /** Value generated for non-existent arguments */
  public static  String EMPTY = "undefined";



        /**
         * Replace VB operators by JS. don't apply to 'strings.'
         * ^ * / \ mod + - & = <> <= >= := >
         * < ! is not and or xor eqv imp like
         *
         * Precedence is mostly implemented
         */
        public static String fixOperators(String token) {
            foreach (var fix in POSSIBLE_FIXES) {
                if (Regex.IsMatch(token,fix.Key)) {
                    return fix.Value;
                }
            }
            return token;
        }

        /**
         * Test whether line is one-line: If ... Then ... [Else ...].
         */
        public static bool  isOneLineIf(String line) {
            return (Regex.IsMatch(line,ONE_LINE_IF_THEN_ELSE) ||
                new Line().parseLine(line).getLine().matches(ONE_LINE_IF_THEN));
        }
    }
}
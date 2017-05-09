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
     * This class contains basic data about a VB file.
     *
     * @author Brian Kernighan
     * @author Nikhil Singhal
     */
     public class TranslationUnit {

        private static  Regex CONTINUATION_PATTERN =  new Regex("(.*)_$");

  /** Current line, in a Line object */
  private Line currentLine;

  private GlobalState globalState;

  /** All input lines as is, no \n. Fixed up in convert() */
  private List<String> lines;

        /** Current line number. Advance happens first, so start at -1 */
        private int currentLineNumber;

        /** Depth of nested constructs */
        private int depth;

        /** Name of function currently being translated */
        private String functionName;

        private int subNestingValue;

        /**
         * User-defined type names. These are used when variables to these types are defined.
         * For normal variables, the type is just erased (Dim x as String => var x; // String), but for
         * types/classes this is not enough, since they will have some variables already bound to them.
         * Those are therefore translated as: Dim x as MyType => var x = new MyType();.
         */
        private HashSet<String> typeNames;

        public TranslationUnit() {
            this.globalState = new GlobalState();
            this.currentLine = new Line(globalState);
            this.currentLineNumber = -1;
            this.lines = new List<string>();

            this.depth = 0;
            this.functionName = "";
            this.subNestingValue = 0;

            this.typeNames = new HashSet<string>();
        }

        public void cleanup(IEnumerable<String> vba) {
            foreach (String line in vba) {
                if (line != null) {
                    lines.Add(line.Trim());
                }
            }

            // Merge continuation lines (ending with _) into one long one
            for (int i = lines.Count - 1; i >= 0; --i) {
                String line = lines[i];
                var continuationMatcher = CONTINUATION_PATTERN.Match(line);
                if (continuationMatcher.Success) {
                    line = continuationMatcher.Groups[1] + lines[i + 1];
                    lines[i] =line;
                    lines.RemoveAt(i + 1);
                }
            }

            // Convert 1-line If's into multi-line
            for (int i = lines.Count - 1; i >= 0; --i) {
                if (ConverterUtil.isOneLineIf(lines[i])) {
                    rewriteOneLineIf(i);
                }
            }

            lines.Add(ConverterUtil.EOF);
        }

        public void addGlobalName(String name) {
            globalState.addGlobalName(name);
        }

        public void addLocalName(String name) {
            globalState.addLocalName(name);
        }

        public void addWithName(String name) {
            globalState.addWithName(name);
        }

        public bool isTypeName(String name) {
            return typeNames.Contains(name);
        }

        public void addTypeName(String name) {
            typeNames.Add(name);
        }

        /**
         * Advance to the next line in lines[]
         */
        public void advance() {
            ++currentLineNumber;
            if (currentLineNumber < lines.Count) {
                currentLine.parseLine(lines[currentLineNumber]);
            }
        }

        /**
         * Entered a Sub/Function.
         */
        public void enterSub() {
            ++subNestingValue;
        }

        /**
         * Left a Sub/Function
         */
        public void leaveSub() {
            --subNestingValue;
            if (subNestingValue == 0) {
                globalState.clearLocalNames();
            }
        }

        public Line getCurrentLine() {
            return currentLine;
        }

        public int getDepth() {
            return depth;
        }

        public String getFunctionName() {
            return functionName;
        }

        public String getLine(int lineNumber) {
            return lines[lineNumber];
        }

        public String getCurrentLineAsString() {
            return currentLine.getLine();
        }

        public int getCurrentLineNumber() {
            return currentLineNumber;
        }

        public int getSubNestingValue() {
            return subNestingValue;
        }

        public String getWithName() {
            return globalState.getWithName();
        }

        public void indent() {
            ++depth;
        }

        public void undent() {
            --depth;
        }

        public bool  isArrayName(String name) {
            return globalState.isArrayName(name);
        }

        public void popWithName() {
            globalState.popWithName();
        }

        public void setFunctionName(String functionName) {
            this.functionName = functionName;
        }

        /**
         * Convert If ... Then ... [Else ...] on one line into multiple lines so translateIf() can
         * handle it. Note: this needs to be case-independent if one is going down that path. Probably
         * should be done with re.I as an argument so it can be adjusted at run time rather than being
         * wired in. (but there's no "do nothing" 3rd arg to re.sub). Not coordinated with
         * Line.toUpperCase(), which tests whether to do conversion.
        */
        private void rewriteOneLineIf(int lineNumber) {
            String original = lines[lineNumber];
            String thenPart;
            String elsePart;
            int where;

            lines[lineNumber]= original.ReplaceFirst("(?i)Then .*", "Then"); // if part
            thenPart = original.ReplaceFirst("(?i).*Then ", "");
            thenPart = thenPart.ReplaceFirst("(?i)Else .*", "").Trim();
            where = lineNumber + 1;
            lines.Insert( where, thenPart );
            if (original.matches("(?i).*Else .+")) {
                elsePart = original.ReplaceFirst("(?i).*Else ", "").Trim();
                where += 1;
                lines.Insert(where, "Else");
                where += 1;
                lines.Insert(where, elsePart);
            }
            lines.Insert(where + 1, "End If");
        }
    }
}
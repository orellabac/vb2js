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

namespace com.google.vb2js {



    /**
     * Simple class to keep track of the global state of the VBA code being
     * converted.
     *
     * @author Brian Kernighan
     * @author Nikhil Singhal
     */
     public class GlobalState {

        /** Stack of names in With */
        private  Stack<String> withNames;

        /** Names of global vars */
        private HashSet<String> globalNames;

        /** Names of local vars */
        private  HashSet<String> localNames;

        public GlobalState() {
            withNames = new Stack<String>();

            globalNames = new HashSet<string>();
            localNames = new HashSet<string>();
        }

        public void addGlobalName(String name) {
            globalNames.Add(name);
        }

        public void addLocalName(String name) {
            localNames.Add(name);
        }

        /** Add a With name to the stack */
        public void addWithName(String name) {
            withNames.Push(name);
        }

        public void clearLocalNames() {
            localNames.Clear();
        }

        /** Get the current With name (ie, the top of the stack) */
        public String getWithName() {
            return withNames.Peek();
        }

        /** Remove the latest With name from the stack (ie, pop) */
        public void popWithName() {
            withNames.Pop();
        }

        public bool  isArrayName(String name) {
            return (localNames.Contains(name) || globalNames.Contains(name));
        }
    }
}
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

namespace com.google.vb2js { 

/**
 * Custom class for exceptions thrown by VbaJsConverter
 *
 * @author Brian Kernighan
 * @author Nikhil Singhal
 */

 class ParseException : Exception {

  /** Line number in source VBA file where exception occurred */
  private  int lineNumber;

  /** Contents of line where error occurred */
  private  String line;

  public ParseException(String message, int lineNumber, String line):base(message) {

    this.lineNumber = lineNumber;
    this.line = line;
  }

  public ParseException(String message):this(message,-1,null) {
    
  }

  int getLineNumber() {
    return lineNumber;
  }

  String getLine() {
    return line;
  }

  public override String ToString() {
            String s = this.Message;

    if (lineNumber != -1) {
      s += " at line " + lineNumber;
    }

    if (line != null) {
      s += " (" + line + ")";
    }

    return s;
  }

  public String getMessage() {
    return ToString();
  }
}
}
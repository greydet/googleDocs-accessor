#!/usr/bin/python
#
# Copyright (C) 2011 Gonzague Reydet.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import getopt
import sys

import spreadsheet

def printUsageAndExit():
    print 'python dump.py --user [username] -pw [password]'
    sys.exit(2)

# Parse command line options
try:
    opts, args = getopt.getopt(sys.argv[1:], "s:w:", ["user=", "pw="])
except getopt.error, msg:
    printUsageAndExit()

user = ''
pw = ''
s = ''
w = ''

#Process options
for o, a in opts:
    if o == "--user":
        user = a
    elif o == "--pw":
        pw = a
    elif o == "-s":
        s = a
    elif o == "-w":
        w = a

if user == '' or pw == '':
    printUsageAndExit()

client = spreadsheet.SpreadsheetClient(user, pw)

if s == '':
    while s == '':
        client.ListSpreadsheets()
        input = raw_input('\nSelection: ')
        if client.GetSpreadsheet(input) != None:
            s = input
        print ''

if w == '':
    sprdSht = client.GetSpreadsheet(s)
    while w == '':
        sprdSht.ListWorksheet()
        input = raw_input('\nSelection: ')
        if sprdSht.GetWorksheet(input) != None:
            w = input
        print ''

client.GetSpreadsheet(s).GetWorksheet(w).ListRows()


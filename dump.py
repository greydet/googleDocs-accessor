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
    print 'python dump.py --user <username> --pw <password> [-s <spreadsheet>] [-w <worksheet>] [-a <action>]'
    sys.exit(2)
#def

# Parse command line options
try:
    opts, args = getopt.getopt(sys.argv[1:], "s:w:a:", ["user=", "pw="])
except getopt.error, msg:
    printUsageAndExit()

user = ''
pw = ''
s = ''
w = ''
act = ''

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
    elif o == "-a":
        act = a
    #if
#for

if user == '' or pw == '':
    printUsageAndExit()
#if

client = spreadsheet.SpreadsheetClient(user, pw)

if s == '':
    while s == '':
        client.ListSpreadsheets()
        input = raw_input('\nSelection: ')
        if client.GetSpreadsheet(input) != None:
            s = input
        #if
        print ''
    #while
#if

if w == '':
    sprdSht = client.GetSpreadsheet(s)
    while w == '':
        sprdSht.ListWorksheet()
        input = raw_input('\nSelection: ')
        if sprdSht.GetWorksheet(input) != None:
            w = input
        #if
        print ''
    #while
#if

while act == '':
    input = raw_input('\nAction: ')
    args = input.split(' ')
    if hasattr(spreadsheet.Worksheet, args[0]):
        act = args[0]
        args.pop(0)
    #if
#while

if act == 'ListRows':
    client.GetSpreadsheet(s).GetWorksheet(w).ListRows()
elif act == 'PrintRow':
    if len(args) == 1:
        client.GetSpreadsheet(s).GetWorksheet(w).PrintRow(int(args[0]))
    else:
        print 'Error: Missing row index argument for the PrintRow action'
    #if
elif act == 'PrintLastRow':
    client.GetSpreadsheet(s).GetWorksheet(w).PrintLastRow()
elif act == 'PrintCells':
    if len(args) == 4:
        client.GetSpreadsheet(s).GetWorksheet(w).PrintCells(args[0], args[1], args[2], args[3])
    else:
        print 'Usage: PrintCells(column, firstRow, lastRow)'
    #if
#if


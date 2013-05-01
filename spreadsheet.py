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

import gdata.spreadsheet.service

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
#def

class Worksheet:
    def __init__(self, client, spreadsheet, feedEntry):
        self.feedEntry = feedEntry
        self.client = client
        self.spreadsheet = spreadsheet
        self.key = client._GetKeyFromEntry(feedEntry)
    #def

    def ListRows(self):
        feed = self.client.gdata.GetListFeed(self.spreadsheet.key, self.key)
        self.client._PrintFeed(feed)
    #def

    def PrintRow(self, i):
        feed = self.client.gdata.GetListFeed(self.spreadsheet.key, self.key)
        print '%s %s' % (feed.entry[i].title.text, feed.entry[i].content.text)
    #def

    def PrintLastRow(self):
        feed = self.client.gdata.GetListFeed(self.spreadsheet.key, self.key)
        nbEntries = len(feed.entry)
        i = nbEntries - 1
        print '%s %s' % (feed.entry[i].title.text, feed.entry[i].content.text)
    #def

    def PrintCells(self, firstCol, lastCol, firstRow, lastRow):
       self.client._PrintFeed(self.GetRecords(firstCol, lastCol, firstRow, lastRow))
    #def

    def GetRecords(self, firstCol, lastCol, firstRow, lastRow):
        if int(lastCol) == -1:
            lastCol = self.feedEntry.col_count.text
        #if
        if int(lastRow) == -1:
            lastRow = self.feedEntry.row_count.text
        #if
        query = gdata.spreadsheet.service.CellQuery()
        query.return_empty = 'true'
        query.min_col = firstCol
        query.max_col = lastCol
        query.min_row = firstRow
        query.max_row = lastRow
        return self.client.gdata.GetCellsFeed(self.spreadsheet.key, self.key, query=query)
    #def
#class

class Spreadsheet:
    def __init__(self, client, feedEntry):
        self.feedEntry = feedEntry
        self.client = client
        self.key = client._GetKeyFromEntry(feedEntry)
    #def

    def ListWorksheet(self):
        feed = self.client.gdata.GetWorksheetsFeed(self.key)
        print 'Worksheets of "%s" spreadsheet:\n' % self.feedEntry.title.text
        self.client._PrintFeed(feed)
    #def

    def GetWorksheet(self, title):
        feed = self.client.gdata.GetWorksheetsFeed(self.key)
        worksheetEntry = self.client._GetEntry(title, feed)
        if worksheetEntry != None:
            return Worksheet(self.client, self, worksheetEntry)
        #if
        return None
    #def
#class

class SpreadsheetClient:
    def __init__(self, email, password):
        self.gdata = gdata.spreadsheet.service.SpreadsheetsService()
        self.gdata.email = email
        self.gdata.password = password
        self.gdata.ProgrammaticLogin()
    #def
    
    def ListSpreadsheets(self):
        feed = self.gdata.GetSpreadsheetsFeed()
        print 'Spreadsheets:'
        self._PrintFeed(feed)
    #def

    def GetSpreadsheet(self, title):
        feed = self.gdata.GetSpreadsheetsFeed()
        spreadsheetEntry = self._GetEntry(title, feed)
        if spreadsheetEntry == None:
            return None
        #if
        return Spreadsheet(self, spreadsheetEntry) 
    #def

    def _GetEntry(self, title, feed):
        entries = self._GetEntries(title, feed)
        if len(entries) > 1:
            print 'Error: Found multiple entries with title: %s' % (title)
            return None
        #if
        if len(entries) == 0:
            print 'Error: No entry found with title: %s' % (title)
            return None
        #if
        return entries[0]
    #def

    def _GetEntries(self, title, feed):
        foundEntries = []
        if is_number(title):
            i = int(title)
            foundEntries.append(feed.entry[i])
        else:
            for i, entry in enumerate(feed.entry):
                if entry.title.text == title:
                    foundEntries.append(entry)
                #if
            #for
        #if
        return foundEntries
    #def

    def _GetKeyFromEntry(self, entry):
        idSegments = entry.id.text.split('/')
        return idSegments[len(idSegments) - 1]
    #def

    def _PrintFeed(self, feed):
        for i, entry in enumerate(feed.entry):
            if isinstance(feed, gdata.spreadsheet.SpreadsheetsCellsFeed):
                print '%s %s\n' % (entry.title.text, entry.content.text)
            elif isinstance(feed, gdata.spreadsheet.SpreadsheetsListFeed):
                print '%s %s %s' % (i, entry.title.text, entry.content.text)
            else:
                print '%s %s' % (i, entry.title.text)
            #if
        #for
    #def
#class


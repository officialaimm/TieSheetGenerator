from itertools import combinations
from random import shuffle
from math import ceil
import xlsxwriter

class TieSheet:
    def __init__(self,playersEachGame,scoreRule,preferredGameEachDay):
        self.playersEachGame = playersEachGame
        self.scoreRule = scoreRule
        self.preferredGameEachDay = preferredGameEachDay
        #CONSTANTS
        #
        #FIXTURE_CONSTANT
        self.FIXTURES_TITLE = "Fixtures"
        self.FIXTURES_INDEX = "Fixture"
        self.FIXTURES_PLAYERS = "Participants"
        self.FIXTURES_OUTSIDERS = "Outsiders"
        self.FIXTURES_POINTS = "Points"
        self.FIXTURES_DAY = "Day"
        #TABLE_CONSTANT
        self.TABLE_TITLE = "Table"
        #RULE_CONSTANT
        self.RULES_TITLE = "Rules"
    def generateFixtures(self):
        sheet = self.xlsxFile.add_worksheet(self.FIXTURES_TITLE)
        fixtures = list(combinations(self.participants,self.playersEachGame))
        shuffle(fixtures)
        row = 0
        sheet.write_row(row,0,[
            self.FIXTURES_INDEX,
            self.FIXTURES_PLAYERS,
            self.FIXTURES_OUTSIDERS,
            self.FIXTURES_POINTS,
            self.FIXTURES_DAY
        ])
        #empty row
        sheet.write_row(row,0,[])
        row+=1
        for index,fixture in enumerate(fixtures):
            #1-indexed
            index = index+1
            days = ceil(index/self.preferredGameEachDay)
            #outsiders
            outsiders = list(filter(lambda participant:participant not in fixture,self.participants))
            #empty row
            sheet.write_row(row,0,[])
            row+=1
            #indexed row
            for playerIndex,player in enumerate(fixture):
                sheet.write_row(row,0,[
                    playerIndex==0 and index or '',
                    player,
                    playerIndex<len(outsiders) and outsiders[playerIndex] or '',
                    '',
                    playerIndex==0 and days or ''
                ])
                row+=1
    def generateTable(self):
        sheet = self.xlsxFile.add_worksheet(self.TABLE_TITLE)
    def generateRules(self):
        sheet = self.xlsxFile.add_worksheet(self.RULES_TITLE)
    def generate(self,participants,filename):
        self.participants = participants
        #create xlsx file with given filename
        self.xlsxFile = xlsxwriter.Workbook(filename)
        #start generating
        self.generateFixtures()
        self.generateTable()
        self.generateRules()
        #close file
        self.xlsxFile.close()
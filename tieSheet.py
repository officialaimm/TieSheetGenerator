from itertools import combinations
from random import shuffle
from math import ceil
import xlsxwriter

class TieSheet:
    def __init__(self,playersEachGame,winnerScore,preferredGameEachDay,scoreRule={}):
        self.playersEachGame = playersEachGame
        self.winnerScore = winnerScore
        self.preferredGameEachDay = preferredGameEachDay
        self.scoreRule = scoreRule
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
        self.TABLE_INDEX = "S.N."
        self.TABLE_PLAYER = "Participant"
        self.TABLE_PLAYED = "Games Played"
        self.TABLE_POINTS = "Points"
        self.TABLE_WINS = "Wins"
        #RULE_CONSTANT
        self.RULES_TITLE = "Rules"
    def generateFixture(self,participants,playersEachGame):
        fixtures = list(combinations(participants,playersEachGame))
        shuffle(fixtures)
        return fixtures
    def generateFixtureSheet(self):
        sheet = self.xlsxFile.add_worksheet(self.FIXTURES_TITLE)
        fixtures = self.generateFixture(self.participants,self.playersEachGame)
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
        #will be handy for formula in table sheet
        self.rowWithPointsStart = row+1
        self.playerCol = "B"
        self.pointsCol = "D"
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
        #will be handy for formula in table sheet
        self.rowWithPointsEnd = row
    def generateTable(self):
        sheet = self.xlsxFile.add_worksheet(self.TABLE_TITLE)
        row = 0
        sheet.write_row(row,0,[
            self.TABLE_INDEX,
            self.TABLE_PLAYER,
            self.TABLE_PLAYED,
            self.TABLE_POINTS,
            self.TABLE_WINS
        ])
        #empty row
        sheet.write_row(row,0,[])
        row+=1
        for index,participant in enumerate(self.participants):
            sheet.write_row(row,0,[
                index+1,
                participant,
                #matches played
                "=COUNT(FILTER({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd},{fixtureSheet}!{fixturePlayerCol}{fixtureRowStart}:{fixturePlayerCol}{fixtureRowEnd}={column}{row}))"
                .format(
                    fixtureSheet=self.FIXTURES_TITLE,
                    fixturePointCol=self.pointsCol,
                    fixturePlayerCol=self.playerCol,
                    fixtureRowStart=self.rowWithPointsStart,
                    fixtureRowEnd=self.rowWithPointsEnd,
                    column="B",
                    row=row+1
                ),
                #points
                "=SUM(FILTER({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd},{fixtureSheet}!{fixturePlayerCol}{fixtureRowStart}:{fixturePlayerCol}{fixtureRowEnd}={column}{row}))"
                .format(
                    fixtureSheet=self.FIXTURES_TITLE,
                    fixturePointCol=self.pointsCol,
                    fixturePlayerCol=self.playerCol,
                    fixtureRowStart=self.rowWithPointsStart,
                    fixtureRowEnd=self.rowWithPointsEnd,
                    column="B",
                    row=row+1
                ),
                #wins
                "=COUNT(FILTER({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd},({fixtureSheet}!{fixturePlayerCol}{fixtureRowStart}:{fixturePlayerCol}{fixtureRowEnd}={column}{row}) * ({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd}={score}) ))"
                .format(
                    fixtureSheet=self.FIXTURES_TITLE,
                    fixturePointCol=self.pointsCol,
                    fixturePlayerCol=self.playerCol,
                    fixtureRowStart=self.rowWithPointsStart,
                    fixtureRowEnd=self.rowWithPointsEnd,
                    column="B",
                    row=row+1,
                    score=self.winnerScore
                ),
            ])
            row+=1
    def generateRules(self):
        sheet = self.xlsxFile.add_worksheet(self.RULES_TITLE)
    def generate(self,participants,filename):
        self.participants = participants
        #create xlsx file with given filename
        self.xlsxFile = xlsxwriter.Workbook(filename)
        #start generating
        self.generateFixtureSheet()
        self.generateTable()
        self.generateRules()
        #close file
        self.xlsxFile.close()
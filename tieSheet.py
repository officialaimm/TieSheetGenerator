# dependency modules
import xlsxwriter

# standard modules
from itertools import combinations
from random import shuffle
from math import ceil


class TieSheet:

    def __init__(self, playersEachGame):
        self._playersEachGame = playersEachGame
        # CONSTANTS
        #
        # FIXTURE_CONSTANT
        self._FIXTURES_TITLE = "Fixtures"
        self._FIXTURES_INDEX = "Fixture"
        self._FIXTURES_PLAYERS = "Participants"
        self._FIXTURES_OUTSIDERS = "Outsiders"
        self._FIXTURES_POINTS = "Points"
        self._FIXTURES_DAY = "Day"
        # TABLE_CONSTANT
        self._TABLE_TITLE = "Table"
        self._TABLE_INDEX = "S.N."
        self._TABLE_PLAYER = "Participant"
        self._TABLE_PLAYED = "Games Played"
        self._TABLE_POINTS = "Points"
        self._TABLE_WINS = "Wins"
        # SORTED_TABLE CONSTANT
        self._SORTED_TABLE_TITLE = "Table(Sorted)"
        # RULE_CONSTANT
        self._RULES_TITLE = "Rules"

    def _generateFixture(self):
        fixtures = list(combinations(
            self._participants, self._playersEachGame))
        shuffle(fixtures)
        return fixtures

    def _generateFixtureSheet(self):
        sheet = self._xlsxFile.add_worksheet(self._FIXTURES_TITLE)
        fixtures = self._generateFixture()
        row = 0
        sheet.write_row(row, 0, [
            self._FIXTURES_INDEX,
            self._FIXTURES_PLAYERS,
            self._FIXTURES_OUTSIDERS,
            self._FIXTURES_POINTS,
            self._FIXTURES_DAY
        ])
        # empty row
        sheet.write_row(row, 0, [])
        row += 1
        # will be handy for formula in table sheet
        self._rowWithPointsStart = row+1
        self._playerCol = "B"
        self._pointsCol = "D"
        for index, fixture in enumerate(fixtures):
            # 1-indexed
            index = index+1
            days = ceil(index/self._preferredGameEachDay)
            # outsiders
            outsiders = list(
                filter(lambda participant: participant not in fixture, self._participants))
            # empty row
            sheet.write_row(row, 0, [])
            row += 1
            # indexed row
            for playerIndex, player in enumerate(fixture):
                sheet.write_row(row, 0, [
                    playerIndex == 0 and index or '',
                    player,
                    playerIndex < len(
                        outsiders) and outsiders[playerIndex] or '',
                    '',
                    playerIndex == 0 and days or ''
                ])
                row += 1
        # will be handy for formula in table sheet
        self._rowWithPointsEnd = row

    def _generateTable(self):
        # Add sorted_table sheet before table
        sheetSorted = self._xlsxFile.add_worksheet(self._SORTED_TABLE_TITLE)
        sheet = self._xlsxFile.add_worksheet(self._TABLE_TITLE)
        row = 0
        sheet.write_row(row, 0, [
            self._TABLE_PLAYER,
            self._TABLE_PLAYED,
            self._TABLE_POINTS,
            self._TABLE_WINS
        ])
        # empty row
        sheet.write_row(row, 0, [])
        row += 1
        for index, participant in enumerate(self._participants):
            sheet.write_row(row, 0, [
                participant,
                # matches played
                "=COUNT(FILTER({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd},{fixtureSheet}!{fixturePlayerCol}{fixtureRowStart}:{fixturePlayerCol}{fixtureRowEnd}={column}{row}))"
                .format(
                    fixtureSheet=self._FIXTURES_TITLE,
                    fixturePointCol=self._pointsCol,
                    fixturePlayerCol=self._playerCol,
                    fixtureRowStart=self._rowWithPointsStart,
                    fixtureRowEnd=self._rowWithPointsEnd,
                    column="B",
                    row=row+1
                ),
                # points
                "=SUM(FILTER({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd},{fixtureSheet}!{fixturePlayerCol}{fixtureRowStart}:{fixturePlayerCol}{fixtureRowEnd}={column}{row}))"
                .format(
                    fixtureSheet=self._FIXTURES_TITLE,
                    fixturePointCol=self._pointsCol,
                    fixturePlayerCol=self._playerCol,
                    fixtureRowStart=self._rowWithPointsStart,
                    fixtureRowEnd=self._rowWithPointsEnd,
                    column="B",
                    row=row+1
                ),
                # wins
                "=COUNT(FILTER({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd},({fixtureSheet}!{fixturePlayerCol}{fixtureRowStart}:{fixturePlayerCol}{fixtureRowEnd}={column}{row}) * ({fixtureSheet}!{fixturePointCol}{fixtureRowStart}:{fixturePointCol}{fixtureRowEnd}={score}) ))"
                .format(
                    fixtureSheet=self._FIXTURES_TITLE,
                    fixturePointCol=self._pointsCol,
                    fixturePlayerCol=self._playerCol,
                    fixtureRowStart=self._rowWithPointsStart,
                    fixtureRowEnd=self._rowWithPointsEnd,
                    column="B",
                    row=row+1,
                    score=self._winnerScore
                ),
            ])
            row += 1
        sheetSorted.write_row(0, 0, [
            "=SORT(SORT({tableSheet}!A:Z,{winsColumnNumber},FALSE),{pointsColumnNumber},FALSE)"
            .format(
                tableSheet=self._TABLE_TITLE,
                winsColumnNumber="5",
                pointsColumnNumber="4"
            ),
        ])

    def _generateRules(self):
        sheet = self._xlsxFile.add_worksheet(self._RULES_TITLE)

    def generate(self, winnerScore, preferredGameEachDay, participants, filename):
        self._winnerScore = winnerScore
        self._preferredGameEachDay = preferredGameEachDay
        self._participants = participants
        # create xlsx file with given filename
        self._xlsxFile = xlsxwriter.Workbook(filename)
        # start generating
        self._generateFixtureSheet()
        self._generateTable()
        self._generateRules()
        # close file
        self._xlsxFile.close()

import sys
reload(sys)
sys.setdefaultencoding( "utf-8" ) # since there can easily be UTF8 characters in the file


import cgitb
cgitb.enable(format='text')


def ColsFromCode(c):
    retval = 0
    if len(c) == 0:
        return 0
    if len(c) == 1:
        return (ord(c) - ord('A'))
    if len(c) == 2:
        return (ord(c[0]) - ord('A') + 1) * 26 + (ord(c[1]) - ord('A'))

PenaltyCodes = {
    '4': "Fourth Minor",
    4.0: "Fourth Minor",
    'B': "Back Block",
    'H': "Block w/ Head",
    'X': "Track Cut",
    'C': "Direction of Play",
    'E': "Elbows",
    'F': "Forearms",
    'A': "High Block",
    'I': "Illegal Procedure",
    'L': "Low Block",
    'M': "Multi-Player Block",
    'O': "Out of Bounds Block",
    'P': "Out of Play Block",
    'S': "Skate Out of Bounds",
    'N': "Insubordination",
    'G': "(Gross) Misconduct",
    'Z': "Delay of Game",
    'PM': "Penalty Minutes",
    'FO': "Foul Out"
}

def CanonSkaterNumber(x):
    """Convert a skater number to a unicode string (not an int or float)"""
    if type(x) == type(1.0):
        return unicode(int(x))
    if type(x) == type(1):
        return unicode(x)
    return unicode(x)
EmptyCellsTypes = []
def IsBlankCell(x):
    """Is the cell blank (this will vary based on xlsx vs xls files)"""
    return x in EmptyCellsTypes or x.value == "" or x.value == " " or x.value == None


class StatBookVersion:
    """An abstraction of the different versions of the statbook"""
    def __init__(self, workbook):
        try:
            self.version = workbook.sheet_by_name("Read ME").cell(2,0).value
        except Exception:
            try:
                self.version = workbook.sheet_by_name("Read_Me").cell(2,0).value
            except Exception:
                self.version = workbook.sheet_by_name("Read Me").cell(2,0).value            
        self.maxNumJams = 20
        self.maxNumPasses = 8
        self.hasNoPivot = False
        self.hasMinors = True
        self.periodTimerSheetName = "Jam Timer"
        self.hasGhostPoints = True
        self.headerRows = 1
        self.lineupColumns = {'p':0,'2':3,'3':6,'4':9,'j':12} # p,b2,b3,b4,j
        self.hasColors = False
        self.hasPassesInLineup = True
        self.hasActionsErrors = True
        if self.version == "Official October 2012 Revision":
            self.maxNumJams = 30
            self.maxNumPasses = 10
            self.hasNoPivot = True
            self.periodTimerSheetName = "Bout Clock"
        elif self.version == "Official January 2013 Release":
            self.maxNumJams = 76/2
            self.maxNumPasses = 10
            self.hasNoPivot = True
            self.hasMinors = False
            self.periodTimerSheetName = "Bout Clock"
            self.hasGhostPoints = False # we don't support the ghost point sheets yet
            self.lineupColumns = {'j':0,'p':3,'2':6,'3':9,'4':12} # j,p,b2,b3,b4
            self.headerRows = 2
            self.hasColors
        elif self.version == "Official March 2013 Release":
            self.maxNumJams = 76/2
            self.maxNumPasses = 10
            self.hasNoPivot = True
            self.hasMinors = False
            self.periodTimerSheetName = "Bout Clock"
            self.hasGhostPoints = False # we don't support the ghost point sheets yet
            self.lineupColumns = {'j':0,'p':3,'2':6,'3':9,'4':12} # j,p,b2,b3,b4
            self.headerRows = 2
            self.hasColors
        elif self.version == "Official March 2014 Release":
            self.maxNumJams = 76/2
            self.maxNumPasses = 10
            self.hasNoPivot = True
            self.hasMinors = False
            self.periodTimerSheetName = "Bout Clock"
            self.hasGhostPoints = False # we don't support the ghost point sheets yet
            self.lineupColumns = {'j':0,'p':4,'2':8,'3':12,'4':16} # j,p,b2,b3,b4
            self.headerRows = 2
            self.hasPassesInLineup = False
            self.hasColors
        elif self.version == "Official April 2014 Release":
            self.maxNumJams = 76/2
            self.maxNumPasses = 10
            self.hasNoPivot = True
            self.hasMinors = False
            self.periodTimerSheetName = "Game Clock"
            self.hasGhostPoints = False # we don't support the ghost point sheets yet
            self.lineupColumns = {'j':0,'p':4,'2':8,'3':12,'4':16} # j,p,b2,b3,b4
            self.headerRows = 2
            self.hasPassesInLineup = False
            self.hasColors
        elif self.version == "Official January 2015 Release":
            self.maxNumJams = 76/2
            self.maxNumPasses = 10
            self.hasNoPivot = True
            self.hasMinors = False
            self.periodTimerSheetName = "Game Clock"
            self.hasGhostPoints = False # there aren't even ghost point sheets
            self.lineupColumns = {'j':0,'p':4,'2':8,'3':12,'4':16} # j,p,b2,b3,b4
            self.headerRows = 2
            self.hasPassesInLineup = False
            self.hasActionsErrors = False # no longer supported
            self.hasExtendedInfo = True # extra bout info
    def __str__(self):
        return self.version
    def Config(self, obj):
        cls = obj.__class__
        configs = {
            'Scorekeeper' : {
                'RowsPerPeriodHeader' : 59 - 2 * 20,
                'ColsPerTeam' : 37
            },
            'Scorekeeper/Official October 2012 Revision' : {
                'RowsPerPeriodHeader' : 9,
                'ColsPerTeam' : 43
            },
            'Scorekeeper/Official January 2013 Release' : {
                'RowsPerPeriodHeader' : 3 + 7,
                'ColsPerTeam' : 19
            },
            'Scorekeeper/Official March 2013 Release' : {
                'RowsPerPeriodHeader' : 3 + 7,
                'ColsPerTeam' : 19
            },
            'Scorekeeper/Official March 2014 Release' : {
                'RowsPerPeriodHeader' : 9,
                'ColsPerTeam' : 19,
                'RowsPerJam' : 1,
                'BaseRow' : 3,
            },
            'Scorekeeper/Official April 2014 Release' : {
                'RowsPerPeriodHeader' : 9,
                'ColsPerTeam' : 19,
                'RowsPerJam' : 1,
                'BaseRow' : 3,
            },
            'Scorekeeper/Official January 2015 Release' : {
                'RowsPerPeriodHeader' : 4,
                'ColsPerTeam' : 19,
                'RowsPerJam' : 1,
                'BaseRow' : 3,
            },
            'PenaltyTrackerNoMinors/Official March 2014 Release' : {
                'BaseRow' : 3,
                'MajorPenaltyCount': 7, # really 9, but not everything supports that yet
                'ColsPerTeam' : ColsFromCode("P"),
                'ColsPerPeriod' : ColsFromCode("AC")
            },
            'PenaltyTrackerNoMinors/Official April 2014 Release' : {
                'BaseRow' : 3,
                'MajorPenaltyCount': 7, # really 9, but not everything supports that yet
                'ColsPerTeam' : ColsFromCode("P"),
                'ColsPerPeriod' : ColsFromCode("AC")
            },
            'PenaltyTrackerNoMinors/Official January 2015 Release' : {
                'BaseRow' : 3,
                'MajorPenaltyCount': 7, # really 9, but not everything supports that yet
                'ColsPerTeam' : ColsFromCode("P"),
                'ColsPerPeriod' : ColsFromCode("AC")
            },
            'LineupTracker/Official October 2012 Revision' : {
                'RowsPerPeriodHeader' : 7,
                'ColsPerTeam' : 19
            },
            'LineupTracker/Official January 2013 Release' : {
                'RowsPerPeriodHeader' : 3 + 5,
                'ColsPerTeam' : 22
            },
            'LineupTracker/Official March 2013 Release' : {
                'RowsPerPeriodHeader' : 3 + 5,
                'ColsPerTeam' : 22
            },
            'LineupTracker/Official March 2014 Release' : {
                'RowsPerPeriodHeader' : 3 + 5,
                'ColsPerTeam' : 26,
                'RowsPerJam' : 1,
                'ColsPerSkater' : 4,
            },
            'LineupTracker/Official April 2014 Release' : {
                'RowsPerPeriodHeader' : 3 + 5,
                'ColsPerTeam' : 26,
                'RowsPerJam' : 1,
                'ColsPerSkater' : 4,
            },
            'LineupTracker/Official January 2015 Release' : {
                'RowsPerPeriodHeader' : 3 + 1,
                'ColsPerTeam' : 26,
                'RowsPerJam' : 1,
                'ColsPerSkater' : 4,
            },
            'JamTimer/Official January 2013 Release' : {
                'RowsPerPeriodHeader' : 13,
            },
            'JamTimer/Official March 2013 Release' : {
                'RowsPerPeriodHeader' : 13,
            },
            'JamTimer/Official October 2012 Revision' : {
                'RowsPerPeriodHeader' : 11,
            },
            'Actions' : { 'BaseRow' : 2 },
            'Actions/Official January 2013 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4,
            },
            'Actions/Official March 2013 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4,
            },
            'Actions/Official March 2014 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4
            },
            'Actions/Official April 2014 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4
            },
            'Errors' : { 'BaseRow' : 2 },
            'Errors/Official January 2013 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4,
            },
            'Errors/Official March 2013 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4,
            },
            'Errors/Official March 2014 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4
            },
            'Errors/Official April 2014 Release' : {
                'RowsPerTeam' : 49,
                'ColsPerPeriod' : 7,
                'BaseRow' : 4
            },
            
        }
        seen = []
        while cls:
            # get either name/version or name
            config = configs.get((cls.__name__+'/'+self.version),configs.get(cls.__name__,{}))
            for key in config:
                if key in seen:
                    continue
                setattr(obj,key,config[key])
                seen.append(key)
            if cls.__bases__:
                cls = cls.__bases__[0]
            else:
                break

    
class StatBook:
    """Class representing the underlying xlsx or xls file"""
    def __init__(self, path):
        """Figure out if we are an xlsx or xls file and load using the appropriate library"""
        self.fileName = path
        import os
        if os.path.splitext(path)[-1] == '.xlsx':
            import xlsx, time
            self.workbook = xlsx.open_workbook(path)
            self.empty_cell = xlsx.empty_cell
            # this is probably off by some amount of time, but for days, we should be OK
            self.dateToTuple = lambda(x): time.localtime(x * 24 * 60 * 60 - 2209096800.0)[0:6]
        else:
            import xlrd
            self.workbook = xlrd.open_workbook(path)
            self.empty_cell = xlrd.empty_cell
            self.dateToTuple = lambda(x): xlrd.xldate_as_tuple(x,self.workbook.datemode)
        global EmptyCellsTypes
        if self.empty_cell not in EmptyCellsTypes:
            EmptyCellsTypes.append(self.empty_cell)
        try:
            self.IBRF = self.workbook.sheet_by_name("IBRF")
        except:
            try:
                self.IBRF = self.workbook.sheet_by_name("IGRF")
            except:
                self.IBRF = self.workbook.sheet_by_name("ITRF")
        self.version = StatBookVersion(self.workbook)
        self.Score = self.workbook.sheet_by_name("Score")
        self.Penalty = self.workbook.sheet_by_name("Penalties")
        self.LineUp = self.workbook.sheet_by_name("Lineups")
        try:
            self.JamTimer = self.workbook.sheet_by_name(self.version.periodTimerSheetName)
        except Exception:
            self.JamTimer = self.workbook.sheet_by_name(self.version.periodTimerSheetName.replace(" ","_"))
        if self.version.hasActionsErrors:
            self.Actions = self.workbook.sheet_by_name("Actions")
            self.Errors = self.workbook.sheet_by_name("Errors")
        self.jamTimer = (JamTimer(self.JamTimer, 0,0), JamTimer(self.JamTimer, 0, 1))
        self.home = HomeTeam(self)
        self.away = AwayTeam(self)
        if self.home.league == self.away.league: # if same league, just use team name
            self.home.teamName = self.home.team
            self.away.teamName = self.away.team
        elif self.home.team == self.away.team: # if both all stars, just use league name
            self.home.teamName = self.home.league
            self.away.teamName = self.away.league
        else:
            self.home.teamName = self.home.league + " " + self.home.team
            self.away.teamName = self.away.league + " " + self.away.team
        if self.version.hasColors:
            self.home.color = self.Score.cell(0,8)
            self.away.color = self.Score.cell(0,27)
        self.hasMinors = self.version.hasMinors
        if self.hasMinors:
            # check to see if there are actually minors
            self.hasMinors = False
            for pt in self.home.penaltyTracker + self.away.penaltyTracker:
                if self.hasMinors: break
                for si in range(0,20):
                    if self.hasMinors: break
                    for mi in range(0,28):
                        m = pt.MinorPenalty(si,mi)
                        if m is not None:
                            self.hasMinors = True
                            break

    def BoutInfo(self):
        """Extract basic bout info from the IBRF"""
        import time
        bout = self.IBRF.cell(2,10)
        if bout in EmptyCellsTypes:
            bout = ""
        elif type(bout.value) == type(1.0):
            bout = int(bout.value)
        else:
            bout = bout.value
        extraInfo = {}
        if self.version.hasExtendedInfo:
            date = self.IBRF.cell(6,1).value
        else:
            date = self.IBRF.cell(4,1).value
        if type(date) == type("") or type(date) == type(u""): # should be YYYY-MM-DD
            date = date.split("-")
            date = (int(date[0]),int(date[1]),int(date[2]),0,0,0)
        else:
            date = self.dateToTuple(date)
        retval = { 'version': self.version, 
            'location': self.IBRF.cell(2,1).value, 
            'city': self.IBRF.cell(2,7).value, 
            'state':self.IBRF.cell(2,9).value,
            'game': self.IBRF.cell(2,10).value,
            'bout':bout, 
            'date': time.strftime("%B %d, %G",date + (0,0,0)), 
            'teamNames': (self.home.teamName, self.away.teamName) }
        if self.version.hasExtendedInfo:
            retval['tournament'] = self.IBRF.cell(4,1).value
            retval['host'] = self.IBRF.cell(4,7).value
            retval['suspension'] = self.IBRF.cell(6,10)
            retval['start'] = self.IBRF.cell(6,7).value
            # end time removed
        else:
            retval['start'] = self.IBRF.cell(4,7).value,
            retval['end'] = self.IBRF.cell(4,10).value
        return retval
    
    def TitleSummary(self):
        bi = self.BoutInfo()
        return bi['date'] +":" + self.home.teamName + "(" + "%d"%self.home.TotalScore() + ") vs " + self.away.teamName + "(" + "%d"%self.away.TotalScore() + ")"
    def HasScoreSheets(self):
        """Is there data entered in the score sheet?"""
        return not IsBlankCell(self.Score.cell(2,0))
    def HasLineupSheets(self):
        """Is there data entered in the lineup sheets?"""
        if self.version.hasNoPivot: # 2012 Oct version auto include the jam number and jammer skater
            # so check to see if everything else is blank
            for ji in range(0,1):
                if not IsBlankCell(self.LineUp.cell(1 + self.version.headerRows + ji * 2, 5)):
                    return True
                if not IsBlankCell(self.LineUp.cell(1 + self.version.headerRows + ji * 2, 8)):
                    return True
                if not IsBlankCell(self.LineUp.cell(1 + self.version.headerRows + ji * 2, 11)):
                    return True
                if not IsBlankCell(self.LineUp.cell(1 + self.version.headerRows + ji * 2, 14)):
                    return True
            return False
        return not IsBlankCell(self.LineUp.cell(2,0))
    def HasPenaltySheets(self):
        """Is there data entered in the penalty tracker sheets?"""
        return True # for now
    def HasActionSheets(self):
        """Is there data entered in the actions sheets?"""
        if not self.hasActiosnErrors:
            return False
        for a in self.home.actions + self.away.actions:
            if not a.IsBlank():
                return True
        return False
    def HasErrorSheets(self):
        """Is there data entered in the errors sheets?"""
        if not self.hasActiosnErrors:
            return False
        for a in self.home.errors + self.away.errors:
            if not a.IsBlank():
                return True
        return False

class Team:
    """Abstract class representing the sections of the various sheets for each team"""
    TeamIndex = 0
    IGRFSkaterFirstRow = 10
    IGRFLeagueRow = 7
    IGRFTeamRow = 8
    def __init__(self, book):
        self.statbook = book
        self.skaters = []
        self.skaterNumbers = []
        self.skaterNames = []
        if self.statbook.version.hasExtendedInfo:
            self.IGRFSkaterFirstRow = 13
            self.IGRFLeagueRow = 9
            self.IGRFTeamRow = 10
        for i in range(0,20):
            num = self.SkaterNum(i)
            name = self.SkaterName(i)
            if IsBlankCell(num) or num.value == "-": # "For any empty spots in either roster, input "-" in Skater # and Skater Name columns."
                break
            self.skaterNumbers.append(CanonSkaterNumber(num.value))
            self.skaterNames.append(name.value)
            self.skaters.append(Skater(self, name.value, num.value))
        self.scorekeeper = (Scorekeeper(self.statbook.Score, self.TeamIndex, 0), Scorekeeper(self.statbook.Score, self.TeamIndex, 1))
        if self.statbook.version.hasMinors:
            self.penaltyTracker = (PenaltyTracker(self.statbook.Penalty, self.TeamIndex, 0), PenaltyTracker(self.statbook.Penalty, self.TeamIndex, 1))
        else:
            self.penaltyTracker = (PenaltyTrackerNoMinors(self.statbook.Penalty, self.TeamIndex, 0), PenaltyTrackerNoMinors(self.statbook.Penalty, self.TeamIndex, 1))
        if book.HasLineupSheets():
            self.lineupTracker = (LineupTracker(self.statbook.LineUp, self.TeamIndex, 0), LineupTracker(self.statbook.LineUp, self.TeamIndex, 1))
        else:
            self.lineupTracker = (EstLineupTracker(self,self.scorekeeper[0], self.penaltyTracker[0]), EstLineupTracker(self,self.scorekeeper[1], self.penaltyTracker[1]))
        if book.version.hasActionsErrors:
            self.actions = (Actions(self.statbook.Actions, self.TeamIndex,0), Actions(self.statbook.Actions, self.TeamIndex, 1))
            self.errors = (Errors(self.statbook.Errors, self.TeamIndex,0), Errors(self.statbook.Errors, self.TeamIndex, 1))
        else:
            self.actions = ()
            self.errors = ()
        self.league = self.statbook.IBRF.cell(self.IGRFLeagueRow, self.IBRFCol).value
        self.team = self.statbook.IBRF.cell(self.IGRFTeamRow, self.IBRFCol).value
        self.teamName = self.league + ":" + self.team
    def Period(self, index):
        """Returns a tuple of the period's scorekeeper, penalty tracker, lineup tracker, actions, 
        errors.
        
        Index is 1 or 2
        
        """
        index = index - 1
        return (self.scorekeeper[index],self.penaltyTracker[index], self.lineupTracker[index],
            self.actions[index], self.errors[index])
    def SkaterNum(self, index):
        """Returns the number for the given skater in the IBRF"""
        return self.statbook.IBRF.cell(index + self.IGRFSkaterFirstRow, self.IBRFCol)
    def SkaterName(self, index):
        """Returns the name for the given skater in the IBRF"""
        return self.statbook.IBRF.cell(index + self.IGRFSkaterFirstRow, self.IBRFCol+1)
    def SkaterForNum(self, num):
        """Gets the skater object for a given skater number on the team"""
        for sk in self.skaters:
            if sk.IsSkaterNumber(num):
                return sk
        return None
    def TotalScore(self, period = None):
        """Get the total score for either both periods, or (if the optional period is passed)
        the total score for just that one period
        
        """
        if period is None:
            return self.scorekeeper[0].TotalScore() + self.scorekeeper[1].TotalScore()
        return self.scorekeeper[period].TotalScore()
    def TotalMajors(self, period = None):
        """Get the total score for either both periods, or (if the optional period is passed)
        the total score for just that one period
        
        """
        if period is None:
            return self.penaltyTracker[0].TotalMajors() + self.penaltyTracker[1].TotalMajors()
        return self.penaltyTracker[period].TotalMajors() 
    def EstimateLineUp(self):
        """Use an estimated lineup for tracker, derived from score and penalty sheets"""
        self.lineupTracker = (EstLineupTracker(self,self.scorekeeper[0], self.penaltyTracker[0]), EstLineupTracker(self,self.scorekeeper[1], self.penaltyTracker[1]))

class HomeTeam(Team):
    """A subclass of Team representing the home team"""
    TeamIndex = 0
    IBRFCol = 1

class AwayTeam(Team):
    """A subclass of Team representing the away team"""
    TeamIndex = 1
    IBRFCol = 7

class Skater:
    """An object representing a skater, with a name, number(s), for a specific team"""
    def __init__(self, team, name, number):
        self.team = team
        self.name = name
        self.number = CanonSkaterNumber(number)
        if self.name:
            self.nameOrNumber = self.name
            self.nameAndNumber = self.name + ", " + str(self.number)
        else:
            self.nameOrNumber = str(self.number)
            self.nameAndNumber = str(self.number)
        self.alias = [self.number]
        self.jamsSkated = None
    def AddAliasNumber(self, number):
        """Allows for more than one number for a given skater, for cross-bout purposes
        where the skater may skate under another number, or if the number is entered
        differently on different sheets
        
        """
        self.alias.append(CanonSkaterNumber(number))
    def IsSkaterNumber(self, number):
        """Does the given number repesent the skater?"""
        return CanonSkaterNumber(number) in self.alias
    def __repr__(self):
        if self.name:
            return self.team.teamName + u" " + unicode(self.number) + u"(" + unicode(self.name) + u")"
        else:
            return self.team.teamName + u" " + unicode(self.number) 
    def JamsSkated(self):
        """Find all jams the given skater skated in"""
        if self.jamsSkated:
            return self.jamsSkated
        retval = []
        for period in range(0,1):
            lu = self.team.lineupTracker[period]
            for jam in range(1,self.version.maxNumJams):
                skaters = lu.LineupForJam(jam)
                if not skaters:
                    continue
                    break
                for skn in skaters:
                    if self.IsSkaterNumber(skn):
                        retval.append((period,jam))
                        break
        self.jamsSkated = retval
        return retval

class PeriodStat:
    """An abstract section (page) of a given sheet representing a given team and period"""
    def __init__(self, sheet, teamIndex, period):
        self.sheet = sheet
        self.teamIndex = teamIndex
        self.period = period
        self.version = StatBookVersion(sheet.book)
        self.version.Config(self)
    BaseRow = 0
    BaseCol = 0
    RowsPerTeam = 0
    ColsPerTeam = 0
    RowsPerJam = 0 #
    RowsPerPeriodHeader = 0
    ColsPerPeriod = 0
    def RowsPerPeriod(self):
        return self.RowsPerPeriodHeader + self.version.maxNumJams * self.RowsPerJam
    def relcell(self, drow, dcol):
        """Get a relative cell in the period stat, supporting the various period/team row/column
        configurations of the different stat sheets
        
        """
        return self.sheet.cell(self.BaseRow + self.RowsPerTeam * self.teamIndex + self.RowsPerPeriod() * self.period + drow,
                               self.BaseCol + self.ColsPerTeam * self.teamIndex + self.ColsPerPeriod * self.period + dcol)

class Scorekeeper(PeriodStat):
    """Scorekeeper sheet information.  Records information based on jam number, with an
    optional parameter flag indicating 'after the star pass' if True.
    
    """
    BaseCol = 0
    BaseRow = 2
    ColsPerTeam = 37
    RowsPerJam = 2
    RowsPerPeriodHeader = 59 - 2 * 20
    ColsPerPass = 3
    def __init__(self, sheet, teamIndex, period):
        PeriodStat.__init__(self, sheet, teamIndex, period)
        self.BaseRow = 1 + self.version.headerRows
        if not self.version.hasGhostPoints:
            self.ColsPerPass = 1
        #self.ColsPerTeam = 37 + (self.version.maxNumPasses - 8) * self.ColsPerPass
        #self.ColsPerTeam = 13 + self.version.maxNumPasses * self.ColsPerPass
    def buildJamIndex(self):
        """Create the 1 based jam index (lazily) to account for star passes"""
        if not hasattr(self,'jamIndex'):
            self.jamIndex = {}
            ji = 1
            for i in range(0,self.version.maxNumJams):
                jamNum = self.relcell(i * self.RowsPerJam, 0)
                if jamNum.value == ji:
                    self.jamIndex[ji] = i * self.RowsPerJam
                    ji = ji + 1
        return self.jamIndex
    def NumberOfJams(self):
        """How many jams are in the given period"""
        return len(self.buildJamIndex())
    def HasStarPass(self, jam):
        """Does the given jam (starting with jam 1) have a star pass"""
        ji = self.buildJamIndex().get(jam)
        if ji is None:
            return False
        v = self.relcell(ji + 1,0).value
        if type(v) == type("") or type(v) == type(u""):
            return v.lower() == "sp"
        return False
    def NumberOfPasses(self, jam, starpass = False):
        """How many passes are in the given jam
        
        Pass True for the extra parameter to get the passes after the star pass"""
        for ps in range(2,self.version.maxNumPasses):
            passpts = self.JamPassPoints(jam,ps)
            if passpts is None:
                return ps - 1
        return self.version.maxNumPasses
    def JamPoints(self, jam, starpass = False):
        """How many points were scored in the given jam (or jam after star pass)"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            # if we return 0, we break brecre, since the jam numbers are wrong
            # if we return None, we break some stat that adds none
            return None
        #return self.relcell(ji + self.RowsPerJam * starpass,4 + self.ColsPerPass * self.version.maxNumPasses).value
        return self.relcell(ji + self.RowsPerJam * starpass,7 + self.ColsPerPass * (self.version.maxNumPasses-1)).value
    def TotalScore(self):
        """Total scores for period represented by this page of the sheet"""
        total = 0
        for ji in range(0, self.NumberOfJams()):
            pts = self.JamPoints(ji+1)
            if pts:
                total = total + pts
            if self.HasStarPass(ji+1):
                pts= self.JamPoints(ji+1, True)
                if pts:
                    total = total + pts
        return total
    def Jammer(self, jam, starpass = False):
        """Who was the jammer for the given jam (or star pass)"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        return CanonSkaterNumber(self.relcell(ji + 2 * starpass,1).value)
    def JamPassPoints(self, jam, pss, starpass = False):
        """How many points were scored in a given pass of a given jam"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        cell = self.relcell(ji + self.RowsPerJam * starpass,7 + (pss - 2) * self.ColsPerPass)
        if IsBlankCell(cell):
            return None
        return cell.value
    def JamPassGhostPoints(self, jam, pss, starpass = False):
        """What ghost points were recorded for a given pass of a given jam
        
        Returns a list of character codes
        
        """
        if not self.version.hasGhostPoints:
            return []
        ji = self.buildJamIndex().get(jam) 
        retval = []
        v = self.relcell(ji + self.RowsPerJam * starpass,7 + (pss - 2) * self.ColsPerPass + 1).value
        if v:
            retval.append(v)
        v = self.relcell(ji + self.RowsPerJam * starpass,7 + (pss - 2) * self.ColsPerPass + 2).value
        if v:
            retval.append(v)
        v = self.relcell(ji + self.RowsPerJam * starpass+1,7 + (pss - 2) * self.ColsPerPass + 1).value
        if v:
            retval.append(v)
        v = self.relcell(ji + self.RowsPerJam * starpass+1,7 + (pss - 2) * self.ColsPerPass + 2).value
        if v:
            retval.append(v)
        return retval
    def JamLost(self, jam, starpass = False):
        """Was the Lost Lead cell checked for the given jam?"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        return len(self.relcell(ji + self.RowsPerJam * starpass, 2).value)
    def JamLead(self, jam, starpass = False):
        """Was the Lead cell checked for the given jam?"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        return len(self.relcell(ji + self.RowsPerJam * starpass, 3).value)
    def JamCall(self, jam, starpass = False):
        """Was the Called cell checked for the given jam?"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        return len(self.relcell(ji + self.RowsPerJam * starpass, 4).value)
    def JamInj(self, jam, starpass = False):
        """Was the Injury cell checked for the given jam?"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        return len(self.relcell(ji + self.RowsPerJam * starpass, 5).value)
    def JamNP(self, jam, starpass = False):
        """Was the No Pass cell checked for the given jam?"""
        ji = self.buildJamIndex().get(jam) 
        if ji is None:
            return None
        return len(self.relcell(ji + self.RowsPerJam * starpass, 6).value)

class PenaltyTracker(PeriodStat):
    """Penalty tracker sheet information.  Records information based the zero based skater index
    and a zero based penalty index (routines return None if no corresponding data is found
    for that skater/penalty)"""
    BaseCol = 0
    BaseRow = 2
    ColsPerPeriod = 39
    RowsPerTeam = 46
    def __init__(self, sheet, teamIndex, period):
        PeriodStat.__init__(self, sheet, teamIndex, period)
        self.BaseRow = 1 + self.version.headerRows
    def Skater(self, skater):
        """Skater number from the PT sheet (just in case it is different from IGRF)
        
        Returns a string with the skater number
        
        """
        number = self.relcell(skater * 2, 0)
        if IsBlankCell(number):
            return None
        return CanonSkaterNumber(number.value)
    def MinorPenalty(self, skater, index):
        """What is the minor penalty for the given skater index and penalty index
        
        Returns a tuple of the code and what jam (or None if the penalty index has nothing)
        
        """
        code = self.relcell(skater * 2, index + 1)
        if IsBlankCell(code):
            return None
        jam = self.relcell(skater * 2 + 1, index + 1)
        if IsBlankCell(jam) or jam.value == "?":
            return (code.value, 0)
        return (code.value, jam.value)
    def MajorPenalty(self, skater, index):
        """What is the major penalty for the given skater index and penalty index
        
        Returns a tuple of the code and what jam (or None if the penalty index has nothing).
        NB: A fourth minor will return the code '4' as a string, not a number
        
        """
        code = self.relcell(skater * 2, index + 30)
        if IsBlankCell(code):
            return None
        if code.value == 4 or code.value == 4.0:
            code.value = '4'
        if type(code.value) != type("") and type(code.value) != type(u""):
            print "Warning: major penalty",skater,index,code.value
        jam = self.relcell(skater * 2 + 1, index + 30)
        if IsBlankCell(jam) or jam.value == "?":
            return (code.value, 0)
        return (code.value, jam.value)
    def Expulsion(self, skater):
        """What is the explusion for the given skater index (if any)
        
        Returns a tuple of the code and what jam (or None if the penalty index has nothing)
        
        """
        code = self.relcell(skater * 2, 37)
        if IsBlankCell(code):
            return None
        jam = self.relcell(skater * 2 + 1, 37)
        if IsBlankCell(jam) or jam.value == "?":
            return (code.value, 0)
        return (code.value, jam.value)
    def TotalMajors(self):
        """Total number of major penalties (including '4's)"""
        total = 0
        for skater in range(0,20):
            for pi in range(0,7):
                if self.MajorPenalty(skater, pi) is not None:
                    total = total + 1
        return total
class PenaltyTrackerNoMinors(PeriodStat):
    """Penalty tracker sheet information.  Records information based the zero based skater index
    and a zero based penalty index (routines return None if no corresponding data is found
    for that skater/penalty).   Designed for "no minors" Jan 2013 (single) PT"""
    BaseCol = 0
    BaseRow = 3
    MajorPenaltyCount = 7
    ColsPerTeam = ColsFromCode("N")
    ColsPerPeriod = ColsFromCode("Y")
    def __init__(self, sheet, teamIndex, period):
        PeriodStat.__init__(self, sheet, teamIndex, period)
    def Skater(self, skater):
        """Skater number from the PT sheet (just in case it is different from IGRF)
        
        Returns a string with the skater number
        
        """
        number = self.relcell(skater * 2, 0)
        if IsBlankCell(number):
            return None
        return CanonSkaterNumber(number.value)
    def MinorPenalty(self, skater, index):
        """What is the minor penalty for the given skater index and penalty index
        
        Returns a tuple of the code and what jam (or None if the penalty index has nothing)
        
        """
        return None
    def MajorPenalty(self, skater, index):
        """What is the major penalty for the given skater index and penalty index
        
        Returns a tuple of the code and what jam (or None if the penalty index has nothing).
        NB: A fourth minor will return the code '4' as a string, not a number
        
        """
        code = self.relcell(skater * 2, index + 1)
        if IsBlankCell(code):
            return None
        if type(code.value) != type("") and type(code.value) != type(u""):
            print "Warning: major penalty",skater,index,code.value
        jam = self.relcell(skater * 2 + 1, index + 1)
        if IsBlankCell(jam) or jam.value == "?":
            return (code.value, 0)
        return (code.value, jam.value)
    def Expulsion(self, skater):
        """What is the explusion for the given skater index (if any)
        
        Returns a tuple of the code and what jam (or None if the penalty index has nothing)
        
        """
        code = self.relcell(skater * 2, self.MajorPenaltyCount + 1)
        if IsBlankCell(code):
            return None
        jam = self.relcell(skater * 2 + 1, self.MajorPenaltyCount + 1)
        if IsBlankCell(jam) or jam.value == "?":
            return (code.value, 0)
        return (code.value, jam.value)
    def TotalMajors(self):
        """Total number of major penalties (including '4's)"""
        total = 0
        for skater in range(0,20):
            for pi in range(0,self.MajorPenaltyCount):
                if self.MajorPenalty(skater, pi) is not None:
                    total = total + 1
        return total

class LineupTracker(PeriodStat):
    """Similar to the Scorekeeper sheet, returns the appropriate information for a given pass
    (with an optional star pass)
    
    """
    BaseCol = 0
    BaseRow = 2
    ColsPerTeam = 18
    RowsPerJam = 2
    ColsPerSkater = 3
    RowsPerPeriodHeader = 57 - 2 * 20
    def __init__(self, sheet, teamIndex, period):
        PeriodStat.__init__(self, sheet, teamIndex, period)
        self.BaseRow = 1 + self.version.headerRows
    def buildJamIndex(self):
        """Build the index of the 1-based jam lazily, to account for the star passes"""
        if not hasattr(self,'jamIndex'):
            self.jamIndex = {}
            ji = 1
            for i in range(0,self.version.maxNumJams):
                jamNum = self.relcell(i * self.RowsPerJam, 0)
                if jamNum.value == ji:
                    self.jamIndex[ji] = i * self.RowsPerJam
                    ji = ji + 1
        return self.jamIndex
    def LineupForJam(self, jam, starpass = False):
        """Returns the lineup (P,B,B,B,J) of the skater numbers in the given jam"""
        ji = self.buildJamIndex().get(jam)
        if ji is None:
            #print "jam",jam,self.buildJamIndex();
            return None
        firstCol = 1
        if self.version.hasNoPivot:
            firstCol = 2
        jammer = self.relcell(ji, firstCol+self.version.lineupColumns['j'])
        if IsBlankCell(jammer):
            #print "blank jammer for ", jam
            return None
        jammer = jammer.value
        pivot = self.relcell(ji, firstCol+self.version.lineupColumns['p']).value
        blocker2 = self.relcell(ji, firstCol+self.version.lineupColumns['2']).value
        blocker3 = self.relcell(ji, firstCol+self.version.lineupColumns['3']).value
        blocker4 = self.relcell(ji, firstCol+self.version.lineupColumns['4']).value
        return (CanonSkaterNumber(pivot), CanonSkaterNumber(blocker2), CanonSkaterNumber(blocker3), CanonSkaterNumber(blocker4), CanonSkaterNumber(jammer))
    def JamHasPivot(self, jam):
        if not self.version.hasNoPivot:
            return True # as far as we know, it is a pivot
        ji = self.buildJamIndex().get(jam)
        if ji is None:
            #print "jam",jam,self.buildJamIndex();
            return None
        return IsBlankCell(self.relcell(ji, 1))
        return False
    def PenaltyPasses(self, jam, starpass = False):
        """Returns a dictionary of the penalty box in/out passes information for the given
        jam number.  The dictionary uses the skater number as the key, whose value is a
        tuple of the enter and exit values
        
        """
        retval = {}
        firstCol = 1
        if self.version.hasNoPivot:
            firstCol = 2
        ji = self.buildJamIndex().get(jam)
        if ji is None:
            return None
        for si in range(0,5):
            sn = self.relcell(ji + self.RowsPerJam * starpass, firstCol + self.ColsPerSkater * si)
            if IsBlankCell(sn):
                continue # if a skater is missing, ignore it
                return None
            sn = CanonSkaterNumber(sn.value)
            retval[sn] = []
            if self.RowsPerJam == 1:
                for pi in range(0,self.ColsPerSkater-1):
                    code = self.relcell(ji + self.RowsPerJam * starpass, firstCol + self.ColsPerSkater * si + 1 + pi)
                    if IsBlankCell(code):
                        continue
                    if code.value == 'S':
                        retval[sn].append(('S',None)) # start, never exit
                    elif code.value == '$':
                        retval[sn].append(('S',1))
                    elif code.value == 'X' or code.value == 'x':
                        retval[sn].append((1,1))
                    elif code.value == '/':
                        retval[sn].append((1,None))
                    elif code.value == '|':
                        retval[sn].append((None,1)) # start in box, exit after star pass
                    elif code.value == '\\':
                        retval[sn].append((None,1)) # entered in jam, exit after star pass
                    elif code.value == '3' or code.value == 3:
                        pass # injury - doesn't impact
                    else:
                        print "Unkown line up code", code.value, "for", sn,"jam", jam
            else:
                for pi in range(0,1):
                    startPass = self.relcell(ji + pi + 2 * starpass, firstCol + self.ColsPerSkater * si + 1)
                    endPass = self.relcell(ji + pi + 2 * starpass, firstCol + self.ColsPerSkater * si + 2)
                    if IsBlankCell(startPass) and IsBlankCell(endPass):
                        continue
                    retval[sn].append((startPass.value, endPass.value))
        return retval

class EstLineupTracker:
    """Provides a estimated lineup tracker, based on information in the scorekeeper
    sheet and penalty tracker sheet
    
    """
    def __init__(self, team, scorekeeper, penaltyTracker):
        self.team = team
        self.scorekeeper = scorekeeper
        self.penaltyTracker = penaltyTracker
    def JamHasPivot(self, jam):
        return True
    def LineupForJam(self, jam, starpass = False):
        """Returns the lineup (P,B,B,B,J) of the skater numbers in the given jam by
        getting the jammer from the score keeper, and finding what skaters got
        penalties in that jam to fill in the rest
        
        """
        jammer = self.scorekeeper.Jammer(jam, starpass)
        if not jammer:
            return None
        pivot = None
        if starpass: # the "pivot" (not a pivot) is the last jammer
            pivot = self.scorekeeper.Jammer(jam, False)
        elif self.scorekeeper.HasStarPass(jam): # look
            pivot = self.scorekeeper.Jammer(jam, True)
        if pivot:
            pivot = CanonSkaterNumber(pivot)
        lineup = [CanonSkaterNumber(jammer)]
        for si in range(0,20):
            for mi in range(0,28):
                m = self.penaltyTracker.MinorPenalty(si,mi)
                if not m:
                    continue
                if m[1] == jam:
                    skater = CanonSkaterNumber(self.team.skaterNumbers[si])
                    if skater not in lineup:
                        lineup = [skater] + lineup
            for mi in range(0,7):
                m = self.penaltyTracker.MajorPenalty(si,mi)
                if not m:
                    continue
                if m[1] == jam:
                    skater = CanonSkaterNumber(self.team.skaterNumbers[si])
                    if len(skater) == 0:
                        continue
                    if skater not in lineup:
                        lineup = [skater] + lineup
        while len(lineup) < 5:
            lineup = ['???'] + lineup
        if pivot:
            if pivot in lineup:
                lineup.remove(pivot)
                lineup = [pivot] + lineup
            else:
                lineup[0] = pivot # lineup[0] better be the pivot
        return (lineup[0], lineup[1], lineup[2], lineup[3], lineup[4])
    def PenaltyPasses(self, jam, starpass = False):
        """Returns a dictionary of the penalty box in/out passes information for the given
        jam number.  Currently returns an empty dictionary if the jam appears to exist in
        the scorekeeper sheet, or None if the jam doesn't
        
        """
        jammer = self.scorekeeper.Jammer(jam, starpass)
        if not jammer:
            return None
        retval = {}
        for sknum in self.LineupForJam(jam,starpass):
            retval[sknum] = []
        return retval

class JamTimer(PeriodStat):
    """Represents the jam timer sheet"""
    BaseCol = 0
    BaseRow = 8
    RowsPerJam = 1
    RowsPerPeriodHeader = 36 - 20
    def buildJamIndex(self):
        """Builds a lazy index of the jams"""
        if not hasattr(self,'jamIndex'):
            self.jamIndex = {}
            ji = 1
            for i in range(0,self.version.maxNumJams):
                jamNum = self.relcell(i, 0)
                if jamNum.value == ji:
                    self.jamIndex[ji] = i
                    ji = ji + 1
        return self.jamIndex
    def JamDuration(self, jam):
        """How long did the given jam last (or None if not found)"""
        ji = self.buildJamIndex()
        if jam > len(ji):
            return None
        jamDuration = self.relcell(self.buildJamIndex()[jam], 1)
        if IsBlankCell(jamDuration):
            return 0
        jamDuration = jamDuration.value
        if jamDuration > 0.5: # 2:00 is 12:02 PM
            jamDuration = round((jamDuration - 0.5) * (24 * 60 * 60)) # fraction of a day
        else:
            jamDuration = jamDuration * 1440 # convert to seconds
            if jamDuration <= 2:
                jamDuration = round(jamDuration * 60)
        return jamDuration
    def JamPackLaps(self, jam):
        """How many pack laps"""
        ji = self.buildJamIndex()
        if jam > len(ji):
            return None
        return self.relcell(self.buildJamIndex()[jam], 2).value

    def JamEndedWith(self, jam):
        """What are the codes for the end of the jam (or None if not found)"""
        ji = self.buildJamIndex()
        if jam > len(ji):
            return
        return self.relcell(self.buildJamIndex()[jam], 3).value, self.relcell(self.buildJamIndex()[jam], 4).value

class Actions(PeriodStat):
    """Represents the action sheet"""
    BaseCol = 2
    BaseRow = 2
    ColsPerPeriod = 9
    RowsPerTeam = 47
    def __init__(self, sheet, teamIndex, period):
        PeriodStat.__init__(self, sheet, teamIndex, period)
    def IsBlank(self):
        """Is the entire page of the sheet blank?"""
        for si in range(0,18):
            for a in self.AssistsForSkater(si):
                if a:
                    #print "Skater",si,"Has asisst",a
                    return False
            for a in self.AttacksForSkater(si):
                if a:
                    #print "Skater",si,"Has attack",a
                    return False
        return True
    def AssistsForSkater(self, skater):
        """Returns the assist actions count for the given skater index"""
        return [self.relcell(skater, 0).value,
            self.relcell(skater, 1).value,
            self.relcell(skater, 2).value,
            self.relcell(skater, 3).value,
            self.relcell(skater, 4).value]
    def AttacksForSkater(self, skater):
        """Returns the assist attack count for the given skater index"""
        if self.teamIndex:
            skater -= self.RowsPerTeam - 21
        else:
            skater += self.RowsPerTeam + 21
        return [self.relcell(skater, 0).value,
            self.relcell(skater, 1).value,
            self.relcell(skater, 2).value,
            self.relcell(skater, 3).value,
            self.relcell(skater, 4).value]
class Errors(PeriodStat):
    """Represents the errors sheet"""
    BaseCol = 2
    BaseRow = 2
    ColsPerPeriod = 9
    RowsPerTeam = 47
    def __init__(self, sheet, teamIndex, period):
        PeriodStat.__init__(self, sheet, teamIndex, period)
        #self.BaseRow = 1 + self.version.headerRows
    def IsBlank(self):
        """Is the entire page of the sheet blank?"""
        for si in range(0,18):
            for a in self.ErrorsForSkater(si):
                if a:
                    return False
            for a in self.JammerActionsForSkater(si):
                if a:
                    return False
        return True
    def ErrorsForSkater(self, skater):
        """Returns the errors count for the given skater index"""
        if self.teamIndex:
            skater -= self.RowsPerTeam - 21
        else:
            skater += self.RowsPerTeam + 21
        return [self.relcell(skater, 0).value,
            self.relcell(skater, 1).value,
            self.relcell(skater, 2).value,
            self.relcell(skater, 3).value,
            self.relcell(skater, 4).value]
    def JammerActionsForSkater(self, skater):
        """Returns the jammer actions count for the given skater index"""
        return [self.relcell(skater, 0).value,
            self.relcell(skater, 1).value,
            self.relcell(skater, 2).value,
            self.relcell(skater, 3).value,
            self.relcell(skater, 4).value]

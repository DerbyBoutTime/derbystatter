#!/bin/python
import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )

import cgitb
cgitb.enable(format='text')

from optparse import OptionParser
parser = OptionParser("usage: %prog [options] arg")
parser.add_option("-f", "--file", dest="fname",
                  help="write report to FILE", metavar="FILE")
parser.add_option("-v", "--verbose",
                  action="store_true", dest="verbose", default=False,
                  help="print detailed progress")

fname = None
(options, args) = parser.parse_args()
if len(args) == 1:
    fname = args[0]
elif options.fname:
    fname = options.fname
else:
    fname = "/Users/gandreas/Downloads/STATS-110513_FairbanksRollerGirls_vs_RageCityRollergirls.xlsx"

if not fname:
    parser.error("Missing stat sheet filename")
verbose = options.verbose

import cgitb
cgitb.enable(format='text')

from statbook import *

NumErrors = 0
NumWarnings = 0
def Error(*args):
    global NumErrors
    NumErrors = NumErrors + 1
    print ">>> Error:",
    for arg in args:
        print arg,
    print
def Warning(*args):
    global NumWarnings
    NumWarnings = NumWarnings + 1
    print "### Warning:",
    for arg in args:
        print arg,
    print

def CheckIBRF(self, skater, period=0,jam=0):
    if skater == '???':
        return
    if skater == '':
        return
    if period and jam:
        msg = "P%dJ%d" % (period, jam)
    else:
        msg = ""
    if skater not in self.skaterNumbers:
        Error("Skater",team.teamName,skater,"not found in IBRF",msg)
Team.CheckIBRF = CheckIBRF    

book = StatBook(fname)
#for row in range(10,18):
    #    for col in range(0,8):
    #   print book.IBRF.cell(row, col),
#print
#print book.home.SkaterNum(0), book.home.SkaterName(0)
print book.home.skaters
print book.away.skaters
if not book.HasLineupSheets() and book.HasPenaltySheets() and book.HasScoreSheets():
    book.home.EstimateLineUp()
    book.away.EstimateLineUp()
    print "(Estimating Lineup)"
print "=== Score"
# Check the Scorekeeper sheet
for team in (book.home, book.away):
    for period in (0,1):
        print "---",team.teamName,"Period",period+1
        sk = team.scorekeeper[period]
        #print "Jam Index",sk.buildJamIndex()
        for i in range(1,sk.NumberOfJams()):
            pt = sk.JamPoints(i)
            if pt is None:
                break
            jammer = sk.Jammer(i)
            # Make sure that jammer in score sheet is in the IBRF
            team.CheckIBRF(jammer,period+1,i)
            # Check lead/call
            if verbose:
                if sk.HasStarPass(i):
                    print "*",i,jammer, pt, sk.JamPoints(i,True)
                else:
                    print i,jammer, pt
            for ps in range(2,9):
                passpts = sk.JamPassPoints(i,ps)
                if passpts is None:
                    if ps == 2:
                        if not sk.JamNP(i):
                            Error("Jam",i,"Pass",ps,"has no recorded points, but NP not checked")
                    continue
                if sk.JamNP(i):
                    if ps == 2 and passpts != "-":
                        Warning("Jam",i,"Pass",ps,"should be '-' if NP checked")
                    elif ps != 2 and passpts:
                        Warning("Jam",i,"Pass",ps,"should be blank if NP checked")   
                if sk.version.hasGhostPoints:
                    ghostpts = sk.JamPassGhostPoints(i,ps)
                    # Check the ghost points to see if they make sense for the score recorded
                    if len(ghostpts) > passpts:
                        Error("Jam",i,"Pass",ps,"has more ghost points than recorded points")
                    if passpts == 5:
                        if len(ghostpts) == 0:
                            Warning("Jam",i,"Pass",ps,"has 5 points, but no recorded ghost points")
                        elif (not "J" in ghostpts) and (not "L" in ghostpts):
                            Error("Jam",i,"Pass",ps,"has 5 points, but no Jammer or Lap ghost points, recorded",ghostpts)
# Check Penalty Tracker
def CheckPenaltyCode(x, skater, major=0):
    codes = "BHXCEFAILMOPSNGZ?"
    if major:
        codes += "4"
    if type(x) == type("") or type(x) == type(u""):
        if x in codes:
            return
    if major and x == 4:
        return
    if major:
        Error("Unknown major error code '" + str(x) + "' for skater " + skater)
    else:
        Error("Unknown minor error code '" + str(x) + "' for skater " + skater)
print "=== Penalty"
for team in (book.home, book.away):
    firstMinor = [0] * 20
    firstMajor = [0] * 20
    lastMajor = [7] * 20
    for period in (0,1):
        print "---",team.teamName,"Period",period+1
        pt = team.penaltyTracker[period]
        for i in range(0,len(team.skaters)):
            fourth = []
            if verbose:
                print team.skaters[i].number,
            for pi in range(firstMinor[i],28):
                penalty = pt.MinorPenalty(i, pi)
                if not penalty:
                    break
                CheckPenaltyCode(penalty[0], team.skaters[i].number)
                if pi % 4 == 3:
                    fourth.append(penalty[1])
                firstMinor[i] = pi
                if verbose:
                    print "[",penalty[0],"/",penalty[1],"]",
            fours = []
            if verbose: 
                print "\t|",
            for pi in range(firstMajor[i],lastMajor[i]):
                penalty = pt.MajorPenalty(i, pi)
                if not penalty:
                    continue
                if penalty[0] == '4' or penalty[0] == 4:
                    fours.append(penalty[1])
                else:
                    CheckPenaltyCode(penalty[0], team.skaters[i].number, True)
                if verbose:
                    print "[",penalty[0],"/",penalty[1],"]",
            if verbose: 
                print
            if len(fourth) > len(fours):
                Error("More fourth minors (" + str(len(fourth)) + ") than served (" + str(len(fours)) + "), skater " + team.skaters[i].number)
            if len(fourth) < len(fours):
                Error("Fewer fourth minors (" + str(len(fourth)) + ") than served (" + str(len(fours)) + "), skater " + team.skaters[i].number)
            for jamnum in fourth:
                if jamnum in fours:
                    fours.remove(jamnum)
                elif jamnum+1 in fours:
                    fours.remove(jamnum+1)
                else:
                    Error("Fourth Minor occured in jam", jamnum, "but no corresponding penalty minute served for skater", team.skaters[i].number)
# Check Lineup Tracker
print "=== Lineup Tracker"
for team in (book.home, book.away):
    lastSkaters = ()
    def startsInBoxOld(x): # old standard was "0" if started in box
        return x == 0 or x == '0'
    def startsInBox(x):
        return x == "s" or x == "S"
    for period in (0,1):
        print "---",team.teamName,"Period",period+1
        lu = team.lineupTracker[period]
        for i in range(1,book.version.maxNumJams):
            skaters = lu.LineupForJam(i)
            if not skaters:
                break
            if verbose:
                print i, skaters
            for skater in skaters:
                team.CheckIBRF(skater,period+1,i)
            knownskaters = []
            for skater in skaters:
                if skater == '???':
                    continue
                if skater == '':
                    continue
                if skater in knownskaters:
                    Error("Duplicate in lineup", i,":",skaters)
                knownskaters.append(skater)
            boxPasses = lu.PenaltyPasses(i)
            if not boxPasses:
                lastSkaters = skaters
                continue
            for skater in skaters:
                if skater == '???' or skater == '':
                    continue
                box = boxPasses[skater]
                if not box:
                    continue
                if verbose: 
                    print i,skater,box
                if startsInBox(box[0][0]):
                    if skater not in lastSkaters:
                        # this could also be the case were the last skater fouled out and we sub
                        Error(skater,"started jam",i,"in the box, but wasn't in previous jam")
            lastSkaters = skaters

# Now make cross-sheet comparisons
print "=== Score vs Lineup Tracker"
for team in (book.home, book.away):
    for period in (0,1):
        print "---",team.teamName,"Period",period+1
        lu = team.lineupTracker[period]
        sk = team.scorekeeper[period]
        #print "Jam Index",sk.buildJamIndex()
        for i in range(1,book.version.maxNumJams):
            pt = sk.JamPoints(i)
            skaters = lu.LineupForJam(i)
            if pt is None and skaters is None:
                break
            if pt is None:
                Error("Jam",i,"exists in the lineup but not score")
                continue
            if skaters is None:
                Error("Jam",i,"exists in the score but not lineup")
                continue
            if skaters[4] != sk.Jammer(i):
                Error("Jam",i,"has",skaters[4],"as jammer in lineup, but",sk.Jammer(i),"in score")
print "=== Penalty vs Lineup Tracker"
for team in (book.home, book.away):
    firstMinor = [0] * 20
    firstMajor = [0] * 20
    lastMajor = [7] * 20
    for period in (0,1):
        print "---",team.teamName,"Period",period+1
        pt = team.penaltyTracker[period]
        lu = team.lineupTracker[period]
        for i in range(0,len(team.skaters)):
            skNum = team.skaters[i].number
            for pi in range(0,28):
                penalty = pt.MinorPenalty(i, pi)
                if not penalty:
                    continue
                jamNumber = penalty[1]
                skaters = lu.LineupForJam(jamNumber)
                if skaters and skNum not in skaters:
                    Error("Skater",skNum,"got minor penalty (",penalty[0],") in jam",jamNumber,"but not in lineup for that jam")
            for pi in range(0,7):
                penalty = pt.MajorPenalty(i, pi)
                if not penalty:
                    continue
                jamNumber = penalty[1]
                skaters = lu.LineupForJam(jamNumber)
                if skaters and skNum not in skaters:
                    Error("Skater",skNum,"got major penalty (",penalty[0],") in jam",jamNumber,"but not in lineup for that jam")
            
            
print "=== Score vs Penalty"

print NumErrors, "Errors"
print NumWarnings, "Warnings"
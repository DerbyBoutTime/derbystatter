[![CircleCI](https://circleci.com/gh/WFTDA/derbystatter.svg?style=svg)](https://circleci.com/gh/WFTDA/derbystatter)

derbystatter: Python tools for working with WFTDA statbook spreadsheets
=======================================================================

These tools can read Excel WFTDA statbook spreadsheets, in either xlsx or xls
format and report back possible errors such as duplicate skaters in lineups,
jam mismatches on penalties, and more.

Requirements
------------

- [xlrd](http://pypi.python.org/pypi/xlrd) - necessary for XLS support
- [py.test](http://pytest.org)
- [tox](https://testrun.org/tox/latest)
- xlsx.py (included) - adds limited xlsx capability

statbook.py
-----------

`statbook.py` provides the basic module to read the file as an object (instance
of class StatBook). The StatBook object includes direct references to the
following sheets:

- IBRF
- Score
- Penalty
- LineUp
- JamTimer
- Actions
- Errors

It creates other objects for the jamTimer (one per period) and the home and
away teams. The teams then hold objects that reference the data contained in
the various sheets (with scorekeeper, penaltyTracker, lineupTracker, actions,
and error, each a tuple of two objects, one per period). If there isn't a
lineup tracker sheet, it will attempt to use the scorekeeper sheet and the
penalty sheets to attempt to estimate the lineup (taking the jammer from the
scorekeeper, and then searching through all the penalties to find any record
skaters in the desired jam).

Note: Teams contain an array of objects for each period (so the period index
will be 0..1 since it is an array subscript). Jam numbers are 1 based (when
passed as a parameter).

Installation
------------

It's recommended to use and create a [virtualenv](https://virtualenv.pypa.io/)
to use with this project. This is easiest with
[virtualenvwrapper](https://virtualenvwrapper.readthedocs.io/) and the included Makefile but not required.

    $ mkvirtualenv statter
    (statter) make deps
    Handling requirements...
    ... installs dependencies ...
    Installing test deps...
    ... installs test dependencies ...

    # Run the tests with make test or tox
    $ make test
    $ tox

To install on your system without virtualenv or make, you may use pip directly.

    $ pip install -r requirements.txt

    # (optional) install the test requirements
    $ pip install -r test-requirements.txt

Usage
-----

To simply check for errors in a statsbook, use the `brecre.py` helper

    python derbystatter/brecre.py --help
    Usage: brecre.py [options] arg

    Options:
      -h, --help            show this help message and exit
      -f FILE, --file=FILE  write report to FILE
      -v, --verbose         print detailed progress

Run a basic statsbook check

    python derbystatter/brecre.py STATS-20160806-ToasterCity-WaterTown.xlsx

Run a verbose check and write the output into a file instead of the Terminal

    python derbystatter/brecre.py --verbose --file my_report.txt STATS-20160806-ToasterCity-WaterTown.xlsx

Contributors
------------

derbystatter is the result of much care and thought by Glen Andreas. A few small improvements have been added by dedi hubbard on behalf of the WFTDA.

License
-------

Released under the GNU GENERAL PUBLIC LICENSE, Version 3, 29 June 2007.

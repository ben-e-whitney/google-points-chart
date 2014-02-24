//Last updated 2014-02-24 by Ben Whitney // ben.e.whitney@post.harvard.edu
var pseudoGlobalsBuilder = function() {
  //Better to deal with a hard-coded value than to crowd the named data ranges with one of these for each cycle.
  this.POINTS_TOTAL_CELL = 'G42';
  this.LINK_TO_CHART = 'http://www.hcs.harvard.edu/~dudcoop/points/';

  //These values are larger than necessary so that new values may be easily added. columnsToPropertyValues calls removeEmptyRows,
  //which deals with the unused rows.
  this.MAX_COOPERS  = 50;
  this.MAX_CHORES   = 50;
  this.MAX_CARRIERS = 10;

  //These are 'actual' rows, like 'row 27' in the Google Spreadsheet, where people sign up on the Chart sheets.
  //Nothing fancy is going on here. I'm just avoiding writing out the array by hand.
  this.DAILY_CHORES_ROWS = [];
  for (var i = 0; i < 13; i++) {
    if (i < 12) {
    //MCU 1 through MCU 2. Chores that happen every day.
      this.DAILY_CHORES_ROWS[i] = 3+2*i;
    } else {
    //Kitchen Deep Clean: Hummus, Knives, etc. Note the change in spacing.
      this.DAILY_CHORES_ROWS[i] = 25+1+3*(i-11);
    }
  }
  //Brunch (currently cancelled and erased) and Bathrooms.
  this.WEEKLY_CHORES_ROWS = [32, 35, 38];
  this.ALL_CHORES_COLS = [];
  for (var col = 2; col < 16; col++) {
    this.ALL_CHORES_COLS.push(col);
  }

  //In a lot of the below we use a top row of 2 to avoid the column headings. This makes our bottom row one lower than
  //you might expect. Feel free to redo this all with named data ranges. They are quite nice, and you don't have to dig
  //in the scripts to find where things are pulled from.
  this.SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

  this.BASIC_DATA_SHEET = this.SPREADSHEET.getSheetByName('Basic Data');
  //This next section is organized according to the columns in the 'Basic Data' sheet, from left to right. This lets us
  //use the lengths of the *_COLUMNS variables to allow for inserted columns. Hopefully this makes things less and not
  //more confusing for you. This part would be simpler with named data ranges.
  //To reduce risk of confusion, keep the entries of the *_COLUMNS variables the exact same as the corresponding column headers.
  var leftCol = 1;
  var numCols = 3;
  this.BASIC_DATA = this.BASIC_DATA_SHEET.getRange(1, leftCol, 1+this.MAX_COOPERS, numCols).getValues();
  this.loanBalanceCol = leftCol+this.BASIC_DATA[0].indexOf('Loan Balance');
  this.BASIC_DATA = POINTS_UTILITIES.headerToPropertyNames(this.BASIC_DATA);

  leftCol += numCols;
  numCols = 4;
  this.CHORES_INFO = POINTS_UTILITIES.headerToPropertyNames(this.BASIC_DATA_SHEET.getRange(1, leftCol, 1+this.MAX_CHORES, numCols).getValues());

  leftCol += numCols;
  numCols = 2;
  this.CARRIERS_INFO = POINTS_UTILITIES.headerToPropertyNames(this.BASIC_DATA_SHEET.getRange(1, leftCol, 1+this.MAX_CARRIERS, numCols).getValues());

  //Because it will be so much easier to work with, going to define another object that is organized differently from BASIC_DATA and CHORES_INFO.
  this.GATEWAYS = {};
  for (var i = 0; i < this.CARRIERS_INFO.length; i++) {
    this.GATEWAYS[this.CARRIERS_INFO[i]['Wireless Carrier']] = this.CARRIERS_INFO[i]['Gateway'];
  }

  //This is used for when pulling the main Chart, as well as the Point Values 'behind' it. It will need to be changed if
  //further rows of chores are added, or a cycle becomes a full month, or something. Note that to simply add another row
  //of charts you should change the DAILY_CHORES_ROWS or WEEKLY_CHORES_ROWS variables. numRows is the number of rows
  //(starting at topRow) that get pulled. It will need to be changed if the newly added chore row lies outside the area
  //getting pulled now between DAILY_CHORES_ROWS and WEEKLY_CHORES_ROWS.
  this.CHART_RANGE_INFO    = {leftCol:1, numCols:15, topRow:1, numRows:42};
  //These are used when writing how many points each co-oper has signed up for or off on.
  this.SIGN_UP_RANGE_INFO  = {leftCol:3, numCols: 1, topRow:2, numRows:this.MAX_COOPERS};
  this.SIGN_OFF_RANGE_INFO = {leftCol:5, numCols: 1, topRow:2, numRows:this.MAX_COOPERS};

  //This next section is dedicated to values needed when sending reminders (not through Google Calendar).
  leftCol = 1;
  numCols = 5;
  this.SETTINGS = POINTS_UTILITIES.headerToPropertyNames(this.SPREADSHEET.getSheetByName('Settings').getRange(1, leftCol, 1+this.MAX_COOPERS, numCols).getValues());
  this.SETTINGS = POINTS_UTILITIES.addProperty(this.SETTINGS, {
    'Sign Up Name (Simplified)':     POINTS_UTILITIES.extractArray(this.SETTINGS, 'Sign Up Name').map(POINTS_UTILITIES.simplify),
    'Sign Off Initials (Simplified)': POINTS_UTILITIES.extractArray(this.SETTINGS, 'Sign Off Initials').map(POINTS_UTILITIES.simplify)
  });

  leftCol = 1;
  numCols = 8;
  this.CONTACT_INFO = POINTS_UTILITIES.headerToPropertyNames(this.SPREADSHEET.getSheetByName('Contact Info').getRange(1, leftCol, 1+this.MAX_COOPERS, numCols).getValues());
  if (this.CONTACT_INFO.length != this.SETTINGS.length) {
    throw new Error('CONTACT_INFO and SETTINGS have different lengths.');
  }
  this.COOPER_INFO = [];
  var indices = [];
  for (var i = 0; i < this.CONTACT_INFO.length; i++) {
    this.COOPER_INFO[i] = POINTS_UTILITIES.mergeObjects(this.CONTACT_INFO[i], this.SETTINGS[i]);
    indices[i] = i;
  }
  this.COOPER_INFO = POINTS_UTILITIES.addProperty(this.COOPER_INFO, {'Index': indices});

  //Increase this value after locking old cycles. As long as findChoreSums is running quickly, you can probably get away with never changing this.
  this.firstUnlockedCycle = 1;
  this.START_DATE = new Date(this.SPREADSHEET.getSheetByName('Chart 1').getRange('B2').getValue());
  this.START_DATE.setHours(0, 0, 0, 0);

  this._Cooper = function(queryOperator, that) {
    //queryOperator is styled after the objects used in querying CacheService.
    queryOperatorKeys = Object.keys(queryOperator);
    queryOperatorKey  = queryOperatorKeys[0];
    if (queryOperatorKeys.length != 1 || Object.keys(that.COOPER_INFO[0]).indexOf(queryOperatorKey) == -1) {
      throw new Error('Inappropriate queryOperator '+String(queryOperator)+'.');
    }
    this.index = POINTS_UTILITIES.objectIndexOf(that.COOPER_INFO, queryOperatorKey, queryOperator[queryOperatorKey]);
    this.valid = this.index != -1;
    if (!this.valid) {
      return this;
    }
    var basicData = that.BASIC_DATA[this.index];
    var info = that.COOPER_INFO[this.index];
    this.originalName     = info['Sign Up Name'];
    this.originalInitials = info['Sign Off Initials'];
    //TODO: this is to prevent the chart from catching all the blank cells by identifying them with one of these
    //co-opers. Throwing an error seems like overkill.
    if (!this.originalName || !this.originalInitials) {
      throw new Error('Co-oper whose '+queryOperatorKey+' is \''+queryOperator[queryOperatorKey]+
                      '\' (index '+String(this.index)+') needs a name or initials.'); 
    }
    this.simplifiedName     = POINTS_UTILITIES.simplify(this.originalName);
    this.simplifiedInitials = POINTS_UTILITIES.simplify(this.originalInitials);
    this.wantsEmailReminders = 'yes' == info['Do you want email reminders?'];
    this.wantsTextReminders  = 'yes' == info['Do you want text reminders?'];
    this.fullName = info['Co-oper'];
    this.emailAddress = info['Email Address'];
    //Remove all non-digit characters.
    this.phoneNumber = String(info['Cell Phone Number']).replace(/\D/g,'');
    this.textAddress = this.phoneNumber+'@'+that.GATEWAYS[info['Wireless Carrier']];
    //TODO: scrappy. Not sure this is the best way to do this. Maybe error handling in addToGoogleContacts?
    //Will this mess up birthday wishes?
    if (info['Birthday']) {
      this.birthday = new Date(info['Birthday']);
    } else {
      this.birthday = null;
    }
    //Not defining these methods in the prototype because I ran into problems trying to get the 'this' keyword to refer to
    //the Cooper object calling the method.
    this.loanBalanceRow = 2+this.index;
    Object.defineProperty(this, 'loanBalance', {
      get: function() {
        return Number(basicData['Loan Balance']);
      }.bind(this),
      set: function(newBalance) {
        that.BASIC_DATA_SHEET.getRange(this.loanBalanceRow, that.loanBalanceCol).setValue(Number(newBalance));
        //Allows us to cheaply get the value (say for display) after setting it without storing it to some variable.
        that.BASIC_DATA[this.index]['Loan Balance'] = newBalance;
      }});
  };

  this._SignUpCell = function(row, col, chart, pointValues, that) {
  //Object to contain anything we might want to know about the sign up cell.
    //Adding these as properties might be a holdover from when I was trying to define some methods in the prototype.
    //Feel free to fiddle with it.
    if (that.ALL_CHORES_COLS.indexOf(col) == -1 ||
      (that.DAILY_CHORES_ROWS.indexOf(row) == -1 && that.WEEKLY_CHORES_ROWS.indexOf(row) == -1)) {
      throw new Error('Invalid row ('+String(row)+') or column ('+String(col)+').');
    }
    this.row = Number(row);
    this.col = Number(col);
    this.choreName = (function() {
      if (that.DAILY_CHORES_ROWS.indexOf(row) != -1) {
      //The shared 1 in row_offset between the two cases is due to differences in indexing.
        if (row <= 25) {
          //MCU 1 through MCU 2.
          var row_offset = 1;
          var col_index  = 0;
        } else {
          //For Kitchen Deep Clean.
          var row_offset = 2;
          var col_index  = col-1;
        }
      } else {
        //that.WEEKLY_CHORES_ROWS.indexOf(row) != -1.
        var row_offset = 2;
        var col_index  = col-1;
      }
      return String(pointValues[row-row_offset][col_index]).replace('\n', ' ');
    })();
    this.choreDate  = (function() {
      var row_index = 1;
      if (row <= 29) {
        var col_index = col-1;
      } else {
        var col_index = 7*(1+Number(col > 8));
      }
      return new Date(chart[row_index][col_index]);
    })();
    this.choreValue = Number(pointValues[row-1][col-1]);
    this.valid = this.choreName != '' && this.choreValue != 0;
    if (!this.valid) {
      return this;
    }
    this.simplifiedNameOnSheet     = POINTS_UTILITIES.simplify(chart[row-1][col-1]);
    this.simplifiedInitialsOnSheet = POINTS_UTILITIES.simplify(chart[row][col-1]);
    this.signUpCooper  = that.Cooper({'Sign Up Name (Simplified)': this.simplifiedNameOnSheet});
    this.signOffCooper = that.Cooper({'Sign Off Initials (Simplified)': this.simplifiedInitialsOnSheet});
    //As in the Cooper constructor, these aren't being defined in the protoype because I couldn't figure out how to
    //specify that the 'this' keyword should refer to the SignUpCell calling the getter.
    this.diagnosis = (function() {
    //Does some analysis and looks for problems in the name/initials pair.
    //Sanity checks:
    //  diagnosis.noName implies !this.signUpCooper.valid (as long as emptyValue isn't in that.SIGN_UP_NAMES).
    //  diagnosis.noInitials implies !this.signOffCooper.valid (as long as emptyValue isn't in that.SIGN_OFF_INITIALS).
    //  diagnosis.sameCooper implies this.signUpCooper.valid and this.signOffCooper.valid.
    //So, if the same person signs up and off on a chore, this.signOffCooper.valid will be true.
      var voidValue = POINTS_UTILITIES.simplify('VOID');
      //Made a variable in case you want a default value like 'Pick Me' in the chart or whatever.
      var emptyValue =  '';
      var diagnosis = {};
      diagnosis.voided     = this.simplifiedNameOnSheet == voidValue || this.simplifiedInitialsOnSheet == voidValue;
      diagnosis.noName     = this.simplifiedNameOnSheet == emptyValue;
      diagnosis.noInitials = this.simplifiedInitialsOnSheet == emptyValue;
      //Can't test this.signUpCooper == this.signOffCooper. See <http://stackoverflow.com/questions/201183/how-do-you-determine-equality-for-two-javascript-objects>.
      diagnosis.sameCooper = this.signUpCooper.valid && this.signOffCooper.valid && this.signUpCooper.index == this.signOffCooper.index;
      return diagnosis;
    }).bind(this)();
  this.chore = that.Chore(this.choreName, this.choreDate);
  };

  this._Chore = function(choreName, choreDate, that) {
    //Note that these values aren't simplified. Just keep the Point Values sheets neat.
    this.name = choreName;
    this.nameIndex = POINTS_UTILITIES.objectIndexOf(that.CHORES_INFO, 'Chore Name', this.name);
    if (this.nameIndex == -1) {
      throw new Error(choreName+" wasn't found in the chores list.");
    }
    var choreInfo = that.CHORES_INFO[this.nameIndex];
    this.startDate = (function() {
      var startTime = new Date(choreDate);
      //.getValues returns a cell entry like '01/01/2000' as new Date('01/01/2000') or something.
      //choreInfo['Start Time']'s day is 1899-12-30, I think. It gets the intended time, which
      //is all we need.
      startTime.setMinutes(choreInfo['Start Time'].getMinutes());
      startTime.setHours(choreInfo['Start Time'].getHours());
      startTime.setDate(startTime.getDate()+choreInfo['Day Offset']);
      return startTime;
      }).apply(this)
    this.stopDate = (function() {
        var stopTime = new Date(this.startDate);
        stopTime.setHours(stopTime.getHours()+Math.floor(choreInfo['Duration']/60));
        stopTime.setMinutes(stopTime.getMinutes()+choreInfo['Duration']%60);
        return stopTime;
      }).apply(this);
  };

  //Important: do not use an additional 'new' keyword when calling these constructors. If you run
  //  var friend = new PSEUDO_GLOBALS.Cooper({'Sign Up Name': 'Alex'});
  //then, to my understanding, 'new' overrules the 'this' binding, meaning you won't be able to access
  //this._Cooper (because 'this' will refer to the new Cooper object, not PSEUDO_GLOBALS).
  this.Cooper = (function(queryOperator) {
    return new this._Cooper(queryOperator, this);
  }).bind(this);
  this.SignUpCell = (function(row, col, chart, pointValues) {
    return new this._SignUpCell(row, col, chart, pointValues, this);
  }).bind(this);
  this.Chore = (function(choreName, choreDate) {
    return new this._Chore(choreName, choreDate, this);
  }).bind(this);

  //  "Unfortunately Apps Script only supports JavaScript 1.6 syntax, which doesn't include the
  //  yield statement (it was introduced in 1.7)."
  //See <http://code.google.com/p/google-apps-script-issues/issues/detail?id=418>.
  this.allCoopers = (function(that) {
    return {
      index: 0,
      hasNext: function() {return this.index < that.COOPER_INFO.length;},
      next:    function() {return that.Cooper({'Co-oper': that.COOPER_INFO[this.index++]['Co-oper']});},
      rewind:  function() {this.index = 0;}
    };
  })(this);

  this.cycleSignUpCells = function(cycleNum) {
    return (function(cycleNum, that) {
      return {
        chart:       POINTS_UTILITIES.retrieveRange('Sign Up Data', cycleNum),
        pointValues: POINTS_UTILITIES.retrieveRange('Point Values', cycleNum),
        daySignUpCells: undefined,
        hasNext: function() {
          //For first execution.
          if (!this.daySignUpCells) {
            this.rewind();
          }
          var x = this.daySignUpCells.hasNext(); 
          if (!(this.j < that.ALL_CHORES_COLS.length-1)) {
          }
          return x || this.j < that.ALL_CHORES_COLS.length-1;},
        next: function() {
          if (this.daySignUpCells.nextValue) {
            return this.daySignUpCells.next();
          } else {
            this.j++;
            if (!(this.j < that.ALL_CHORES_COLS.length)) {
            } 
            this.daySignUpCells = that.daySignUpCells(that.ALL_CHORES_COLS[this.j], this.chart, this.pointValues);
            //Remember that we must always call hasNext before calling next!
            this.daySignUpCells.hasNext();
            return this.next();
          }
        },
        rewind: function() {
          this.j = 0;
          this.daySignUpCells = that.daySignUpCells(that.ALL_CHORES_COLS[this.j], this.chart, this.pointValues);
        },
      };
    })(cycleNum, this);
  };

  this.daySignUpCells = function (col, chart, pointValues) {
    return (function(col, chart, pointValues, that) {
      //Two bits of trickiness go on here: invalid SignUpCells are skipped, and weekly chores are included
      //if the column corresponds to a Saturday.
      return {
        orderedPairs: (function() {
          pairs = [];
          that.DAILY_CHORES_ROWS.forEach(function(row) {
            pairs.push([row, col]);
          });
          //Check whether it is a Saturday.
          if (col == 8 || col == 15) {
            that.WEEKLY_CHORES_ROWS.forEach(function(row) {
              for (var col_offset = 0; col_offset > -7; col_offset--) {
                pairs.push([row, col+col_offset]);
              }
            });
          }
          return pairs;
        })(),
        //hasNext will call findNext, which will increment i by 1.
        i: -1,
        nextValue: undefined,
        nextValueChanged: false,
        hasNext: function() {
          if (typeof this.i === 'undefined') {
            this.rewind();
          }
          this.nextValue = this.findNext();
          if (!this.nextValue) {
          }
          this.nextValueChanged = true;
          return Boolean(this.nextValue);
        },
        next: function() {
          if (!this.nextValueChanged) {
            throw new Error('nextValue not advanced between calls to next.');
          }
          this.nextValueChanged = false;
          return this.nextValue;
        },
        findNext: function() {
          this.i++;
          if (this.orderedPairs.length > 30 && this.i > 30) {
          }
          if (!(this.i < this.orderedPairs.length)) {
            return false;
          }
          var pair = this.orderedPairs[this.i];
          if (!pair[1]) {
          }
          var signUpCell = that.SignUpCell(pair[0], pair[1], chart, pointValues);
          if (signUpCell.valid) {
            return signUpCell;
          } else {
            return this.findNext();
          }
        },
        rewind: function() {
          this.i = -1;
          this.nextValue = undefined;
          this.nextValueChanged = false;
        }
      };
    })(col, chart, pointValues, this);}
  this.POINTS_STEWARD = this.Cooper({'Stewardship': 'Points'});
  //TODO: check that POINTS_STEWARD has a vaild email address. This is relied on in
  //POINTS_UTILITIES.sendEmail.
  return this;
};

function pseudoGlobalsFetcher() {
  function updateTimeSensitiveProperties(pseudoGlobals) {
    pseudoGlobals.todayIs = new Date();
    //Be very careful with Dates. First off, if you do
    //  var todayIs = new Date();
    //  var alt_todayIs = new Date(todayIs);
    //then todayIs != alt_todayIs because they are different objects. And if we do
    //  todayIs.setHours(0, 0, 0, 0);
    //  alt_todayIs.setHours(0, 0, 0, 0);
    //then todayIs.getTime() == alt_todayIs.getTime() and todayIs.valueOf() == alt_todayIs.valueOf(), but the Google Scripts
    //debugger is saying that (right now, at least) todayIs == Date (70349511) and alt_todayIs == Date (70349512). So it would
    //seem that their milliseconds differ. Previously I got this working with Utilities.formatDate, so try that if you have a problem.
    //'Zero out' the date times to midnight (same is done to PSEUDO_GLOBALS.START_DATE).
    pseudoGlobals.todayIs.setHours(0, 0, 0, 0);
    //Find the difference in milliseconds between todayIs and START_DATE, and convert that to difference in fortnights. Find the floor
    //and add it to 1 (we start counting the Charts at 1).
    //TODO: figure out what you want to do here. There will be a lost of this, or there will be a ton of reshuffling.
    //JSON doesn't support Dates, so START_DATE is a string at the moment.
    pseudoGlobals.START_DATE = new Date(pseudoGlobals.START_DATE);
    pseudoGlobals.currentCycleNum = parseInt(1+Math.floor((pseudoGlobals.todayIs.getTime()-pseudoGlobals.START_DATE.getTime())/(1000*60*60*24*14)));
    return pseudoGlobals;
  };

  //TODO: get rid of this once you decide how this should work.
  return updateTimeSensitiveProperties(new pseudoGlobalsBuilder());

  var publicCache = CacheService.getPublicCache();
  //Important: Whenever we write to the spreadsheet, we MUST make the accompanying changes to the actual arrays!! Search absolutely everywhere for .setValues.
  //I've checked over all this -- include a HUGE note about it, but you seem to be good.
  //TODO: document this issue.
  if (typeof PSEUDO_GLOBALS !== 'undefined') {
    //We have to periodically write back into the cache so that the key/value pair doesn't expire.
    publicCache.put('PSEUDO_GLOBALS', Utilities.jsonStringify(PSEUDO_GLOBALS), 6*60*60);
    return updateTimeSensitiveProperties(PSEUDO_GLOBALS);
  }
  PSEUDO_GLOBALS = publicCache.get('PSEUDO_GLOBALS')
  if (PSEUDO_GLOBALS) {
    PSEUDO_GLOBALS = Utilities.jsonParse(PSEUDO_GLOBALS);
  } else {
    PSEUDO_GLOBALS = new pseudoGlobalsBuilder();
    //The third argument is how long (in seconds) before the key/value pair expires. Six hours is the maximum.
    //It doesn't matter that we haven't change the time-sensitive properties, since we will update them whenever
    //we retrieve this value from the cache.
    publicCache.put('PSEUDO_GLOBALS', Utilities.jsonStringify(PSEUDO_GLOBALS), 6*60*60);
  }
  return updateTimeSensitiveProperties(PSEUDO_GLOBALS);
}

var UI_BUILDERS = {
  DEFAULTS: {
    submitButtonHeight: 25,
    submitButtonWidth: 75,
    shortUserInputLabelWidth: 100,
    mediumUserInputLabelWidth: 200,
    longUserInputLabelWidth: 300,
    textBoxWidth: 150,
    numberTextBoxWidth: 25,
    radioButtonYesNoWidth: 50,
    hexColors: {
      red   : '#FF0000',
      purple: '#C400FF',
      black : '#000000',
      grey  : '#d3d3d3'
    }
  },

  getCalendarPreferences: function() {
    var googleReminderTypes = [
      {name:'popup', defaultValue:5},
      {name:'text',  defaultValue:15},
      {name:'email', defaultValue:60}
    ];
    var overallWidth = 1.25*(this.DEFAULTS.mediumUserInputLabelWidth+2*this.DEFAULTS.radioButtonYesNoWidth);
    var overallHeight = 250;

    var uiAppInstance = UiApp.createApplication()
      .setWidth(overallWidth)
      .setHeight(1.1*overallHeight);
    var formPanel = uiAppInstance.createFormPanel()
      .setSize(String(overallWidth)+'px', String(overallHeight)+'px');
    var organizingPanel = uiAppInstance.createVerticalPanel()
      .setWidth(String(overallWidth)+'px')
      .setHeight(String(overallHeight)+'px')
      .setSpacing(10);
    var inputGrid = uiAppInstance.createGrid(1+2*googleReminderTypes.length, 2)
      .setWidth(String(overallWidth)+'px')
      .setCellSpacing(0)
      .setCellPadding(0);
    var submitAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(1.3*this.DEFAULTS.submitButtonHeight)+'px');
    var submitButton = uiAppInstance.createSubmitButton('Submit')
      .setSize(String(this.DEFAULTS.submitButtonWidth)+'px', String(this.DEFAULTS.submitButtonHeight)+'px');
    var hiddenIdentifier = uiAppInstance.createTextBox()
      .setName('identifier')
      .setValue('UI_BUILDERS.getCalendarPreferences')
      .setVisible(false);
    var reminderWarning = uiAppInstance.createLabel('Reminders must come 5 minutes to 1 day before the chore.')
      .setWidth(String(this.DEFAULTS.mediumUserInputLabelWidth)+'px')
      .setStyleAttribute('color', 'red')
      .setVisible(false);
    submitAbsolutePanel.add(submitButton, overallWidth-1.2*this.DEFAULTS.submitButtonWidth, 0.2*this.DEFAULTS.submitButtonHeight)
      .add(reminderWarning, 0, 0)
      .add(hiddenIdentifier, 0, 0);

    inputGrid.setWidget(0, 0, uiAppInstance.createLabel("What's your name?")
        .setWidth(String(this.DEFAULTS.mediumUserInputLabelWidth)+'px'));
    //TODO: see what happens if you drop the width setting here.
    var nameListBox = uiAppInstance.createListBox()
      .setWidth(String(2*this.DEFAULTS.radioButtonYesNoWidth)+'px')
      .setName('userName');
    PSEUDO_GLOBALS.SETTINGS.forEach(function(entry) {nameListBox.addItem(entry['Sign Up Name']);});
    inputGrid.setWidget(0, 1, nameListBox);

    var minutesLabels    = [];
    var minutesTextBoxes = [];
    var radioMiniGrids   = [];
    var changeHandlers   = [];
    var radioRecorders   = [];
    for (var i = 0; i < googleReminderTypes.length; i++) {
      var remType = googleReminderTypes[i];
      inputGrid.setWidget(1+2*i, 0, uiAppInstance.createLabel('Would you like '+remType.name+' reminders?')
        .setWidth(String(this.DEFAULTS.mediumUserInputLabelWidth)+'px')
        .setWordWrap(true));
      minutesLabels[i] = uiAppInstance.createLabel('How many minutes in advance of your chores would you like '+
        remType.name+' reminders?')
        .setWidth(String(this.DEFAULTS.mediumUserInputLabelWidth)+'px')
        .setStyleAttribute('color', UI_BUILDERS.DEFAULTS.hexColors.grey)
        .setWordWrap(true);
      minutesTextBoxes[i] = uiAppInstance.createTextBox()
        .setValue(remType.defaultValue)
        .setName(remType.name+'TextBox')
        .setWidth(String(2*this.DEFAULTS.radioButtonYesNoWidth)+'px')
        .setStyleAttribute('textAlign', 'right')
        .setEnabled(false)
        //For when an invalid value has been fixed: enable the submit button and hide the error message.
        .addChangeHandler(uiAppInstance.createClientHandler()
          .forTargets(submitButton)
          .setEnabled(true)
          .forTargets(reminderWarning)
          .setVisible(false));
      //See <https://developers.google.com/apps-script/reference/calendar/calendar-event#addEmailReminder(Integer)> for range restriction.
      //I am going to be more strict that that and set the upper bound to 1 day. I think that sending reminders more than 1 day in
      //advance is asking for confusion. If someone wants, though, we can change it.
      changeHandlers[i] = uiAppInstance.createClientHandler()
        .validateNotRange(minutesTextBoxes[i], 5, 24*60)
        .forTargets(submitButton)
        .setEnabled(false)
        .forTargets(reminderWarning)
        .setVisible(true);
      minutesTextBoxes[i].addChangeHandler(changeHandlers[i]);
      inputGrid.setWidget(2+2*i, 0, minutesLabels[i]);
      inputGrid.setWidget(2+2*i, 1, minutesTextBoxes[i]);
      //TODO: if this works, explain why the workaround.
      //<https://code.google.com/p/google-apps-script-issues/issues/detail?id=506>
      radioRecorders[i] = uiAppInstance.createTextBox()
        .setName(remType.name+'RadioRecorder')
        //The radio buttons default to 'No', so this needs to default to 'no'. We're using a ClickHandler, so the default value will
        //be submitted if neither button is ever clicked.
        .setValue('no')
        .setVisible(false);
      submitAbsolutePanel.add(radioRecorders[i], 0, 0);
      radioMiniGrids[i] = uiAppInstance.createGrid(1, 2)
        .setWidth(String(2*this.DEFAULTS.radioButtonYesNoWidth)+'px')
        .setCellSpacing(0)
        .setCellPadding(0);
      radioMiniGrids[i].setWidget(0, 0, uiAppInstance.createRadioButton(remType.name+'RadioGroup', 'Yes')
          .setName(remType.name+'Group')
          .setWidth(String(this.DEFAULTS.radioButtonYesNoWidth)+'px')
          .setValue(false)
          .addClickHandler(uiAppInstance.createClientHandler()
            .forTargets(minutesLabels[i])
            .setStyleAttribute('color', UI_BUILDERS.DEFAULTS.hexColors.black)
            .forTargets(minutesTextBoxes[i])
            .setEnabled(true)
            .forTargets(radioRecorders[i])
            .setText('yes')))
        .setWidget(0, 1, uiAppInstance.createRadioButton(remType.name+'RadioGroup', 'No')
          .setName(remType.name+'Group')
          .setWidth(String(this.DEFAULTS.radioButtonYesNoWidth)+'px')
          .setValue(true)
          .addClickHandler(uiAppInstance.createClientHandler()
            .forTargets(minutesLabels[i])
            .setStyleAttribute('color', UI_BUILDERS.DEFAULTS.hexColors.grey)
            .forTargets(minutesTextBoxes[i])
            .setEnabled(false)
            .forTargets(radioRecorders[i])
            .setText('no')))
      inputGrid.setWidget(1+2*i, 1, radioMiniGrids[i]);
    }

    organizingPanel.add(inputGrid)
      .add(submitAbsolutePanel);
    formPanel.add(organizingPanel);
    uiAppInstance.add(formPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  getMessage: function(askingMessage, askingMessageHeight) {
    var textAreaHeight = 120;
    var overallWidth = UI_BUILDERS.DEFAULTS.longUserInputLabelWidth;
    var overallHeight = askingMessageHeight+textAreaHeight+1.2*UI_BUILDERS.DEFAULTS.submitButtonHeight+35;
    var uiAppInstance = UiApp.createApplication().setWidth(overallWidth)
      .setHeight(overallHeight);
    var formPanel = uiAppInstance.createFormPanel();
    var verticalPanel = uiAppInstance.createVerticalPanel()
      .setSpacing(10);
    verticalPanel.add(uiAppInstance.createLabel(askingMessage)
        .setHeight(String(askingMessageHeight)+'px')
        .setWidth(String(overallWidth)+'px')
        .setWordWrap(true))
      .add(uiAppInstance.createTextArea()
        .setName('textArea')
        .setWidth(String(0.9*overallWidth)+'px')
        .setHeight(String(textAreaHeight)+'px'));
    //verticalPanel.add(uiAppInstance.createSubmitButton("Submit"));
    var submitAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(1.2*UI_BUILDERS.DEFAULTS.submitButtonHeight)+'px');
    var submitButton = uiAppInstance.createSubmitButton('Submit')
      .setSize(String(UI_BUILDERS.DEFAULTS.submitButtonWidth)+'px', String(UI_BUILDERS.DEFAULTS.submitButtonHeight)+'px');
    var hiddenIdentifier = uiAppInstance.createTextBox()
      .setName('identifier')
      .setValue('UI_BUILDERS.getMessage')
      .setVisible(false);
    submitAbsolutePanel.add(submitButton, overallWidth-1.2*this.DEFAULTS.submitButtonWidth, 0.2*this.DEFAULTS.submitButtonHeight)
      .add(hiddenIdentifier, 0, 0);
    verticalPanel.add(submitAbsolutePanel);
    formPanel.add(verticalPanel);
    uiAppInstance.add(formPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  userSplitsName: function(namePieces) {
    switch (namePieces.length) {
    case 1:
      var firstNameGuess = namePieces[0];
      var middleNameGuess = '';
      var lastNameGuess = '';
      break;
    default:
      var firstNameGuess  = namePieces[0];
      var middleNameGuess = namePieces.slice(1, namePieces.length-1).join(' ');
      var lastNameGuess   = namePieces[namePieces.length-1];
    }
    var formRows = [
      {'labelName':'First Name(s):',  'textBoxName':'firstName',  'textBoxValue':firstNameGuess},
      {'labelName':'Middle Name(s):', 'textBoxName':'middleName', 'textBoxValue':middleNameGuess},
      {'labelName':'Last Name(s):',   'textBoxName':'lastName',   'textBoxValue':lastNameGuess}
    ];

    var introLabelHeight = 40;
    var overallWidth = 1.25*(this.DEFAULTS.shortUserInputLabelWidth+this.DEFAULTS.textBoxWidth);
    var overallHeight = 145;

    var uiAppInstance = UiApp.createApplication()
      .setWidth(overallWidth)
      .setHeight(1.1*overallHeight);
    var formPanel = uiAppInstance.createFormPanel()
      .setSize(String(overallWidth)+'px', String(overallHeight)+'px');
    var formPanel = uiAppInstance.createFormPanel()
      .setSize(String(overallWidth)+'px', String(overallHeight)+'px');
    var organizingGrid = uiAppInstance.createGrid(3, 1)
      .setWidth(String(overallWidth)+'px')
      //TODO: delete this if you can.
      .setHeight(String(overallHeight)+'px')
      .setCellSpacing(0)
      .setCellPadding(0);
    var introAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(introLabelHeight)+'px');
    var nameInputGrid = uiAppInstance.createGrid(3, 2)
      .setWidth(String(overallWidth)+'px')
      .setCellSpacing(0)
      .setCellPadding(0);
    var introLabel = uiAppInstance.createLabel("I'm not sure how to split "+namePieces.join(' ')+"'s name up. Please help.")
      .setSize(String(overallWidth)+'px', String(introLabelHeight)+'px');
    introAbsolutePanel.add(introLabel, 0, 0);

    for (var i = 0; i < formRows.length; i++) {
      var formRow = formRows[i];
      nameInputGrid.setWidget(i, 0, uiAppInstance.createLabel(formRow.labelName)
          .setWidth(String(this.DEFAULTS.shortUserInputLabelWidth)+'px'));
      //Text boxes look better IMO if the height is left to the default value.
      nameInputGrid.setWidget(i, 1, uiAppInstance.createTextBox()
          .setName(formRow.textBoxName)
          .setValue(formRow.textBoxValue)
          .setWidth(String(this.DEFAULTS.textBoxWidth)+'px'));
    }

    var submitAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(1.2*this.DEFAULTS.submitButtonHeight)+'px');
    var submitButton = uiAppInstance.createSubmitButton('Submit')
      .setSize(String(this.DEFAULTS.submitButtonWidth)+'px', String(this.DEFAULTS.submitButtonHeight)+'px');
    var hiddenIdentifier = uiAppInstance.createTextBox()
      .setName('identifier')
      .setValue('UI_BUILDERS.userSplitsName')
      .setVisible(false);
    submitAbsolutePanel.add(submitButton, overallWidth-this.DEFAULTS.submitButtonWidth, 0.2*this.DEFAULTS.submitButtonHeight)
      .add(hiddenIdentifier, 0, 0);

    organizingGrid.setWidget(0, 0, introAbsolutePanel)
      .setWidget(1, 0, nameInputGrid)
      .setWidget(2, 0, submitAbsolutePanel);
    formPanel.add(organizingGrid);
    uiAppInstance.add(formPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  getNameAndCycle: function(name, cycleNum) {
    var validCycleNums = [];
    for (var i = 1; i <= PSEUDO_GLOBALS.currentCycleNum; i++) {
      validCycleNums.push(i);
    }

    var overallWidth  = 2*1.5*this.DEFAULTS.shortUserInputLabelWidth;
    var overallHeight = 100;
    var uiAppInstance = UiApp.createApplication()
      .setWidth(overallWidth)
      .setHeight(overallHeight);
    var formPanel = uiAppInstance.createFormPanel()
      .setSize(String(overallWidth)+'px', String(overallHeight)+'px');
    var organizingPanel = uiAppInstance.createVerticalPanel()
      .setWidth(String(overallWidth)+'px')
      .setHeight(String(overallHeight)+'px')
      .setSpacing(10);
    var inputGrid = uiAppInstance.createGrid(2, 2)
      .setWidth(String(overallWidth)+'px')
      .setCellSpacing(0)
      .setCellPadding(0);

    inputGrid.setWidget(0, 0, uiAppInstance.createLabel("What's your name?")
        .setWidth(String(1.5*this.DEFAULTS.shortUserInputLabelWidth)+'px'));
    var nameListBox = uiAppInstance.createListBox()
      //.setWidth(String(2*this.DEFAULTS.radioButtonYesNoWidth)+'px')
      .setName('userName');
    PSEUDO_GLOBALS.SETTINGS.forEach(function(entry) {nameListBox.addItem(entry['Sign Up Name']);});
    if (name) {
      nameListBox.setSelectedIndex(POINTS_UTILITIES.objectIndexOf(PSEUDO_GLOBALS.SETTINGS, 'Sign Up Name', name));
    }
    inputGrid.setWidget(0, 1, nameListBox);

    inputGrid.setWidget(1, 0, uiAppInstance.createLabel('Pick a cycle.')
      .setWidth(String(1.5*this.DEFAULTS.shortUserInputLabelWidth)+'px'));
    var cycleNumListBox = uiAppInstance.createListBox()
      //.setWidth(String(2*this.DEFAULTS.radioButtonYesNoWidth)+'px')
      .setName('cycleNum')
    validCycleNums.forEach(function (validCycleNum) {cycleNumListBox.addItem(String(validCycleNum), validCycleNum);});
    if (cycleNum) {
      cycleNumListBox.setSelectedIndex(validCycleNums.indexOf(Number(cycleNum)));
    }
    inputGrid.setWidget(1, 1, cycleNumListBox);

    var submitAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(1.3*this.DEFAULTS.submitButtonHeight)+'px');
    var submitButton = uiAppInstance.createSubmitButton('Next')
      .setSize(String(this.DEFAULTS.submitButtonWidth)+'px', String(this.DEFAULTS.submitButtonHeight)+'px');
    submitAbsolutePanel.add(submitButton, overallWidth-1.2*this.DEFAULTS.submitButtonWidth, 0.2*this.DEFAULTS.submitButtonHeight);
    var hiddenIdentifier = uiAppInstance.createTextBox()
      .setName('identifier')
      .setValue('UI_BUILDERS.getNameAndCycle')
      .setVisible(false);
    submitAbsolutePanel.add(hiddenIdentifier, 0, 0);

    organizingPanel.add(inputGrid)
      .add(submitAbsolutePanel);
    formPanel.add(organizingPanel);
    uiAppInstance.add(formPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  getLoanInfo: function(cooper, cycleNum, oldLoan) {
    //The second term can't be this.DEFAULTS.numberTextBoxWidth because the ListBox is bigger than that.
    var overallWidth  = 1.25*(this.DEFAULTS.shortUserInputLabelWidth+this.DEFAULTS.mediumUserInputLabelWidth);
    var overallHeight = 150;
    var uiAppInstance = UiApp.createApplication()
      .setWidth(overallWidth)
      .setHeight(overallHeight);
    var formPanel = uiAppInstance.createFormPanel()
      .setSize(String(overallWidth)+'px', String(overallHeight)+'px');
    var organizingPanel = uiAppInstance.createVerticalPanel()
      .setWidth(String(overallWidth)+'px')
      .setHeight(String(overallHeight)+'px')
      .setSpacing(10);
    var introLabel = uiAppInstance.createLabel('OK '+String(cooper.originalName)+'. For Cycle '+String(cycleNum)+
        ' you are currently '+POINTS_UTILITIES.getLoanVerb(oldLoan)+' '+String(Math.abs(oldLoan))+' points. Your total balance is '+
        String(cooper.loanBalance)+' points.');

    var submitAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(1.3*this.DEFAULTS.submitButtonHeight)+'px');
    var buttonRecorder = uiAppInstance.createTextBox()
      .setName('buttonPressed')
      .setValue('')
      .setVisible(false);
    var previousSubmitButton = uiAppInstance.createSubmitButton('Previous')
      .setSize(String(this.DEFAULTS.submitButtonWidth)+'px', String(this.DEFAULTS.submitButtonHeight)+'px')
      .addClickHandler(uiAppInstance.createClientHandler()
        .forTargets(buttonRecorder)
        .setText('previous'));
    var nextSubmitButton = uiAppInstance.createSubmitButton('Next')
      .setSize(String(this.DEFAULTS.submitButtonWidth)+'px', String(this.DEFAULTS.submitButtonHeight)+'px')
      .addClickHandler(uiAppInstance.createClientHandler()
        .forTargets(buttonRecorder)
        .setText('next'));
    var hiddenIdentifier = uiAppInstance.createTextBox()
      .setName('identifier')
      .setValue('UI_BUILDERS.getLoanInfo')
      .setVisible(false);
    var loanAmountWarning = uiAppInstance.createHTML('Give a value in [0, 40].')
      .setStyleAttribute('color', UI_BUILDERS.DEFAULTS.hexColors.red)
      .setVisible(false);
    submitAbsolutePanel.add(previousSubmitButton, 0, 0.2*this.DEFAULTS.submitButtonHeight)
      //Setting the horizontal position to overallWidth-1*this.DEFAULTS.submitButtonWidth wasn't working.
      //Couldn't figure out why.
      .add(nextSubmitButton, overallWidth-1.2*this.DEFAULTS.submitButtonWidth, 0.2*this.DEFAULTS.submitButtonHeight)
      .add(hiddenIdentifier, 0, 0)
      .add(buttonRecorder, 0, 0)
      .add(loanAmountWarning, 0.25*overallWidth, 0.3*this.DEFAULTS.submitButtonHeight);

    var inputGrid = uiAppInstance.createGrid(2, 2)
      .setWidth(String(overallWidth)+'px')
      .setCellSpacing(0)
      .setCellPadding(0);
    var loanActionListBox = uiAppInstance.createListBox()
      .setName('loanAction')
      .setWidth(String(this.DEFAULTS.shortUserInputLabelWidth)+'px')
      .addItem('Take a loan')
      .addItem('Repay a loan');
    var loanAmountTextBox =  uiAppInstance.createTextBox()
      .setWidth(String(this.DEFAULTS.shortUserInputLabelWidth)+'px')
      .setName('loanAmount')
      .setStyleAttribute('textAlign', 'right')
      .setValue('10');
    var loanAmountValid = uiAppInstance.createClientHandler()
      .validateRange(loanAmountTextBox, 0, 40)
      .forTargets(loanAmountWarning)
      .setVisible(false)
      .forTargets(nextSubmitButton)
      .setEnabled(true);
    var loanAmountInvalid = uiAppInstance.createClientHandler()
      .validateNotRange(loanAmountTextBox, 0, 40)
      .forTargets(loanAmountWarning)
      .setVisible(true)
      .forTargets(nextSubmitButton)
      .setEnabled(false);
    loanAmountTextBox.addValueChangeHandler(loanAmountInvalid)
      .addValueChangeHandler(loanAmountValid);
    inputGrid.setWidget(0, 0, uiAppInstance.createLabel('What would you like to do?')
        .setWidth(String(this.DEFAULTS.mediumUserInputLabelWidth)+'px'))
      .setWidget(0, 1, loanActionListBox)
      .setWidget(1, 0, uiAppInstance.createLabel('How many points?')
        .setWidth(String(this.DEFAULTS.mediumUserInputLabelWidth)+'px'))
      .setWidget(1, 1, loanAmountTextBox);

    organizingPanel.add(introLabel)
      .add(inputGrid)
      .add(submitAbsolutePanel);
    formPanel.add(organizingPanel);
    uiAppInstance.add(formPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  getAbsenceInfo: function() {
    var validCycleNums = [];
    //TODO: it might be nice to bump this up so that you can set it before the next cycle. If not, note limitation.
    for (var i = 1; i <= PSEUDO_GLOBALS.currentCycleNum; i++) {
      validCycleNums.push(i);
    }

    var overallWidth  = 2*1.5*this.DEFAULTS.shortUserInputLabelWidth;
    var overallHeight = 130;
    var uiAppInstance = UiApp.createApplication()
      .setWidth(overallWidth)
      .setHeight(overallHeight);
    var formPanel = uiAppInstance.createFormPanel()
      .setSize(String(overallWidth)+'px', String(overallHeight)+'px');
    var organizingPanel = uiAppInstance.createVerticalPanel()
      .setWidth(String(overallWidth)+'px')
      .setHeight(String(overallHeight)+'px')
      .setSpacing(10);
    var inputGrid = uiAppInstance.createGrid(3, 2)
      .setWidth(String(overallWidth)+'px')
      .setCellSpacing(0)
      .setCellPadding(0);

    var nameLabel = uiAppInstance.createLabel("What's your name?")
      .setWidth(String(1.5*this.DEFAULTS.shortUserInputLabelWidth)+'px');
    var nameListBox = uiAppInstance.createListBox()
      .setName('userName');
    PSEUDO_GLOBALS.SETTINGS.forEach(function(entry) {nameListBox.addItem(entry['Sign Up Name']);});

    var cycleNumLabel = uiAppInstance.createLabel('Pick a cycle.')
      .setWidth(String(1.5*this.DEFAULTS.shortUserInputLabelWidth)+'px');
    var cycleNumListBox = uiAppInstance.createListBox()
      .setName('cycleNum');
    validCycleNums.forEach(function (validCycleNum) {cycleNumListBox.addItem(String(validCycleNum), validCycleNum);});

    var absenceLabel = uiAppInstance.createLabel('How many days will you be gone?')
      .setWidth(String(1.5*this.DEFAULTS.shortUserInputLabelWidth)+'px');
    var absenceAmountTextBox = uiAppInstance.createTextBox()
      .setWidth(String(this.DEFAULTS.shortUserInputLabelWidth)+'px')
      .setName('absenceAmount')
      .setValue('0')
      .setStyleAttribute('textAlign', 'right');
    var absenceAmountWarning = uiAppInstance.createLabel('Please give a value between 0 and 14.')
      .setWidth(String(1.5*this.DEFAULTS.shortUserInputLabelWidth)+'px')
      .setStyleAttribute('color', 'red')
      .setVisible(false);
    var submitButton = uiAppInstance.createSubmitButton('Submit')
      .setSize(String(this.DEFAULTS.submitButtonWidth)+'px', String(this.DEFAULTS.submitButtonHeight)+'px');

    var absenceAmountValid = uiAppInstance.createClientHandler()
      .validateRange(absenceAmountTextBox, 0, 14)
      .forTargets(absenceAmountWarning)
      .setVisible(false)
      .forTargets(submitButton)
      .setEnabled(true);
    var absenceAmountInvalid = uiAppInstance.createClientHandler()
      .validateNotRange(absenceAmountTextBox, 0, 14)
      .forTargets(absenceAmountWarning)
      .setVisible(true)
      .forTargets(submitButton)
      .setEnabled(false);
    absenceAmountTextBox.addValueChangeHandler(absenceAmountInvalid)
      .addValueChangeHandler(absenceAmountValid);

    inputGrid.setWidget(0, 0, nameLabel)
      .setWidget(0, 1, nameListBox)
      .setWidget(1, 0, cycleNumLabel)
      .setWidget(1, 1, cycleNumListBox)
      .setWidget(2, 0, absenceLabel)
      .setWidget(2, 1, absenceAmountTextBox);

    var submitAbsolutePanel = uiAppInstance.createAbsolutePanel()
      .setSize(String(overallWidth)+'px', String(1.3*this.DEFAULTS.submitButtonHeight)+'px');
    var hiddenIdentifier = uiAppInstance.createTextBox()
      .setName('identifier')
      .setValue('UI_BUILDERS.getAbsenceInfo')
      .setVisible(false);
    submitAbsolutePanel.add(submitButton, overallWidth-1.2*this.DEFAULTS.submitButtonWidth, 0.2*this.DEFAULTS.submitButtonHeight)
      .add(hiddenIdentifier, 0, 0)
      .add(absenceAmountWarning, 0, 0);

    organizingPanel.add(inputGrid)
      .add(submitAbsolutePanel);
    formPanel.add(organizingPanel);
    uiAppInstance.add(formPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  displayMessage: function(message, height) {
    var uiAppInstance = UiApp.createApplication()
      .setWidth(this.DEFAULTS.longUserInputLabelWidth)
      .setHeight(height);
    uiAppInstance.add(uiAppInstance.createLabel(message))
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  },

  displayMessageAndLink: function(message, linkText, linkURL, height) {
    var uiAppInstance = UiApp.createApplication().setWidth(this.DEFAULTS.longUserInputLabelWidth)
      .setHeight(height);
    var verticalPanel = uiAppInstance.createVerticalPanel()
      .setSpacing(10);
    verticalPanel.add(uiAppInstance.createLabel(message)
        .setWordWrap(true));
    var calendarAnchor = uiAppInstance.createAnchor(linkText, linkURL); 
    verticalPanel.add(calendarAnchor)
      .setCellHorizontalAlignment(calendarAnchor, UiApp.HorizontalAlignment.CENTER)
      .setCellVerticalAlignment(calendarAnchor, UiApp.VerticalAlignment.MIDDLE);
    uiAppInstance.add(verticalPanel);
    SpreadsheetApp.getActiveSpreadsheet().show(uiAppInstance);
  }
};

var POINTS_UTILITIES = {
  jsonNumber: function(input) {
    var defaultNum = 0;
    return isNaN(input) ? defaultNum : Number(input);
  },

  mergeObjects: function(objectOne, objectTwo) {
    var keysOne = Object.keys(objectOne);
    var keysTwo = Object.keys(objectTwo);
    var objectThree = {};
    keysOne.forEach(function(key) {
      if (key in keysTwo && objectOne[key] != objectTwo[key]) {
        throw new Error('Conflict at property '+String(key)+". Objects can't be merged.");
      }
      objectThree[key] = objectOne[key];
    });
    keysTwo.forEach(function(key) {objectThree[key] = objectTwo[key];});
    return objectThree;
  },

  addProperty: function(inputArray, propertyObject) {
    for (var i = 0; i < inputArray.length; i++) {
      for (var j = 0; j < Object.keys(propertyObject).length; j++) {
        var key = Object.keys(propertyObject)[j];
        inputArray[i][key] = propertyObject[key][i];
      }
    }
    return inputArray;
  },

  getLoanVerb: function(loan) {
    if (loan < 0) {
      return 'paying back';
    } else {
      return 'borrowing';
    }
  },

  removeEmptyRows: function(inputArray) {
  //Returns array with the 'empty' entries removed. An entry is considered 'empty' if it is an array of empty strings.
  //This corresponds to a row in the spreadsheet with all empty cells.
  //Apply to arrays of arrays.
    outputArray = [];
    for (var i = 0; i < inputArray.length; i++) {
      if (inputArray[i].some(Boolean)) {
        outputArray.push(inputArray[i]);
      }
    }
    return outputArray;
  },

  simplify: function(input) {
  //Strips whitespace and converts to lowercase.
    return input.trim().toLowerCase();
  },

  extractArray: function(inputArray, position) {
  //Slices multidimensional arrays, such as those produced by the .getValues() method.
  //Can also be used with arrays of objects. Then position should be a property name.
    outputArray = [];
    for (var i = 0; i < inputArray.length; i++) {
      outputArray[i] = inputArray[i][position];
    }
    return outputArray;
  },

  objectIndexOf: function(inputArray, propertyName, givenValue) {
  //Looks for an entry of inputArray (an object) whose propertyName value matches givenValue.
  //Mimics behavior of Array.indexOf method.
    return POINTS_UTILITIES.extractArray(inputArray, propertyName).indexOf(givenValue);
  },

  columnsToPropertyValues: function(inputArray, columnNames) {
  //First, calls removeEmptyRows. Then turns the array of arrays into an array of objects,
  //so there's less hard-coding strewn about.
    inputArray = POINTS_UTILITIES.removeEmptyRows(inputArray);
    var outputArray = [];
    var numProperties = columnNames.length;
    for (var i = 0; i < inputArray.length; i++) {
      if (inputArray[i].length != numProperties) {
        throw new Error('Row '+String(i)+' has the wrong number of entries.');
      } else {
        outputArray[i] = {};
        for (var j = 0; j < numProperties; j++) {
          outputArray[i][columnNames[j]] = inputArray[i][j];
        }
      }
    }
    return outputArray;
  },

  headerToPropertyNames: function(inputArray) {
    return this.columnsToPropertyValues(inputArray.slice(1), inputArray[0]);
  },

  truncatedISO: function(date) {
  //Adapted from the Mozilla Developer Network webpage on the Date object (long URL).
    var pad = function(n) {return n<10 ? '0'+n : n;};
    return date.getUTCFullYear()+'-'+pad(date.getUTCMonth()+1)+'-'+pad(date.getUTCDate());
  },

  retrieveRange: function(rangeOption, cycleNum) {
  //It'd be great to replace this with a function that defines all these ranges 'up the
  //scope chain', in the scope of whatever function is calling retrieveRange, but the
  //methods I've seen are non-standard (Function.caller) or deprecated (arguments.caller).
  //So we'll just have to call this once for each variable we want to define.
    switch (rangeOption) {
    case 'Sign Up Totals':
      var sheetNamePrefix = 'Loads';
      var rangeInfo = PSEUDO_GLOBALS.SIGN_UP_RANGE_INFO;
      //I'm changing variable names from apply_remove_empty_rows to applyRemoveEmptyRows. Since the function
      //is named removeEmptyRows I might be making things less clear, but I'm going to think of it as a
      //capitalization rule and then it's fine.
      var applyRemoveEmptyRows = true;
      break;
    case 'Sign Off Totals':
      var sheetNamePrefix = 'Loads';
      var rangeInfo = PSEUDO_GLOBALS.SIGN_OFF_RANGE_INFO;
      var applyRemoveEmptyRows = true;
      break;
    case 'Sign Up Data':
      var sheetNamePrefix = 'Chart';
      var rangeInfo = PSEUDO_GLOBALS.CHART_RANGE_INFO;
      var applyRemoveEmptyRows = false;
      break;
    case 'Point Values':
      var sheetNamePrefix = 'Point Values';
      var rangeInfo = PSEUDO_GLOBALS.CHART_RANGE_INFO;
      var applyRemoveEmptyRows = false;
      break;
    default:
      throw new Error('Invalid range option '+String(rangeOption)+'.');
      return;
      break;
    }
    //This function should only be called inside another that has fetched PSEUDO_GLOBALS.
    var fetchedValues = PSEUDO_GLOBALS.SPREADSHEET.getSheetByName(sheetNamePrefix+' '+String(cycleNum)).getRange(
      rangeInfo.topRow, rangeInfo.leftCol, rangeInfo.numRows, rangeInfo.numCols).getValues();
    if (applyRemoveEmptyRows) {
      fetchedValues = POINTS_UTILITIES.removeEmptyRows(fetchedValues);
    }
    return fetchedValues;
  },

  retrieveSheet: function(sheetOption, cycleNum) {
  //Analogous to retrieveRange. Again, this function must be called in an environment where PSEUDO_GLOBALS is defined.
    var valid_sheetOptions = ['Chart', 'Point Values', 'Loads'];
    if (valid_sheetOptions.indexOf(sheetOption) == -1) {
      throw new Error('Invalid sheet option '+String(sheetOption)+'.');
    }
    return PSEUDO_GLOBALS.SPREADSHEET.getSheetByName(sheetOption+' '+String(cycleNum));
  },

  //This section should be viewed as my (Ben Whitney's) best guess as to what is going on. The behavior you get
  //is very strange and the error messages are not very helpful. Be careful and back up your work if you mess
  //around with this stuff! Deep breaths. :)
  //Important: this function was written because of the following scenario.
  //  var x = 0;
  //  UI_BUILDERS.userSplitsName(); //A function that creates a UiInstance and has a SubmitButton.
  //  x += 1;
  //Before the user hits the SubmitButton, x will be incremented! This causes problems when we expect to be able
  //to access the responses (via ScriptDb) in the lines following the UI_BUILDERS.userSplitsName call.
  waitForUIResponse: function(identifierValue) {
    //As of right now, identifierValue isn't actually used. It's a holdover from when we used ScriptDb.
    //These values are in milliseconds.
    var maxWaitTime = 2*60*1000;
    var waitBetweenChecks = 500;
    var maxChecks = Math.floor(maxWaitTime/waitBetweenChecks);
    var userResponses;
    var numChecks = 0;
    var privateCache = CacheService.getPrivateCache();
    privateCache.remove('userResponses');
    while (!userResponses && numChecks < maxChecks) {
      //Slight pause so we're not hitting ScriptDb constantly.
      Utilities.sleep(waitBetweenChecks);
      userResponses = privateCache.get('userResponses');
      numChecks++;
    }
    if (numChecks == maxChecks) {
      throw new Error('User took too long to respond.');
    } else {
      return Utilities.jsonParse(userResponses);
    }
    return;
  },

  sendEmail: function(recipient, subject, body, options) {
    //Wrapper for MailApp.sendEmail. Catches errors in sending the message. If one occurs,
    //emails the Points Steward with an error message. This relies on the Points Steward
    //having a valid email address! TODO: which is not currently checked.
    var email_prefix = '[Points] '
    for (var i = 0, divider = ''; i < 25; i++) {
      divider += '-';
    }
    try {
      MailApp.sendEmail(recipient, email_prefix+subject, body, options);
    }
    catch (e) {
      MailApp.sendEmail(PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress,
        email_prefix+'Error in Sending Email',
        'ERROR: '+String(e)+'\n'+divider+'\n'+
        'RECIPIENT: '+recipient+'\n'+divider+'\n'+
        'SUBJECT: '+subject+'\n'+divider+'\n'+
        'BODY: '+body+'\n'+divider+'\n'+
        'OPTIONS: '+String(options),
        {});
    }
  }
};

//The onOpen function is executed automatically every time a Spreadsheet is loaded.
function onOpen() {
  //An entry of null makes a line separator.
  menuEntries = [
    {name:'Grant permissions'              , functionName:'grantPermissions'}     , null,
    {name:'Report an absence'              , functionName:'processAbsence'}       , null,
    {name:'Take or repay a loan'           , functionName:'processLoan'}          , null,
    {name:'Send a reminder'                , functionName:'userInitiatedReminder'}, null,
    {name:'Add chores to Google Calendar'  , functionName:'addToGoogleCalendar'}  , null,
    {name:'Add co-opers to Google Contacts', functionName:'addToGoogleContacts'}  , null,
    {name:'Force point counting'           , functionName:'findChoreSums'}
  ];
  //Not using PSEUDO_GLOBALS.SPREADSHEET potentially saves us some time here. It would be reasonable to call
  //pseudoGlobalsBuilder (including saving it to the cache) after adding the menu, though.
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Co-op', menuEntries);
}

function onEdit() {
  findChoreSums();
}

function userInitiatedReminder() {
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  var selectedCell = PSEUDO_GLOBALS.SPREADSHEET.getActiveCell();
  if (selectedCell.getSheet().getName() != 'Chart '+String(PSEUDO_GLOBALS.currentCycleNum)) {
    UI_BUILDERS.displayMessage('Please select a cell on the current chart.', 20);
    return;
  }
  var row = selectedCell.getRow();
  var col = selectedCell.getColumn();
  var signUpCell = PSEUDO_GLOBALS.SignUpCell(row, col,
    POINTS_UTILITIES.retrieveRange('Sign Up Data', PSEUDO_GLOBALS.currentCycleNum),
    POINTS_UTILITIES.retrieveRange('Point Values', PSEUDO_GLOBALS.currentCycleNum));
  if (!signUpCell.valid) {
    UI_BUILDERS.displayMessage('Please select a cell corresponding to the sign up entry for a chore.', 35);
    return;
  }
  if (!signUpCell.signUpCooper.valid) {
    UI_BUILDERS.displayMessage('That is not a valid sign up name.', 20);
    return;
  }
  UI_BUILDERS.getMessage('Please type a nice message for me to send:', 20);
  userResponses = POINTS_UTILITIES.waitForUIResponse('UI_BUILDERS.getMessage');
  POINTS_UTILITIES.sendEmail(signUpCell.signUpCooper.textAddress, signUpCell.choreName+' Reminder', userResponses.message,
    {cc: PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress});
  UI_BUILDERS.displayMessage('Just sent the message to '+signUpCell.signUpCooper.originalName+'. Thanks.', 20);
}

function addToGoogleContacts() {
  function formatPhoneNumber(phoneNumber) {
  //Sort of based on <http://en.wikipedia.org/wiki/North_American_Numbering_Plan>.
    return [phoneNumber.slice(0, 3), phoneNumber.slice(3, 6), phoneNumber.slice(6)].join('-');
  }

  function addCooper(cooper, contactGroup) {
    var namePieces = cooper.fullName.split(' ');
    var firstName, middleName, lastName;
    if (namePieces.length == 2) {
      //So that we can use userResponses.firstName and so on everywhere.
      var userResponses = {
        firstName: namePieces[0],
        lastName:  namePieces[1]
      }
    } else {
      UI_BUILDERS.userSplitsName(namePieces);
      var userResponses = POINTS_UTILITIES.waitForUIResponse('UI_BUILDERS.userSplitsName');
    }
    //ContactsApp complains if either userResponses.firstName or userResponses.lastName isn't given, so we can't
    //deal here with the case where the co-oper has only one name or something.
    var contact = ContactsApp.createContact(userResponses.firstName, userResponses.lastName, cooper.emailAddress);
    contact.addPhone(ContactsApp.Field.MOBILE_PHONE, formatPhoneNumber(cooper.phoneNumber));
    //If middleName is null or '', the conditional block won't be reached.
    if (userResponses.middleName) {
      contact.setMiddleName(userResponses.middleName);
    }
    //Date's getMonth method returns a value between 0 and 11, inclusive, so we don't have to make adjustments
    //when using it for array indexing.
    //TODO: scrappy data checking here. Figure out how to best do this. Right now cooper.birthday is `null`
    //if their entry on 'Basic Data' was blank.
    if (cooper.birthday) {
      contact.addDate(ContactsApp.Field.BIRTHDAY, monthValues[cooper.birthday.getMonth()],
        cooper.birthday.getDate(), cooper.birthday.getFullYear());
    }
    contactGroup.addContact(contact);
  }

  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  //There has to be a better way to do this, but the Contact's addDate method wants one of these, not an integer.
  var monthValues = [
    ContactsApp.Month.JANUARY, ContactsApp.Month.FEBRUARY, ContactsApp.Month.MARCH, ContactsApp.Month.APRIL,
    ContactsApp.Month.MAY, ContactsApp.Month.JUNE, ContactsApp.Month.JULY, ContactsApp.Month.AUGUST,
    ContactsApp.Month.SEPTEMBER, ContactsApp.Month.OCTOBER, ContactsApp.Month.NOVEMBER, ContactsApp.Month.DECEMBER
  ];
  if (PSEUDO_GLOBALS.todayIs.getMonth() > 5) {
    //Remember that getMonth returns 0 for January, 1 for February, etc.
    var fallYear = PSEUDO_GLOBALS.todayIs.getFullYear();
  } else {
    var fallYear = PSEUDO_GLOBALS.todayIs.getFullYear()-1;
  }
  var contactGroupName = 'Co-op '+String(fallYear)+'–'+String(fallYear+1);
  //ContactApp.getContactGroup returns null if the named group does not exist. In this case, because
  //See <https://developers.google.com/apps-script/reference/contacts/contact#addDate(Object,Month,Integer,Integer)> and
  //<https://developers.google.com/apps-script/reference/base/month>/
  //Boolean(null) == false, || will return the value of the second expression.
  var contactGroup = ContactsApp.getContactGroup(contactGroupName) || ContactsApp.createContactGroup(contactGroupName);

  //At this time there isn't a way to create all the contacts separately and then add them to a group all at once. That would
  //be faster than this approach.
  while (PSEUDO_GLOBALS.allCoopers.hasNext()) {
    addCooper(PSEUDO_GLOBALS.allCoopers.next(), contactGroup);
    //So we don't hit ContactsApp too frequently.
    Utilities.sleep(0.5*1000);
  }
  PSEUDO_GLOBALS.allCoopers.rewind();
  UI_BUILDERS.displayMessageAndLink('All done. You might want to click on the link below to merge '+
    'duplicates or make any changes.', 'Go to Contacts', 'https://www.google.com/contacts', 100);
}

function processLoan() {
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  var nextFunction = UI_BUILDERS.getNameAndCycle;
  var userResponses = {};
  var changesMade = false;
  while (!changesMade) {
    //This switch is kind of unfortunate (I've just watched <http://www.youtube.com/watch?v=4F72VULWFvc>) but
    //I really don't want to futz with the individual functions or doPost. Things buzzing around my head:
    //  doPost must return the current UiAppInstance.
    //  doPost is executed in a brand-new instance (word?), and you have to use ScriptDb/CacheServer/Spreadsheet
    //    (or some other trick) to pass data from the original instance to this one.
    //  Currently I have the UI-building functions (e.g. UI_BUILDERS.getNameAndCycle) returning nothing, and I'm not sure
    //    when they would be returned (execution seems to blow by the UI-building functions if you don't stop it,
    //    as I've been doing with waitForUIResponse) or if the value would be returned at all. I'm not in the mood
    //    to fiddle with that right now.
    //Apologies for the lack of sources. The reason I'm wary of touching this is that these are issues where the
    //documentation has not been very clear to me. The situation I am trying to avoid is that you, with your great
    //coding skills, see this ugly code and tear it down (fine so far) without understanding that some you might have
    //to deal with some weird behavior afterwards. Feel free to modify!
    switch (nextFunction) {
    case UI_BUILDERS.getNameAndCycle:
      //At first userResponses doesn't have these attributes. UI_BUILDERS.getNameAndCycle can deal with this.
      UI_BUILDERS.getNameAndCycle(userResponses.originalName, userResponses.cycleNum);
      userResponses = POINTS_UTILITIES.waitForUIResponse('UI_BUILDERS.getNameAndCycle');
      nextFunction = UI_BUILDERS.getLoanInfo;
      break;
    case UI_BUILDERS.getLoanInfo:
      //We need to save this value in case the user presses the 'Previous' button.
      var savedCycleNum = userResponses.cycleNum;
      var cooper = PSEUDO_GLOBALS.Cooper({'Sign Up Name': userResponses.originalName});
      //Pull the Loads sheet, look at the first row (the header), and find the 'Loan' column.
      var loadsSheet = POINTS_UTILITIES.retrieveSheet('Loads', userResponses.cycleNum);
      var loanCol = 1+loadsSheet.getDataRange().getValues()[0].indexOf('Loan');
      var loanRange = loadsSheet.getRange(2+cooper.index, loanCol);
      var oldLoan = Number(loanRange.getValue());
      UI_BUILDERS.getLoanInfo(cooper, userResponses.cycleNum, oldLoan);
      userResponses = POINTS_UTILITIES.waitForUIResponse('UI_BUILDERS.getLoanInfo');
      switch (userResponses.buttonPressed) {
      case 'previous':
        nextFunction = UI_BUILDERS.getNameAndCycle;
          userResponses.originalName = cooper.originalName;
          userResponses.cycleNum = savedCycleNum;
        break;
      case 'next':
        nextFunction = UI_BUILDERS.displayMessage;
        break;
      }
      break;
    case UI_BUILDERS.displayMessage:
      var newLoan = oldLoan+userResponses.loanAmount;
      cooper.loanBalance += userResponses.loanAmount;
      loanRange.setValue(newLoan);
      //TODO: point/points plural thing here as well.
      POINTS_UTILITIES.sendEmail(PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress, 'Loan Change for '+cooper.originalName,
      cooper.originalName+' is '+POINTS_UTILITIES.getLoanVerb(newLoan)+' '+String(Math.abs(userResponses.loanAmount))+
        ' points for a total loan of '+String(newLoan)+' points in Cycle '+String(savedCycleNum)+
        '. Total balance is now '+String(cooper.loanBalance)+' points.',
        {name:'Points', replyTo:PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress, cc:cooper.emailAddress});
      UI_BUILDERS.displayMessage('OK. Now you are '+POINTS_UTILITIES.getLoanVerb(newLoan)+' '+String(Math.abs(newLoan))+
        ' points and your total balance is '+String(cooper.loanBalance)+
        " points. I've emailed you and "+PSEUDO_GLOBALS.POINTS_STEWARD.originalName+'.', 50);
      changesMade = true;
      break;
    }
  }
}

function grantPermissions() {
  return;
}

function processAbsence() {
//Analogous to processLoan. Notably, though, the user input replaces, is not added to, the old absence amount.
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  UI_BUILDERS.getAbsenceInfo();
  var userResponses = POINTS_UTILITIES.waitForUIResponse('UI_BUILDERS.getAbsenceInfo');
  var cooper = PSEUDO_GLOBALS.Cooper({'Sign Up Name': userResponses.originalName});
  //Pull the Loads sheet, look at the first row (the header), and find the 'Loan' column.
  var loadsSheet = POINTS_UTILITIES.retrieveSheet('Loads', userResponses.cycleNum);
  //Add 1 to the index because .getValues returns an array indexed at 0, while .getRange below takes indexing starting at 1.
  var absenceCol = 1+loadsSheet.getDataRange().getValues()[0].indexOf('Days out');
  //Add 2 to the nameIndex: 1 for the header, and 1 because PSEUDO_GLOBALS.SIGN_UP_NAMES_ORIGINAL is indexed starting at 0.
  loadsSheet.getRange(2+cooper.index, absenceCol).setValue(Number(userResponses.absenceAmount));

  //TODO: deal with day/days plural thing.
  POINTS_UTILITIES.sendEmail(PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress, 'Absence Change for '+cooper.originalName,
    cooper.originalName+' is going to be out for '+String(userResponses.absenceAmount)+' days in Cycle '+
    String(userResponses.cycleNum)+'.',
      {name:'Points', replyTo:cooper.emailAddress, cc:cooper.emailAddress});
  UI_BUILDERS.displayMessage('OK. Now you are marked down as out for '+String(userResponses.absenceAmount)+' days for Cycle '+
    String(userResponses.cycleNum)+". I've emailed you and "+PSEUDO_GLOBALS.POINTS_STEWARD.originalName+'.', 50);
}

function happyBirthday() {
  var sameMonthAndDay = function(dateOne, dateTwo) {
    if (!dateOne) {
      //TODO: this is a scrappy way of checking whether the co-oper's birthday is defined or not.
      //Relies on us always putting the co-oper's birthday first.
      return false;
    }
    return dateOne.getMonth() == dateTwo.getMonth() && dateOne.getDate() == dateTwo.getDate();
  };
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  //This would probably be faster if we first iterated just over the birthdays and then constructed co-opers only for
  //those whose birthday it was. Not necessary here.
  while (PSEUDO_GLOBALS.allCoopers.hasNext()) {
    var cooper = PSEUDO_GLOBALS.allCoopers.next();
    if (sameMonthAndDay(cooper.birthday, PSEUDO_GLOBALS.todayIs)) {
      MailApp(cooper.emailAddress, 'Happy Birthday '+cooper.originalName+'!', ':)\n\nThe Points Chart',
        {name:'Points', replyTo:PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress});
    }
  }
  PSEUDO_GLOBALS.allCoopers.rewind();
}

function findChoreSums(cycleNum) {
//This function cycles over all the chores, counting up how many points
//each co-oper has signed up and off for. It loads everything into arrays,
//does the calculations there (beware of indices!), and then writes it at the end.
//It also gets a total number of chore points for the cycle (VOIDed chores are excluded)
//and does some formatting.
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  if (typeof cycleNum === 'undefined') {
    //TODO: fix this hardcoded thing.
    //This is so it will do the next cycle when people are signing up for chores.
    var lookForward = PSEUDO_GLOBALS.currentCycleNum < 8 ? 1 : 0;
    for (var cycleNum = PSEUDO_GLOBALS.currentCycleNum+lookForward; cycleNum >= PSEUDO_GLOBALS.firstUnlockedCycle; cycleNum--) {
      findChoreSums(cycleNum);
    }
    return;
  }

  function conditionalColoring(signUpCell) {
  //Do some conditional formatting.
    function setNameColor(cssColor) {
      signUpSheet.getRange(signUpCell.row, signUpCell.col).setFontColor(cssColor);
    }
    function setInitialsColor(cssColor) {
      signUpSheet.getRange(signUpCell.row+1, signUpCell.col).setFontColor(cssColor);
    }
    if (signUpCell.diagnosis.allValid) {
        setNameColor(UI_BUILDERS.DEFAULTS.hexColors.black);
        setInitialsColor(UI_BUILDERS.DEFAULTS.hexColors.black);
        return;
    }
    if (signUpCell.diagnosis.voided) {
        setNameColor(UI_BUILDERS.DEFAULTS.hexColors.purple);
        setInitialsColor(UI_BUILDERS.DEFAULTS.hexColors.purple);
        return;
    }
    if (signUpCell.diagnosis.sameCooper) {
        //Implies this.signUpCooper.valid && this.signOffCooper.valid.
        setNameColor(UI_BUILDERS.DEFAULTS.hexColors.black);
        setInitialsColor(UI_BUILDERS.DEFAULTS.hexColors.red);
        return;
    }
    if (signUpCell.signUpCooper.valid || signUpCell.diagnosis.noName) {
        setNameColor(UI_BUILDERS.DEFAULTS.hexColors.black);
    } else {
        setNameColor(UI_BUILDERS.DEFAULTS.hexColors.red);
    }
    if (signUpCell.signOffCooper.valid || signUpCell.diagnosis.noInitials) {
        setInitialsColor(UI_BUILDERS.DEFAULTS.hexColors.black);
    } else {
        setInitialsColor(UI_BUILDERS.DEFAULTS.hexColors.red);
    }
  return;
  }

  var signUpSheet      = POINTS_UTILITIES.retrieveSheet('Chart', cycleNum);
  var pointValuesSheet = POINTS_UTILITIES.retrieveSheet('Point Values', cycleNum);
  var chart       = POINTS_UTILITIES.retrieveRange('Sign Up Data', cycleNum);
  var pointValues = POINTS_UTILITIES.retrieveRange('Point Values', cycleNum);

  var signUpCounts  = [];
  var signOffCounts = [];
  for (var i = 0; i < PSEUDO_GLOBALS.COOPER_INFO.length; i++) {
    signUpCounts[i]  = 0;
    signOffCounts[i] = 0;
  }
  var pointsTotal = 0;

  //This loop goes over all the days for a daily chore (e.g. MCU). For weekly chores it has the effect of
  //going over each of the chores (e.g. the different bathrooms). Sometimes it will be looking at blank
  //cells. A lot easier than somehow avoiding these blank cells is ensuring that adding their contents to
  //points_value does what we expect: no change. All we have to do is convert the cell contents to a Number.
  //Here is the situation we are trying to avoid: to start, points_value is 123, cell_contents_1 is '', and
  //cell_contents_2 is '1'. Then,
  //  points_value += cell_contents_1; points_value += cell_contents_2;
  //will result in points_value being '1231' instead of 124. Simple enough, but it took me a bit to figure out
  //what the problem was when I was getting a final points_value of like '87623471897623348763'.
  var cycleSignUpCells = PSEUDO_GLOBALS.cycleSignUpCells(cycleNum);
  while (cycleSignUpCells.hasNext()) {
    var signUpCell = cycleSignUpCells.next();
    if (!signUpCell.valid) {
      continue;
    }
    if (signUpCell.signUpCooper.valid && !signUpCell.diagnosis.voided) {
      signUpCounts[signUpCell.signUpCooper.index] += signUpCell.choreValue;
      if (signUpCell.signOffCooper.valid && !signUpCell.diagnosis.sameCooper) {
          signOffCounts[signUpCell.signUpCooper.index] += signUpCell.choreValue;
      }
    }
    if (!signUpCell.diagnosis.voided) {
      pointsTotal += signUpCell.choreValue;
    }
    conditionalColoring(signUpCell);
  }
  cycleSignUpCells.rewind();

  pointValuesSheet.getRange(PSEUDO_GLOBALS.POINTS_TOTAL_CELL).setValue(pointsTotal);
  //Convert the arrays of integers back to arrays of one-dimensional arrays.
  for (var i = 0; i < PSEUDO_GLOBALS.COOPER_INFO.length; i++) {
    signUpCounts[i]  = [signUpCounts[i]];
    signOffCounts[i] = [signOffCounts[i]];
  }
  //Write sign up and sign off totals to the corresponding columns on the spreadsheet. We're
  //only writing to rows that correspond to co-opers (probably a smaller range than the one
  //we initially pulled, because usually PSEUDO_GLOBALS.COOPER_INFO.length < PSEUDO_GLOBALS.MAX_COOPERS).
  var loadsSheet = POINTS_UTILITIES.retrieveSheet('Loads', cycleNum);
  //Kind of gross. Sorry. Keeping it because it gets wordy.
  var toWrites   = [signUpCounts, signOffCounts];
  var rangeInfos = [PSEUDO_GLOBALS.SIGN_UP_RANGE_INFO, PSEUDO_GLOBALS.SIGN_OFF_RANGE_INFO];
  for (var i = 0; i < toWrites.length; i++) {
    loadsSheet.getRange(rangeInfos[i].topRow, rangeInfos[i].leftCol,
      PSEUDO_GLOBALS.COOPER_INFO.length, rangeInfos[i].numCols).setValues(toWrites[i]);
  }
}

function addToGoogleCalendar() {
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  var chart       = POINTS_UTILITIES.retrieveRange('Sign Up Data', PSEUDO_GLOBALS.currentCycleNum);
  var pointValues = POINTS_UTILITIES.retrieveRange('Point Values', PSEUDO_GLOBALS.currentCycleNum);
  var choresCalName = 'Co-op Chores';
  UI_BUILDERS.getCalendarPreferences();
  var userResponses = POINTS_UTILITIES.waitForUIResponse('UI_BUILDERS.getCalendarPreferences');
  userResponses.wantsPopupReminders = 'yes' == userResponses.wantsPopupReminders; 
  userResponses.wantsTextReminders  = 'yes' == userResponses.wantsTextReminders; 
  userResponses.wantsEmailReminders = 'yes' == userResponses.wantsEmailReminders;
  CalendarApp.getAllCalendars().forEach(function(calendar) {
    if (calendar.getName() == choresCalName) {
      calendar.deleteCalendar();
    }
  });
  var choresCalendar = CalendarApp.createCalendar(choresCalName, {color: CalendarApp.Color.PLUM});

  var choresThisCycle = [];
  var cycleSignUpCells = PSEUDO_GLOBALS.cycleSignUpCells(PSEUDO_GLOBALS.currentCycleNum);
  while (cycleSignUpCells.hasNext()) {
    var signUpCell = cycleSignUpCells.next();
    //Safe to do this because signUpCells.signUpCooper is only checked in the event that signUpCell.valid is true.
    if (!signUpCell.valid || !signUpCell.signUpCooper.valid) {
      continue;
    }
    if (signUpCell.signUpCooper.simplifiedName == userResponses.simplifiedGivenName) {
      choresThisCycle.push(signUpCell.chore);
    }
  }
  cycleSignUpCells.rewind();

  choresThisCycle.forEach(function(chore) {
    var event = choresCalendar.createEvent(chore.name, chore.startDate, chore.stopDate);
    userResponses.wantsPopupReminders && event.addPopupReminder(userResponses.popupReminderTime);
    userResponses.wantsTextReminders  && event.addSmsReminder(userResponses.textReminderTime);
    userResponses.wantsEmailReminders && event.addEmailReminder(userResponses.emailReminderTime);
  });
}

function sendReminders() {
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  var chart = POINTS_UTILITIES.retrieveRange('Sign Up Data', PSEUDO_GLOBALS.currentCycleNum);
  var pointValues = POINTS_UTILITIES.retrieveRange('Point Values', PSEUDO_GLOBALS.currentCycleNum);
  var todayCol = -1;
  for (var i = 0; i < 14; i++) {
    if (PSEUDO_GLOBALS.todayIs.getTime() == chart[1][i+1].getTime()) {
      todayCol = i+2;
      break;
    }
  }
  //Shoule we check for it being undefined/null instead?
  if (todayCol == -1) {
    throw new Error('Unable to find column corresponding to today in Cycle '+
                    String(PSEUDO_GLOBALS.currentCycleNum)+'.');
  }

  function sortChores(col, lookingAt, chart, pointValues) {
    var daySignUpCells = PSEUDO_GLOBALS.daySignUpCells(col, chart, pointValues);
    var choresPerCooper = [];
    while (daySignUpCells.hasNext()) {
      var signUpCell = daySignUpCells.next();
      if (!signUpCell.signUpCooper.valid) {
        continue;
      }
      if (!choresPerCooper[signUpCell.signUpCooper.index]) {
        choresPerCooper[signUpCell.signUpCooper.index] = {'cooper': signUpCell.signUpCooper, 'chores': []}; 
      }
      switch (lookingAt) {
      case 'Today':
        choresPerCooper[signUpCell.signUpCooper.index]['chores'].push(signUpCell.choreName);
        break;
      case 'Yesterday':
        if (!signUpCell.diagnosis.voided &&
          (!signUpCell.signOffCooper.valid || signUpCell.diagnosis.sameCooper)) {
          choresPerCooper[signUpCell.signUpCooper.index]['chores'].push(signUpCell.choreName);
        }
        break;
      }
    }
    daySignUpCells.rewind();
    return choresPerCooper;
  }

  function makeCommaList(chores) {
    var numChores = chores.length;
    switch (numChores) {
    case 1:
      var outputList = chores[0];
      break;
    case 2:
      var outputList = chores[0]+' and '+chores[1];
      break;
    default:
      var outputList = '';
      for (var i = 0; i < numChores-1; i++) {
        outputList += chores[i]+', ';
      }
      outputList += 'and '+chores[numChores-1];
    }
    return outputList;
  }

  function sendReminder(choresObject, messageSubject, emailLeadIn, remType) {
    switch (remType) {
    case 'Text':
      var messageBody = makeCommaList(choresObject.chores)+'.';
      break;
    case 'Email':
      var messageBody = 'Hello '+choresObject.cooper.originalName+',\n\n'+emailLeadIn+makeCommaList(choresObject.chores)+'.'+
        '\n\nHere is the link to the chart:\n'+PSEUDO_GLOBALS.LINK_TO_CHART+'\n\nThis email was automatically generated. '+
        'Please tell '+PSEUDO_GLOBALS.POINTS_STEWARD.originalName+' if there has been a mistake.';
      break;
    }
    POINTS_UTILITIES.sendEmail(choresObject.cooper[remType.toLowerCase()+'Address'], messageSubject, messageBody,
      {name:'Points', replyTo:PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress});
  }

  function lookAtToday(todayCol, chart, pointValues) {
    var messageSubject = 'Chores '+POINTS_UTILITIES.truncatedISO(PSEUDO_GLOBALS.todayIs);
    var emailLeadIn = 'Today you are signed up for ';
    var choresPerCooper = sortChores(todayCol, 'Today', chart, pointValues);
    var hasChoresToday = [];
    choresPerCooper.forEach(function(choresObject) {
      if (!choresObject) {
        return;
      }
      hasChoresToday.push(choresObject.cooper.originalName);
      if (choresObject.cooper.wantsEmailReminders) {
        sendReminder(choresObject, messageSubject, emailLeadIn, 'Email');
      }
      if (choresObject.cooper.wantsTextReminders) {
        sendReminder(choresObject, messageSubject, emailLeadIn, 'Text')
      }
    });
    return hasChoresToday;
  }

  //TODO: this should look back at all previous days.
  function lookAtYesterday(todayCol, chart, pointValues) {
    //Column 2 corresponds to the first Sunday of a cycle.
    if (todayCol == 2) {
      if (PSEUDO_GLOBALS.currentCycleNum == 1) {
        return;
      }
      //Column 15 corresponds to the last Sunday of a cycle.
      todayIndex = 15;
      var chart       = POINTS_UTILITIES.retrieveRange('Sign Up Data', PSEUDO_GLOBALS.currentCycleNum-1);
      var pointValues = POINTS_UTILITIES.retrieveRange('Point Values', PSEUDO_GLOBALS.currentCycleNum-1);
    } else {
      todayCol -= 1;
    }
    var messageSubject = 'Reminder to Sign Off';
    var emailLeadIn = 'Please remember to get someone to sign off on ';
    var choresPerCooper = sortChores(todayCol, 'Yesterday', chart, pointValues);
    var needSignOff = [];
    choresPerCooper.forEach(function(choresObject) {
      //As of now choresObject will still exist if the person has no chores. choresObject.chores will be [],
      //because sortChores will recognize that they've signed off on everything.
      if (!choresObject.chores.length) {
        return;
      }
      needSignOff.push(choresObject.cooper.originalName);
      if (choresObject.cooper.wantsEmailReminders) {
        sendReminder(choresObject, messageSubject, emailLeadIn, 'Email');
      }
    });
    return needSignOff;
  }

  var hasChoresToday = lookAtToday(todayCol, chart, pointValues);
  var needSignOff = lookAtYesterday(todayCol, chart, pointValues);
  var messageSubject = 'Chores Report '+POINTS_UTILITIES.truncatedISO(PSEUDO_GLOBALS.todayIs);
  //TODO: deal with case that no one needs sign offs.
  POINTS_UTILITIES.sendEmail(PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress, messageSubject, makeCommaList(hasChoresToday)+
    ' have chores today.\n\n'+makeCommaList(needSignOff)+' need sign offs on chores from yesterday.',
    {name:'Points', replyTo:PSEUDO_GLOBALS.POINTS_STEWARD.emailAddress});
}

function propagateSheetChanges() {
  PSEUDO_GLOBALS = pseudoGlobalsFetcher();
  var modelCycleNum = 2;
  var maxCycleNum = 8;
  var sheetNamePrefixes = ['Loads', 'Point Values', 'Chart'];
  for (var i = modelCycleNum+1; i <= maxCycleNum; i++) {
    sheetNamePrefixes.forEach(function(sheetNamePrefix) {
      var sheetName = sheetNamePrefix+' '+String(i);
      var modelName = sheetNamePrefix+' '+String(modelCycleNum);
      var oldSheet = PSEUDO_GLOBALS.SPREADSHEET.getSheetByName(sheetName);
      if (oldSheet) {
        PSEUDO_GLOBALS.SPREADSHEET.deleteSheet(oldSheet); 
      }
      var copiedSheet = PSEUDO_GLOBALS.SPREADSHEET.getSheetByName(modelName).copyTo(PSEUDO_GLOBALS.SPREADSHEET).setName(sheetName);
      if (sheetNamePrefix != 'Loads') {
        copiedSheet.hideSheet();
      }
    });
  }
}

function doPost(e) {
  //Important (from <https://developers.google.com/apps-script/reference/ui/server-handler>):
  //  When a ServerHandler is invoked, the function it refers to is called on the Apps Script server in a "fresh" script.
  //  This means that no variable values will have survived from previous handlers or from the initial script that loaded the app.
  //  Global variables in the script will be re-evaluated, which means that it's a bad idea to do anything slow
  //  (like opening a Spreadsheet or fetching a Calendar) in a global variable.
  //This is why you need to use CacheService or ScriptDb or something similar to save the values in here. Writing to the spreadsheet
  //will also work, but changing the value of a 'global' variable (we're in a new 'globe' here) won't!

  //var scriptDatabase = ScriptDb.getMyDb();
  var toSave = {identifier: e.parameter.identifier};
  switch (e.parameter.identifier) {
  case 'UI_BUILDERS.userSplitsName':
    toSave.firstName  = e.parameter.firstName;
    toSave.middleName = e.parameter.middleName;
    toSave.lastName   = e.parameter.lastName;
    break;
  case 'UI_BUILDERS.getCalendarPreferences':
    //TODO: include note either here or (probably better) in documentation so that people don't mess up with JSON.
    //Then again, not sure that it's JSON that was causing problems here.
    toSave.wantsPopupReminders = e.parameter.popupRadioRecorder;
    toSave.wantsTextReminders  = e.parameter.textRadioRecorder;
    toSave.wantsEmailReminders = e.parameter.emailRadioRecorder;
    toSave.popupReminderTime = POINTS_UTILITIES.jsonNumber(e.parameter.popupTextBox);
    toSave.textReminderTime  = POINTS_UTILITIES.jsonNumber(e.parameter.textTextBox);
    toSave.emailReminderTime = POINTS_UTILITIES.jsonNumber(e.parameter.emailTextBox);
    toSave.simplifiedGivenName = POINTS_UTILITIES.simplify(e.parameter.userName);
    break;
  case 'UI_BUILDERS.getNameAndCycle':
    toSave.originalName = e.parameter.userName;
    toSave.cycleNum = POINTS_UTILITIES.jsonNumber(e.parameter.cycleNum);
    break;
  case 'UI_BUILDERS.getLoanInfo':
    toSave.buttonPressed = e.parameter.buttonPressed;
    //See <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Conditional_Operator>.
    toSave.loanAmount = POINTS_UTILITIES.jsonNumber(e.parameter.loanAmount)*(e.parameter.loanAction == 'Take a loan' ? 1 : -1);
    break;
  case 'UI_BUILDERS.getAbsenceInfo':
    toSave.originalName = e.parameter.userName;
    toSave.cycleNum = POINTS_UTILITIES.jsonNumber(e.parameter.cycleNum);
    toSave.absenceAmount = POINTS_UTILITIES.jsonNumber(e.parameter.absenceAmount);
    break;
  case 'UI_BUILDERS.getMessage':
    toSave.message = e.parameter.textArea;
    break;
  default:
    //TODO: this is going to cause an "unexpected error" so you should write to some prominent cell on the spreadsheet. 
    throw new Error('Unrecognized UI identifier.');
    break;
  }
  CacheService.getPrivateCache().put('userResponses', Utilities.jsonStringify(toSave), 60);
  var currentApp = UiApp.getActiveApplication();
  currentApp.close();
  return currentApp;
}

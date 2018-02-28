from pyexcel_ods import save_data, get_data
import re
import json


#-------------#
#  Utilities  #
#-------------#
def check_row (row):
  if   len(row) == 0: return "empty"
  elif len(row) <= 2: return "section"
  elif len(row) <  5: return "ignored"
  elif sum([is_number(c) for c in row]) > 5: return "series"
  elif sum([is_check(c) for c in row])  > 5: return "check"
  elif sum([is_date(c) for c in row])   > 5: return "date"
  return "unknown"

def is_date (val):
  if type(val) not in (str, unicode): return False # Dates must be string or unicode
  match = re.match('\d{2,4}/\d{2,4}', val)         # Fiscal dates (e.g. 07/08, 2007/2008, or 2007/08)
  return match is not None

def is_number (val):
  return type(val) in (int, float)

def is_check (val):
  return val == "OK"

def is_text (val):
  return type(val) in (unicode, str)



class Series (object):
  def __init__ (self, row, row_num, source, section, dates):
    self.row      = row
    self.row_num  = row_num
    self.warnings = []
    self.source   = source
    self.section  = section
    self.name     = self.get_name(row)
    self.data     = self.get_data(row, dates)

  def get_name (self, row):
    name = None
    cells = [c for c in row if c and is_text(c)]
    for c in cells:
      if name: self.warn("Additional text found: " + c)
      else   : name = c
    if not name: self.warn("Invalid series (no label)")
    return name

  # Data should be a dictionary of dates:values
  # e.g. { "06/07" : 100, "07/08" : 102, ... }
  def get_data (self, row, dates):
    data = {}
    cells = [(k, c) for k, c in enumerate(row) if is_number(c)]
    for k, c in cells:
      try:
        date = dates[k]
        data[date] = float(c)
      except KeyError:
        self.warn("Date not found on col " + str(k))
    return data

  def data_match (self, data, threshold):
    matches = 0
    for date in sorted(data):
      if date not in self.data: continue
      a = self.data[date]
      b = data[date]
      if a and b:
        diff = abs(a / b - 1)
        if diff <= threshold: matches += 1
    return matches

  def data_diff (self, data):
    diff = 0
    for date in sorted(data):
      if date not in self.data: continue
      a = self.data[date]
      b = data[date]
      if a and b:
        diff += abs(a / b - 1)
    return diff

  def show (self):
    return "(" + self.source + " row " + str(self.row_num) + ") " + self.name

  def warn (self, msg):
    self.warnings.append((msg, self.row_num, self.row))



class SuperSeries (object):
  def __init__ (self, s):
    self.names     = {} # Frequency dictionary of names
    self.series    = {} # Dictionary of sources -> series from that source
    self.values    = {} # Dictionary of dates -> frequency dictionary of tables
    self.consensus = {} # Consensus values
    self.add_series(s)

  # Add a series to this SuperSeries
  def add_series (self, s):
    if s.source in self.series:
      print "--- ERROR ---!"
      self.explain_match(s)
      raise Exception("Trying to add another series from the same source!")
    if s.name in self.names: self.names[s.name] += 1 # Increase frequency of name
    else                   : self.names[s.name]  = 1 # Register new name
    self.series[s.source] = s
    self.add_data(s.data)

  def add_data (self, data):
    values = self.values
    for k in data:
      v = data[k]
      if not k in values   : values[k] = {}
      if not v in values[k]: values[k][v] = 0
      values[k][v] += 1
      self.consensus[k] = max(values[k])

  def name_match (self, s):
    return s.name in self.names

  def data_match (self, s, threshold):
    return s.data_match(self.consensus, threshold)

  def data_diff (self, s):
    return s.data_diff(self.consensus)

  def explain_match (self, s):
    for k in self.series: print "Super  :", self.series[k].show()
    print "Series :", s.show()
    print "Match  :", self.data_match(s, 0.0005), "of", len(s.data)
    for date in sorted(s.data):
      if date not in self.consensus: continue
      cv = self.consensus[date]            # Consensus value
      nv = s.data[date]                    # New value
      if cv and nv:
        diff = abs(cv / nv - 1)            # Calculate difference with consensus value
        print str(date).ljust(16), str(cv).ljust(16), str(nv).ljust(16), str(diff).ljust(8)



class Mutator (object):
  def __init__ (self, data, inputs):
    print "Starting Mutator..."
    self.warnings = []
    self.ss = []
    for opt in inputs:
      print "------------------------"
      print "Scenario:", opt["name"]
      sheet  = self.get_sheet(data, opt)
      rows   = self.get_rows(sheet, opt)
      dates  = self.get_dates(rows["dates"])
      series = self.get_series(rows["series"], dates, opt)
      self.parse_series(series)
      self.report(opt)
    print "Mutator finished."


  #---------------#
  #  Data reader  #
  #---------------#
  def get_sheet (self, data, opt):
    try:
      sheet = data[opt["sheet"]]
      print "Sheet has", len(sheet), "rows"
      return sheet
    except KeyError:
      raise Exception("Sheet not found!")

  # Group rows by type
  def get_rows (self, sheet, opt):
    section  = ""
    ignored  = []
    dates    = []
    series   = []
    unknown  = []
    for k, row in enumerate(sheet):
      row_type = check_row(row)
      if   row_type is "section"  : section = row[0]
      if   k in opt["ignore_rows"]: ignored.append((k, row))
      elif row_type is "empty"    : ignored.append((k, row))
      elif row_type is "check"    : ignored.append((k, row))
      elif row_type is "section"  : ignored.append((k, row))
      elif row_type is "date"     : dates.append((k, row))
      elif row_type is "series"   : series.append((k, row, section))
      else:
        self.warn("Unknown row type", k, row)
        unknown.append((k, row))
    print "Rows parsed:", len(series), "series,", len(ignored), "ignored,", len(dates), "dates,", len(unknown), "unknown"
    return { "ignored" : ignored, "dates" : dates, "series" : series, "unknown" : unknown }

  # Extract date row into a column -> date lookup dictionary
  # e.g. { 23 : "06/07", 24 : "07/08", ... } (where the key is the column number)
  def get_dates (self, date_rows):
    if not date_rows:
      raise Exception("Date row not found!")
    elif len(date_rows) > 1:
      for row in date_rows: print row
      raise Exception("Multiple date rows found!")
    dates = {}
    date_row = date_rows[0][1]
    cells = [(k, c) for k, c in enumerate(date_row) if is_date(c)]
    for k, c in cells: dates[k] = c
    print "Date range", cells[0][1], "to", cells[-1][1]
    return dates

  def get_series (self, series_rows, dates, opt):
    series  = []
    ignored = []
    for r in series_rows:
      s = Series(r[1], r[0], opt["name"], r[2], dates)
      if s.name in opt["ignore_series"]:
        ignored.append(s.row)
      else:
        series.append(s)
        self.warnings += s.warnings
    print "Series parsed:", len(ignored), "ignored,", len(series), "remaining"
    return series


  #------------------#
  #  Series joiners  #
  #------------------#
  def find_ss (self, s, ss_list, threshold=0.0, match_count=6):
    matches = [ss for ss in ss_list if ss.data_match(s, threshold) > match_count]
    if len(matches) == 1: return matches[0]
    # Refine with name match if there are multiple results
    matches = [ss for ss in matches if ss.name_match(s)]
    if len(matches) == 1: return matches[0]
    if matches:
      self.warn("Multiple matches!", s.row_num, s.row)
      for ss in matches: print ss.names, ss.consensus
#       raise Except("oh")

  def parse_series (self, series):
    duplicate = []
    exact     = []
    fuzzy     = []
    new       = []
    s_list    = [s for s in series]
    ss_list   = [ss for ss in self.ss]

    # Deduplicate
    for s in s_list:
      for t in s_list:
        if s is t: continue
        if s.data_match(t.data, 0.0) is len(t.data):
          s_list.remove(t)
          duplicate.append(t)
          break

    # Look for exact matches
    for s in s_list:
      match = self.find_ss(s, ss_list, 0.0, 6)
      if match:
        match.add_series(s)
        ss_list.remove(match)
        s_list.remove(s)
        exact.append(s)

    # Look for fuzzy matches
    for s in s_list:
      match = self.find_ss(s, ss_list, 0.0005, 8)
      if match:
        match.add_series(s)
        ss_list.remove(match)
        s_list.remove(s)
        fuzzy.append(s)

    # Create new SuperSeries for the rest
    for s in s_list:
      self.ss.append(SuperSeries(s))
      new.append(s)

    for s in new + exact + fuzzy:
      self.warnings += s.warnings

    print "Series integrated:", len(duplicate), "duplicates,", len(exact), "exact matches,", len(fuzzy), "fuzzy matches,", len(new), "unmatched new series"
    # for t in joined: t[0].explain_match(t[1])


  #-------------------#
  #  Error reporting  #
  #-------------------#
  def warn (self, msg, row_num, row):
    self.warnings.append((msg, row_num, row))

  def report (self, opt):
    warnings = self.warnings
    self.warnings = []
    no_warn  = None
    bad_rows = set()
    if "ignore_warnings" in opt:
      no_warn = opt["ignore_warnings"]
    if no_warn:
      suppressed = [w for w in warnings if w[0] in no_warn]
      warnings   = [w for w in warnings if w[0] not in no_warn]
    if warnings:
      warnings = sorted(warnings, key=lambda w: w[1])
      for w in warnings:
        # (message, SuperSeries, Series)
        if type(w[1]) is SuperSeries:
          print "WARNING on row", w[2].row_num, "-", w[0]
          print w[2]
          bad_rows.add(w[2].row_num)
        # (message, row_num, row)
        else:
          print "WARNING on row", w[1], "-", w[0]
          print w[2]
          bad_rows.add(w[1])
      print "Warnings on rows", sorted(bad_rows)
    if no_warn and suppressed:
      print len(suppressed), "warnings suppressed"

  def jsonify (self):
    out = []
    # ss_list = [ss for ss in self.ss if len(ss.series) > 1]
    ss_list = sorted(self.ss, key=lambda ss: len(ss.series))
    for ss in ss_list:
      out.append({
          "name"   : max(ss.names),
          "names"  : ss.names,
          "series" : [{ "source" : source, "data" : ss.series[source].data } for source in ss.series]
      })
    return out




# Load settings
with open('settings.json', 'rw') as settings_file:
  settings = json.load(settings_file)

# Load spreadsheet
data = get_data(settings["path"])
print len(data), "sheets imported"

# Parse mutations
m = Mutator(data, settings["inputs"])
print "Mutator has", len(m.ss), "SuperSeries,", len([ss for ss in m.ss if len(ss.series) > 1]), "of these are linked"

def jsonify (self):
  out = []
  # ss_list = [ss for ss in self.ss if len(ss.series) > 1]
  ss_list = filter(lambda ss: len(ss.series) > 1, self.ss)
  ss_list = sorted(self.ss, key=lambda ss: len(ss.series))
  for ss in ss_list:
    out.append({
        "name"   : max(ss.names),
        "names"  : ss.names,
        "series" : [{ "source" : source, "data" : ss.series[source].data } for source in ss.series]
      })
  return out

with open('out.json', 'w') as outfile:
  json.dump(jsonify(m), outfile)

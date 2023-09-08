import openpyxl
import datetime
import json
from collections import OrderedDict


class Trap():
    """ Class for holding trap data"""
    def __init__(self, name, id, location, gps, model, admin, entries):
        self.name = name
        self.id = id
        self.location = location
        self.gps = gps
        self.model = model
        self.admin = admin
        self.entries = entries
        self.active_days = OrderedDict()


def get_entries(raw_entries, dates):
    """ Create dictionary from table rows and dates 

    Example:
    raw_entries = [Cell(value=0), Cell(value=1), Cell(value=2)]
    dates = [2023-01-01, 2023-01-02, 2023-01-03]
    result = {2023-01-01: 0,
              2023-01-02: 1,
              2023-01-03: 2}
    """
    return { dates[n]: raw_entries[n].value
             for n in range(len(raw_entries))
           }


def is_date_in_interval(check_date, start_date, end_date):
    """ Check if date is in interval defined by two other dates """
    # Copy year from checked date
    start_date = start_date.replace(year=check_date.year)
    end_date = end_date.replace(year=check_date.year)

    # Handle cases where end_date is earlier than start_date
    if end_date < start_date:
        if start_date <= check_date <= datetime.datetime(check_date.year, 12, 31) or \
           datetime.datetime(check_date.year, 1, 1) <= check_date <= end_date:
            return True
        else:
            return False

    if start_date <= check_date <= end_date:
        return True
    else:
        return False


class TrapActivity():
    """ Class for processing trap data """
    def __init__(self):
        self.traps = {}
        self.header_data = OrderedDict()
        self.dates = []

        # load seasons
        with open('seasons.json') as fo:
            seasons = json.load(fo)

        # Convert start/end to datetime
        self.seasons = {}
        for sname, season in seasons.items():
            self.seasons[sname] = {'start': datetime.datetime.strptime(season['start'], '%m-%d'),
                                   'end': datetime.datetime.strptime(season['end'], '%m-%d')}

    def load(self, file):
        """ Load trap data from xlsx file 
        
        This can be called multiple times with defferent files and the data will be merged.
        """
        wb = openpyxl.load_workbook(file)

        first = True
        dates = []

        for sheet in wb:
            for row in sheet:
                if first:
                    for r in row[7:]:
                        date = r.value.split('.')
                        pdate = datetime.datetime(int(date[2]), int(date[1]), int(date[0]))
                        season_year = self._get_season_year(pdate)
                        self.header_data.setdefault(season_year, OrderedDict())
                        self.header_data[season_year][self._get_season(pdate)] = None
                        self.dates.append(pdate)
                        dates.append(pdate)
                    first =  False
                else:
                    trap_id = row[1].value
                    if trap_id in self.traps.keys():
                        self.traps[trap_id].entries.update(get_entries(row[7:], dates))
                    else:
                        self.traps[trap_id] = \
                            Trap(row[0].value,
                                row[1].value,
                                row[2].value,
                                (row[3].value, row[4].value),
                                row[5].value,
                                row[6].value,
                                get_entries(row[7:], dates)
                            )

    def _get_season(self, date):
        """ Get season name for specified date """
        for season_name, season in self.seasons.items():
            if is_date_in_interval(date, season['start'], season['end']):
                return season_name

    def _get_season_year(self, date):
        """ Get season year for specified date """
        # pre-mating starts at the end of previous year
        start_date = self.seasons['pre-mating']['start'].replace(year=date.year)
        if start_date <= date <= datetime.datetime(date.year, 12, 31):
            return date.year + 1
        return date.year

    def _get_table_header(self):
        """ Returns header for new csv table"""
        header = []
        for year, seasons in self.header_data.items():
            for season in seasons:
                header.append(f'{season} {year}')

        return f'Lokalita;ID lokality;Oblast;GPS - šířka;GPS - délka;Model fotopasti;Správce;{";".join(header)}\n'

    def process_data(self):
        """ Go through trap entries and count active days for each season """
        # Add 0 value to missing dates - this handeles case when processing data from multiple sheets and a trap is not in all of the sheets
        for date in self.dates:
            for trap in self.traps.values():
                if date not in trap.entries.keys():
                    trap.entries[date] = 0

        # For each seasson
        for season_name, season in self.seasons.items():
            # For each trap
            for trap in self.traps.values():
                # For each entry
                for date, state in trap.entries.items():
                    # Get season year of that entry
                    season_year = self._get_season_year(date)
                    # Add data sctructure for this season_year to trap active_days if it is not there yet
                    trap.active_days.setdefault(season_year, OrderedDict())
                    # If the date is in currently checked season
                    if is_date_in_interval(date, season['start'], season['end']):
                        # Add default value for currently checked season to data sctructure of this season_year if it is not there yet
                        trap.active_days[season_year].setdefault(season_name, 0)
                        # If the value of checked entry is 1 count up the active_days
                        if state == 1:
                            trap.active_days[season_year][season_name] += 1

    def format_data(self):
        """ Create csv table from processed trap data """
        result = self._get_table_header()
        
        for trap in self.traps.values():
            active_days = []
            for seasons in trap.active_days.values():
                for season in seasons.values():
                    active_days.append(str(season))

            result += f'{trap.name};{trap.id};{trap.location};{trap.gps[0]};{trap.gps[1]};{trap.model};{trap.admin};{";".join(active_days)}\n'
        
        return result

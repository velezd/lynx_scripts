import openpyxl
import datetime
import json


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
        self.active_days = {}


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
                        dates.append(datetime.datetime(int(date[2]), int(date[1]), int(date[0])))
                    first =  False
                else:
                    self.traps[row[1].value] = \
                        Trap(row[0].value,
                             row[1].value,
                             row[2].value,
                             (row[3].value, row[4].value),
                             row[5].value,
                             row[6].value,
                             get_entries(row[7:], dates)
                        )

    def process_data(self):
        """ Go through trap entries and count active days for each season """
        for season_name, season in self.seasons.items():
            for trap in self.traps.values():
                if season_name not in trap.active_days:
                    trap.active_days[season_name] = 0
                for date, state in trap.entries.items():
                    if state == 1 and is_date_in_interval(date, season['start'], season['end']):
                        trap.active_days[season_name] += 1

    def format_data(self):
        """ Create csv table from processed trap data """
        result = f'Lokalita;ID lokality;Oblast;GPS - šířka;GPS - délka;Model fotopasti;Správce;{";".join(self.seasons.keys())}\n'
        
        for trap in self.traps.values():
            active_days = ';'.join([ str(trap.active_days[season]) for season in self.seasons.keys() ])
            result += f'{trap.name};{trap.id};{trap.location};{trap.gps[0]};{trap.gps[1]};{trap.model};{trap.admin};{active_days}\n'
        
        return result

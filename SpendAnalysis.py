from datetime import datetime, date, time
from typing import Union

import pandas


class SpendAnalysis:
    alias: str = None
    start_date: datetime = None
    end_date: datetime = None
    summary_by_category: pandas.DataFrame = None
    usedItems: pandas.DataFrame = None
    unfinishedItems: pandas.DataFrame = None
    
    def __init__(self, start_date: Union[datetime, date], end_date: Union[datetime, date], alias: str):
        if type(start_date == date):
            start_date = datetime.combine(start_date, time.min)
        if type(end_date == date):
            end_date = datetime.combine(end_date, time.min)
        self.start_date = start_date
        self.end_date = end_date
        self.alias = alias

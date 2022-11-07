from datetime import datetime

import pandas


class SpendAnalysis:
    alias: str = None
    start_date: datetime = None
    end_date: datetime = None
    summary_by_category: pandas.DataFrame = None
    usedItems: pandas.DataFrame = None
    unfinishedItems: pandas.DataFrame = None
    
    def __init__(self, start_date: datetime, end_date: datetime, alias: str):
        self.start_date = start_date
        self.end_date = end_date
        self.alias = alias

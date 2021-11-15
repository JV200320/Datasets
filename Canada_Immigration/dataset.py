import pandas as pd
import numpy as np

class Dataset():

    def __init__(self) -> None:
        immigr, reg_immigr, total_continent_immigr = self._check_for_excel_files_in_dir()
        if immigr.empty:
            self.immigr = self._get_immigration_data()
        else:
            self.immigr = immigr
        if reg_immigr.empty:
            self.reg_immigr = self._get_immigr_per_region()
        else:
            self.reg_immigr = reg_immigr
        if total_continent_immigr.empty:
            self.total_continent_immigr = self._get_total_immgr_per_continent()
        else:
            self.total_continent_immigr = total_continent_immigr

    def _check_for_excel_files_in_dir(self):
        try:
            immigr = pd.read_excel('Immigration.xlsx')
        except:
            immigr = pd.DataFrame()
        try:
            reg_immigr = pd.read_excel('RegionImmigration.xlsx', index_col=[0, 1])
        except:
            reg_immigr = pd.DataFrame()
        try:
            total_continent_immigr = pd.read_excel('TotalContinentImmigration.xlsx', index_col='AreaName')
        except:
            total_continent_immigr = pd.DataFrame()
        return immigr, reg_immigr, total_continent_immigr

    def _get_immigration_data(self):
        df = pd.read_excel('Canada.xlsx')
        df_immigr = self._clean_dataframe(df)
        df_immigr = self._create_northern_america_total(df_immigr)
        df_immigr = df_immigr[df_immigr['AreaName'] != 'Unknown']
        df_immigr.to_excel('Immigration.xlsx', index=False)
        return df_immigr

    def _get_immigr_per_region(self):
        reg_immigr =  self.immigr[self.immigr['AreaName'].apply(lambda area_name: not 'Total' in area_name)]
        reg_immigr = reg_immigr.set_index(['AreaName', 'RegName'])
        reg_immigr.to_excel('RegionImmigration.xlsx')
        return reg_immigr

    def _get_total_immgr_per_continent(self):
        df_continent_immigr = self.immigr.drop('RegName',axis=1)
        df_total_continent_immigr =  df_continent_immigr[df_continent_immigr['AreaName'].apply(lambda area_name: 'Total' in area_name)].set_index('AreaName')
        df_total_continent_immigr.to_excel('TotalContinentImmigration.xlsx')
        return df_total_continent_immigr

    def _clean_dataframe(self, dataset: pd.DataFrame) -> pd.DataFrame:
        df_immigr = dataset.iloc[20:,2:-1].rename(columns=dataset.iloc[19])
        df_immigr = df_immigr.replace('..', 0)
        return df_immigr.drop(20)

    def _create_northern_america_total(self, dataframe: pd.DataFrame) -> pd.DataFrame:
        row = dataframe[dataframe['AreaName'] == 'Northern America']
        row.at[row.index[0],'RegName'] = np.nan
        row.at[row.index[0], 'AreaName'] = 'Northern America Total'
        dataframe = dataframe.append(row, ignore_index=True)
        return dataframe

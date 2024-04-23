import regex as re
import win32com.client as win32
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter

class ParkRunData():
    '''
      A class containing methods to extract and visualise my park run results
    '''

    def __init__(self):
        self.data: dict[str, list[str]] = {
            'run_date': [],
            'time_achieved': []
        }


    def extract(self):
        '''
          extracts results data from park run emails in my Outlook account
        '''

        outlook = win32.Dispatch('Outlook.Application').GetNameSpace('MAPI')
        root_folder = outlook.Folders.Item(1)

        # get the results emails from my Outlook archive folder
        for folder in root_folder.Folders:

            if folder.Name == 'Archive':
                archive = folder.Items

                for mail in archive:

                    search_terms: list[str] = ['result', 'parkrun']
                    if all(term in mail.Subject for term in search_terms):
                        contents: str = mail.Body

                        re_pattern = r'Your time was ([\d|:]+)'
                        match_str = re.search(re_pattern, contents)

                        if match_str is not None:
                            result = match_str.group(1)

                            self.data['time_achieved'].append(result)
                            received_date = mail.ReceivedTime.strftime("%m-%d-%y")
                            self.data['run_date'].append(received_date)

    def visualise(self):
        '''
          produces a visulisation of the park run results data
        '''

        # create a dataframe from the results data & clean the data
        df = pd.DataFrame(self.data)

        if not df.empty:
            df['run_date'] = pd.to_datetime(df['run_date'], format="%m-%d-%y")
            df['time_achieved'] = pd.to_datetime(df['time_achieved'], format="%H:%M:%S")
            df = df.sort_values(by='run_date')

            # create the visulisation
            ax = plt.subplot()
            ax.plot(df['run_date'], df['time_achieved'], marker='x', linewidth=3)
            ax.yaxis.set_major_formatter(DateFormatter('%H:%M'))

            font = {'family': 'serif', 'color': 'darkred'}

            plt.xlabel('Date', fontdict=font, fontsize=15)
            plt.ylabel('Time', fontdict=font, fontsize=15)
            plt.title('My parkrun results', fontdict=font, loc='left', fontsize=20)
            plt.show()

import pandas as pd
import xlrd
import xlsxwriter

print('*********************************************'
      '\n    ADD CONTACT INFORMATION TO FILE\n'
      '**********************************************\n\n')

a = input('Enter the name of your CRM Contacts File: ')
a = 'Source_Files/' + a + '.xlsx'

def main():

    crm = pd.read_excel(a)
    crm['Account'] = crm['Account Name']
    crm = pd.DataFrame(data=crm, columns=['Account', 'First Name',
                                            'Last Name', 'Title',
                                            'Email Address', 'Office Phone'])

    inbound = input('Enter the file name: ')
    inbound = 'Source_Files/' + inbound + '.xlsx'
    try:
        target_report = pd.read_excel(inbound)
        group = input('Do you want to include the Corporate Group Y/N: ')
        if group == 'N' or group == 'n':
            target_report = target_report.drop(['Corp Group Name'], axis=1)
        else:
            target_report = pd.DataFrame(data= target_report, columns= ['Account', 'Machine Type', 'Last Scan Date',
                                                                        'Called By', 'Start Date and Time', 'Duration',
                                                                        'Status', 'Subject', 'Description', 'Direction']
                                         )
        target_report['Account'] = target_report['Dealer']
        target_report = target_report.drop(['Dealer Name'], axis=1)
    except:
        "FileNotFoundError:"
        print('Ooops - no files by that name.')


        #Output File#

    output = pd.merge(crm, target_report, on="Account")
    output = output[output['Title'].str.startswith('IT') == False]
    name = input('Please name your file: ')
    name = 'Output_Folder/' + name + '.xlsx'
    output.to_excel(name , sheet_name='sheet_1', engine='xlsxwriter')

    print(output)

if __name__ == '__main__':
    main()
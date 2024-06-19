import PyPDF2
import re
import pdfplumber
import pandas as pd
import os

def find_headings(name, file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)
        
        heading_pattern = re.compile(r'\[\d{6}\]\s.*')

        start_page_number = -1
        end_page_number = 0

        for page_num in range(num_pages):
            page_obj = reader.pages[page_num]
            text = page_obj.extract_text()
            
            if text:
                if name in text:
                    start_page_number = page_num
                elif heading_pattern.search(text) and start_page_number != -1:
                    end_page_number = page_num
                    break
        
        if end_page_number == 0:
            end_page_number = start_page_number + 1
        
        if start_page_number != -1: 
            table_data = extract_table(start_page_number, end_page_number, file_path, name)
            return table_data
        else:
            return None

def extract_table(start, end, file_path, name):
    all_data = pd.DataFrame()

    with pdfplumber.open(file_path) as pdf:
        for page_num in range(start, end + 1):
            page = pdf.pages[page_num]
            tables = page.extract_tables()

            for table in tables:
                table_data = pd.DataFrame(table)
                if all_data.empty:
                    all_data = table_data
                else:
                    if table_data.shape[1] == all_data.shape[1]:
                        all_data = pd.concat([all_data, table_data.iloc[1:]], ignore_index=True)
                    else:
                        all_data = pd.concat([all_data, table_data], ignore_index=True)
    all_data = check_present_column(name, all_data)

    return all_data

def check_present_column(name, all_data):
    if name == '[110000] Balance sheet':
        balance_sheet = ["Property, plant and equipment", "Other intangible assets", "Non-current investments", 
                         "Loans, non-current", "Total non-current financial assets", "Total non-current assets", 
                         "Inventories", "Current investments", "Trade receivables, current", 
                         "Cash and cash equivalents", "Loans, current", "Total current financial assets", 
                         "Total current assets", "Total assets", "Equity share capital", "Other equity", 
                         "Total equity", "Borrowings, non-current", "Total non-current financial liabilities", 
                         "Provisions, non-current", "Total non-current liabilities", "Borrowings, current", 
                         "Trade payables, current", "Total current financial liabilities", 
                         "Provisions, current", "Total current liabilities", "Total equity and liabilities"]
        
        all_data.columns = all_data.iloc[0]  
        all_data = all_data[1:]  
        all_data = all_data[all_data.iloc[:, 0].isin(balance_sheet)] 

    elif name == '[210000] Statement of profit and loss':
        Profit_and_Loss = ["Revenue from operations","Other income","Total income","Cost of materials consumed",
                           "Changes in inventories of finished goods, work-in-progress and stock-in-trade","Employee benefit expense",
                           "Finance costs","Depreciation, depletion and amortisation expense","Other expenses","Total expenses","Total profit before tax",
                           "Total tax expense","Total profit (loss) for period from continuing operations","Total profit (loss) for period",
                           "Whether company has other comprehensive income OCI components presented net of tax","Total comprehensive income",
                           "Comprehensive income OCI components presented before tax [Abstract]","Earnings per share [Abstract]","Whether company has comprehensive income OCI components presented before tax",]


        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Profit_and_Loss)] 

    elif name == "[320000] Cash flow statement, indirect":
        cash_flow =["Whether cash flow statement is applicable on company","Profit before tax"]
        
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(cash_flow)] 

    
    elif name == "[400100] Notes - Equity share capital":
        Equity_share = ["Disclosure of shareholding more than five per cent in company [Abstract]"]

        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Equity_share)]

    elif name == "[400600] Notes - Property, plant and equipment":
        Property_plant_equipment = ["Whether property, plant and equipment are stated at revalued amount"]
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Property_plant_equipment)]

    elif name == "[400700] Notes - Investment property":
        Investment_property = ["Depreciation method, investment property, cost model","Useful lives or depreciation rates, investment property, cost model"]
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Investment_property)]

    elif name == "[400900] Notes - Other intangible assets":
        intangible_assets = ["Total increase (decrease) in Other intangible assets"," Other intangible assets at end of period"]
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(intangible_assets)]

    elif name == "[401000] Notes - Biological assets other than bearer plants":
        Biological_assets = ["Depreciation method, biological assets other than bearer plants, at cost","Useful lives or depreciation rates, biological assets other than bearer plants, at cost"]
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Biological_assets)]

    elif name == "[401100] Notes - Subclassification and notes on liabilities and assets":
        Subclassification = ["Advances, non-current","Fixed deposits with banks","Total balance with banks","Cash on hand","Total cash and cash equivalents",
                             "Total cash and bank balances","Total balances held with banks to extent held as margin money or security against borrowings, guarantees or other commitments",
                             "Bank deposits with more than 12 months maturity","Interest accrued on borrowings",
                             "Interest accrued on public deposits","Interest accrued others","Unpaid dividends",
                             "Total application money received for allotment of securities and due for refund and interest accrued thereon",
                             "Unpaid matured deposits and interest accrued thereon","Unpaid matured debentures and interest accrued thereon","Debentures claimed but not paid",
                             "Public deposit payable, current","Current liabilities portion of share application money pending allotment"]
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Subclassification)]
    
    elif name == "[401200] Notes - Additional disclosures on balance sheet":
        additional_balance_sheet = ['Total contingent liabilities','Total contingent liabilities and commitments','Deposits accepted or renewed during period','Deposits matured and claimed but not paid during period',
                                    'Deposits matured and claimed but not paid','Deposits matured but not claimed','Interest on deposits accrued and due but not paid',
                                    'Share application money received during year','Share application money paid during year',
                                    'Amount of share application money received back during year','Amount of share application money repaid returned back during year',
                                    'Number of person share application money paid during year','Number of person share application money received during year',
                                    'Number of person share application money paid during year','Share application money received and due for refund',
                                    'Net worth of company','Unclaimed share application refund money','Unclaimed matured debentures',
                                    'Unclaimed matured deposits','Interest unclaimed amount','Investment in subsidiary companies',
                                    'Investment in government companies','Amount due for transfer to investor education and protection fund (IEPF)',
                                    'Gross value of transactions with related parties','Number of warrants converted into equity shares during period','Number of warrants converted into preference shares during period',
                                    'Number of warrants converted into preference shares during period','Number of warrants converted into debentures during period',
                                    'Number of warrants issued during period (in foreign currency)','Number of warrants issued during period (INR)',
                                    'Details of specified bank Notes held and transacted during the period from 8 November 2016 to 30 December 2016 [Table]']
        
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(additional_balance_sheet)]


    elif name == "[500100] Notes - Subclassification and notes on income and expenses":
        Subclassification_expenses = ['Total revenue from operations','Total interest income','Total dividend income','Total other income',
                                      'Total interest expense','Total finance costs','Salaries and wages','Total remuneration to directors',
                                      'Total managerial remuneration','Total employee benefit expense','Total depreciation, depletion and amortisation expense',
                                      'Consumption of stores and spare parts','Power and fuel','Rent','Repairs to building',
                                      'Repairs to machinery','Insurance','Total rates and taxes excluding taxes on income','Directors sitting fees',
                                      'Loss on disposal of intangible Assets','Loss on disposal, discard, demolishment and destruction of depreciable property plant and equipment',
                                      'Total payments to auditor','CSR expenditure','Miscellaneous expenses']
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Subclassification_expenses)]

    elif name == "[500200] Notes - Additional information statement of profit and loss":
        statement_of_profit_and_loss = ['Total changes in inventories of finished goods, work-in-progress and stock-in-trade','Total revenue from sale of products',
                                        'Total revenue from sale of services','Gross value of transaction with related parties','Bad debts of related parties']
        
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(statement_of_profit_and_loss)]


    elif name == "[610200] Notes - Corporate information and statement of IndAs compliance":
        Corporate_info = ['Statement of Ind AS compliance [TextBlock]','Whether there is any departure from Ind AS',
                          'Whether there are reclassifications to comparative amounts']
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Corporate_info)] 


    elif name == "[610300] Notes - Accounting policies, changes in accounting estimates and errors":
        Accounting_policies = ['Whether initial application of an Ind AS has an effect on the current period or any prior period','Whether there is any voluntary change in accounting policy',
                               'Whether there are changes in acounting estimates during the year']
        
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Accounting_policies)] 

    elif name == "[610500] Notes - Events after reporting period":
        Events_after_reporting = ['Whether there are non adjusting events after reporting period']

        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Events_after_reporting)]

    elif name == "[610700] Notes - Business combinations":
        Bussiness_combination_1 = ['Whether there is any business combination','Whether there is any goodwill arising out of business combination',
                                   'Whether there are any acquired receivables from business combination','Whether there are any contingent liabilities in business combination']
        
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Bussiness_combination_1)]

    elif name == "[610800] Notes - Related party":
        Related_party = ['Whether there are any related party transactions during year','Whether entity applies exemption in Ind AS 24.25']
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Related_party)]


    elif name == "[610900] Notes - First time adoption":
        First_time_adoption = ['Whether company has adopted Ind AS first time']
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(First_time_adoption)]


    elif name == "[611000] Notes - Exploration for and evaluation of mineral resources":
        evalution_of_mineral = ['Whether there are any exploration and evaluation activities']
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(evalution_of_mineral)]

    elif name == "[611200] Notes - Fair value measurement":
        Fair_value_measurement = ['Whether assets have been measured at fair value',' Whether liabilities have been measured at fair value']
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Fair_value_measurement)]

    elif name == "[611500] Notes - Interests in other entities":
        Intrest_in_other_entities = ['Whether company has subsidiary companies which are yet to commence operations',
                                     ' Whether company has subsidiary companies liquidated or sold during year','Whether company has invested in associates',
                                     'Whether company has associates which are yet to commence operations','Whether company has invested in joint ventures','Whether company has joint ventures which are yet to commence operations',
                                     'Whether company has joint ventures liquidated or sold during year','Whether there are unconsolidated structured entities',
                                     'Whether there are unconsolidated subsidiaries','Whether there are unconsolidated structured entities controlled by investment entity']
        
        all_data.columns = all_data.iloc[0] 
        all_data = all_data[1:]  
        
        
        all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
        
        all_data = all_data[all_data.iloc[:, 0].isin(Intrest_in_other_entities)]






    return all_data




names = ['[110000] Balance sheet','[210000] Statement of profit and loss','[320000] Cash flow statement, indirect','[400100] Notes - Equity share capital','[400600] Notes - Property, plant and equipment',
        '[400700] Notes - Investment property','[400900] Notes - Other intangible assets','[401000] Notes - Biological assets other than bearer plants',
        '[401100] Notes - Subclassification and notes on liabilities and assets','[401200] Notes - Additional disclosures on balance sheet','[500100] Notes - Subclassification and notes on income and expenses',
        '[500200] Notes - Additional information statement of profit and loss','[610200] Notes - Corporate information and statement of IndAs compliance',
        '[610300] Notes - Accounting policies, changes in accounting estimates and errors','[610500] Notes - Events after reporting period','[610700] Notes - Business combinations',
        '[610800] Notes - Related party','[610900] Notes - First time adoption','[611000] Notes - Exploration for and evaluation of mineral resources','[611200] Notes - Fair value measurement',
        '[611500] Notes - Interests in other entities']

file_path = '/Users/geetanshjoshi/Desktop/Findata/pdf/TATA 2.pdf'


output_dir = "extracted_tables"
os.makedirs(output_dir, exist_ok=True)

for name in names:
    table_data = find_headings(name, file_path)
    if table_data is not None:
        output_file_csv = os.path.join(output_dir, f"{name.replace(' ', '_').replace('[', '').replace(']', '')}.csv")
        table_data.to_csv(output_file_csv, index=False)
        print(f"Extracted data for {name} saved to {output_file_csv}")
    else:
        print(f"No matching heading found for {name}")
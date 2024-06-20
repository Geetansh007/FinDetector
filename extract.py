import PyPDF2
import re
import pdfplumber
import pandas as pd
import os
from xlsxwriter import Workbook

class PDFExtractor:
    def __init__(self, file_path, output_dir):
        self.file_path = file_path
        self.output_dir = output_dir
        self.names = ['[110000] Balance sheet','[210000] Statement of profit and loss','[320000] Cash flow statement, indirect',
                      '[400100] Notes - Equity share capital','[400600] Notes - Property, plant and equipment',
                      '[400700] Notes - Investment property','[400900] Notes - Other intangible assets',
                      '[401000] Notes - Biological assets other than bearer plants','[401100] Notes - Subclassification and notes on liabilities and assets',
                      '[401200] Notes - Additional disclosures on balance sheet','[500100] Notes - Subclassification and notes on income and expenses',
                      '[500200] Notes - Additional information statement of profit and loss','[610200] Notes - Corporate information and statement of IndAs compliance',
                      '[610300] Notes - Accounting policies, changes in accounting estimates and errors','[610500] Notes - Events after reporting period',
                      '[610700] Notes - Business combinations','[610800] Notes - Related party','[610900] Notes - First time adoption',
                      '[611000] Notes - Exploration for and evaluation of mineral resources','[611200] Notes - Fair value measurement',
                      '[611500] Notes - Interests in other entities','[611700] Notes - Other provisions, contingent liabilities and contingent assets',
                      '[611900] Notes - Accounting for government grants and disclosure of government assistance','[612000] Notes - Construction contracts',
                      '[612100] Notes - Impairment of assets','[612200] Notes - Leases','[612300] Notes - Transactions involving legal form of lease',
                      '[612400] Notes - Service concession arrangements','[612500] Notes - Share-based payment arrangements',
                      '[612600] Notes - Employee benefits','[612800] Notes - Borrowing costs','[612900] Notes - Insurance contracts',
                      '[613000] Notes - Earnings per share','[613100] Notes - Effects of changes in foreign exchange rates',
                      '[613300] Notes - Operating segments','[613400] Notes - Consolidated Financial Statements',
                      '[700100] Notes - Key managerial personnels and directors remuneration and other information','[700200] Notes - Corporate social responsibility',
                      '[700300] Disclosure of general information about company','[700400] Disclosures - Auditors report',
                      '[700500] Disclosures - Signatories of financial statements','[700600] Disclosures - Directors report',
                      "[700700] Disclosures - Secretarial audit report"]



    def shorten_sheet_name(self, name):
        return ''.join(e for e in name if e.isalnum())[:31]

    

    def extract_all_tables(self):
        os.makedirs(self.output_dir, exist_ok=True)
        
        for name in self.names:
            try:
                table_data = self.find_headings(name)
                if table_data is not None:
                    file_name = f"{self.shorten_sheet_name(name)}.xlsx"
                    excel_path = os.path.join(self.output_dir, file_name)
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        table_data.to_excel(writer, sheet_name='Sheet1', index=False)
                    print(f"Extracted data for {name} saved to {file_name}")
                else:
                    print(f"No matching heading found for {name}")
            except Exception as e:
                print(f"Error processing {name}: {e}")

            
    def find_headings(self, name):
        try:
            with open(self.file_path, 'rb') as file:
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
                    return self.extract_table(start_page_number, end_page_number, name)
                else:
                    return None
        except Exception as e:
            print(f"Error in find_headings for {name}: {e}")
            return None

    def extract_table(self, start, end, name):
        try:
            all_data = pd.DataFrame()
            with pdfplumber.open(self.file_path) as pdf:
                for page_num in range(start, end + 1):
                    if page_num < len(pdf.pages):
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
            return self.check_present_column(name, all_data)
        except Exception as e:
            print(f"Error in extract_table for {name}: {e}")
            return pd.DataFrame()

    def check_present_column(self, name, all_data):
        try:
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


            elif name == "[611700] Notes - Other provisions, contingent liabilities and contingent assets":
                other_provisions = ['Whether there are any contingent liabilities']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(other_provisions)]


            elif name == "[611900] Notes - Accounting for government grants and disclosure of government assistance":
                goverment_grants = ['Whether company has received any government grant or government assistance']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(goverment_grants)]

            elif name == "[612000] Notes - Construction contracts":
                Construction_contracts = ['Whether there are any construction contracts']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(Construction_contracts)]


            elif name == "[612100] Notes - Impairment of assets":
                Impairment = ['Whether there is any impairment loss or reversal of impairment loss during the year','Whether impairment loss recognised or reversed for individual Assets or cash-generating unit']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(Impairment)]



            elif name == "[612200] Notes - Leases":
                Leases = ['Whether company has entered into any lease agreement','Whether any operating lease has been converted to financial lease or vice-versa']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(Leases)]

            elif name == "[612300] Notes - Transactions involving legal form of lease":
                Transactions_involving = ['Whether there are any arrangements involving legal form of lease']

                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(Transactions_involving)]


            elif name == "[612400] Notes - Service concession arrangements":
                arrangements = ['Whether there are any service concession arrangments']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(arrangements)]


            elif name == "[612500] Notes - Share-based payment arrangements":
                share_based = ['Whether there are any share based payment arrangement']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(share_based)]



            elif name == "[612600] Notes - Employee benefits":
                benefits = ['Whether there are any defined benefit plans']
                all_data.columns = all_data.iloc[0] 
                all_data = all_data[1:]  
                
                
                all_data.iloc[:, 0] = all_data.iloc[:, 0].apply(lambda x: x.replace('\n', ' ') if isinstance(x, str) and '\n' in x else x)
                
                all_data = all_data[all_data.iloc[:, 0].isin(benefits)]


            elif name == '[612800] Notes - Borrowing costs':
                costs = ['Whether any borrowing costs has been capitalised during the year']  
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(costs)]

            elif name  == "[612900] Notes - Insurance contracts":
                contracts = ['Whether there are any insurance contracts as per Ind AS 104']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(contracts)]

            elif name  == "[613000] Notes - Earnings per share":
                per_share = ['Profit (loss), attributable to ordinary equity holders of parent entity','Profit (loss), attributable to ordinary equity holders of parent entity including dilutive effects',
                            'Weighted average number of ordinary shares outstanding']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(per_share)]



            elif name == "[613100] Notes - Effects of changes in foreign exchange rates":
                foregin_exchange = ['Whether there is any change in functional currency during the year','Description of presentation currency']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(foregin_exchange)]


            elif name == '[613300] Notes - Operating segments':
                operating_segments = ['Whether there are any reportable segments','Whether there are any major customers']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(operating_segments)]


            elif name == "[613400] Notes - Consolidated Financial Statements":
                Financial_Statements =['Whether consolidated financial statements is applicable on company']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(Financial_Statements)]

            elif name == "[700100] Notes - Key managerial personnels and directors remuneration and other information":
                Key_managerial = ['Name of key managerial personnel or director',' Director identification number of key managerial personnel or director','Shares held by key managerial personnel or director',
                                'Salary key managerial personnel or director','Total key managerial personnel or director remuneration']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(Key_managerial)]




            elif name =="[700200] Notes - Corporate social responsibility":
                corporate_social = ['Whether provisions of corporate social responsibility are applicable on company']
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(corporate_social)]

            elif name == "[700300] Disclosure of general information about company":
                genral_information = ['Name of company','Corporate identity number','Permanent account number of entity',
                                    'Address of registered office of company','Type of industry','Date of board meeting when final accounts were approved',
                                    'Date of start of reporting period','Date of end of reporting period','Nature of report standalone consolidated',
                                    'Content of report','Description of presentation currency','Level of rounding used in financial statements','Whether company is maintaining books of account and other relevant books and papers in electronic form']

                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(genral_information)]

            elif name == "[700400] Disclosures - Auditors report":
                Auditors_report =['Whether companies auditors report order is applicable on company',"Whether auditors' report has been qualified or has any reservations or contains adverse remarks","Category of auditor",
                                'Name of audit firm','Name of auditor signing report','Firms registration number of audit firm','Membership number of auditor','Address of auditors',"Permanent account number of auditor or auditor's firm",
                                "SRN of form ADT-1"," Date of signing audit report by auditors",'Date of signing of balance sheet by auditors',
                                'Whether companies auditors report order is applicable on company',"Whether auditors' report has been qualified or has any reservations or contains adverse remarks"]
                
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(Auditors_report)]

            elif name == "[700500] Disclosures - Signatories of financial statements":
                Signatories = ['First name of director','Middle name of director','Last name of director','Designation of director','Director identification number of director',
                            'Date of signing of financial statements by director']
                
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(Signatories)]

            elif name == "[700600] Disclosures - Directors report":
                Director_report = ['Disclosure in board of directors report explanatory [TextBlock]','Description of state of companies affair','Disclosure relating to amounts if any which is proposed to carry to any reserves',
                                'Disclosures relating to amount recommended to be paid as dividend','Details regarding energy conservation','Details regarding technology absorption','Details regarding foreign exchange earnings and outgo',
                                'Disclosures in director’s responsibility statement','Details of material changes and commitment occurred during period affecting financial position of company',
                                'Particulars of loans guarantee investment under section 186 [TextBlock]','Particulars of contracts/arrangements with related parties under section 188(1) [TextBlock]',"Whether there are contracts/arrangements/transactions not at arm's length basis",
                                "Whether there are material contracts/arrangements/transactions at arm's length basis","Disclosure of extract of annual return as provided under section 92(3) [TextBlock]"," Name of main product/service","Description of main product/service",
                                "NIC code of product/service","Percentage to total turnover of company","Disclosure for companies covered under section 178(1) on directors appointment and remuneration including other matters provided under section 178(3) [TextBlock]",
                                "Disclosure of statement on development and implementation of risk management policy [TextBlock]","Details on policy development and implementation by company on corporate social responsibility initiatives taken during year",
                                "Disclosure of financial summary or highlights [TextBlock]","Disclosure of change in nature of business [TextBlock]","Details of directors or key managerial personnels who were appointed or have resigned during year [TextBlock]"," Disclosure of companies which have become or ceased to be its subsidiaries, joint ventures or associate companies during year [TextBlock]",
                                "Details relating to deposits covered under chapter v of companies act [TextBlock]","Details of deposits which are not in compliance with requirements of chapter v of act [TextBlock]","Details of significant and material orders passed by regulators or courts or tribunals impacting going concern status and company’s operations in future [Text Block]",
                                "Details regarding adequacy of internal financial controls with reference to financial statements [Text Block]","Disclosure of appointment and remuneration of director or managerial personnel if any, in the financial year [Text Block]","Number of meetings of board","First name of director",
                                "Middle name of director","Last name of director","Designation of director","Director identification number of director","Date of signing board report"]
                
                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(Director_report)]


            elif name == "[700700] Disclosures - Secretarial audit report":
                Secretarial_audit = ["Whether secretarial audit report is applicable on company"]

                all_data.columns = all_data.iloc[0]
                all_data = all_data[1:]

                all_data.iloc[:,0]= all_data.iloc[:,0].apply(lambda x:x.replace('\n',' ') if isinstance(x,str) and '\n' in x else x)
                all_data = all_data[all_data.iloc[:, 0].isin(Secretarial_audit)]
            return all_data

        except Exception as e:
            print(f"Error in check_present_column for {name}: {e}")
            return pd.DataFrame()


if __name__ == "__main__":
    extractor = PDFExtractor(file_path="", output_dir="")
    extractor.extract_all_tables()
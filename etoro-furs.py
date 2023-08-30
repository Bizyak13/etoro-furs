# !/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import csv
import json
import os.path
import argparse
import xml.etree.ElementTree
from urllib import request
from datetime import datetime
from datetime import timedelta
import configparser

import openpyxl

bsi_rate_url = 'https://www.bsi.si/_data/tecajnice/dtecbs-l.xml'

csv_1st_line = '#FormCode;Version;TaxPayerID;TaxPayerType;DocumentWorkflowID;;;;;;'
csv_2nd_line = 'DOH-DIV;3.9;xxxxxxxx;FO;O;;;;;;'
csv_3rd_line = '#datum prejema dividende;davčna številka izplačevalca dividend;identifikacijska  številka izplačevalca dividend;naziv izplačevalca dividend;naslov izplačevalca dividend;država izplačevalca dividend;vrsta dividende;znesek dividend;tuji davek;država vira;uveljavljam oprostitev po mednarodni pogodbi'


def get_rounded_float(number) -> str:
    number = str(number)
    if ',' in number and '.' in number:
        if number.index(',') < number.index('.'):
            number = number.replace(',', '')
        else:
            number = number.replace('.', '')            
    number = float(number)
    number = f'{number:.2f}'
    return str(number).replace('.', '#').replace(',', '.').replace('#', ',')


def remove_offending_rows(table, keep) -> openpyxl.worksheet:
    for row in range(table.max_row+1, 2, -1):
        if table.cell(row=row, column=2).value != keep:
            table.delete_rows(row, 1)
    
    return table


def get_config() -> configparser.RawConfigParser:
    config = configparser.RawConfigParser()
    config.read('CONFIG.cfg')
    return config


def get_config_value(section, key) -> any:
    value = config[section][key]
    if value is not None:
        return config[section][key]
    else: 
        raise SystemExit(f'ERROR: Missing config value for {section}:{key}')
        


def parse_args() -> list:
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help='Input file from etoro. Must be in .xlsx format.', type=str)
    parser.add_argument('output', help='Output file from csv. Must be in .csv format.', type=str)
    parser.add_argument('-v', '--verbose', help='Verbose output', action='store_true')
    args = parser.parse_args()

    if args.input.split('.')[-1] != 'xlsx':
        raise SystemExit('ERROR: Provided input file not in the right format')

    if os.path.exists(args.input) == False:
        raise SystemExit('ERROR: Provided input file does not exist')

    return args


def get_conversion_rate_file() -> dict:
    filename = 'currency-rates.xml'
    if not os.path.exists(filename):
        request.urlretrieve(bsi_rate_url, filename)

    conversion_file = xml.etree.ElementTree.parse(filename).getroot()

    rates = {}
    for d in conversion_file:
        date = d.attrib["datum"].replace("-", "")
        rates[date] = {}
        for r in d:
            currency = r.attrib["oznaka"]
            rates[date][currency] = r.text
    return rates


def get_conversion_rate_on_date(rates, date, currency) -> any:
    date = date.strftime('%Y%m%d')
    if date in rates and currency in rates[date]:
        rate = rates[date][currency]
    else:
        rate = 0
    return rate


def get_config_taxid() -> int:
    tax_id = get_config_value('TAX_ID','tax_id')
    if len(tax_id) == 8 and int(tax_id):
        return tax_id
    else:
        raise SystemExit('ERROR: Tax ID is not in the right format')


def parse_input_file(rates) -> dict:
    data = {}
    workbook = openpyxl.load_workbook(args.input)
    dividends = workbook['Dividends']
    activity = workbook['Account Activity']
    activity = remove_offending_rows(activity, 'Dividend')

    max_col = dividends.max_column
    max_row = dividends.max_row

    for i in range(2, max_row+1):
        rate = 0
        data_row = {}
        date = datetime.strptime(dividends.cell(row=i, column=1).value, '%d/%m/%Y')
        # get price data for every row, based on the date
        for j in range(1, max_col + 1):
            data_row[dividends.cell(row=1, column=j).value] = dividends.cell(row=i, column=j).value
            data_row['Date of Payment FURS'] = date.strftime('%d. %m. %Y')
            if j == 2:
                # get data about the company and append it here
                symbol_curr = activity.cell(row=i, column=3).value.split('/')
                data_row['Symbol'] = symbol_curr[0]
                data_row['Currency'] = symbol_curr[1]
                company_json = json.load(open('companies.json', encoding='utf-8'))
                for company in company_json['companies']:
                    if company['symbol'] == activity.cell(row=i, column=3).value.split('/')[0]:
                        data_row['Company Name'] = company['name']
                        data_row['Company Address'] = company['address']
                        data_row['Company Country'] = company['country']
                        data_row['Company TAX ID'] = company['taxNumber']

            if j == 3:
                if date.strftime('%Y%m%d') in rates:
                    rate = float(rates[date.strftime('%Y%m%d')][activity.cell(row=i, column=3).value.split('/')[1]])
                else:
                    for k in range(1, 10):
                        last_working = (date - timedelta(days=k)).strftime('%Y%m%d')
                        if last_working in rates:
                            rate = float(rates[last_working][data_row['Currency']])
                            data_row['Conversion rate date'] = (date - timedelta(days=k)).strftime('%Y-%m-%d')
                            break
                    if rate == 0:
                        raise SystemExit('ERROR: No exchange rate found for this date')
                data_row[f'Conversion rate (EUR/{data_row["Currency"]})'] = rate
                data_row['Net Dividend Received (EUR)'] = get_rounded_float(dividends.cell(row=i, column=3).value/rate)
                
            if j == 5:
                data_row['Withholding Tax Amount (EUR)'] = get_rounded_float(dividends.cell(row=i, column=5).value/rate)
        
        data[i-1] = data_row
    return data


def create_output_file(data) -> str:
    
    if 'csv' not in args.output:
        output = args.output + '.csv'
    else:
        output = args.output
    
    csv_output = open(output, 'w', encoding='utf-8')
    csv_writer = csv.writer(csv_output, lineterminator='\n', delimiter=';')
    csv_writer.writerow(csv_1st_line.split(';'))
    csv_writer.writerow(csv_2nd_line.replace('xxxxxxxx', get_config_taxid()).split(';'))
    csv_writer.writerow(csv_3rd_line.split(';'))
    for l in range(1, len(data)+1):
        line = data[l]
        dividend_type = get_config_value('TAX_ID', 'DIVIDEND_TYPE')
        if not dividend_type:
            dividend_type = 1
        csv_writer.writerow([line['Date of Payment FURS'], '', line['Company TAX ID'], line['Company Name'], line['Company Address'], line['Company Country'], dividend_type, line['Net Dividend Received (EUR)'], line['Withholding Tax Amount (EUR)'], line['Company Country'], ''])
    csv_output.close()

    return output


if __name__ == '__main__':
    print('etoro-furs: Running etoro-furs')
    args = parse_args()
    config = get_config()
    rates = get_conversion_rate_file()
    print('etoro-furs: Rates loaded')
    
    print('etoro-furs: Parsing input file')
    data = parse_input_file(rates)
    print('etoro-furs: Generating csv file')
    file = create_output_file(data)
    print(f'etoro-furs: DONE! -> {file}')

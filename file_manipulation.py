import pandas as pd
import xlrd
import datetime
from bs4 import BeautifulSoup
import requests
import time
import dash
import dash_bootstrap_components
from dash import Dash, html, dash_table

start_time = time.time()/60

files = ['Stock Level - for Marketing.xls', 'Polymer Shipment - Export.xls', 'Polymer Shipment - Local.xls', 'Plan.xls']
file_dataframes = []
file_name = 0
now = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m')
params_to = datetime.datetime.strftime(datetime.datetime.now() - datetime.timedelta(days=1), '%Y-%m-%d')

for file in files:
    path = xlrd.open_workbook(file)
    df = pd.read_excel(path)
    if file == files[0]:
        df = df.dropna(axis=0, how='all').reset_index(drop=True).drop([0, 1, 2]).reset_index(drop=True).dropna(axis=1, how='all')
    else:
        df = df.dropna(axis=0, how='all').reset_index(drop=True).drop([0, 1, 2, 3]).reset_index(drop=True).dropna(axis=1, how='all')
    df = df.dropna(subset=['Unnamed: 0']).reset_index(drop=True)
    df.insert(0, "Grade", '', True)
    df = pd.concat([df, pd.DataFrame(columns=[
        'SA',
        'SB',
        'SC',
        'spec',
        'Total',
        # '1',
        'Total All Sorts',
        # '2',
        'Prod.Plan',
        'PI',
        'Local Sales',
        'Overall Local Sales',
        # '3',
        'Actual Local Shipment',
        'Actual Export Shipment',
        '4',
        'Rem.Production',
        'Rem.Offtakers Shipment',
        'Rem.Local Shipment',
        'Expected Stock'])])

    for ind, value in df['Unnamed: 0'].items():
        df.at[ind, 'Unnamed: 0'] = str(value).replace('-', '')
        df.at[ind, 'Grade'] = str(value).replace('-', '')
        if value == 'On spec' or value == 'Off spec SA' or value == 'Off spec SB' or value == 'Off spec SC':
            grade_sort = f"{df.at[ind - 1, 'Grade']} {value}".split(' ')
            df.at[ind, 'Grade'] = f"{grade_sort[0]} {grade_sort[-1]}"
        else:
            if file == files[3]:
                df.at[ind, 'Grade'] = f"{str(value).replace('-', '')} spec"

    # file_name += 1
    # df.to_excel(f"{file_name}.xlsx")
    file_dataframes.append(df)


df_coords = {
    'Seller': [3, 15],
    'Date': [1, 15],
    'Vol': [7, 15],
    'Price': [8, 15],
    'Amnt': [9, 15],
    'Status': [13, 15]
}
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Connection': 'keep-alive',
    # 'Cookie': f'ASP.NET_SessionId={cookie[0]}; IsViewedSwal12; IsViewedSwal12=1',
    'Referer': 'https://broker.uzex.uz/profile/orders?mode=custom',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}

params = {
    'from': '2023-01-01',
    'to': params_to,
    'excel': 'no',
    'mode': 'custom',
}

payload = {
    'Username': 'bk877',
    'Password': 'az777*07az'
}

login = requests.post("https://broker.uzex.uz", data=payload, allow_redirects=True)
with requests.Session() as session:
    post = session.post("https://broker.uzex.uz", data=payload)
    response = session.get("https://broker.uzex.uz/profile/orders?mode=custom", headers=headers, params=params)
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('tbody')
    data_dict = {}
    data_list = [row.get_text() for row in table.find_all('td')]
    grade_list = [row.attrs.get('data-content') for row in table.find_all('a')]
    for i, v in df_coords.items():
        data_dict[i] = [seller for seller in data_list[v[0]::v[1]]]
    data_dict['Grade1'] = [grade for grade in grade_list[1::7]]
    df = pd.DataFrame(data_dict)
    df['Seller'] = df['Seller'].str.upper()
    df['Grade1'] = df['Grade1'].str.upper()
    df = (df[(df['Seller'] == 'СП ООО "UZ-KOR GAS CHEMICAL"') & (df['Status'] != '5. Аннулирован') & ((df['Grade1'].str.match(r'ПОЛИПРОПИЛЕН')) | (df['Grade1'].str.match(r'ПОЛИЭТИЛЕН')))]).reset_index(drop=True)
    df['Vol'] = df['Vol'].replace(regex=[r' тонна', r' килограмм', '\n'], value='')
    df['Price'] = df['Price'].replace(regex=r',', value='.')
    df['Amnt'] = df['Amnt'].replace(regex=r',', value='.')
    df['Grade1'] = df['Grade1'].replace(regex=[r'СП ООО UZ-KOR GAS CHEMICAL', r'СП ООО "UZ-KOR GAS CHEMICAL"', r'-', r' '], value='')
    df = df.astype({"Vol": "float", "Amnt": "float", "Price": "float"})
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True).dt.normalize()
    df_monthly = ((df[df['Date'].dt.strftime('%Y-%m') == str(now)]).reset_index(drop=True).groupby(['Date', 'Grade1'], as_index=False)[['Vol', 'Amnt']].sum())
    df_overall = (pd.concat([pd.read_excel('Rem.vol.xlsx', dtype='object'), df.groupby(['Grade1'], as_index=False)[['Vol', 'Amnt']].sum()]))
    df.groupby(['Grade1'], as_index=False)[['Vol', 'Amnt']].sum()
    # df.to_excel('df.xlsx')
    # df_monthly.to_excel('df_monthly.xlsx')
    # df_overall.to_excel('df_overall.xlsx')
i = 0
for new_df in [df_monthly, df_overall]:
    try:
        for idx, val in file_dataframes[0]['Grade'].items():
            if val != file_dataframes[0].at[idx, 'Unnamed: 0']:
                new_df.loc[(new_df['Grade1'].str.contains(val.split(' ')[0]) & (new_df['Grade1'].str.contains(val.split(' ')[-1]) | new_df['Grade1'].str.contains('GСОРТ'))), 'Grade'] = val
            else:
                new_df.loc[(new_df['Grade1'].str.contains(val.split(' ')[0])), 'Grade'] = f"{val} spec"
        new_df = new_df.groupby(['Grade'], as_index=False)[['Vol', 'Amnt']].sum()
        new_df['Price'] = new_df['Amnt'].div(new_df['Vol'])
    except:
        pass
    finally:
        file_dataframes.append(new_df)

    # i += 50
    # new_df.to_excel(f"{i}.xlsx")


file_dataframes[0] = file_dataframes[0].fillna(0)
for ind, value in file_dataframes[0]['Grade'].items():
    file_dataframes[0].at[ind, 'Unnamed: 4'] /= 1000
    file_dataframes[0].at[ind, 'Unnamed: 3'] /= 1000
    file_dataframes[0].at[ind, 'Total All Sorts'] = file_dataframes[0].at[ind, 'Unnamed: 4']

    if 'SA' in value or 'SB' in value or 'SC' in value or 'spec' in value:
        if not file_dataframes[4].empty:
            sales_df_filter = file_dataframes[4].loc[file_dataframes[4]['Grade'] == value]
            if value in sales_df_filter['Grade'].values:
                idx = sales_df_filter['Grade'].index.values[0]
                file_dataframes[0].at[ind, 'Local Sales'] = file_dataframes[4].at[idx, 'Vol']
        else:
            file_dataframes[0].at[ind, 'Local Sales'] = 0

        overall_sales_df_filter = file_dataframes[5].loc[file_dataframes[5]['Grade'] == value]
        if value in overall_sales_df_filter['Grade'].values:
            idx = overall_sales_df_filter['Grade'].index.values[0]
            file_dataframes[0].at[ind, 'Overall Local Sales'] = file_dataframes[5].at[idx, 'Vol']

        plan_df_filter = file_dataframes[3].loc[file_dataframes[3]['Grade'] == value]
        if value in plan_df_filter['Grade'].values:
            idx = plan_df_filter['Grade'].index.values[0]
            file_dataframes[0].at[ind, 'Prod.Plan'] = file_dataframes[3].at[idx, 'Prod.Plan']
            file_dataframes[0].at[ind, 'PI'] = file_dataframes[3].at[idx, 'Lotte'] + file_dataframes[3].at[idx, 'Samsung']

        exp_ship_df_filter = file_dataframes[1].loc[file_dataframes[1]['Grade'] == value]
        if value in exp_ship_df_filter['Grade'].values:
            idx = exp_ship_df_filter['Grade'].index.values[0]
            file_dataframes[0].at[ind, 'Actual Export Shipment'] = file_dataframes[1].at[idx, 'Unnamed: 3'] / 1000

        loc_ship_df_filter = file_dataframes[2].loc[file_dataframes[2]['Grade'] == value]
        if value in loc_ship_df_filter['Grade'].values:
            idx = loc_ship_df_filter['Grade'].index.values[0]
            file_dataframes[0].at[ind, 'Actual Local Shipment'] = file_dataframes[2].at[idx, 'Unnamed: 3'] / 1000

        file_dataframes[0].at[ind, 'Rem.Production'] = file_dataframes[0].at[ind, 'Prod.Plan'] - file_dataframes[0].at[ind, 'Unnamed: 3']
        file_dataframes[0].at[ind, 'Rem.Offtakers Shipment'] = file_dataframes[0].at[ind, 'PI'] - file_dataframes[0].at[ind, 'Actual Export Shipment']
        file_dataframes[0].at[ind, 'Rem.Local Shipment'] = file_dataframes[0].at[ind, 'Overall Local Sales'] - file_dataframes[0].at[ind, 'Actual Local Shipment']
        file_dataframes[0].at[ind, 'Unnamed: 4'] = file_dataframes[0].at[ind, 'Unnamed: 4'] - file_dataframes[0].at[ind, 'Rem.Local Shipment']

        idx = file_dataframes[0][file_dataframes[0]['Unnamed: 0'] == value.split(' ')[0]].index.values[0]
        file_dataframes[0].at[idx, value.split(' ')[-1]] = file_dataframes[0].at[ind, 'Unnamed: 4']
        if 'spec' in value:
            for col in ['Prod.Plan', 'PI', 'Local Sales', 'Overall Local Sales', 'Rem.Production', 'Rem.Offtakers Shipment', 'Rem.Local Shipment']:
                file_dataframes[0].at[idx, col] = file_dataframes[0].at[ind, col]

file_dataframes[0]['Expected Stock'] = file_dataframes[0]['spec'] + file_dataframes[0]['Rem.Production'] - file_dataframes[0]['Rem.Offtakers Shipment']
file_dataframes[0]['Total'] = file_dataframes[0][['SA', 'SB', 'SC', 'spec']].sum(axis=1)

for ind, value in file_dataframes[0]['spec'].items():
    if value == 0 and file_dataframes[0].at[ind, 'SC'] == 0 and file_dataframes[0].at[ind, 'SB'] == 0 and file_dataframes[0].at[ind, 'SA'] == 0:
        file_dataframes[0] = file_dataframes[0].drop([ind])
file_dataframes[0] = file_dataframes[0].drop(columns=['Overall Local Sales', 'Actual Local Shipment', 'Actual Export Shipment', '4', 'Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4']).rename({'spec': 'ONSPEC'}, axis='columns').reset_index(drop=True)

pp_total_ind = file_dataframes[0]['Grade'][file_dataframes[0]['Grade'] == 'Y130'].index.tolist()[0]
hdpe_total_ind = file_dataframes[0]['Grade'].index[-1]

file_dataframes[0].loc[pp_total_ind + 0.5] = file_dataframes[0].loc[:pp_total_ind].sum(numeric_only=True, axis=0)
file_dataframes[0].at[pp_total_ind + 0.5, 'Grade'] = 'TOTAL'

file_dataframes[0].loc[hdpe_total_ind + 0.5] = file_dataframes[0].loc[(pp_total_ind+1):hdpe_total_ind].sum(numeric_only=True, axis=0)
file_dataframes[0].at[hdpe_total_ind + 0.5, 'Grade'] = 'TOTAL'

file_dataframes[0] = file_dataframes[0].sort_index().reset_index(drop=True).set_index('Grade')
file_dataframes[0].loc['GRAND TOTAL'] = file_dataframes[0].loc['TOTAL'].sum()


file_dataframes[0] = file_dataframes[0].round(3).replace(0, '-')
file_dataframes[0] = file_dataframes[0].reset_index()
# file_dataframes[0].to_excel('d1gdfgdgdgdf.xlsx')

app = Dash(__name__)
server = app.server
app.layout = html.Div([
    html.H1(children="Stock Level", style={'textAlign': 'center'}),
    dash_table.DataTable(
        data=file_dataframes[0].to_dict('records'),
        columns=[{'id': c, 'name': c} for c in file_dataframes[0].columns],
        style_cell={
            'textAlign': 'center',
            'padding': '10px'
        },
        style_cell_conditional=[
            {
                'if': {'column_id': 'Grade'},
                'textAlign': 'left'
            }
        ],

        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': 'rgb(220, 220, 220)'
            }
        ],

        style_as_list_view=True,
        style_header={
            'backgroundColor': 'rgb(210, 210, 210)',
            'color': 'black',
            'fontWeight': 'bold'
        },
    )
])

if __name__ == '__main__':
    app.run_server(debug=True)

# print(f"Finished in {round((time.time()/60-start_time), ndigits=2)} minutes")

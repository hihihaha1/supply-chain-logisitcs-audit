import pandas as pd


path = 'Supply chain logisitcs problem.xlsx'

orders = pd.read_excel(path, sheet_name='OrderList')
rates = pd.read_excel(path, sheet_name='FreightRates')
plants = pd.read_excel(path, sheet_name='PlantPorts')


def progress_logistics():
    orders = pd.read_excel(path, sheet_name='OrderList')
    rates = pd.read_excel(path, sheet_name='FreightRates')

    rates = rates.rename(columns={
        'orig_port_cd' : 'Origin Port',
        'dest_port_cd' : 'Destination Port',
        'svc_cd' : 'Service Level'
    })

    df = pd.merge(orders, rates, on=['Origin Port', 'Destination Port', 'Carrier', 'Service Level'], how='left')

    df['is_ok'] = ((df['Weight'] >= df['minm_wgh_qty']) & (df['Weight'] <= df['max_wgh_qty']))

    df = df.sort_values(by=['Order ID', 'is_ok'], ascending=[True, False])

    df['max_rate'] = df.groupby('Order ID')['rate'].transform('max')

    df = df.drop_duplicates(subset=['Order ID'], keep='first')


    df['Status'] = 'OK'
    mask_uw = (df['rate'].notnull()) & (df['Weight'] < df['minm_wgh_qty'])
    df.loc[mask_uw, 'Status'] = 'Underweight'

    mask_ow_stat = (df['rate'].notnull()) & (df['Weight'] > df['max_wgh_qty'])
    df.loc[mask_ow_stat, 'Status'] = 'Overweight'

    mask_rate = df['rate'].isnull()
    df.loc[mask_rate, 'Status'] = 'No Rate'

    plant_merge = pd.merge(df, plants, on=['Plant Code'], how='left')

    plant_merge['Port_Match'] = plant_merge['Origin Port'] == plant_merge['Port']

    mismatch = plant_merge[plant_merge['Port_Match'] == False]
    if len(mismatch) > 0:
        print(mismatch[['Order ID', 'Plant Code', 'Origin Port', 'Port', 'Status']].head())


    df['Estimated_Cost'] = 0.0
    mask_ok = df['Status'] == 'OK'
    df.loc[mask_ok, 'Estimated_Cost'] = df['Weight'] * df['rate']

    mask_ow = df['Status'] == 'Overweight'
    df.loc[mask_ow, 'Estimated_Cost'] = df['max_rate'] * df['Weight'] * 1.2

    mask_nr = df['Status'] == 'No Rate'
    df.loc[mask_nr, 'Estimated_Cost'] = 0

    df = df[(df['minm_wgh_qty'] <= df['Weight']) & (df['max_wgh_qty'] >= df['Weight'])]

    return df

if __name__ == "__main__":

    final_df = progress_logistics()

    final_report = final_df[[
    'Order ID',
    'Order Date',
    'Plant Code',
    'Origin Port',
    'Destination Port',
    'Carrier',
    'Service Level',
    'Weight',
    'rate',
    'Status',
    'Estimated_Cost'
    ]].copy()

    final_report['Estimated_Cost'] = final_report['Estimated_Cost'].round(2)

    final_report.to_excel('Logistics_Final_Report.xlsx', index=False)
    print('Все сохранено в файл: "Logistics_Final_Report.xlsx"')
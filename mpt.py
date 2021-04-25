import pandas as pd
import numpy as np
import scipy.stats as stats
import investpy as ipy
import xlsxwriter
from datetime import datetime,timedelta

#https://www.machinelearningplus.com/machine-learning/portfolio-optimization-python-example/#optimal-risky-portfolio
# https://pypi.org/project/investpy/

string_tickers = input('Enter Tickers seperated by space only:\n')

country = input('Enter country of stocks:\n')

frequency = input('Enter Frequency of Data:\n')

period = input('Enter Number of Years:\n')

which_column = [str(input('Which Data (Select Number):\n1. Open\n2. High\n3. Low\n4. Close\n5. Adj Close\n6. Volume\n'))]

omit_last_line = [str(input('Omit last line? (Y/N)\n'))]

given_weights = [str(input('Weights? (Y/N)\n'))]

list_tickers = string_tickers.split(' ')

rf = str(input('What is the risk free rate? (10% --> .1)\n'))

if 'y' in given_weights:
    weights_string = input('Enter weights seperated by space only:\n').split(' ')
    weights = []
    
    for ticker in weights_string:
        weights.append(int(ticker))
    weights_df = pd.DataFrame([weights], columns=list_tickers)   

elif 'n' in given_weights:
    weights = []
    for tickers in list_tickers:
        weights.append(1/len(list_tickers))
    weights_df = pd.DataFrame([weights], columns=list_tickers)

today = str(datetime.now())[8:10]+'/'+str(datetime.now())[5:7]+'/'+str(datetime.now())[0:4]
start = str(datetime.today() - timedelta(days=(int(period)*365)))
start_date = start[8:10]+'/'+start[5:7]+'/'+start[0:4]
frequency_num = 0

if 'monthly' or 'Monthly' in frequency:
    frequency_num += 12
elif 'weekly' or 'Weekly' in frequency:
    frequency_num += 52
elif 'daily' or 'Daily' in frequency:
    frequency_num += 365
else:
    frequency_num += 1

column_chosen = []

if '1' in which_column:
    column_chosen.append('Open')
elif '2' in which_column:
    column_chosen.append('High')
elif '3' in which_column:
    column_chosen.append('Low')
elif '4' in which_column:
    column_chosen.append('Close')
elif '5' in which_column:
    column_chosen.append('Adj Close')
elif '6' in which_column:
    column_chosen.append('Volume')



first_data = ipy.stocks.get_stock_historical_data(list_tickers[0],
                                                  country=country, from_date=start_date , to_date=today, as_json=False,
                                                  order='ascending', interval=frequency)
compiled_data = first_data[column_chosen[0]]

i=1
while i < len(list_tickers):
    to_add = ipy.stocks.get_stock_historical_data(list_tickers[i], country=country, from_date=start_date , to_date=today, as_json=False,
                                                  order='ascending', interval=frequency)
    compiled_data = pd.merge(compiled_data, to_add[column_chosen[0]], how='outer', left_index=True, right_index=True)
    i+=1
    
compiled_data.columns = list_tickers

if 'y' in omit_last_line:
    compiled_data = compiled_data.drop(compiled_data.index[-1])
    
percent_change_data = compiled_data.pct_change()

expected_return = {}
std_dev = {}
variance = {}
    
for column in percent_change_data:
    expected_return[column] = round(percent_change_data[column].mean(), 5)
    std_dev[column] = round(percent_change_data[column].std(), 5)
    variance[column] = round(std_dev[column]**2, 5)

annual_expected_return = {}
annual_std_dev = {}
annual_variance = {}

for column in percent_change_data:
    annual_expected_return[column] = round(expected_return[column]*frequency_num, 5)
    annual_std_dev[column] = round(std_dev[column]*(frequency_num**.5), 5)
    annual_variance[column] = round(annual_std_dev[column]**2, 5)

labels = ['Expected Return'+' '+frequency, 'Std Dev'+' '+frequency, 'Variance'+' '+frequency,
          'Expected Return Annual', 'Std Dev Annual', 'Variance Annual']

historicals = pd.DataFrame([expected_return, std_dev, variance, annual_expected_return,
                    annual_std_dev, annual_variance], index=labels).T


ind_er = (compiled_data.pct_change().mean())*frequency_num #annual returns
ann_sd = (compiled_data.pct_change().std())*np.sqrt(frequency_num) # annual sd
ann_var = ann_sd**2
port_er = (weights*ind_er).sum()
assets = pd.concat([ind_er, ann_sd], axis = 1)
assets.columns = ['Returns', 'Volatility']
p_ret = []
p_vol = []
p_weights = []
num_assets = len(list_tickers)
num_port = 15000


###########################    

def var_covar_frequency():
    
    var_covar = percent_change_data.cov()

    return var_covar

def var_covar_annual():
    
    return var_covar_frequency()*frequency_num

def var_port_annual():
    var = var_covar_annual()
    p1 = var.dot(weights_df.T)
    varp = weights_df.dot(p1)
    return varp.loc[0,0]

def port_annual_return(): #call on actual value?
    returns_list = list(annual_expected_return.values())
    returns_df = pd.DataFrame([returns_list], columns=list_tickers)
    annual_returns = returns_df.dot(weights_df.T)
    return annual_returns.loc[0,0]

def var_port_frequency(): #make std as well
    var = var_covar_frequency()
    p1 = var.dot(weights_df.T)
    varp = weights_df.dot(p1)
    return varp.loc[0,0]

def port_frequency_return():
    returns_list = list(expected_return.values())
    returns_df = pd.DataFrame([returns_list], columns=list_tickers)
    returns = returns_df.dot(weights_df.T)
    return returns.loc[0,0]

def weights_mvp():
    
    for portfolio in range(num_port):
        weights = np.random.random(num_assets)
        weights = weights/np.sum(weights)
        p_weights.append(weights)
        returns = np.dot(weights, ind_er)

        p_ret.append(returns)
        var = var_covar_annual().mul(weights, axis=0).mul(weights, axis=1).sum().sum() # port var
        ann_sd = np.sqrt(var)
        p_vol.append(ann_sd)

    data = {'Returns':p_ret, 'Volatility':p_vol}

    for counter, symbol in enumerate(list_tickers):
        data[symbol] = [w[counter] for w in p_weights]

    portfolios = pd.DataFrame(data)

    return portfolios.iloc[portfolios['Volatility'].idxmin()]

def market_portfolio():
    for portfolio in range(num_port):
        weights = np.random.random(num_assets)
        weights = weights/np.sum(weights)
        p_weights.append(weights)
        returns = np.dot(weights, ind_er)

        p_ret.append(returns)
        var = var_covar_annual().mul(weights, axis=0).mul(weights, axis=1).sum().sum() # port var
        ann_sd = np.sqrt(var)
        p_vol.append(ann_sd)

    data = {'Returns':p_ret, 'Volatility':p_vol}

    for counter, symbol in enumerate(list_tickers):
        data[symbol] = [w[counter] for w in p_weights]

    portfolios = pd.DataFrame(data)

    return portfolios.iloc[((portfolios['Returns'] - int(float(rf)))/portfolios['Volatility']).idxmax()]

    
def excel():
    writer = pd.ExcelWriter('portfolio data.xlsx', engine='xlsxwriter')
    compiled_data.to_excel(writer, sheet_name='stock data')
    percent_change_data.to_excel(writer, sheet_name='pcnt change data')
    historicals.to_excel(writer, sheet_name='historical data')
    var_covar_annual().to_excel(writer, sheet_name='var covar')
    assets.to_excel(writer, sheet_name='assets')
    weights_mvp().to_excel(writer, sheet_name='mvp')
    market_portfolio().to_excel(writer, sheet_name='market_portfolio')
    writer.save()









        
        

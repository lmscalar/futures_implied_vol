# -*- coding: utf-8 -*-
"""
ethanol.py
@author: lmolina
2021-06-01

"""

import pandas as pd
import numpy as np
import os
import statsmodels.tsa.stattools as ts
import statsmodels.api as sm
from statsmodels.sandbox.regression.predstd import wls_prediction_std
from numpy.linalg import svd
from numpy import sqrt
from scipy import diag
from scipy.stats import norm
import xlwings as xw
from xlwings import Workbook, Range, Sheet, sub

import matplotlib.pyplot as plt
import seaborn as sb
import datetime as dt
sb.set()

specs = {'corn': 5000, 'corn_bbg':50, 'ethanol': 42000}

yr_basis = 252.


fPath = 'C:/Risk/'
file1 = 'corn_ethanol.xlsm'

def spectral_decomp( tradeDate, C, rtns, I=10000):
    '''Singular value decomposition, C= covariance or correaltion matrix.'''
    hold_prd = 1
     
    C = C.loc[tradeDate,:,:]     # covariance panel
    
    [ S, V, D ] = svd( C )    # singular value decomposiiton
    A = np.dot( S, np.sqrt( diag( V )))    
    z = np.random.randn( I, len(C) )   # random normals to be correlated throug spectral decomp
    dr = hold_prd*(rtns.iloc[-1,:])/yr_basis
    
    # correlated random normals
    X = np.dot(z,A.T)
    X = pd.DataFrame(X,columns=C.columns)
    
    dQ = X + dr
                                        
    return dQ
    
def spectral_decomp_VaR( tradeDate, pos_df, vcov, rtns, I, CI, hold_prd ):
    '''Singular value decomposition, C= covariance or correaltion matrix.'''
    
    pos_weights = pos_df['pos_weights']
    C = vcov.loc[tradeDate,:,:] 
    [ S, V, D ] = svd( C )    
    A = S @ sqrt(diag(V)) 
    z = np.random.randn( I, len( C) )   # random normals to be correlated throug spectral decomp
    
    dr = hold_prd*( rtns.loc[tradeDate,:]) / yr_basis   # current returns
    
    # correlated random normals
    X = np.dot(z,A.T)
    X = pd.DataFrame( X, columns = C.columns )
    
    dQ = X + dr
    
    sim_rtns = np.dot( dQ, pos_weights.T)
    # monte_carlo VaR 
    
    abs_pos_value = pos_df['abs_pos_value'].sum()
        
    CI = CI*100. 
    # calculate VaR from simulated returns
    VaR = abs_pos_value *np.percentile( sim_rtns, CI )*np.sqrt( hold_prd )
    
    #print("Position value: %s" % abs_pos_value)
    #print('VaR in basis pts: %0.5f' % np.percentile( sim_rtns, CI ))
    #print("VaR: ", VaR )
                                                  
    return VaR 
    
def cointegration(Y,X, sym1, sym2 ):
    """Stationarity test for X and Y."""
    
    wk = Workbook.active()
    
    dates = [ dt.datetime.strptime(d,'%Y-%m-%d') for d in X.index ]
    X.index = dates
    Y.index = dates
    
    X = sm.add_constant(X)    
    model = sm.OLS(Y,X)
    results = model.fit()
        
    print( results.summary())
    print('Parameters: ', results.params)
    print('Standard errors: ', results.bse)
    #print('Predicted values: ', results.predict())
    print('P Value_F: ', results.f_pvalue)
    print('P Value_T: ', results.pvalues)
    print('RSQ: ', results.rsquared)
    print('RSQ_Adj: ', results.rsquared_adj)
    
    z = results.resid
    stdev = np.std(z)
    up_std = 1.5*stdev
    down_std = -1.5*stdev
    
    mn = z.mean()
    '''
    fig = plt.figure()
    plt.title('%s vs. %s' % (sym1, sym2))
    plt.plot(X.index, z )
    plt.axhline(up_std, color='r')
    plt.axhline(down_std, color='g')
    plt.axhline(y= mn, color='g')
    plt.show()
    '''
    
    sym = sym1 + ' vs. ' + sym2
    
    stat_arb = pd.DataFrame( z, index=X.index, columns=[sym])
    stat_arb['up_std'] = up_std
    stat_arb['down_std'] = down_std
    stat_arb['mean'] = mn
    
    Range('parameters','graph_data').clear_contents()
    Range('parameters','stat').value = stat_arb.tail(500)
    

    #------------------------------------------
    #http://statsmodels.sourceforge.net/devel/generated/statsmodels.tsa.stattools.adfuller.html
    ##http://www.quantstart.com/articles/Basics-of-Statistical-Mean-Reversion-Testing
       
    x = z
    result = ts.adfuller(x, 1) # maxlag is now set to 1
    print(result)
    
    
def position_weights( pos_df, tradeDate ):
    """Calculate positions weights, w."""
    
    # calculate weights for the VaR calculation
    position_value = pos_df[tradeDate]*pos_df['positions']*pos_df['contractSize']
    abs_position_value = position_value.abs()
    position_weights = position_value/abs_position_value.sum()   #position weights
    
    pos_df['pos_value'] = position_value
    pos_df['abs_pos_value'] = abs_position_value    
    pos_df['pos_weights'] = position_weights
    
    return pos_df    
    
    
# VaR calculation: position values, covariance matrix, conf_int, hold_prd , sqrt(w'Vw)*a   
def calc_VaR( pos_df, vcov, conf_int, hold_prd, tradeDate ):
    
    C = vcov.loc[tradeDate,:,:]
    # theta are the position weights
    theta = pos_df['pos_weights']
    pos_value = pos_df['abs_pos_value'].sum()
    
    # print(len(theta))
    dn = np.dot(theta.T,C)
    VaR = sqrt( np.dot(dn , theta ))*norm.ppf( conf_int )*sqrt( hold_prd ) * pos_value
    
    return VaR


def read_data(interval=90):
    
    prices = pd.read_excel(r'C:\Risk\corn_ethanol.xlsm','data',header=0, index_col=0)
    pos = pd.read_excel(r'C:\Risk\corn_ethanol.xlsm','positions',header=0, index_col=0)    
    rtns = np.log(prices/prices.shift(1))
    rtns = rtns.dropna()
    dates = [ d.strftime('%Y-%m-%d') for d in prices.index ]
    prices.index = dates

    
    abs_returns = prices/prices.iloc[0,:]
    abs_returns = abs_returns.dropna()
    
    vol = rtns.rolling(window=interval,center=False).std()*sqrt(252)
    vol = vol.dropna()
    
    covar = rtns.rolling(window=interval).cov().dropna()
    corr = rtns.rolling(window=interval).corr().dropna()    
     
    dates = [ d.strftime('%Y-%m-%d') for d in covar.items] 
     
    return prices, dates, covar, corr, pos, rtns, abs_returns, vol
    
    
def position_prices(pos, tradeDate, prices):  
    
    pos = pos.join(prices.loc[tradeDate]).fillna(0)
    
    return pos
    
    
def main():

    df, dates, covar, corr, pos, rtns, returns, vol = read_data()

    return df, dates, covar, corr, pos, rtns, returns

def plot_graphs(interval):
    
    prices, dates, covar, corr, pos_, rtns, abs_returns, vols = read_data(interval)
    
    # graph results
    corn_vol = vols[['C 1','C 2','C 3','C 4','C 5']]
    ethanol_vol = vols[['CUA1','CUA2','CUA3','CUA4','CUA5']]
   
    plt.title("Monte Carlo and Parametric VaR")
   
    plt.figure()
    plt.plot(corn_vol,lw=2)
    plt.title("Corn 90-day Historical Vol")
    plt.legend(corn_vol.columns)
    plt.show()
    
    plt.figure()
    plt.plot( ethanol_vol, lw=2 )
    plt.title("Ethanol 90-day Historical Vol")
    plt.legend(ethanol_vol.columns)
    plt.show()
    
    plt.figure()
    plt.plot(corr.loc[:,'C 1','CUA1'],lw=2)
    plt.plot(corr.loc[:,'C 2','CUA2'],lw=2)
    plt.plot(corr.loc[:,'C 3','CUA3'],lw=2)
    plt.legend(['C 1/CUA1','C 2/CAU2','C 3/CUA3'])
    plt.title("Correlation Corn vs. Ethanol")
    plt.show()
    
    # plot corn and ethanol prices
    prices_df = prices.copy()
    prices_df.index = [ dt.datetime.strptime(d,'%Y-%m-%d') for d in prices.index ]
    corn_prices =  prices_df[['C 1','C 2','C 3','C 4','C 5']]
    ethanol_prices = prices_df[['CUA1','CUA2','CUA3','CUA4','CUA5']]
    plt.figure()
    plt.plot(corn_prices,lw=1.5)
    plt.title("Corn Prices")
    plt.legend(['C 1','C 2','C 3','C 4','C 5'])
    plt.show()
    plt.figure()
    plt.plot(ethanol_prices,lw=1.5)
    plt.title("Ethanol prices")
    plt.legend(['CUA1','CUA2','CUA3','CUA4','CUA5'])
    plt.show()

@sub
def stat_arb():
    
    I = 10000
    CI = .95
    hold_prd =1 
    
    wk = Workbook.active()
    
    Range('parameters','run').value = "Running stat Arb..."
    
    interval = Range('parameters','var_window').value
    interval = int(interval)
       
    prices, dates, covar, corr, pos_, rtns, abs_returns, vols = read_data(interval)   
   
    var_df = []
    mc_var = []
    for tradeDate in  dates:  
        pos = position_prices(pos_, tradeDate, prices)
        pos = position_weights(pos, tradeDate)
        Q = spectral_decomp_VaR( tradeDate, pos, covar, rtns, I, CI, hold_prd )
        mc_var.append(Q)
        para = calc_VaR( pos, covar, CI, hold_prd, tradeDate )
        var_df.append(para) 
        
    df = pd.DataFrame({"MC_VaR":mc_var,"Para_VaR": var_df}, index=dates)
    
    
    Range('vols','A1').value = vols
    Range('VaR','A1').value = df
    Range('VaR','E2').value = pos
    Range('Covariance','A1').value = covar.loc[tradeDate,:,:]
    Range('Correlations','A1').value = corr.loc[tradeDate,:,:]
    
    # plot graphs for data
    #plot_graphs(interval)   
    
    sym1 = Range('parameters','product1').value 
    sym2 = Range('parameters','product2').value 
    
    # cointegration of corn vs. ethanol
    #sym1 = 'CUA2'    # " sym1 Sell high, buy low sym2
    #sym2 = 'SB1' 
    
    #y = np.log(prices.loc['2014-12-01':, sym1 ]/prices.loc['2014-12-01', sym1 ])   # corn
    #x = np.log(prices.loc['2014-12-01':, sym2 ]/prices.loc['2014-12-01', sym2 ])   # ethanol
    y = np.log(prices.loc['2014-12-01':, sym1 ]*pos.loc[sym1,'contractSize'])   # corn
    x = np.log(prices.loc['2014-12-01':, sym2 ]*pos.loc[sym2,'contractSize'])   # ethanol
    
    #y = np.log(prices.loc['2014-12-01':, sym1 ])   # corn
    #x = np.log(prices.loc['2014-12-01':, sym2 ])   # ethanol
               
    cointegration(y,x, sym1, sym2)   # run cointegration on spred
    
    '''
    spread = y - x 
    plt.figure()
    spread.plot(lw=2)
    plt.title('%s vs. %s' % (sym1, sym2))
    plt.show()
    '''    
    
    Range('parameters','run').value = "Run Completed!"
    wk.save()
    
    
    
if __name__ == '__main__':
    
    stat_arb()
    
    #wk.close()
    #os.system('taskkill /f /im EXCEL.EXE')
    
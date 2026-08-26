# Gold / CHF Quantitative Analysis – Excel & VBA

## Overview

This project was carried out as part of an **Excel VBA project at ESILV**.

The objective was to analyse the relationship between **Gold and the Swiss Franc (CHF)**, two assets often considered as safe-haven assets during periods of economic and geopolitical uncertainty.

The analysis was performed using **Excel and VBA**, with historical data from 2020 to 2025.

## Data

The data comes from **Yahoo Finance**:

* Gold Futures: `GC=F`
* USD/CHF: `USDCHF=X`

The USD/CHF exchange rate was converted into CHF/USD so that the two assets could be compared consistently.

## Analysis

The project includes several steps:

### 1. Data Preparation

The VBA tool cleans and prepares the raw financial data:

* Cleaning and formatting the data
* Synchronising the Gold and CHF series
* Converting USD/CHF into CHF/USD
* Preparing the data for the analysis

### 2. Return Calculation

Logarithmic returns are calculated from the cleaned price series.

This allows us to analyse the relative changes in the assets rather than their raw prices.

### 3. Statistical Analysis

For both Gold and CHF, we calculate:

* Mean return
* Volatility
* Minimum return
* Maximum return

We also analyse the relationship between the two assets using:

* Correlation
* Linear regression
* Beta coefficient
* R²

### 4. Excel & VBA

The analysis was implemented in two ways:

* **Excel formulas** for the statistical calculations
* **VBA macros** to automate the different steps of the analysis

The workbook includes a main interface with buttons to prepare the data, calculate returns, generate results and create graphs.

## Results

The results show a **weak positive relationship** between Gold and CHF returns, with a correlation of around **0.07**.

Gold is significantly more volatile than the Swiss Franc, while the average returns of both assets are close to zero.

The regression also shows a small positive beta and a very low R², suggesting that CHF returns have limited explanatory power for Gold returns over the period studied.

## Conclusion

This project allowed us to use **Excel and VBA for financial data analysis** and to study the relationship between two safe-haven assets.

The results suggest that Gold and CHF have some similarities as defensive assets, but their short-term returns are only weakly related.

The analysis could be extended by using a longer period, other safe-haven currencies such as JPY, or more advanced and non-linear models.


## Tools

* Microsoft Excel
* VBA
* Yahoo Finance
* Statistical analysis
* Linear regression

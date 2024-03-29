## excel-duckdb

Excel-duckdb joins the most powerful analytics platform on the planet with duckdb. Excel-duckdb is a poorly written vba-add-in that can be added to any excel file.

Execute sql queries using duckdb against local or remote files, databases, or object storage all from excel! 

Jupyter notebooks no more with excel X duckdb!

## Installation

Start by installing duckdb-odbc and create a new 64-bit ODBC dsn for DuckDB.


1. Install: [64-bit ODBC](https://duckdb.org/docs/installation/?version=stable&environment=odbc&platform=win)
1. Create a new ODBC Data Source Name (DSN): [Msft ODBC](https://support.microsoft.com/en-us/office/administer-odbc-data-sources-b19f856b-5b9b-48c9-8b93-07484bfab5a7)
    - System DSN >> Add >> DuckDB Driver

1. Create a DSN called: `excel-duckdb`



Once the setup is complete, you can include the add-in in any excel workbook by searching for `Add-ins` in the global workbook search, or navigating to the `Developer` tab >> `Excel Add-ins`. 

From here, click browse and add the `excel_duckdb.xlam` file as an add-in to raw dog some sql straight into the formula bar for `q4 Financials - v2 Final - final - copy.xlsx` that you spent all night working on!

Additionally, you will most likely have to add the location of the `excel_duckdb.xlam` file as a trusted location for the macro to work. This
can be accomplished by: 

1. File >> Options >> Trust Center >> Trust Center Settings
1. Choose the Trusted Locations tab
1. Add new location 
1. Browse to the folder the `excel_duckdb.xlam` file is located 

You may need to modify other permission, or settings to enable macros. Mileage will vary depending on your setup. 


## Usage

Query parquet files in object storage, local files on your machine, or even [MotherDuck](https://motherduck.com/)!

TODO: add credential management that is not storing credentials in plain text in VBA. 

Users will be limited to 255 characters per query.

```
=duckdb("
    select *
    from 'https://d37ci6vzurychx.cloudfront.net/trip-data/green_tripdata_2022-02.parquet'
    where passenger_count > 1 and lpep_pickup_datetime::date = '2022-02-15'
    order by 2 desc
")
```

The response is a dynamic array which can then be shared, or further analyzed with the full power and might of, microsoft excel. 


## Issues 

No testing has been done. 

If you run into issues, try toggling off the `Protechted View` check boxes in the trust center, and 
ensure that `Microsoft ActiveX Data Objects 6.1 Lirbary` has been toggled as a reference from the developer tab. 

- `Developer` Ribbon >> Visual Basic (`alt + F11`)
- Tools >> References >> `Microsoft ActiveX Data Objects 6.1 Lirbary`

## Road Map

- Caching 
- Credential management for object storage 

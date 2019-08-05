# README

This's excel script. Copy columns (which you choose) from xlsx file to csv. :computer:

# Setup

Make sure you have Ruby installed.

♦️ Ruby version: 2.5.3

Change line 7 (write yor path to example.xlsx):
```
workbook = Roo::Spreadsheet.open '/your_path/script_excel/example.xlsx'
```

Run in your terminal

```
git clone git@github.com:Maddalenap/script_excel.git

cd script_excel

bundle install

rails db:migrate
```
And run with
```
ruby /your_path/script_excel/example.xlsx
```

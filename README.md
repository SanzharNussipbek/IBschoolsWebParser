# README

# IB Schools Web Scraper

This program is a web scraper of IB schools with Diploma Programme and English Language of instruction. The data is parsed from official [IBO website](https://www.ibo.org). 

The goal is to collect all available information on IB schools, e.g. school name, country name, coordinator(salutation, name, title, phone), school address, website url, school page url, programmes.

There are three different regions to be covered: Americas, Asia Pacific and Africa, Europe and Middle East.

Even though this program collects data for DP programme schools with English language of instruction, it can be easily modified by changing the url which is to be parsed.

## How to use it?

- Run the program and wait for the Excel file to appear
- In order to modify the program and change the specifications like programme and language of instruction, just search the same thing in the official website, copy the url and change the url in the dictionary of urls in the main function.

    However, the country_list also needs to be changed manually: inputting the country code from the url, number of pages (one page has 20 schools) and country name.

This project was made out of laziness and desire to automate the routine of collecting the data manually for a job. It is indeed needs to be improved, and hopefully, with more knowledge in the future, it will be able to do everything by itself.
---
layout: post
title:  "Excel Parsing"
date:   2019-07-03
excerpt: "Excel Parsing with POI Library with Java"
tag:
- POI
- Excel
comments: true
---

# JAVA_Excel_Parsing
POI Library by Apache

## Before get starting...
I had a Excel file that has to be upload to my database. 
Each cells were mixed by  imformations, blanks,
 and formulas. And I had to parse ExcelFile in forms
  like "1","apple","red","","Round"...etc. So first, 
  I decided to use Java and I found POI Library.

## What Library?
There were some Libraries to use when 
we handle Excel files with Java. When I 
surf in Internet most people used POI Library. 
Because there are versions of Excel we could use. 
Most of people use higher version Excel, and I have 2007 version so I used POI Library. 

### POI Library
- Before we use library
    Check this out if you have problem with using xssfWorkbook.\
    [Can Apache be compiled / used Java 10 or newer?](https://poi.apache.org/help/faq.html#faq-java10)\
    Because I had problem using xssfworkbook in new Mac book. I installed java 12 version in mac. 
    So my code worked in windows notebook, but it had problem in compiling in Mac.
    If you think there are no problems in your code and adding jar files, and have compile problem like this
    ![스크린샷 2019-07-02 오후 10 00 38](https://user-images.githubusercontent.com/32008149/60514655-e5186a00-9d14-11e9-9f5a-eab1df34fae1.png)
    
    Try use maven project. You could easily copy and paste dependency in there.

- XSSF / HSSF / SS
  
  Name | Feature 
  ----- | ------      
  HSSF | Excel 97 ~ 2003
  XSSF | Excel 2007 ~
  SS   | XSSF Straming version (Low memory and fits to mass data)  
   
- Max Data

    Excel 2003 | Excel 2007 
    ----- | -----
    265 Column | 16,384 Column
    65,536 Line | 1,048,576 Line

- Download here : https://poi.apache.org \
Window OS : Download Zip\
Linux or Unix OS : Download tar

    ![스크린샷 2019-07-02 오후 11 16 52](https://user-images.githubusercontent.com/32008149/60519967-886e7c80-9d1f-11e9-8bf6-d4a4221b5afe.png)
- We have to Add Blue Rectangular files in BuildPath(Libraries)\
if you are going to use xlsx file you should add ooxml-lib directory files.
    ![스크린샷 2019-07-02 오후 11 13 54](https://user-images.githubusercontent.com/32008149/60519806-3cbbd300-9d1f-11e9-960a-b9e905540000.png)

- In Maven, you can add dependency to POM.xml (Beacareful for the Version)\
    [[Download Link]](https://mvnrepository.com/artifact/org.apache.poi/poi/3.17)
    ![POI3](https://user-images.githubusercontent.com/32008149/60109311-1937dc00-97a5-11e9-8ef5-db98598edaad.PNG)
    if you want to handle xslsx file, you have to put ooxml dependency to pom.xml\
    [[Download Link]](https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/3.17)
    ![스크린샷 2019-07-02 오후 11 34 02](https://user-images.githubusercontent.com/32008149/60521268-ed2ad680-9d21-11e9-96e0-d499a487d272.png)
    
- Make new Maven Project\
    ![스크린샷 2019-07-02 오후 10 09 22](https://user-images.githubusercontent.com/32008149/60515203-180f2d80-9d16-11e9-881e-960d3d0c3fe4.png)

- Put dependency in to pom.xml and install maven.(Becareful for the version, try to use one that is mostly used.) 
    ![스크린샷 2019-07-02 오후 10 46 19](https://user-images.githubusercontent.com/32008149/60517857-5b1fcf80-9d1b-11e9-93f2-946ca674d862.png)

- Check Server is alright!\
    ![스크린샷 2019-07-02 오후 10 50 31](https://user-images.githubusercontent.com/32008149/60518118-d6818100-9d1b-11e9-88bc-716f5217d214.png)

- Put code in jsp file.\
    ![스크린샷 2019-07-03 오전 10 27 44](https://user-images.githubusercontent.com/32008149/60556660-57249980-9d7d-11e9-9f95-0fc77ab8e28d.png)

- Result!\
    ![스크린샷 2019-07-03 오전 10 30 05](https://user-images.githubusercontent.com/32008149/60556762-aa96e780-9d7d-11e9-9578-7c722fc5a40a.png)

    
Then I had to import csv file to my DataBase.\
But there were serious Problems in importing.

## CSV
I had serious problem with importing csv file to Database.
There were many ways to import Data to DataBase, but I need
 to import csv File because I need a function that has to be 
 imported by button in Web. So First, I thought csv is same with Excel file.
So I parse xlsx file and made new xlsx file in it's own form. Then I just change extension of file, 
xlsx to csv. But then, problem occur.
Database shows how it will works if I import csv files. But, it seems something wrong.
So I tried to change Column delimiter, Quote char, and Quote char. 
But that doesn't work at all. I tried to think what makes problem. 
Then I thought what is CSV?

### What is CSV?
CSV (Comma separated version) is file that has been separated 
by comma. It is not same as xls,xlsx. First, I thought it's same 
as xls or xlsx. But that is wrong. It is not Excel file. Every files have meta 
data in it. In program I made xlsx file. So it's xlsx file, even 
if I changed extension to csv. And what is CSV? It is text file. 
So I changed some code in program. And I made csv file in program. 
Then importing works! 

### Import CSV to Database(MariaDB)
I used DBeaver tool to handle MariaDB. And There were some 
problems to think about. 
- First, I need to change "o", "x", " " text to "Y", or "N".  
So I found index that has to be converted.(Code can be weired about "A,B,C .." values can be "Y", But Data that I 
received does not need to think about that issues.)
  ~~~
  if(columnindex==11 || columnindex==15){                   		 
         if(value=="x" || value.isEmpty()){
              buff.append("\"N\","); 
         }else{
              buff.append("\"Y\",");                 			
         } 
  ~~~
- Second, there are some formulas, in data. When I parse data from file, it brings formula such as "x1 + y1".
What I want was result but, it brings me formula. So I changed value to get Numericvalue. 
    ```
    switch (cell.getCellType()){  
    
         case XSSFCell.CELL_TYPE_FORMULA:                        
    
    	 value=cell.getFormulaCellValue()+"";
    	// we should use cell.getNumericCellValue!
    
    	break;
    		                        
    }
   ``` 
- Third, if Formula cell's value has problem that makes "#value", 
         problem occurs. So I decided to choose NumericValue rather than to 
         show both String Value and NumericValue by "if else" syntax. Because what is formula? It has to make
         NumericValue. I can make them to show "-1" if there are problem in Formula, but 
         
- Lastly, date has to be converted. First I didn't check the date, but when I tried to import to Database, 
I saw weired number.Then I realized I should convert to date format. 
    ```
    case XSSFCell.CELL_TYPE_NUMERIC:
                        	if(columnindex==14 || columnindex==27){
                        		SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
                        		value = format.format(cell.getDateCellValue());
                        		break;
                        	}
                            value=cell.getNumericCellValue()+"";
                            break;

    ```
DB Import was successful, but I need to think about some issues.


# job Postings Analysis

## Introduction

Using a dataset of data science job postings from 2023, collected and compiled by Luke Barousse from different job posting platforms, I created the interactive salary dashboard shown below.


## Data Science Salary Dashboard
![job salary dashboard](images/salary_dashboard_demo.gif)  
This dashboard consists of two main components:

* **Charts**, which make it easy to observe and compare the median salaries of data science jobs based on job title, country, and job schedule type.
* **Cards**, which display the median salary, the top job posting platform, and the total number of job postings based on the selected filters.

## Excel Skills Used
### Data Validation:
A dedicated Data Validation worksheet was created to provide filtered lists of valid values for the Job Title, Country, and Job Schedule Type fields. These lists are used as data validation rules, ensuring that users can select only predefined values. This helps prevent invalid input and unexpected errors.
### Formulas and Functions:
- **Median**: The median salary is calculated based on the values selected in the option cells using the following formula:

    ```excel
    = MEDIAN(
    IF(
    (data_job_posts[job_title_short]=A2)*
    (data_job_posts[job_country]=country)*
    ISNUMBER(SEARCH(type,data_job_posts[job_schedule_type]))*
    (data_job_posts[salary_year_avg]<>0),
    data_job_posts[salary_year_avg]
    ))
    ```

- **Count:** The number of job postings that match the selected options is calculated using the following formula:

    ```excel
    = COUNT(
    IF(
    (data_job_posts[job_title_short]=Data_Validation!$A2)*
    (data_job_posts[job_country]=country)*
    ISNUMBER(SEARCH(type,data_job_posts[job_schedule_type])), 1
    ))
    ```
    This count is used to create the Job Count card. It is also used to sort the values in the Job Title data validation list in descending order based on the number of matching job postings.

- **XLOOKUP:** To retrieve the values displayed on the dashboard cards, the XLOOKUP function is used. For example:

    ```excel
    =IF(ISNUMBER($I$2),XLOOKUP(title,D2:D11,E2:E11),"No Data Found")
    ```
    The **IF** function prevents errors from being displayed on the dashboard. If no matching data is found for the selected options, the formula returns "No Data Found" instead of an error value.

- **FILTER:** The FILTER function is used to generate the list of valid Job Schedule Type values in the Data Validation worksheet. It filter out schedule types that contain more than one type by eliminating those that contain "and", as all multi-type job schedule specifications include the word "and".

    ```excel
    =FILTER(M2#,NOT(ISNUMBER(SEARCH("and",M2#)))*(M2#<>0))
    ```

- **SORT:** The SORT function is used to arrange data in descending order, making bar charts easier to interpret, ranking job titles by popularity, and retrieving desired values much quicker.

    ```excel
    =SORT(A2:B11,2,-1)
    ```
    This formula sorts the job titles by their job counts in descending order.

- **SUBSTITUTE:** It used to clean the **Top Job Platform** value by removing the prefix **"via "** from the platform name.
    
    ```excel
    =IF(E2<>0,SUBSTITUTE($D$2,"via ",""),"No Platform Found")
    ```
    If no matching platform is found, the formula returns "No Platform Found" instead of displaying invalid value.


### Charts:
- **Bar Charts:**   
    <br>
    ![job salary dashboard](images/title_median_table.png)  

    The median salary for each job title is calculated based on the selected option values and then sorted in ascending order to create the following bar chart:  
    <br> 
    ![bar chart of median salaries by job title](images/title_bar_chart.gif) 

    As above, the median salary is calculated for each job schedule type and then sorted in ascending order to create the following bar chart:  
    <br>
    ![bar chart of median salaries by job schedule type](images/type_bar_chart.gif)  

- **Map Chart:**
    The median salary for each country in the dataset is calculated based on the selected option values and displayed on the following map chart:  
    <br>

    ![Chart Map](images/map_chart.gif)


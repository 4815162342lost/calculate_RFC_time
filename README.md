### Calculate the spent time for RFC lifecycle

Unfortunately this project was not approval by my Change manager and currently in frozen state. Script was not tested very well and was not implemented on production.

But anyway, the main approach of small project:
I am working on huge outsourcing project and we are follow to ITIL approach. Any huge and small changes on our infrastructure called 'RFC'. For example, server decommission, hardware upgrade, new databases installation, implement new it solution -- RFC. 
And we have multiple separate teams in project (Linux team, Windows team, Virtualization team and etc.). Each team which participate on RFC should provide assessment and plan. 

Customer pay for our work (for all time which was spent for working with RFC). RFC time is contain assessment, implementation, planning, internal meeting phases. Also cost depend on engineer's grade. For high grade engineer customer pays more money. EN1 -- lowest grade, EN4 -- highest. All engineers should log spending time in excel file in 'Resource involved' sheet as follow:

![screen](https://raw.githubusercontent.com/4815162342lost/calculate_RFC_time/master/screens/Selection_622.png)

For each RFC we have separate file. And on end of month we need to create the report which calculate total time which spend to all RFC. And this solution automatize this process. 

Script read each file in ./RFC folder and create one report which contain:
1) Total time for spending for each RFC
2) Total spending time for each RFC grouping by grade
3) Total spending time for each RFC grouping by engineer
4) Total spending time for each RFC grouping by team
5) Total spending time for each RFC grouping by task

**Example:**
![screen](https://raw.githubusercontent.com/4815162342lost/calculate_RFC_time/master/screens/Selection_668.png)

On **"Total_metrics"** sheet we can see summarizing metriks for all RFC:
![screen](https://raw.githubusercontent.com/4815162342lost/calculate_RFC_time/master/screens/Selection_669.png)

**How to run:**
1) Copy all RFC form to ./RFC directory:
![screen](https://raw.githubusercontent.com/4815162342lost/calculate_RFC_time/master/screens/Selection_670.png)
2) Run the script and read output:
![screen](https://raw.githubusercontent.com/4815162342lost/calculate_RFC_time/master/screens/Selection_671.png)
3) Open the report file and enjoy.

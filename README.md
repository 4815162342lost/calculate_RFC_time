### Calculate the spent time for RFC lifecycle

Unfortunately this project was not approval by my manager and currently in frozen state. Scrip was not tested very well and was not implemented on production.

But anyway, the main approach of small project:
I am working on huge outsoursing project and we are follow to ITIL approach. Any huge and small changes on our infrastructure called 'RFC'. For example, server decomission, hardware upgrade, new databases installation, implement new it solution -- RFC. 
And we have multiple separate teams in project (Linux team, Windows team, Virtualization team and etc.). Each team which participate on RFC should provide assessment and plan. 

Customer pay for our work (for all time which was spent for working with RFC). RFC time is contain assessment, implementation, planning, internal meeting phases. Also cost depend on engineer's grade. For high grade engineer customer pay more money. EN1 -- lowest grade, EN4 -- higher. All enheneer should log spendint time in excl file in 'Resource involved' sheet as follow:

![screen](https://raw.githubusercontent.com/4815162342lost/calculate_RFC_time/master/screens/Selection_622.png)

For each RFC we have separate file. And on end of month we need to create the report which calculate total time which spend to RFC. And tis solution automatize this process. 

Script read each file in ./RFC folder and create one report which contain:

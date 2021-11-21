#  Schedule and orders combiner

The app has very specific use case, though, it saves up to 30 minutes a day of manual work for two people at the moment. The main reason behind using Node was my lack of knowledge of Visual Basic as I am sure that can be easily achived with VB marcos.

The interface that provides us with orders .xlsx files has some issue and gives us wrong courier's company name. But we can download another .xlsx file called schedules.xlsx and combine them into one that gives us all the information required for our workflow during the day.

The result will be valid if you put two files into Schedules_Orders folder one that begings with "orders" and and another that begings with "schedules", as there might be endings variations when you have multiple files on your system, e.g. "schedules (39)" or "orders (20)".

The console message will desplay an array of couriers that had been substituted and had 0 orders.


# ScrapeDataWebVBA

In this tutorial gonna Scrape Data from Web page using VBA and saving the results in excel sheets.

## Libraries (Requirements)
First of all if you need to browse on web page with Macros you need to add some references :

. [microsoft html object library](https://arkham46.developpez.com/articles/office/officeweb/?page=page_1)

. [Internet Explorer Object](https://riptutorial.com/vba/example/27772/internet-explorer-object)

To add those references you need to :
1. Go to  Visual Basic Editor
2. Select  tools
3. Select references
4. Serach for the two references and select them, then click  OK

## Code Part

To use my VBA Code when you download it, you need first to let all the Macros works on your Machine
That's why the first step is : 
* Going to Macros security : &nbsp;&nbsp;

![image](https://github.com/tahajadid/ScrapeDataWebVBA/blob/main/img/securite1.jpg)

* Go to Macros settings, And Select the Last one (Active all the Macros..), Then select the Last checkBox as shown :&nbsp;

![image](https://github.com/tahajadid/ScrapeDataWebVBA/blob/main/img/securite2.jpg) &nbsp;

And now go into VBA editor to see the code (Module), And put your target web site as *IE.navigate* content : &nbsp;

![image](https://github.com/tahajadid/ScrapeDataWebVBA/blob/main/img/target%20web%20page.jpg)

* Youn need to know the number of rows of your table and put the number as content of *RowNum* content.(in my case i have only 5 rows) &nbsp;

And Finally you can click on *Plot* button that execute the Code (Module) to return the following result (I puted a random information):&nbsp;

![image](https://github.com/tahajadid/ScrapeDataWebVBA/blob/main/img/Resultats.jpg)


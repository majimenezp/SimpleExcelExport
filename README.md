SimpleExcelExport
=================

SimpleExcelExport it is a small library to export to excel(using NPOI) a list of objects in your program, in order to add export capabilities to your application in a short time.

I already using in some projects, and maybe you can give a try if have a lot of modules where you need to exporte some data to excel.

Soon i will to upload in nuget for a easy install.

##Dependencies:

- NPOI (Installing using nuget: Install-Package NPOI)

##How to use

- Add the reference in your project.

- Create a list of object an populate.

       var persons = new List<Person>();
       
- Pass the list to SimpleExcelExport:

	var result=SimpleExcelExport.ExportToExcel.ListToExcel<Person>(persons);

- The library return the generated excel as a byte array.

##Column names and columns order
The columns names and order by default are the property name and alphabetical order, if you want to set a custom name for the columns and order, you need to add an attribute to your properties in your POCO class, for example:

	[ExcelExport("Last Name", order = 2,HFontBold=true,HFontColor ="White",HBackColor="Red")]
        public string LastName { get; set; }

        [ExcelExport("day of birth", order = 3)]
        public DateTime BirthDay { get; set; }

Where you set the column name with the first parameter and the column order with the named parameter "order".
Also(thanks to @rivuc) you can define font weight(bold), cell background color and font(foreground) color.

Any comment or idea for a new feature(or even better to send a pull request) contact me at:
twitter: [@majimenezp](http://twitter.com/majimenezp)

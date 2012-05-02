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

Any comment or idea for a new feature(or even better to send a pull request) contact me at:
twitter: [@majimenezp](http://twitter.com/majimenezp)

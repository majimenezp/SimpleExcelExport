Feature: List of objects export to excel
	In order to export a list of objetcs
	As user
	I want to export to excel

Scenario: Export list of person
	Given a this list of persons
	| Name  | LastName | BirthDay   | Country | Sex    | NumberOfChildren | Height |
	| Cosme | Fulanito | 01/01/1980 | Mexico  | Male   | 0                | 1.70   |
	| Maria | Gomez    | 02/02/1980 | Panama  | Female | 2                | 1.60   |
	| John  | Doe      | 03/03/1975 | USA     | Male   | 1                | 1.90   |
	Then export the list to a excel file located in:'C:\output.xls'
	Then export the list with header format to a excel file located in:'C:\outputh.xls'

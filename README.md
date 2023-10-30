# VBA Templates

#A Quick Intro on Adding Custom Functions

Save your Excel file as an .xlsm, and go to the Developer Tab. Hit Visual Basic.

![VisualBasic](https://github.com/lsbravo/VBA_Templates/assets/121823541/12e014ab-0241-4c2c-bcf7-e30b770e4475)

You'll get a new window called Microsoft Visual Basic for Applications. In that, hit Insert and then Module.

![Add a Module](https://github.com/lsbravo/VBA_Templates/assets/121823541/ba861d8f-834e-4849-b535-03664df106c6)

You can double click that module to get a space for your Function code. This script is customized to take one parameter (the project name) and build a path from it.

![CustomFunction](https://github.com/lsbravo/VBA_Templates/assets/121823541/152f21b9-c4ee-4622-b517-dd93a260f5ec)

Hit save.

To use the custom fuction, you'll type it in like any other--make the cell =ProjLink() and include the project name between the parentheses. If you're operating inside a table, you can use [@ColumnName] to make the formula more readable.

![TheFormula](https://github.com/lsbravo/VBA_Templates/assets/121823541/972ea747-982e-4b85-ad99-7577e569c7ef)


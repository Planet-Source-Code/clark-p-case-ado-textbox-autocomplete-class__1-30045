AutoComplete.cls
By Robert Rowe

   Coal Software & Systems
   robert_rowe@yahoo.com	
   12/07/00

(ADO Adaptation by Clark Case)

Based on code from By MauTheMan - 1999

This class adds AutoComplete functionality to a textbox.
You start typing and the class searches for a record that starts
with the letters you typed, if it finds one, it completes the rest
of the textbox with the record. 

HOW TO USE IT:

Requires: Microsoft Active Data Ojects 2.5

Add the class module "AutoComplete.cls" to your project. 
Declare an instance of the class in the Declarations section of a form
In the Form Load event, set the DatabasePathName, TableName, FieldName 
and theTextBox properties (don't forget to use the Set keyword when 
assigning your Textbox to theTextBox).
Call the CheckKey method from the textbox keydown event passing the KeyCode.
Call the AutoComplete method from the textbox change event.
If you want to have more than one AutoComplete textbox in the same form, use
multiple instances of the class.

Enjoy.

Any problem, E-mail me: robert_rowe@yahoo.com
			
Clark Case cpcase@msn.com

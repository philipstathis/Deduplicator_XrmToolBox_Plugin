#####Deduplicator Plugin for the XrmToolBox

This code is meant to be installed and executed as a plugin for the popular toolbox for Microsoft Dynamics CRM, [XrmToolBox](http://www.xrmtoolbox.com/)

![Deduplicator SnapShot](http://i.imgur.com/6JDBaZM.jpg)

#####Description

This code was created by the need to have a CRM administrator manually intervene and perform a 1-time merge and delete of duplicated records in an organization.
Using the full power and modularity of XrmToolBox, this sample has no custom code to connect to CRM environments or store credentials.
It it focused on helping the administrator set up rules that would determine a set of records as duplicates, then allow the user to merge dependant 1:N records and purge all duplicates minus the selected surviving record.

#####How To Install
* Download [XrmToolBox](http://www.xrmtoolbox.com/)
* Navigate to Releases and download https://github.com/philipstathis/Deduplicator_XrmToolBox_Plugin/releases/download/v1.0/Deduplicator.dll
* Place Deduplicator.dll in the Plugins Directory
* Launch XrmToolBox!

#####How To Build and Run
* Depending on your development machine setup you may need to install CRM 2015 SDK (Available from msdn: http://www.microsoft.com/en-us/download/details.aspx?id=44567)
* Resolve References to the CRM 2015 Assemblies
* Download [XrmToolBox](http://www.xrmtoolbox.com/)
* Compile Solution using .NET 4.5.2
* Place Deduplicator.dll in the Plugins Directory
* Launch XrmToolBox!

#####How To Use
* Steps are outlined to take the user through the merge and delete duplicate process in the application
* Step 1: Select Entity
* Step 2: Pick Attributes that indicate unique records (can be more than one selection so the entire set will be unique)
* Step 3: Pick 1:N related Entity
* Step 4: View Duplicates
* Step 5: Select Row from Step 4 and view detailed records (here Step 2 columns can be tweaked to show more data)
* Step 6: Same Row from Step 4 is queried in terms of related entity
* Step 7: Select Row from Step 5 to be the surviving record and update all rows from Step 6
* Step 8: Delete Rows from Step 5 that are toggled as 'Marked for Deletion' 

#####Features
* Works with July Release of XrmToolBox!
* Connects to any CRM instance supported by XrmToolBox (2011/2013/2015) and retrieves entity and attribute metadata
* User has full control to specify entity and criteria to determine duplicate records
* User can also choose to update a single related 1:N entity to use the surviving record in step 3 (optional)
* Configurable display of records to be deleted so you can add columns as needed and re-query

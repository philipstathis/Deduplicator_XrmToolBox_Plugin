#####Deduplicator Plugin for the XrmToolBox

This code is meant to be installed and executed as a plugin for the popular toolbox for Microsoft Dynamics CRM, [XrmToolBox](http://www.xrmtoolbox.com/)

#####Description

This code was created by the need to have a CRM administrator manually intervene and perform a 1-time merge and delete of duplicated records in an organization.
Using the full power and modularity of XrmToolBox, this sample has no custom code to connect to CRM environments or store credentials.
It it focused on helping the administrator set up rules that would determine a set of records as duplicates, then allow the user to merge dependant 1:N records and purge all duplicates minus the selected surviving record.

#####How To Use

#####Features
* Connects to any CRM instance supported by XrmToolBox (2011/2013/2015) and retrieves entity and attribute metadata
* User has full control to specify entity and criteria to determine duplicate records
* User can also choose to update a single related 1:N entity to use the surviving record in step 3 (optional)
* Configurable display of records to be deleted so you can add columns as needed and re-query

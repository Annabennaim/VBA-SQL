# Pricing on monetary notes using VBA and SQL

This project was decomposed in three main parts:

1) First, we had to retrieve data from the interface Bloomberg. We created an excel file named "MONETARY_INDEX_DATA" with the values of multiple monetay index such as EONIA or ESTER, from 01/01/2016 to 12/31/2023. We used the data from Bloomberg. We then created an other file called "NOTES_DATA", which contains the principal caracteristics of the notes given. Finally, we created a final file named " NOTE_RATE_DATA".
2) Secondly, with alle the date we retrieve, we had to construct a database on SQL with 4 different tables: "MONETARY_INDEX", "ISSUANCE_STATIC_DATA", "ISSUANCE_RATE_DATA" AND "ISSUANCE_DYNAMIC_DATA". This last table was fill with the data of the other tables.
3) Lastly, we had to create an other file with just one sheet "Note"

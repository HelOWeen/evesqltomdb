CREATE TABLE [invGroups] 
(   [groupID] INTEGER,
   [categoryID] INTEGER,
   [groupName] TEXT(100),
   [description] MEMO,
   [useBasePrice] YESNO,
   [allowManufacture] YESNO,
   [allowRecycler] YESNO,
   [anchored] YESNO,
   [anchorable] YESNO,
   [fittableNonSingleton] YESNO,
   [published] YESNO 

)

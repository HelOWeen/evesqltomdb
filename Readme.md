# EVESqlToMdb
Export/convert EVE's static data export (SDE) to an _(existing)_ MS Access database.    
    
_(Strictly speaking, this tool potentially allows to export/convert data from any to any DBMS, provided 
a) there's an ADO provider available for the DBMS and b) you get the data type mapping right. But 
I've only tested it with MS SQL Server -> MS Access (MDB))_

---

## Purpose
As you may or may not know, I'm the developer of the EVE 3rd party tool [EVEWalletAware](http://eve.basicaware.de/evewalletaware/index.html), nicknamed "EWA". 
EWA uses a MS Access database as its local storage. CCP provides its SDE (mainly) as an SQL dump. The obvious task at hand: 
copy the necessary data from SQL database to the Access database. With a few standard DB tools, this is 
a task taking ~ 15-20 minutes. Which was all good while CCP had its 2x a year release cycle. Now with the 
6 week release cycle, those few minutes do add up. Long story short - time to _"toolize"_(tm) that task.  

Which brings us to __EVESqlToMdb__. Sure, I could quickly throw together some script, hardcoding all the tables and columns. 
But experience shows that those quick'n'dirty hacks come back to bite you sooner or later and that fixing this stuff 
takes longer in the long run than doing it right in the first place.  

EVESqlToMdb therefore was designed in a way so that I can cope with database changes by just changing its 
XML configuration/mapping file and be done with it, once the application itself works OK.

---

## Prerequisits / Recommended tools
If you're working with the SDE, you're obviously using one version of Microsoft's SQL Server. If you're just starting with EVE's database 
and are in need of an SQL Server: MS provides its Express editions of their [SQL Server for free](http://www.microsoft.com/en-us/server-cloud/products/sql-server-editions/sql-server-express.aspx). 
While you're there, you should also grab a copy of the SQL Server Management Studio (also free).

As CCP moves more and more data out of the SDE into various other data sources (YAML, SQLite), you might want to grab a copy of Desmont McCallock's 
_(of EVEMon)_ [SDEExternalsToSql](https://forums.eveonline.com/default.aspx?g=posts&t=444535), which puts all those various data sources back into the SDE.

So my workflow for creating [EWA's](http://eve.basicaware.de/evewalletaware/index.html) database looks like
- Grab the SDE from CCP's [developer page](Static Database Export), restore the SQL backup included in the SDE to my SQL server installation.
- Use Desmont's [SDEExternalsToSql](https://forums.eveonline.com/default.aspx?g=posts&t=444535) to import the various YAML files and SQLite DBs which 
came with the SDE back into the SQL database.
- Run EVESqlToMdb to create EWA's database from there.


## XML configuration
A sample XML configuration files looks like this:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<!--
 *** SQL to MS Access table/columns mapping ***
 Possible values of the attributes TypeSource/TypeTarget are the equivalent ADO DataTypes enum values, i.e. 11 = adBoolean
 See http://www.w3schools.com/asp/ado_datatypes.asp
 -->
<Settings>
    <DBSettings>
        <DBSetting Name="Source">
            <Connection>Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=EVEDB;User Id=eve;Password=eve1;</Connection>
        </DBSetting>
        <DBSetting Name="Target">
            <Connection>Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DATA\EVEWalletAware\EVEWalletAware.mdb;User Id=admin;Password=;</Connection>
        </DBSetting>
    </DBSettings>
    <Tables>
        <Table Name="invCategories" Query="published=1">
            <Columns>
                <Column Name="categoryID" TypeTarget="3" IsIndex="true" IsPrimary="true" IsUnique="true" />
                <Column Name="categoryName" TypeTarget="202" Size="100" />
                <Column Name="published" TypeTarget="11" />
            </Columns>
        </Table>
        <Table Name="invGroups">
            <Columns>
                <Column Name="groupID" IsIndex="true" IsPrimary="true" IsUnique="true" />
                <Column Name="categoryID" IsIndex="true" />
                <Column Name="groupName" TypeTarget="202" Size="100" />
                <!-- Memo columns in Access need the size defined as 65535 -->
                <Column Name="description" TypeTarget="203" Size="65535" AllowNull="true" />
                <Column Name="useBasePrice" TypeTarget="11" />
                <Column Name="allowManufacture" TypeTarget="11" />
                <Column Name="allowRecycler" TypeTarget="11" />
                <Column Name="anchored" TypeTarget="11" />
                <Column Name="anchorable" TypeTarget="11" />
                <Column Name="fittableNonSingleton" TypeTarget="11" />
                <Column Name="published" TypeTarget="11" />
            </Columns>
        </Table>
    </Tables>
    <!-- Universal data type mappings -->
    <Mappings>
        <Mapping TypeSource="200" TypeTarget="202" />
    </Mappings>
</Settings>
```

Let's run through the nodes and their meaning.

### Node _&lt;DBSetting&gt; and &lt;Connection&gt;_
```xml
<DBSetting Name="Source">
    <Connection>Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=EVEDB;User Id=eve;Password=eve1;</Connection>
</DBSetting>
```
A configuration needs to provide exactly __2__ &lt;DBSetting&gt; nodes. One named _"Source"_ and one _"Target"_. 
As you can see, both are regular ADO connection strings. The _gotcha_ here is that the target database _needs_ to exist already.

If you're unsure how to construct the required ADO connection string, have a look at 
[http://www.connectionstrings.com/](http://www.connectionstrings.com/).

### Node _&lt;Table&gt;_
Each &lt;Table&gt; node contains the definition of the columns we want to copy from this table and - if necessary, 
the data type mappings. ___Note:___ if this table is already present in the _target_ database, __IT WILL BE DROPPED!__ 
... and recreated according to the columns definition present here.

You may optionally specific a _Query_ for limiting/controlling which records will be copied from the 
source database. The _Query_-Attribute is the __WHERE__ clause of a SQL statement. For example if 
you would use the statement ...
```sql
SELECT typeID, typeName FROM invTypes WHERE published=1;
```
... to limit the exported items to those that are published _(=ingame available)_, the &lt;Table&gt; 
node should look like this:
```xml
<Table Name="invTypes" Query="published=1">
```
In theory every valid __WHERE__ clause should work.


### Node _&lt;Column&gt;_
The &lt;Column&gt; nodes define the needed columns and if necessary their data type in the target DB. 
The following attributes are available to control the table/column creation in the target DB:
- __Name__ _(mandatory)_    
Obviously the name of the column
- __TypeTarget__    
If the data type of the target column differs fro some reason from the source column's data type, put 
the needed data type in here. This number is an _ADO DataType Enum_ (see below).
- __Size__    
The size of the column, typically used in conjunction with string data types.
- __AllowNull__    
What you expect it to be - allow NULL values for this column.
- __IsIndex__    
Create an index for this column.
- __IsPrimary__    
Set this column as the primary key.
- __IsUnique__    
Require unique values.

### Node _&lt;Mapping&gt;_
These nodes let you define overall mapping rules, which should apply for all columns of that type, except 
where the _Column_ definition overwrites these rules.

This should be pretty self-explanatory: columns of data type _TypeSource_ should be created as columns 
_TypeTarget_ in the target database.

---

## ADO DataType Enum
These are the [ADO datatypes as listed by Microsoft](https://msdn.microsoft.com/en-us/library/ms675318%28v=vs.85%29.aspx). 
The values are used to define column mappings, if the source column's data type doesn't "translate" well 
into the target's DBMS column type, i.e. SQL's adVarChar _(200)_ throws an error when attempting to 
create a column in a MS Access database. A working datatype for MS Access in that case is adVarWChar _(202)_.
 
<table border="1">
    <tr>
        <th>Name</th>
        <th>Value _(decimal)_</th>
        <th>Description</th>
    </tr>
    <tr>
        <td>adBigInt</td>
        <td align="right">20</td>
        <td>
            Indicates an eight-byte signed integer (DBTYPE_I8).
        </td>
    </tr>
    <tr>
        <td>adBinary</td>
        <td align="right">128</td>
        <td>
            Indicates a binary value (DBTYPE_BYTES).
        </td>
    </tr>
    <tr>
        <td>adBoolean</td>
        <td align="right">11</td>
        <td>
            Indicates a Boolean value (DBTYPE_BOOL).
        </td>
    </tr>
    <tr>
        <td>adBSTR</td>
        <td align="right">8</td>
        <td>
            Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
        </td>
    </tr>
    <tr>
        <td>adChapter</td>
        <td align="right">136</td>
        <td>
            Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
        </td>
    </tr>
    <tr>
        <td>adChar</td>
        <td align="right">129</td>
        <td>
            Indicates a string value (DBTYPE_STR).
        </td>
    </tr>
    <tr>
        <td>adCurrency</td>
        <td align="right">6</td>
        <td>
            Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
        </td>
    </tr>
    <tr>
        <td>adDate</td>
        <td align="right">7</td>
        <td>
            Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
        </td>
    </tr>
    <tr>
        <td>adDBDate</td>
        <td align="right">133</td>
        <td>
            Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
        </td>
    </tr>
    <tr>
        <td>adDBTime</td>
        <td align="right">134</td>
        <td>
            Indicates a time value (hhmmss) (DBTYPE_DBTIME).
        </td>
    </tr>
    <tr>
        <td>adDBTimeStamp</td>
        <td align="right">135</td>
        <td>
            Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
        </td>
    </tr>
    <tr>
        <td>adDecimal</td>
        <td align="right">14</td>
        <td>
            Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
        </td>
    </tr>
    <tr>
        <td>adDouble</td>
        <td align="right">5</td>
        <td>
            Indicates a double-precision floating-point value (DBTYPE_R8).
        </td>
    </tr>
    <tr>
        <td>adEmpty</td>
        <td align="right">0</td>
        <td>
            Specifies no value (DBTYPE_EMPTY).
        </td>
    </tr>
    <tr>
        <td>adError</td>
        <td align="right">10</td>
        <td>
            Indicates a 32-bit error code (DBTYPE_ERROR).
        </td>
    </tr>
    <tr>
        <td>adFileTime</td>
        <td align="right">64</td>
        <td>
            Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
        </td>
    </tr>
    <tr>
        <td>adGUID</td>
        <td align="right">72</td>
        <td>
            Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
        </td>
    </tr>
    <tr>
        <td>adIDispatch</td>
        <td align="right">9</td>
        <td>
            Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH).<br />Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
        </td>
    </tr>
    <tr>
        <td>adInteger</td>
        <td align="right">3</td>
        <td>
            Indicates a four-byte signed integer (DBTYPE_I4).
        </td>
    </tr>
    <tr>
        <td>adIUnknown</td>
        <td align="right">13</td>
        <td>
            Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN).<br />Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
        </td>
    </tr>
    <tr>
        <td>adLongVarBinary</td>
        <td align="right">205</td>
        <td>
            Indicates a long binary value.
        </td>
    </tr>
    <tr>
        <td>adLongVarChar</td>
        <td align="right">201</td>
        <td>
            Indicates a long string value.
        </td>
    </tr>
    <tr>
        <td>adLongVarWChar</td>
        <td align="right">203</td>
        <td>
            Indicates a long null-terminated Unicode string value.
        </td>
    </tr>
    <tr>
        <td>adNumeric</td>
        <td align="right">131</td>
        <td>
            Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
        </td>
    </tr>
    <tr>
        <td>adPropVariant</td>
        <td align="right">138</td>
        <td>
            Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
        </td>
    </tr>
    <tr>
        <td>adSingle</td>
        <td align="right">4</td>
        <td>
            Indicates a single-precision floating-point value (DBTYPE_R4).
        </td>
    </tr>
    <tr>
        <td>adSmallInt</td>
        <td align="right">2</td>
        <td>
            Indicates a two-byte signed integer (DBTYPE_I2).
        </td>
    </tr>
    <tr>
        <td>adTinyInt</td>
        <td align="right">16</td>
        <td>
            Indicates a one-byte signed integer (DBTYPE_I1).
        </td>
    </tr>
    <tr>
        <td>adUnsignedBigInt</td>
        <td align="right">21</td>
        <td>
            Indicates an eight-byte unsigned integer (DBTYPE_UI8).
        </td>
    </tr>
    <tr>
        <td>adUnsignedInt</td>
        <td align="right">19</td>
        <td>
            Indicates a four-byte unsigned integer (DBTYPE_UI4).
        </td>
    </tr>
    <tr>
        <td>adUnsignedSmallInt</td>
        <td align="right">18</td>
        <td>
            Indicates a two-byte unsigned integer (DBTYPE_UI2).
        </td>
    </tr>
    <tr>
        <td>adUnsignedTinyInt</td>
        <td align="right">17</td>
        <td>
            Indicates a one-byte unsigned integer (DBTYPE_UI1).
        </td>
    </tr>
    <tr>
        <td>adUserDefined</td>
        <td align="right">132</td>
        <td>
            Indicates a user-defined variable (DBTYPE_UDT).
        </td>
    </tr>
    <tr>
        <td>adVarBinary</td>
        <td align="right">204</td>
        <td>
            Indicates a binary value.
        </td>
    </tr>
    <tr>
        <td>adVarChar</td>
        <td align="right">200</td>
        <td>
            Indicates a string value.
        </td>
    </tr>
    <tr>
        <td>adVariant</td>
        <td align="right">12</td>
        <td>
            Indicates an Automation Variant (DBTYPE_VARIANT).<br />Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
        </td>
    </tr>
    <tr>
        <td>adVarNumeric</td>
        <td align="right">139</td>
        <td>
            Indicates a numeric value.
        </td>
    </tr>
    <tr>
        <td>adVarWChar</td>
        <td align="right">202</td>
        <td>
            Indicates a null-terminated Unicode character string.
        </td>
    </tr>
    <tr>
        <td>adWChar</td>
        <td align="right">130</td>
        <td>
            Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
        </td>
    </tr>
</table>


---

# Version history


### Version 1.0.15, 21.10.2015
#### Bugfix

-  

#### New

-  

#### Change

-  

#### Misc.

- __Initial release.__  

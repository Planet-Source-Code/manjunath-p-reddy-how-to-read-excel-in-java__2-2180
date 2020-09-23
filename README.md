<div align="center">

## How to read Excel in java


</div>

### Description

This article shows a simple ways of accessing spread sheets (such as microsoft excel) in java
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Manjunath P Reddy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/manjunath-p-reddy.md)
**Level**          |Intermediate
**User Rating**    |4.7 (186 globes from 40 users)
**Compatibility**  |Java \(JDK 1\.2\)
**Category**       |[Databases/ JDBC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-jdbc__2-61.md)
**World**          |[Java](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/java.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/manjunath-p-reddy-how-to-read-excel-in-java__2-2180/archive/master.zip)





### Source Code

<br><br><P>
This sample code was an outcome of a friend's requirement for reading an excel file dynamically. He had an EJB layer for
a flight management system for read/writes to database. It so happened that the data also started coming from various
departments in a spread sheet format. So either he had to import the data to his oracle database manually or re-design
his EJB's for accomadating this new data input.
So what i designed for him was a simple facade pattern classes run by a daemon process which makes use of the existing
enterprise java beans which enabled him to treat the spreadsheet data as no different. The scope of the design is beyond
this article.
<p>
So what i will illustrate in this article is a simple way of accessing spreadsheets as if they were a database. This
article holds good for java running on windows-servers. The access itself is through a jdbc-odbc bridge.
<P>
okie..so here we go...
<P>
Open the odbc data administrator console and click on system dsn. select add and add the Microsoft Excel driver from the
list of drivers and give a name to the dsn(say exceltest) and select the workbook.
<P>
Then all one needs is to connect through this dsn just like connecting to the database and accessing records.
<P>
Here's the sample code
<P><P>
<PRE>
import java.io.*;
import java.sql.*;
public class ExcelReadTest{
	public static void main(String[] args){
		Connection connection = null;
		try{
			Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
			Connection con = DriverManager.getConnection( "jdbc:odbc:exceltest" );
			Statement st = con.createStatement();
			ResultSet rs = st.executeQuery( "Select * from [Sheet1$]" );
			ResultSetMetaData rsmd = rs.getMetaData();
			int numberOfColumns = rsmd.getColumnCount();
			while (rs.next()) {
				for (int i = 1; i <= numberOfColumns; i++) {
					if (i > 1) System.out.print(", ");
					String columnValue = rs.getString(i);
					System.out.print(columnValue);
				}
				System.out.println("");
			}
			st.close();
			con.close();
		} catch(Exception ex) {
			System.err.print("Exception: ");
			System.err.println(ex.getMessage());
		}
	}
}
</PRE>
<br>
<b>Answer to Comments:</b><br>
Hey john ur comment is right and well recieved. The reason for the first row not being printed is that the jdbc-bridge assumes the first row to be akin to column names in the database. Hence the first available row is the name of the column...which explains why the third row is printed if the first row is missing.
<br>
Hope this helps and thanks for the comment
<br>
<br>
Standard sql queries like
<pre>
Select column_name1,column_name2 from [Sheet1$] where column_name3 like '%bob%';
</pre>
can be used on the spreadsheet.
<br>
You can use the following snippet to print the column names(which is the first row of the spread sheet).<br>
<pre>
			for (int i = 1; i <= numberOfColumns; i++) {
				if (i > 1) System.out.print(", ");
				String columnName = rsmd.getColumnName(i);
				System.out.print(columnName);
			}
			System.out.println("");
</pre>


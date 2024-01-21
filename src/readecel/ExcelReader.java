package readecel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;


public class ExcelReader {
    public static void main(String[] args) throws InvalidFormatException, ParseException {
        try {
	            File file = new File("C:\\Users\\Raj\\Downloads\\Assignment_Timecard.xlsx");
	            FileInputStream fis = new FileInputStream(file);
	            XSSFWorkbook workbook = new XSSFWorkbook(fis);
	            XSSFSheet sheet = workbook.getSheetAt(0); 
	            
	            int q1=0,q2=1,q3=2;
	            List<List<String>> memberlist=sevenCon(sheet);
	            
	            getName(sheet,memberlist,q1);
	            
	            List<List<String>> memberlist1=lessthanten(sheet);
	            getName(sheet,memberlist1,q2);
	            
	            List<List<String>> memberlist2=greaterthanfourteen(sheet);
	            getName(sheet,memberlist2,q3);
	            
	            workbook.close();
	        } catch (IOException e) 
		        {
		            e.printStackTrace();
		        }
                
    }

    //greater than fourteen finction
    private static List<List<String>> greaterthanfourteen(XSSFSheet sheet) {
		
    	List<List<String>> rowData = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            List<String> rowValues = new ArrayList<>();
            for (int j = 0; j <= sheet.getRow(i).getLastCellNum(); j++) {
                XSSFCell cell = sheet.getRow(i).getCell(j);
                if (cell != null) {
                    rowValues.add(cell.toString());
                } else {
                    rowValues.add("");  
                }
            }
            rowData.add(rowValues);
        }
        
        
        //grouping data by position & time in
        HashMap<String, List<String>> groupedData = new HashMap<>(); 
        for(int i = 0; i < rowData.size(); i++) {
            String primaryKey = rowData.get(i).get(0); 
            String column3Data = rowData.get(i).get(2); 
            if(groupedData.containsKey(primaryKey)) {
                groupedData.get(primaryKey).add(column3Data);
            } else {
                List<String> dataList = new ArrayList<>();
                dataList.add(column3Data);
                groupedData.put(primaryKey, dataList);
            }
        }
        
        
        
      //grouping data by position & time card
        HashMap<String, List<String>> groupedData2 = new HashMap<>(); 
        for(int i = 0; i < rowData.size(); i++) {
            String primaryKey = rowData.get(i).get(0); 
            String column4Data = rowData.get(i).get(4); 
            if(groupedData2.containsKey(primaryKey)) {
                groupedData2.get(primaryKey).add(column4Data);
            } else {
                List<String> dataList = new ArrayList<>();
                dataList.add(column4Data);
                groupedData2.put(primaryKey, dataList);
            }
        }
        
        
        List<List<String>> greaterthanmember = new ArrayList<>();
        List<String> greatermember = new ArrayList<>();
        
        
        
        Set<String> keys=groupedData.keySet();
        for(String key:keys)
        {
        	//System.out.println("key="+key+""+"values="+groupedData.get(key));
        	List<String> values = (groupedData.get(key));
        	int c=0;
        	Date next = null;
        	LocalTime timenext=null;
        	for(String value:values)
        	{
        		
        		if(value != "") 
        		{
        		 
        			//converting string to date formate
        		 String dateStr = value;
        	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        	        try {
        	        	
        	        		//converting string to time formate HH:MM
            	            Date date = dateFormat.parse(dateStr); //11
            	            
            	            String timeString = (groupedData2.get(key).get(c));
            	            String[] parts = timeString.split(":");
            	            int hours = Integer.parseInt(parts[0]);
            	            int minutes = Integer.parseInt(parts[1]);
            	            LocalTime time = LocalTime.of(hours, minutes);
            	            
            	            if(timenext == null) {
            	                timenext = time;
            	              }
            	            if(next == null) {
            	                next = date;
            	              }
            	            if(next.compareTo(date)==0) //
            	            {
            	            	LocalTime result = time.plusHours(timenext.getHour()).plusMinutes(timenext.getMinute());
            	            	 if (result.isAfter(LocalTime.of(14, 0))) {
            	                     greatermember.add(key);
            	                     break;
            	                 } 
  
            	            }
            	            else
            	            {
            	            	next=date;//10
            	            	timenext=time;
            	            }
            	            //System.out.println(c);
            	            c++;
            	            
            	
            	            

        	        	} catch (ParseException e) 
	            	        {
	            	            e.printStackTrace();
	            	        }
        		}
        		else
        		{
        			continue;
        		}
        		
        		
        	}
        	greaterthanmember.add(greatermember);
        
        
        }
        return greaterthanmember;

    	
	}//greaterthanfourteen


	//lessthanten function
    private static List<List<String>> lessthanten(XSSFSheet sheet) 
    {
        List<List<String>> rowData = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            List<String> rowValues = new ArrayList<>();
            for (int j = 0; j <= sheet.getRow(i).getLastCellNum(); j++) {
                XSSFCell cell = sheet.getRow(i).getCell(j);
                if (cell != null) {
                    rowValues.add(cell.toString());
                } else {
                    rowValues.add("");  
                }
            }
            rowData.add(rowValues);
        }
        
        
        //grouping data by position & time in
        HashMap<String, List<String>> groupedData = new HashMap<>(); 
        for(int i = 0; i < rowData.size(); i++) {
            String primaryKey = rowData.get(i).get(0); 
            String column3Data = rowData.get(i).get(2); 
            if(groupedData.containsKey(primaryKey)) {
                groupedData.get(primaryKey).add(column3Data);
            } else {
                List<String> dataList = new ArrayList<>();
                dataList.add(column3Data);
                groupedData.put(primaryKey, dataList);
            }
        }
        
        
        
      //grouping data by position & time card
        HashMap<String, List<String>> groupedData2 = new HashMap<>(); 
        for(int i = 0; i < rowData.size(); i++) {
            String primaryKey = rowData.get(i).get(0); 
            String column4Data = rowData.get(i).get(4); 
            if(groupedData2.containsKey(primaryKey)) {
                groupedData2.get(primaryKey).add(column4Data);
            } else {
                List<String> dataList = new ArrayList<>();
                dataList.add(column4Data);
                groupedData2.put(primaryKey, dataList);
            }
        }
        
        
        List<List<String>> lessthanmember = new ArrayList<>();
        List<String> lessmember = new ArrayList<>();
        
        
        
        Set<String> keys=groupedData.keySet();
        for(String key:keys)
        {
        	//System.out.println("key="+key+""+"values="+groupedData.get(key));
        	List<String> values = (groupedData.get(key));
        	int c=0;
        	Date next = null;
        	LocalTime timenext=null;
        	for(String value:values)
        	{
        		
        		if(value != "") 
        		{
        		 
        			//converting string to date formate
        		 String dateStr = value;
        	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        	        try {
        	        	
        	        		//converting string to time formate HH:MM
            	            Date date = dateFormat.parse(dateStr); 
            	            
            	            
            	            String timeString = (groupedData2.get(key).get(c));
            	            String[] parts = timeString.split(":");
            	            int hours = Integer.parseInt(parts[0]);
            	            int minutes = Integer.parseInt(parts[1]);
            	            LocalTime time = LocalTime.of(hours, minutes);
            	            
            	            if(timenext == null) {
            	                timenext = time;
            	              }
            	            if(next == null) {
            	                next = date;
            	              }
            	            if(next.compareTo(date)==0) //
            	            {
            	            	LocalTime result = time.plusHours(timenext.getHour()).plusMinutes(timenext.getMinute());
            	            	 if (result.isAfter(LocalTime.of(1, 0)) && result.isBefore(LocalTime.of(10, 0))) {
            	                     lessmember.add(key);
            	                     break;
            	                 } 
  
            	            }
            	            else
            	            {
            	            	next=date;//10
            	            	timenext=time;
            	            }
            	            //System.out.println(c);
            	            c++;
            	            
            	
            	            

        	        	} catch (ParseException e) 
	            	        {
	            	            e.printStackTrace();
	            	        }
        		}
        		else
        		{
        			continue;
        		}
        		
        		
        	}
        	lessthanmember.add(lessmember);
        
        
        }
        return lessthanmember;
	}//lessthanten


	//getname function
	private static void getName(XSSFSheet sheet, List<List<String>> memberlist, int q) {
		 List<List<String>> rowData = new ArrayList<>();
	        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	            List<String> rowValues = new ArrayList<>();
	            for (int j = 0; j <= sheet.getRow(i).getLastCellNum(); j++) {
	                XSSFCell cell = sheet.getRow(i).getCell(j);
	                if (cell != null) {
	                    rowValues.add(cell.toString());
	                } else {
	                    rowValues.add("");  
	                }
	            }
	            rowData.add(rowValues);
	        }
	        
	      //grouping data
	        HashMap<String, List<String>> groupedData = new HashMap<>(); 
	        for(int i = 0; i < rowData.size(); i++) {
	            String primaryKey = rowData.get(i).get(0); 
	            String column7Data = rowData.get(i).get(7); 
	            if(groupedData.containsKey(primaryKey)) {
	                groupedData.get(primaryKey).add(column7Data);
	            } else {
	                List<String> dataList = new ArrayList<>();
	                dataList.add(column7Data);
	                groupedData.put(primaryKey, dataList);
	            }
	        }
	        if(q==0)
	        {
	        	System.out.println("here is the detail of the employee who came 7 consecutive days.\n");
	        	Set<String> keys=groupedData.keySet();
		        
	        	for(List<String> member:memberlist)
	            {
	            	for(String mem:member) 
	            	{
	            		System.out.println("NAME= "+groupedData.get(mem).get(1)+" with Position ID= "+mem+" has worked for 7 or more consutive Days:-\n");
	            	}
	            	break;
	            }
	        }
	        if(q==1)
	        {
	        	System.out.println("here is the detail of the employee who came 10 hours of time between shifts but greater than 1 hour:-\n");
	        	Set<String> keys=groupedData.keySet();
	        	for(List<String> member:memberlist)
	            {
	            	for(String mem:member) 
	            	{
	            		System.out.println("NAME= "+groupedData.get(mem).get(1)+" with Position ID= "+mem+" has worked for 10 hours of time between shifts but greater than 1 hour.\n");
	            	}
	            	break;
	            }
	        }
	        if(q==2)
	        {
	        	System.out.println("here is the detail of the employee who has worked for more than 14 hours in a single shift:-\n");
	        	Set<String> keys=groupedData.keySet();
	        	for(List<String> member:memberlist)
	            {
	            	for(String mem:member) 
	            	{
	            		System.out.println("NAME= "+groupedData.get(mem).get(1)+" with Position ID= "+mem+" has worked for more than 14 hours in a single shift.\n");
	            	}
	            	break;
	            }
	        }
	        
	        
		
	}//getname
	
	
	
	
	
	//sevencon
	private static List<List<String>> sevenCon(XSSFSheet sheet) {
        List<List<String>> rowData = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            List<String> rowValues = new ArrayList<>();
            for (int j = 0; j <= sheet.getRow(i).getLastCellNum(); j++) {
                XSSFCell cell = sheet.getRow(i).getCell(j);
                if (cell != null) {
                    rowValues.add(cell.toString());
                } else {
                    rowValues.add("");  
                }
            }
            rowData.add(rowValues);
        }
        
        
        
        //grouping data by position & time in
        HashMap<String, List<String>> groupedData = new HashMap<>(); 
        for(int i = 0; i < rowData.size(); i++) {
            String primaryKey = rowData.get(i).get(0); 
            String column3Data = rowData.get(i).get(2); 
            if(groupedData.containsKey(primaryKey)) {
                groupedData.get(primaryKey).add(column3Data);
            } else {
                List<String> dataList = new ArrayList<>();
                dataList.add(column3Data);
                groupedData.put(primaryKey, dataList);
            }
        }
        

        
        List<List<String>> sevendaymember = new ArrayList<>();
        List<String> member = new ArrayList<>();
        
        Set<String> keys=groupedData.keySet();
        for(String key:keys)
        {
        	//System.out.println("key="+key+""+"values="+groupedData.get(key));
        	List<String> values = (groupedData.get(key));
        	int c=0;
        	Date next = null;
        	for(String value:values)
        	{
        		
        		if(value != "") 
        		{
        		 
        		 String dateStr = value;
        	        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        	        try {
        	        	
        	        		
            	            Date date = dateFormat.parse(dateStr); 
            	            
            	            if(next == null) {
            	                next = date;
            	              }
            	            if(next.compareTo(date)==0) 
            	            {
            	            	c++;
            	            }
            	            else
            	            {
            	            	next=date;
            	            }
            	            
            	            
            	            Calendar cal = Calendar.getInstance();
            	            cal.setTime(date);
            	            cal.add(Calendar.DATE, 1);
            	            next=(cal.getTime());//11
            	            

        	        	} catch (ParseException e) 
	            	        {
	            	            e.printStackTrace();
	            	        }
        		}
        		else
        		{
        			continue;
        		}
        		
        		
        	}
        	if(c>=7)
        	{
        		
        		member.add(key);
        	}
        	sevendaymember.add(member);
        	
        }
		return sevendaymember;
		
		
	}//sevencon

}//class
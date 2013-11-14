package poi.app.builder.ExcelBuilder;


import java.io.FileOutputStream;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author SushantKafle
 *
 */
public class excelBuilder {
	
	private XSSFCellStyle header;
	private XSSFCellStyle subheader;
	private XSSFCellStyle body;
	private String headerDistinguisher="~~";
	private String subheaderDistinguisher="~";
	private Workbook wb;
	
	public excelBuilder()
	{
		wb = new XSSFWorkbook();
	}
	
	
	/**
	 * <ul>
	 * 	<li>
	 * 		<h1>build</h1>
			 * The build function is the prime function that builds the
			 * excel file
	 * 	</li>
	 * </ul>
	 * <br>
	 * @param data
	 * - represents the two dimensional data to be dumped in the
	 * excel file.
	 */
	public void buildNewSheet(String sheetName,Object data[][])
	{
			//Create a sheet
			Sheet sheet = wb.createSheet(sheetName);
			
			//Initialize the styles
			initStyles(wb);
			
			//Dump all data to the sheet
			sheet=addData(sheet,data);
			
			//Get the regions to merge
			ArrayList<int[]>MergeData = new ArrayList<int[]>();
			MergeData=getMergeData(data);
			
			//Merge the required cells
			sheet=mergeSheet(sheet,MergeData);
			
			//AutoSize the columns
			autosizeColumnsFromSheet(sheet,0,data.length);	
			
			System.out.println("[Sucess] New Sheet "+sheetName+" created!!");
	}
	
	
	/**
	 * <ul>
	 * 	<li> <h1>saveAs</h1>
	 * saves the excel file to a filepath
	 * 	</li>
	 * </ul>
	 * @param filepath
	 * - represents the output file location
	 */
	public void saveAs(String filepath) 
	{
		try
		{
			//Write to a file
			FileOutputStream fileOut = null;
			fileOut = new FileOutputStream(filepath);
			wb.write(fileOut);
			fileOut.close();
			
			System.out.println("[Sucess] Saving to file Completed!");
			
		}catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	
	
	/**
	 * @param distinguisher
	 * - is a identifier that is used to identify a Header 
	 */
	public void setHeaderDistinguisher(String distinguisher)
	{
		headerDistinguisher = distinguisher;
	}
	
	
	/**
	 * @param distinguisher
	 * - is a identifier that is used to identify a subHeader
	 */
	public void setsubHeaderDistinguisher(String distinguisher)
	{
		subheaderDistinguisher = distinguisher;
	}
	
	
	//Initializing the styles
	private void initStyles(Workbook wb)
	{
		header = (XSSFCellStyle) createBorderedStyle(wb);
		subheader = (XSSFCellStyle) createBorderedStyle(wb);
		body = (XSSFCellStyle) createBorderedStyle(wb);
		
		Font f = wb.createFont();
		f.setFontName("Arial");
		f.setFontHeightInPoints((short)8);
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		f.setColor(IndexedColors.WHITE.getIndex());
		header.setFont(f);
		
		Font f1 = wb.createFont();
		f1.setFontName("Arial");
		f1.setFontHeightInPoints((short)8);
		f1.setColor(IndexedColors.BLACK.getIndex());
		body.setFont(f1);
		subheader.setFont(f1);
		
		header.setFillForegroundColor(new XSSFColor(new java.awt.Color(55,96,145)));
		header.setFillPattern(CellStyle.SOLID_FOREGROUND);
		header.setAlignment(HorizontalAlignment.CENTER);
        header.setVerticalAlignment(VerticalAlignment.CENTER);
        
        subheader.setAlignment(HorizontalAlignment.CENTER);
        subheader.setVerticalAlignment(VerticalAlignment.CENTER);
	}
	
	//Defining Styles
	private static CellStyle createBorderedStyle(Workbook wb){
	    CellStyle style = wb.createCellStyle();
	    style.setBorderRight(CellStyle.BORDER_THIN);
	    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
	    style.setBorderBottom(CellStyle.BORDER_THIN);
	    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	    style.setBorderLeft(CellStyle.BORDER_THIN);
	    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	    style.setBorderTop(CellStyle.BORDER_THIN);
	    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
	    return style;
	}
	
	
	private void autosizeColumnsFromSheet(final Sheet excelSheet, final int fromColumn, final int toColumn) {
        for (int i = fromColumn; i <= toColumn; i++) {
            excelSheet.autoSizeColumn(new Short(String.valueOf(i)));
            try {
                excelSheet.setColumnWidth(i, excelSheet.getColumnWidth(i) + 1300);
            } catch (final Exception e) {
                
            }
        }
    }
	
	
	/*
	 * TrimString remove the unwanted distinguishers
	*/
	private String trimString(String value)
	{
		if(value.startsWith(headerDistinguisher))
			return value.substring(2);
		else if(value.startsWith(subheaderDistinguisher))
			return value.substring(1);
		
		return value;
	}	
	

	private Sheet mergeSheet(Sheet sheet,ArrayList<int[]> MergeData)
	{
		for(int i=0;i<MergeData.size();i++)
		{
			int value[] = new int[4];
			value=MergeData.get(i);
			sheet.addMergedRegion(new CellRangeAddress(value[0],value[2],value[1],value[3]));
		}
		return sheet;
	}
	
	
	private boolean isMergable(Object obj)
	{
		if(obj instanceof String)
		{
			String ob = (String)obj;
			if(ob.startsWith(headerDistinguisher) || ob.startsWith(subheaderDistinguisher))
				return true;
		}
		
		return false;
	}
	
	
	private boolean isHeader(Object obj)
	{
		if(obj instanceof String)
		{
			String ob = (String)obj;
			if(ob.startsWith(headerDistinguisher))
				return true;
		}
		
		return false;
	}
	
	
	/*
	 * Add all data to the Sheet
	*/
	private Sheet addData(Sheet sheet,Object data[][])
	{
		int col=data[0].length;
		int row=data.length;
		
		for(int y=0;y<row;y++)
		{
			Row r=sheet.createRow(y);
			Cell c=null;
			for(int x=0;x<col;x++)
			{
				c=r.createCell(x);
				if(isHeader(data[y][x]))
					c.setCellStyle(header);
				else if(isMergable(data[y][x]))
					c.setCellStyle(subheader);
				else
					c.setCellStyle(body);

				//Update Required
				if(data[y][x] instanceof String)
					c.setCellValue(trimString((String)data[y][x]));
				else if(data[y][x] instanceof Integer)
					c.setCellValue((Integer)data[y][x]);
			}
		}
		
		return sheet;
	}


	
	/*
	 * Returns range of cell to merge [Critical Funtion]
	 */
	private ArrayList<int[]> getMergeData(Object data[][])
	{
		
		ArrayList<int[]> MergeData = new ArrayList<int[]>();
		
		int columns = data[0].length;
		int rows=data.length;
		
		//Defines a data-type called MergedSection
		ArrayList<int[]> MergedSection = new ArrayList<int[]>();
		
		
		int val[] = new int[2];
		
		val[0] = 0;
		val[1] = columns-1;
		MergedSection.add(val);
		
		
		for(int i=0;i<rows;i++)
		{
			
			ArrayList<int[]> tempSection = new ArrayList<int[]>();
			
			while(!MergedSection.isEmpty())
			{
				//Retrieve the first element from MergedSection
				int value[] = new int[2];
				value= MergedSection.get(0);
				MergedSection.remove(0);
				
				int startAt= value[0];
				int endAt=value[1];
				
				String previous="-1";
				
				int rep[] = new int[2];
				rep[0]=-1;
				
				for(int itr=startAt;itr<=endAt;itr++)
				{
					if(!isMergable(data[i][itr]))
						break;
					
					//Check for Merged Section in Array data
					if(data[i][itr].equals(previous))
					{
						//triggered at the first pass
						if(rep[0] == -1)
						{
							rep[0]=itr-1;
						}
						
						//for the last pass
						if((itr == endAt) && rep[0] != -1)
						{
							rep[1]=itr;
							int temp[] = new int[2];
							
							temp[0]=rep[0];
							temp[1]=rep[1];
							
							MergeData.add(new int[]{i,temp[0],i,temp[1]});
							
							tempSection.add(temp);
						}
						
						
					}else if(rep[0] != -1)
					{
						rep[1]=itr-1;

						//add the range in tempSection
						int temp[] = new int[2];
						temp[0]=rep[0];
						temp[1]=rep[1];
						
						tempSection.add(temp);
						
						MergeData.add(new int[]{i,temp[0],i,temp[1]});
						
						//Clear rep values for next pass
						rep[0]=-1;
					}
					previous=(String)data[i][itr];
				}
			}
			MergedSection = tempSection;
		}
		
		MergedSection.clear();
		val[0] = 0;
		val[1] = rows-1;
		MergedSection.add(val);
		
		for(int i=0;i<columns;i++)
		{
			
			ArrayList<int[]> tempSection = new ArrayList<int[]>();
			
			while(!MergedSection.isEmpty())
			{
				//Retrieve the first element from MergedSection
				int value[] = new int[2];
				value= MergedSection.get(0);
				MergedSection.remove(0);
				
				int startAt= value[0];
				int endAt=value[1];
				
				
				
				String previous="-1";
				
				int rep[] = new int[2];
				rep[0]=-1;
				
				for(int itr=startAt;itr<=endAt;itr++)
				{
					if(!isMergable(data[itr][i]))
						break;
					
					//Check for Merged Section in Array data
					if(data[itr][i].equals(previous))
					{
						
						//triggered at the first pass
						if(rep[0] == -1)
						{
							rep[0]=itr-1;
						}
						
						//for the last pass
						if((itr == endAt) && rep[0] != -1)
						{
							rep[1]=itr;
							int temp[] = new int[2];
							
							temp[0]=rep[0];
							temp[1]=rep[1];
							
							tempSection.add(temp);
							MergeData.add(new int[]{temp[0],i,temp[1],i});
						}
						
						
					}else if(rep[0] != -1)
					{
						rep[1]=itr-1;

						//add the range in tempSection
						int temp[] = new int[2];
						temp[0]=rep[0];
						temp[1]=rep[1];
						 
						tempSection.add(temp);
						MergeData.add(new int[]{temp[0],i,temp[1],i});
			
						//Clear rep values for next pass
						rep[0]=-1;
					}
					previous=(String)data[itr][i];
				}
			}
			MergedSection = tempSection;
		}
		
		return MergeData;
	}
	

}

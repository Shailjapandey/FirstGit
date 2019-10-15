package dataTable;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class BussinessLogic
{
  XSSFWorkbook wb;
  XSSFSheet xs;
  XSSFRow xrow;
  XSSFCell xhcell;
  XSSFCell xdcell;
  XSSFCell cell;
  private File folder = null;
  private int i;
  private ArrayList<String> al1 = new ArrayList();
  private Set<String> hs = new HashSet();
  
  public BussinessLogic() {}
  
  public void createExcel(String FilePath, String SheetName, List<String> header) throws Exception
  {
    int rowidx = 0;
    int cellidx = 0;
    
    wb = new XSSFWorkbook();
    xs = wb.createSheet(SheetName);
    xrow = xs.createRow(rowidx);
    
    XSSFFont font = wb.createFont();
    font.setBold(true);
    font.setFontHeight(10.0D);
    font.setFontName("Arial");
    font.setColor(IndexedColors.BLUE_GREY.getIndex());
    xs.autoSizeColumn(100000);
    

    for (cellidx = 0; cellidx < header.size() - 1; cellidx++)
    {
      XSSFCell cell = xrow.createCell(cellidx);
      CellStyle style = wb.createCellStyle();
      style.setWrapText(true);
      style.setFont(font);
      cell.setCellStyle(style);
      cell.setCellValue((String)header.get(cellidx + 1));
      xs.autoSizeColumn(cellidx);
    }
    
    FileOutputStream fos = new FileOutputStream(new File(FilePath));
    wb.write(fos);
    fos.close();
  }
  




  public void writeDatainExitingExcel(List<String> fileListName, List<String> data, String updateFilePath)
    throws Exception
  {
    FileInputStream in = new FileInputStream(updateFilePath);
    FileOutputStream fos = null;
    wb = new XSSFWorkbook(in);
    xs = wb.getSheetAt(0);
    


    for (int rowno = 1; rowno <= fileListName.size(); rowno++)
    {
      XSSFRow row = xs.createRow(rowno);
      for (int cellno = 0; cellno < data.size() - 1; cellno++)
      {
        XSSFCell cell = row.createCell(cellno);
        if (cellno == 2)
        {
          cell.setCellValue(((String)fileListName.get(rowno - 1)).toString());
          xs.autoSizeColumn(rowno);
        }
        else {
          cell.setCellValue((String)data.get(cellno + 1)); }
        xs.autoSizeColumn(cellno);
      }
    }
    in.close();
    fos = new FileOutputStream(new File(updateFilePath));
    wb.write(fos);
    fos.close();
  }
  


  public void writeDatainExitingExcel(List<String> missingFile, String updateFilePath, String SheetName)
    throws Exception
  {
    FileInputStream in = new FileInputStream(updateFilePath);
    FileOutputStream fos = null;
    wb = new XSSFWorkbook(in);
    xs = wb.createSheet(SheetName);
    
    XSSFFont font = wb.createFont();
    font.setBold(true);
    font.setFontHeight(10.0D);
    font.setFontName("Arial");
    font.setColor(IndexedColors.BLUE_GREY.getIndex());
    xs.autoSizeColumn(100000);
    
    xrow = xs.createRow(0);
    CellStyle style = wb.createCellStyle();
    style.setWrapText(true);
    style.setFont(font);
    cell = xrow.createCell(0);
    cell.setCellStyle(style);
    cell.setCellValue("Missing Files");
    System.out.println("mfile.size()" + missingFile.size());
    
    for (int rowno = 1; rowno <= missingFile.size(); rowno++)
    {
      xrow = xs.createRow(rowno);
      cell = xrow.createCell(0);
      cell.setCellValue(((String)missingFile.get(rowno - 1)).toString());
      xs.autoSizeColumn(0);
    }
    in.close();
    fos = new FileOutputStream(new File(updateFilePath));
    wb.write(fos);
    fos.close();
  }
  



  public ArrayList<String> compareFiles(String folderPath)
  {
    ArrayList<String> myList = new ArrayList();
    ArrayList<String> uniquefiles = new ArrayList();
    String fileendswith = "";
    String filetocompare = "";
    folder = new File(folderPath);
    File[] filelist = folder.listFiles();
    System.out.println(filelist.length);
    
    for (i = 0; i < filelist.length; i += 1) {
      myList.add(filelist[i].getName().toUpperCase());
    }
    

    for (String temp : myList)
    {
      fileendswith = temp.substring(temp.lastIndexOf('.') + 1);
      String filewithoutextension = temp.substring(0, temp.lastIndexOf("."));
      System.out.println("file end with  " + fileendswith);
      if (fileendswith.equalsIgnoreCase("xml")) {
        filetocompare = (filewithoutextension + ".xlsx").toUpperCase();
        System.out.println("filetocompare inside xml  " + filetocompare);
        

        if (!myList.contains(filetocompare))
        {





          uniquefiles.add(filetocompare);
        }
      } else {
        filetocompare = (filewithoutextension + ".xml").toUpperCase();
        System.out.println("filetocompare inside xlsx  " + filetocompare);
        if (!myList.contains(filetocompare))
        {



          uniquefiles.add(filetocompare);
        }
      }
    }
    
    for (String temp1 : uniquefiles) {
      System.out.println("unique file  is " + temp1);
    }
    

    return uniquefiles;
  }
  



  public ArrayList<String> readFilename(String folderPath)
    throws Exception
  {
    folder = new File(folderPath);
    File[] filelist = folder.listFiles();
    System.out.println("No of Xmlx files " + filelist.length);
    for (i = 0; i < filelist.length; i += 1) {
      al1.add(filelist[i].getName().substring(0, filelist[i].getName().lastIndexOf(".")));
    }
    
    hs.addAll(al1);
    al1.clear();
    al1.addAll(hs);
    for (i = 0; i < al1.size(); i += 1) {
      System.out.println((String)al1.get(i));
    }
    return al1;
  }
  

  public Map<String, List<String>> readExcelFile(String filePath, int SheetIndex)
    throws Exception
  {
    FileInputStream in = new FileInputStream(filePath);
    File file = new File(filePath);
    ArrayList<String> header1 = new ArrayList();
    ArrayList<String> data = new ArrayList();
    
    ArrayList<String> h = new ArrayList();
    ArrayList<String> d = new ArrayList();
    
    Map<String, List<String>> exceldatalist = new HashMap();
    
    if ((file.isFile()) && (file.exists()))
    {
      wb = new XSSFWorkbook(in);
      xs = wb.getSheetAt(SheetIndex);
      
      int totalRows = xs.getLastRowNum();
      int hcell = 0;
      int dcell = 1;
      
      for (int row = 0; row < totalRows; row++) {
        xrow = xs.getRow(row);
        if (xrow != null) {
          xhcell = xrow.getCell(hcell);
          xdcell = xrow.getCell(dcell);
          if (xhcell != null) {
            header1 = cellvalueAString(xhcell, header1);
          }
          if (xdcell != null) {
            data = cellvalueAString(xdcell, data);
          }
        }
      }
    }
    exceldatalist.put("headerlist", header1);
    exceldatalist.put("datalist", data);
    
    in.close();
    return exceldatalist;
  }
  


  private ArrayList<String> cellvalueAString(Cell cell, ArrayList<String> arr1)
    throws IOException
  {
    String cellvalue = "";
    switch (cell.getCellType())
    {
    case 1: 
      cellvalue = cell.getStringCellValue();
      arr1.add(cellvalue);
      break;
    


    case 0: 
      cellvalue = Integer.toString((int)cell.getNumericCellValue());
      arr1.add(cellvalue);
      break;
    

    case 3: 
      cellvalue = " ";
      arr1.add(cellvalue);
    }
    
    
    return arr1;
  }
}
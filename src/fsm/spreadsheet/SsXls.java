package fsm.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import fsm.common.Log;
import fsm.spreadsheet.SsCell.Type;

public class SsXls extends SsFile
{
   @Override
   public int sheetNameToIndex(String name)
   {
      return wb_.getSheetIndex(name);
   }
   
   @Override
   public String sheetIndexToName(int index)
   {
      return wb_.getSheetName(index);
   }

   @Override
   public void openSheet(int sheetIndex) throws Exception
   {
      sheet_ = wb_.getSheetAt(sheetIndex); 
      if ( sheet_ == null )
      {
         throw new Exception("Sheet index " + sheetIndex + " not found");
      }
   }

   @Override
   public void close()
   {
      try
      {
         if ( wb_ != null )
         {
            wb_.close();
         }
      }
      catch (IOException e)
      {
         Log.severe("Problems closing", e);
      }
      finally
      {
         wb_ = null;
         formulaEvaluator_ = null;
      }      
   }
   
   // --- PROTECTED
   
   // "C:\\Users\\Kurt\\Documents\\Kurt\\eclipse.workspace\\ExcelToolkitScratch\\temp.xlsx"
   SsXls(String filename)
   { 
      super(new File(filename));
   }
   
   @Override
   protected void readImp(FileInputStream fis) throws FileNotFoundException
   {
      try
      {
         wb_ = new HSSFWorkbook(fis);

         // Evaluate cell type
         formulaEvaluator_= wb_.getCreationHelper().createFormulaEvaluator();
      }
      catch (IOException e)
      {
         Log.severe("Failed to read XLS " + getFile().getAbsolutePath(), e);
         try
         {
            close();
         }
         catch (Exception e1)
         {
         }
      }      
   }
   
   @Override
   protected void writeImp(FileOutputStream fos) throws FileNotFoundException
   {
      try
      {
         wb_.write(fos);
      }
      catch (IOException e)
      {
         Log.severe("Failed to read XLS " + getFile().getAbsolutePath(), e);
         try
         {
            close();
         }
         catch (Exception e1)
         {
         }
      }      
   }
   

   @Override
   protected void getTableImp(int[] rows, int[] cols, SsCell[][] table) throws Exception
   {      
      for ( int row=0; row < rows.length; row++ )
      {
         HSSFRow  ssRow = sheet_.getRow(rows[row]);
         if ( ssRow == null )
         {
            continue;
         }
         for ( int col=0; col < cols.length; col++ )
         {
            Cell cell = ssRow.getCell(cols[col]);
            if ( cell == null )
            {
               continue;
            }
            String stringValue = "";
            double numericValue = 0.0;
            Type type = Type.READONLY;
            switch (formulaEvaluator_.evaluateInCell(cell).getCellTypeEnum()) 
            { 
            case STRING: 
               type = Type.STRING;
               stringValue = cell.getStringCellValue();
               break;
            case NUMERIC:
               type = Type.NUMERIC;
               numericValue = cell.getNumericCellValue();
               stringValue = Double.toString(numericValue);
               break;
            case BOOLEAN: 
               type = Type.NUMERIC;
               numericValue = cell.getBooleanCellValue()?1.0:0.0;
               stringValue = Double.toString(numericValue);
            case BLANK:
               type = Type.NUMERIC;
               numericValue = 0.0;
               stringValue = "";
            case ERROR:
            case FORMULA:
            case _NONE:
            default:
               break; 
            } 
            table[row][col].fill(stringValue, numericValue, type);
         }
      }
      return;
   }

   @Override
   public void setTableImp(int[] rows, int[] cols, SsCell[][] cells) throws Exception
   {
      if ( sheet_ == null )
      {
         throw new Exception("Sheet needs to be opened first");
      }
      for ( int row=0; row < rows.length; row++ )
      {
         HSSFRow ssRow = sheet_.getRow(rows[row]);
         if ( ssRow == null )
         {
            continue;
         }
         for ( int col=0; col < cols.length; col++ )
         {
            Cell cell = ssRow.getCell(cols[col]);
            if ( cell == null )
            {
               continue;
            }
            if (cells[row][col].getType() == SsCell.Type.NUMERIC )
            {
               cell.setCellValue(cells[row][col].getValue());
            }
         }
      }
      XSSFFormulaEvaluator.evaluateAllFormulaCells(wb_);
   }
   
   @Override
   public void duplicateSheetImp(String newName) throws Exception
   {     
      if ( sheet_ == null )
      {
         throw new Exception("Sheet needs to be opened first");
      }
      HSSFSheet sheet = wb_.cloneSheet(wb_.getSheetIndex(sheet_.getSheetName()));
      if ( sheet == null )
      {
         throw new Exception("Failed to duplicate sheet");
      }
      
      String sheetName = sheet.getSheetName();
      wb_.setSheetOrder(sheetName, 0);
      int offset = 0;
      String newSheetName = new String(newName);
      while (true)
      {
         if ( offset == 0 )
         {
            if ( wb_.getSheet(newSheetName) == null )
            {
               break;
            }
         }
         else
         {
            newSheetName = newName + "-" + offset;
            if ( wb_.getSheet(newSheetName) == null )
            {
               break;
            }
         }
         offset++;
      }
      wb_.setSheetName(0, newSheetName);
      wb_.setFirstVisibleTab(0);
      wb_.setSelectedTab(0);
      wb_.setActiveSheet(0);
   }
   @Override
   public void setDateCellImp(String cellLocation, Calendar date) throws Exception
   {       
      if ( sheet_ == null )
      {
         throw new Exception("Sheet needs to be opened first");
      }
      CellReference cr = new CellReference(cellLocation);
      HSSFRow row = sheet_.getRow(cr.getRow());
      HSSFCell cell = row.getCell(cr.getCol());
      double d = DateUtil.getExcelDate(date, false); //get double value f
      cell.setCellValue((int)d);  //get int value of the double
      //cell.setCellValue(date);
      XSSFFormulaEvaluator.evaluateAllFormulaCells(wb_);
   }
   
   // --- PRIVATE
   
   private HSSFWorkbook wb_;
   private HSSFSheet sheet_;
   private FormulaEvaluator formulaEvaluator_;
}

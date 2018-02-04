package fsm.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import fsm.common.Log;
import fsm.spreadsheet.SsCell.Type;

public class SsXlsx extends SsFile
{
   @Override
   public void open() throws FileNotFoundException
   {
      try
      {
         fis_ = new FileInputStream(getFile());
         wb_ = new XSSFWorkbook(fis_);

         // Evaluate cell type
         formulaEvaluator_= wb_.getCreationHelper().createFormulaEvaluator();
      }
      catch (IOException e)
      {
         Log.severe("Failed to read XLSX " + getFile().getAbsolutePath(), e);
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
   public boolean isOk()
   {
      return ( formulaEvaluator_ != null );
   }
   
   @Override
   protected int sheetNameToIndex(String name)
   {
      return wb_.getSheetIndex(name);
   }
   
   @Override
   protected String sheetIndexToName(int index)
   {
      return wb_.getSheetName(index);
   }

   @Override
   public void close() throws Exception
   {
      try
      {
         wb_.close();
         fis_.close();
      }
      catch (IOException e)
      {
      }
      finally
      {
         fis_ = null;
         wb_ = null;
         formulaEvaluator_ = null;
      }
      
   }
   
   // --- PROTECTED
   
   // "C:\\Users\\Kurt\\Documents\\Kurt\\eclipse.workspace\\ExcelToolkitScratch\\temp.xlsx"
   SsXlsx(String filename)
   { 
      super(new File(filename));
   }

   @Override
   protected void readTableImp(int sheetIndex, int[] rows, int[] cols, SsCell[][] table)
   {
      XSSFSheet sheet = wb_.getSheetAt(sheetIndex); 
      
      for ( int row=0; row < rows.length; row++ )
      {
         XSSFRow  ssRow = sheet.getRow(rows[row]);
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
   public void writeTableImp(int sheet, int[] rows, int[] cols, SsCell[][] cells_)
   {
      // TODO Implement writeTableImp()     
   }
   
   // --- PRIVATE
   
   private FileInputStream fis_;
   private XSSFWorkbook wb_;
   private FormulaEvaluator formulaEvaluator_;

}

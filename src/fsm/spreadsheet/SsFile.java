package fsm.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;

import fsm.common.Log;

public abstract class SsFile implements AutoCloseable
{
   public static SsFile Create(String filename)
   {
      if ( filename.toLowerCase().endsWith("xlsx") )
      {         
         return new SsXlsx(filename);
      }
      else
      {
         return new SsXls(filename);
      }
   }
   
   public SsFile(File file)
   {
      file_ = file;
      
      // If file doesn't exist, potentially bad path, then try the file at the current working directory
      // - user.dir is updated in Window.processNewParameterFile() 
      if ( !file_.exists() )
      {
         String workingDir = System.getProperty("user.dir", ".");
         file_ = new File (workingDir + File.separator + file_.getName());
      }
   }
   
   public File getFile()
   {
      return file_;
   }
   
   public void read() throws FileNotFoundException, IOException
   {
      try ( FileInputStream fis = new FileInputStream(file_) )
      {
         readImp(fis);
         fis.close();
      }
   }
   
   public void write() throws FileNotFoundException, IOException
   {
      try ( FileOutputStream fos = new FileOutputStream(file_) )
      {
         writeImp(fos);
         fos.close();
      }
   }
   
   public abstract int sheetNameToIndex(String name);
   public abstract String sheetIndexToName(int index);
   public abstract void openSheet(int sheetIndex) throws Exception;
   public abstract int getNumberOfRows() throws Exception;
   public abstract  String[] getSheets();
   public abstract String getSheetName();
   
   public void getTable(SsTable ssTable)
   {
      int[] rows = ssTable.getRowsInSheet();
      int[] cols = ssTable.getColsInSheet();
      ssTable.cells_ = new SsCell[rows.length][cols.length];
      for ( int row=0; row < rows.length; row++ )
      {
         for ( int col=0; col < cols.length; col++ )
         {
            ssTable.cells_[row][col] = new SsCell();
         }
      }
      try
      {
         getTableImp(rows, cols, ssTable.cells_);
      }
      catch (Exception e)
      {
         Log.severe("Error reading excel", e);
      }
      return;
   }
   
   public void setTable(SsTable ssTable)
   {
      try
      {
         setTableImp(ssTable.getRowsInSheet(), ssTable.getColsInSheet(), ssTable.cells_);
      }
      catch (Exception e)
      {
         Log.severe("Error writing excel", e);
      }
      return;      
   }
   
   public void setDateCell(String cellLocation, Calendar date)
   {
      try
      {
         setDateCellImp(cellLocation, date);
      }
      catch (Exception e)
      {
         Log.severe("Error writing excel", e);
      }
   }
   
   public void duplicateSheet(String name)
   {
      try
      {
         duplicateSheetImp(name);
      }
      catch (Exception e)
      {
         Log.severe("Error duplicating sheet excel", e);
      }
      return;          
   }
   
   public abstract void close();
   
   
   // --- PROTECTED

   protected abstract void readImp(FileInputStream fis) throws FileNotFoundException;
   protected abstract void writeImp(FileOutputStream fos) throws FileNotFoundException;

   protected abstract void getTableImp(int[] row,
                                       int[] col,
                                       SsCell[][] cells) throws Exception;

   protected abstract void setTableImp(int[] rows,
                                       int[] cols,
                                       SsCell[][] cells) throws Exception;
   public abstract void setDateCellImp(String cellLocation, Calendar date) throws Exception;
   public abstract void duplicateSheetImp(String name) throws Exception;
   
   // --- PRIVATE
   
   private File file_;

}

package fsm.spreadsheet;

import java.io.File;
import java.io.FileNotFoundException;

import fsm.common.Log;

public abstract class SsFile
{
   public static SsFile Create(String filename)
   {
      return new SsXlsx(filename);
   }
   
   public SsFile(File file)
   {
      file_ = file;
   }
   
   public File getFile()
   {
      return file_;
   }
   
   public SsCell[][] readTable(int sheet, int[] rows, int[] cols)
   {
      SsCell[][] table = new SsCell[rows.length][cols.length];
      for ( int row=0; row < rows.length; row++ )
      {
         for ( int col=0; col < cols.length; col++ )
         {
            table[row][col] = new SsCell();
         }
      }
      try
      {
         open();
         readTableImp(sheet,
                      rows, 
                      cols, 
                      table);
      }
      catch (Exception e)
      {
         Log.severe("Error reading excel", e);
      }
      finally
      {
         try
         {
            close();
         }
         catch (Exception e)
         {
         }
      }
      return table;
   }
   
   public SsCell[][] readTable(String sheet, int[] rows, int[] cols)
   {
      SsCell[][] table = new SsCell[rows.length][cols.length];
      for ( int row=0; row < rows.length; row++ )
      {
         for ( int col=0; col < cols.length; col++ )
         {
            table[row][col] = new SsCell();
         }
      }
      if ( !isOk() ) return table;
      try
      {
         open();
         readTableImp(sheetNameToIndex(sheet),
                      rows, 
                      cols, 
                      table);
      }
      catch (Exception e)
      {
      }
      finally
      {
         try
         {
            close();
         }
         catch (Exception e)
         {
         }
      }
      return table;
   }
   public void writeTable(int[] rows, int[] cols, SsCell[][] cells_)
   {
      
   }
   
   
   // --- PROTECTED

   protected abstract boolean isOk();
   protected abstract void open() throws FileNotFoundException;
   protected abstract int sheetNameToIndex(String name);
   protected abstract void close() throws Exception;

   protected abstract void readTableImp(int sheetIndex,
                                        int[] row,
                                        int[] col,
                                        SsCell[][] cells);

   protected abstract void writeTableImp(int sheetIndex,
                                         int[] rows,
                                         int[] cols,
                                         SsCell[][] cells);
   
   // --- PRIVATE
   
   private File file_;

}

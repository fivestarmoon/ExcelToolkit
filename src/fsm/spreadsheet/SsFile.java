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
      readSheet_ = "";
   }
   
   public void setFileModifiedListener(SsFileModifiedListener l)
   {
      fileModifiedListener_ = l;
   }
   
   public File getFile()
   {
      return file_;
   }
   
   public String getReadSheet()
   {
      return readSheet_;
   }
   
   public SsCell[][] readTable(int sheet, SsTable ssTable)
   {
      int[] rows = ssTable.getRowsInSheet();
      int[] cols = ssTable.getColsInSheet();
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
         if ( fileMonitor_ != null )
         {
            fileMonitor_.stop();
         }
         fileMonitor_ = new FileMonitor(file_, fileModifiedListener_);
         open();
         readSheet_ = this.sheetIndexToName(sheet);
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
   
   public SsCell[][] readTable(String sheet, SsTable ssTable)
   {
      int[] rows = ssTable.getRowsInSheet();
      int[] cols = ssTable.getColsInSheet();
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
         if ( fileMonitor_ != null )
         {
            fileMonitor_.stop();
         }
         fileMonitor_ = new FileMonitor(file_, fileModifiedListener_);
         open();
         readSheet_ = sheet;
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
   
   public void stopFileListener() throws Exception
   {
      if ( fileMonitor_ != null )
      {
         fileMonitor_.stop();
      }
   }
   
   
   // --- PROTECTED

   protected abstract boolean isOk();
   protected abstract void open() throws FileNotFoundException;
   protected abstract int sheetNameToIndex(String name);
   protected abstract String sheetIndexToName(int index);
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
   
   private static class FileMonitor implements Runnable
   {
      public FileMonitor(File file, SsFileModifiedListener l)
      {
         file_ = file;
         fileModifiedListener_ = l;
         if ( fileModifiedListener_ != null )
         {
            Thread thread = new Thread(this);
            thread.setDaemon(true);
            thread.start();
         }
      }
      @Override
      public void run()
      {
         long timeUpdated = file_.lastModified();
         while ( true)
         {
            if ( fileModifiedListener_ == null )
            {
               break;
            }
            long newModified = file_.lastModified();
            if ( newModified != timeUpdated )
            {
               break;
            }
            try
            {
               Thread.sleep(3000);
            }
            catch (InterruptedException e)
            {
            }
         }
         SsFileModifiedListener fileModifiedListener = fileModifiedListener_;
         if ( fileModifiedListener != null )
         {
            fileModifiedListener.fileModified();
         }
         
      }  
      public void stop()
      {
         fileModifiedListener_ = null;
      }    
      private File file_;
      private SsFileModifiedListener fileModifiedListener_;
   }
   
   private File file_;
   private String readSheet_;
   private SsFileModifiedListener fileModifiedListener_;
   private FileMonitor fileMonitor_;

}

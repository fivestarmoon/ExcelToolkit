package fsm.excelToolkit.loading;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.common.utils.FileModifiedListener;
import fsm.common.utils.FileModifiedMonitor;
import fsm.excelToolkit.hmi.table.TableSpreadsheet.TableException;
import fsm.excelToolkit.loading.LoadingMonth.LoadingMonthException;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

class LoadingSpreadSheet implements FileModifiedListener
{      
   /**
    * @param weekEndingCell 
    * 
    */
   LoadingSpreadSheet(
      LoadingSummaryPanel panel, 
      Reader reader,
      String[] resources,
      boolean enableAutoReload)
   {    
      status_ = "JSON error";
      readSheetName_ = "-";    
      enableAutoReload_ = enableAutoReload;
      isModified_ = false;
      try
      {
         parentPanel_ = panel;
         reader.setVerbose(false);
         
         filename_ = reader.getStringValue("path", "");
         if ( filename_.length() == 0 ) throw new Exception("Missing parameter 'path'");
         
         name_ = reader.getStringValue("name", "");
         if ( name_.length() == 0 ) throw new Exception("Missing parameter 'name'");
         
         sheet_ = reader.getStringValue("sheet", "");
         if ( sheet_.length() == 0 ) throw new Exception("Missing parameter 'sheet'");
         
         startRow_ = -1 + (int)reader.getLongValue("startRow", 0);
         if ( startRow_ < 0 ) throw new Exception("Missing parameter 'startRo'w");
         
         endRow_ = -1 + (int)reader.getLongValue("endRow", 0);
         if ( endRow_ < 0 ) throw new Exception("Missing parameter 'endRow'");
         
         startCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("startCol", "0"));
         if ( startCol_ < 0 ) throw new Exception("Missing parameter 'startCol'");
         
         endCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("endCol", "0"));
         if ( endCol_ < 0 ) throw new Exception("Missing parameter 'endCol'");
         
         if ( (endCol_-startCol_+1) <= 0 )
         {
            throw new Exception("Missing parameter 'endCol' must be >= 'startCol'");
         }
         
         try
         {
            startColDate_ = new LoadingMonth(reader.getStringValue("startColDate", ""));            
         }
         catch (LoadingMonthException e)
         {
            throw new TableException("Failed to read 'startColDate'", e);
         }

         resourceMapping_ = new HashMap<String,ArrayList<String>>();
         for ( String label : resources )
         {
            resourceMapping_.put(label, new ArrayList<String>());
         }
         int numAdded = 0;
         Reader[] readers = reader.structArray("resource");
         for ( Reader resReader : readers )
         {
            String name = resReader.getStringValue("name", "");
            String label = resReader.getStringValue("label", "");
            ArrayList<String> nameRes = resourceMapping_.get(label);
            if ( nameRes == null )
            {
               Log.info("Spreadsheet resource '" + name + "' did not match any 'resourceLabel'[]");
               continue;
            }
            nameRes.add(name);
            numAdded++;
         }
         if ( numAdded == 0 ) throw new Exception("Missing parameter resource");
         
         resourceCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("resourceCol", "0"));
         if ( resourceCol_ < 0 ) throw new Exception("Missing parameter 'resourceCol'");
         
         status_ = "Loading...";
         readSheetName_ = "";
      }
      catch ( Exception e )
      {
         filename_ = null;
         Log.severe("Failed to parse JSON structure from 'shreadsheets'[]", e);
      }
   }
   
   public String getFilename()
   {
      return filename_;
   }

   public String getName()
   {
      return name_;
   }

   public String getStatus()
   {
      return status_;
   }

   public boolean isTableValid()
   {
      return table_ != null;
   }
   
   public boolean isFileModified()
   {
      return isModified_;
   }

   public ArrayList<SsCell[]> getLoading(String res, LoadingMonth[] months)
   {
      // Sum up all the rows into spreadsheet and global resources
      ArrayList<SsCell[]> hmiCellArray = new ArrayList<SsCell[]>();
      
      // Get the labels for the named resource
      ArrayList<String> labels = resourceMapping_.get(res);
      if ( labels == null || labels.size() == 0 )
      {
         return hmiCellArray;
      }
      
      for ( int row : table_.getRowIterator() )
      {
         SsCell[] cells = table_.getCellsForRow(row);
         
         // Find the row for the named resource (loop for multiple resources for the same label)
         for ( String label : labels )
         {
            if ( !label.equalsIgnoreCase(cells[0].toString()) )
            {
               continue;
            }
            
            // Create the cells for this resource entry (1 per spreadsheet row)
            SsCell[] hmiCells = new SsCell[HmiColumns.length(months.length)];
            hmiCells[HmiColumns.LABEL.getIndex()] = new SsCell(res);
            hmiCells[HmiColumns.RESOURCE.getIndex()] = new SsCell(cells[0].toString());
            double rowSum = 0.0;
            for ( int ii=0; ii<months.length; ii++ )
            {
               for ( int jj=0; jj<(cells.length-1); jj++ )
               {
                  try
                  {
                  if ( months[ii].equals(startColDate_, jj) )
                  {
                     int cellIdxForMonth = jj+1; //  + 1 to skip resource string
                     rowSum += cells[cellIdxForMonth].getValue();
                     hmiCells[HmiColumns.getMonthIndexToColumnIdx(ii)] = new SsCell(cells[cellIdxForMonth].getValue());
                     break;
                  }
                  }
                  catch (Exception e)
                  {
                     Log.severe("Translating table failed", e);
                  }
               }
               if ( hmiCells[HmiColumns.getMonthIndexToColumnIdx(ii)] == null )
               {
                  hmiCells[HmiColumns.getMonthIndexToColumnIdx(ii)] = new SsCell(0.0);
               }
            }
            hmiCells[HmiColumns.ROW_SUM.getIndex()] = new SsCell(rowSum);
            hmiCellArray.add(hmiCells);
         }
      }
      return hmiCellArray;
   }

   public void destroy()
   {    
      parentPanel_ = null;
      if ( monitor_ != null )
      {
         monitor_.stop();
         monitor_ = null;
         isModified_ = false;
      }
   }

   public void loadBG()
   {
      if ( filename_ == null )
      {
         return;
      }
      table_ = null;
      status_ = "Loading...";
      display(null);
      Thread bgThread = new Thread(new Runnable()
      {
         @Override
         public void run()
         {
            readAndDisplayTable();
         }

      });
      bgThread.setDaemon(true);
      bgThread.start(); 
   }  

   public String getSheetName()
   {
      return readSheetName_;
   }

   @Override
   public void fileModified()
   {
      if ( parentPanel_ == null )
      {
         Log.info("file modified ignored beacuse panel destroyed!");
         return;
      }
      isModified_ = true;
      if ( enableAutoReload_ )
      {
         loadBG();
      }
      else
      {
         display(table_);         
      }
   }

   protected void readAndDisplayTable()
   { 
      SsTable table = null;
      try ( SsFile file = SsFile.Create(filename_) )
      {       
         if ( monitor_ != null )
         {
            monitor_.stop();
            monitor_ = null;
         }
         isModified_ = false;
         file.read();
         file.openSheet(file.sheetNameToIndex(sheet_));
         
         int startRow = startRow_;
         int endRow = endRow_;

         table = new SsTable();
         table.addRow(startRow,  endRow);
         table.addCol(resourceCol_);
         table.addCol(startCol_, endCol_);
         file.getTable(table);     
         file.close();
         status_ = "";
         readSheetName_ = sheet_;
      }
      catch (Exception e)
      {
         Log.severe("Failed read spreadsheet", e);
         status_ = "Spreadsheet error";
         table = null;
      }
      display(table);
   }

   private void display(final SsTable table)
   {
      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         @Override
         public void run()
         {
            table_ = table;
            LoadingSummaryPanel panel = LoadingSpreadSheet.this.parentPanel_;
            if ( panel != null )
            {
               panel.displaySpreadSheet(filename_, LoadingSpreadSheet.this);
               if ( table != null )
               {
                  monitor_ = new FileModifiedMonitor(new File(filename_), LoadingSpreadSheet.this);
                  isModified_ = false;
               }
            }
         }         
      });
   }
   
   private LoadingSummaryPanel parentPanel_;

   private String name_;
   private String filename_;
   private String sheet_;
   private int    startRow_;
   private int    endRow_;
   private int    startCol_;
   private int    endCol_;
   private int    resourceCol_;
   private LoadingMonth startColDate_;
   private SsTable table_;

   private String status_;
   private String readSheetName_;
   private HashMap<String,ArrayList<String>> resourceMapping_;   
   private FileModifiedMonitor monitor_;
   private boolean enableAutoReload_;
   private boolean isModified_;
}
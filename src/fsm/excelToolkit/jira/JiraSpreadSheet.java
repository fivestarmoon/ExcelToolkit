package fsm.excelToolkit.jira;

import java.io.File;

import fsm.common.Log;
import fsm.common.utils.FileModifiedListener;
import fsm.common.utils.FileModifiedMonitor;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

class JiraSpreadSheet implements FileModifiedListener
{ 
   JiraSpreadSheet(
      JiraSummaryPanel panel, 
      String file,
      String resourceCol,
      String chargeCodeCol,
      String budgetCol,
      String actualCol,
      String estimateCol,
      double timeConversion)
   {            
      parentPanel_ = panel;
      filename_ = file;
      resourceCol_= SsCell.ColumnLettersToIndex(resourceCol);
      chargeCodeCol_ = SsCell.ColumnLettersToIndex(chargeCodeCol);
      budgetCol_ = SsCell.ColumnLettersToIndex(budgetCol);
      actualCol_ = SsCell.ColumnLettersToIndex(actualCol);
      etcCol_ =  SsCell.ColumnLettersToIndex(estimateCol); 
      timeConversion_ = timeConversion;
      status_ = "Loading...";
      readSheetName_ = "";
      readSheets_ = new String[0];
      startRow_ = -1;
      endRow_ = -1;
   }

   public SsTable getTable()
   {
      return table_;
   }

   public String getFileName()
   {
      try 
      {
         File file = new File(filename_);
         if ( file != null )
         {
            return file.getName();
         }
      }
      catch (Exception e)
      {         
      }
      return "bad";
   }

   public String getStatus()
   {
      return status_;
   }

   public void destroy()
   {    
      parentPanel_ = null;
      if ( monitor_ != null )
      {
         monitor_.stop();
         monitor_ = null;
      }
   }

   public void loadBG()
   {
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

   public void setSheetName(String sheet)
   {
      readSheetName_ = sheet;
   }
   
   public String[] getSheets()
   {      
      return readSheets_;
   }

   @Override
   public void fileModified()
   {
      if ( parentPanel_ == null )
      {
         Log.info("file modified ignored beacuse panel destroyed!");
         return;
      }
      loadBG();
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
         file.read();
         if ( readSheetName_.length() == 0 )
         {
            file.openSheet(0);
            readSheetName_ = file.sheetIndexToName(sheetOffset_);
         }
         else
         {
            file.openSheet(file.sheetNameToIndex(readSheetName_));
         }
         readSheets_ = file.getSheets();
         
         startRow_ = -1;
         String shortName = file.getFile().getName();
         {
            table = new SsTable();
            table.addRow(0, file.getNumberOfRows()-1);
            table.addCol(resourceCol_);
            file.getTable(table);
            endRow_ = table.getRowsInSheet().length;
            for ( int row : table.getRowIterator() )
            {
               SsCell[] cells = table.getCellsForRow(row);
               if ( cells[0].toString().contains("Assignee") )
               {
                  startRow_ =  table.getRowsInSheet()[row] + 1;
               }
               if ( cells[0].toString().contains("Generated at") )
               {
                  endRow_ = table.getRowsInSheet()[row] - 1;
               }
            }
            Log.info(shortName + " auto detect rows " + (startRow_+1) + " to " + (endRow_+1));
         }
         if ( startRow_ == -1 )
         {
            throw new Exception("Failed to find start row where resourceCol column contains 'Assignee'");
         }
         if ( endRow_ == -1 )
         {
            throw new Exception("Failed to find start row where first column contains 'Generated at'");
         }
         table = new SsTable();
         table.addRow(startRow_,  endRow_);
         table.addCol(resourceCol_);
         table.addCol(chargeCodeCol_);
         table.addCol(budgetCol_);
         table.addCol(actualCol_);
         table.addCol(etcCol_);
         file.getTable(table);     
         file.close();
         status_ = "";

         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            cells[SsColumns.BUDGET.getIndex()].update(
               cells[SsColumns.BUDGET.getIndex()].getValue()/timeConversion_);
            cells[SsColumns.ACTUAL.getIndex()].update(
               cells[SsColumns.ACTUAL.getIndex()].getValue()/timeConversion_);
            cells[SsColumns.ETC.getIndex()].update(
               cells[SsColumns.ETC.getIndex()].getValue()/timeConversion_);
         }
      }
      catch (Exception e)
      {
         Log.severe("Failed read spreadsheet", e);
         status_ = "error";
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
            JiraSpreadSheet.this.parentPanel_.displaySpreadSheet();
            if ( table != null )
            {
               monitor_ = new FileModifiedMonitor(new File(filename_), JiraSpreadSheet.this);
            }
         }         
      });
   }
   
   private JiraSummaryPanel parentPanel_;
   private String filename_;
   private int    sheetOffset_;
   private int    startRow_;
   private int    endRow_;
   private int    resourceCol_;
   private int    chargeCodeCol_;
   private int    budgetCol_;
   private int    actualCol_;
   private int    etcCol_; 
   private double timeConversion_;
   
   private SsTable table_;
   private String  status_;
   private String  readSheetName_;   
   private String[] readSheets_;
   private FileModifiedMonitor monitor_;
}
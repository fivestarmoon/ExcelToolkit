package fsm.excelToolkit.jira;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.common.utils.FileModifiedListener;
import fsm.common.utils.FileModifiedMonitor;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsCell.Type;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

class ActualsSpreadSheet implements FileModifiedListener
{ 
   ActualsSpreadSheet(
      JiraSummaryPanel panel,
      Reader           reader,
      String[]         jiraResources)
   {            
      parentPanel_ = panel;
      filename_ = reader.getStringValue("file", "");
      sheetOffset_ = 0;
      projCodeCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("projCodeCol", "A"));
      chargeCodeCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("chargeCodeCol", "A"));
      resourceCol_= SsCell.ColumnLettersToIndex(reader.getStringValue("resourceCol", "A"));
      actualCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("actualCol", "A"));
      resourceMapping_ = new HashMap<String,ArrayList<String>>();
      for ( String jiraResource : jiraResources )
      {
         resourceMapping_.put(jiraResource, new ArrayList<String>());
      }
      
      // Allow more than one resource to map to the JIRA resource
      Reader[] readers = reader.structArray("resource");
      for ( Reader resReader : readers )
      {
         String actualResource = resReader.getStringValue("actualResource", "");
         String jiraResource = resReader.getStringValue("jiraResource", "");
         ArrayList<String> jiraRes = resourceMapping_.get(jiraResource);
         if ( jiraRes == null )
         {
            Log.info("Actual spreadsheet resource '" + actualResource + "' did not match any 'resourceOrder'[]");
            continue;
         }
         jiraRes.add(actualResource);
      }
      
      status_ = "Loading...";
      readSheetName_ = "";
      readSheets_ = new String[0];
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
   
   public double getActuals(SsTable actualsTable, String jiraChargeCodeIn, String jiraResource)
   {          
      // Trim trailing zero from the JIRA (not required for infoshare)
      StringBuilder jiraChargeCode = new StringBuilder(jiraChargeCodeIn.trim());
      for ( int ii=jiraChargeCode.length()-1; ii>= 0; ii-- )
      {
         if (jiraChargeCode.charAt(ii) == '0')
         {
            jiraChargeCode.setCharAt(ii, ' ');
         }
         else
         {
            break;
         }
      }
      jiraChargeCodeIn = jiraChargeCode.toString().trim();  
      
      // Attempt to find the charge code and resource combination in the spreadsheet
      for ( int row : actualsTable.getRowIterator() )
      {
         // Create the charge code ????-????
         SsCell[] cells = actualsTable.getCellsForRow(row);
         String projCodeStr = cells[PROJ_CODE_COL].toString();
         if ( cells[PROJ_CODE_COL].getType() == Type.NUMERIC )
         {
            projCodeStr = Integer.toString((int)(cells[PROJ_CODE_COL].getValue()));
         }
         String chargeCodeStr = cells[CHARGE_CODE_COL].toString();
         if ( cells[CHARGE_CODE_COL].getType() == Type.NUMERIC )
         {
            chargeCodeStr = Integer.toString((int)(cells[CHARGE_CODE_COL].getValue()));
         }
         String actualChargeCode = (projCodeStr + "-" + chargeCodeStr).trim();
         
         
         // Match charge code
         if ( !jiraChargeCodeIn.equalsIgnoreCase(actualChargeCode) )
         {
            continue;
         }
         
         // Match resource, where multiple resources are allowed to map to a single JIRA resource
         boolean resourceMatch = false;
         ArrayList<String> jiraRes = resourceMapping_.get(jiraResource);
         for ( String res : jiraRes )
         {
            if ( res.equalsIgnoreCase(cells[RESOURCE_COL].toString()) )
            {
               resourceMatch = true;
               break;
            }
         }         
         if ( !resourceMatch )
         {
            continue;
         }
         
         // Match, return the actuals
         return cells[ACTUAL_COL].getValue();
         
      }
      
      // Not found, return zero
      return 0;
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
         table = new SsTable();
         table.addRow(0,  file.getNumberOfRows()-1);
         table.addCol(projCodeCol_);   // 0 PROJ_CODE_COL
         table.addCol(chargeCodeCol_); // 1 CHARGE_CODE_COL
         table.addCol(resourceCol_);   // 2 RESOURCE_COL
         table.addCol(actualCol_);     // 3 ACTUAL_COL
         file.getTable(table);     
         file.close();
         status_ = "";
      }
      catch (Exception e)
      {
         Log.severe("Failed read spreadsheet", e);
         status_ = "error";
         table = null;
      }
      display(table);
   }
   
   // --- PRIVATE

   private void display(final SsTable table)
   {
      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         @Override
         public void run()
         {
            table_ = table;
            ActualsSpreadSheet.this.parentPanel_.displaySpreadSheet();
            if ( table != null )
            {
               monitor_ = new FileModifiedMonitor(new File(filename_), ActualsSpreadSheet.this);
            }
         }         
      });
   }
   
   private static final int PROJ_CODE_COL = 0;
   private static final int CHARGE_CODE_COL = 1;
   private static final int RESOURCE_COL = 2;
   private static final int ACTUAL_COL = 3;
   
   private JiraSummaryPanel parentPanel_;
   private String filename_;
   private int    sheetOffset_;
   private int    projCodeCol_;
   private int    chargeCodeCol_;
   private int    resourceCol_;
   private int    actualCol_;
   private HashMap<String,ArrayList<String>> resourceMapping_;   
   
   private SsTable table_;
   private String  status_;
   private String  readSheetName_;   
   private String[] readSheets_;
   private FileModifiedMonitor monitor_;
}
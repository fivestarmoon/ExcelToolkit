package fsm.excelToolkit;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.common.utils.FileModifiedListener;
import fsm.common.utils.FileModifiedMonitor;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsCell.Type;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

public class ActualsSpreadSheet implements FileModifiedListener
{ 
   public ActualsSpreadSheet(
      TableSpreadsheet panel,
      Reader           reader,
      String[]         parentResources,
      String           parentResourceKey)
   {            
      parentPanel_ = panel;
      filename_ = reader.getStringValue("file", "");
      if ( reader.isKeyForValue("assumeSheetNameMatchJira") )
      {
         assumeSheetNameMatchParent_ = reader.getBooleanValue("assumeSheetNameMatchJira", false);
      }
      if ( reader.isKeyForValue("assumeSheetNameMatchParent") )
      {
         assumeSheetNameMatchParent_ = reader.getBooleanValue("assumeSheetNameMatchParent", false);
      }
      sheetOffset_ = 0;
      projCodeCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("projCodeCol", "A"));
      chargeCodeCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("chargeCodeCol", "A"));
      resourceCol_= SsCell.ColumnLettersToIndex(reader.getStringValue("resourceCol", "A"));
      actualCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("actualCol", "A"));
      resourceMapping_ = new HashMap<String,ArrayList<String>>();
      for ( String parentResource : parentResources )
      {
         resourceMapping_.put(parentResource, new ArrayList<String>());
      }
      
      // Allow more than one resource to map to the parent's resource
      Reader[] readers = reader.structArray("resource");
      for ( Reader resReader : readers )
      {
         String actualResource = resReader.getStringValue("actualResource", "");
         String parentResource = resReader.getStringValue(parentResourceKey, "");
         ArrayList<String> parentRes = resourceMapping_.get(parentResource);
         if ( parentRes == null )
         {
            Log.info("Actual spreadsheet resource '" + actualResource + "' did not match any 'resourceOrder'[]");
            continue;
         }
         parentRes.add(actualResource);
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
   
   public boolean getAssumeSheetNameMatchParent()
   {
      return assumeSheetNameMatchParent_;
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
   
   public double getActuals(SsTable actualsTable, String parentChargeCodeIn, String parentResource)
   {          
      // Trim trailing zero from the parent charge code (not required for infoshare)
      String temp = parentChargeCodeIn.trim().replaceAll("0+$", "");  
      parentChargeCodeIn = temp;
      
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
         
         // Trim trailing zeros
         temp = actualChargeCode.trim().replaceAll("0+$", "");  
         actualChargeCode = temp;
         
         
         // Match charge code
         if ( !parentChargeCodeIn.equalsIgnoreCase(actualChargeCode) )
         {
            continue;
         }
         
         // Match resource, where multiple resources are allowed to map to a single parent resource
         boolean resourceMatch = false;
         ArrayList<String> parentRes = resourceMapping_.get(parentResource);
         if ( parentRes != null )
         {
            for ( String res : parentRes )
            {
               if ( res.equalsIgnoreCase(cells[RESOURCE_COL].toString()) )
               {
                  resourceMatch = true;
                  break;
               }
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
            try
            {
               file.openSheet(file.sheetNameToIndex(readSheetName_));
            }
            catch (Exception e)
            {
               Log.severe("Did not find sheet [" + readSheetName_ + "] in the actuals spreadsheet");
               file.openSheet(0);
               readSheetName_ = file.sheetIndexToName(sheetOffset_);
            }
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
            TableSpreadsheet parentPanel = ActualsSpreadSheet.this.parentPanel_;
            if ( parentPanel != null )
            {
               parentPanel.displaySpreadSheet();
            }
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
   
   private TableSpreadsheet parentPanel_;
   private String filename_;
   private boolean assumeSheetNameMatchParent_;
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
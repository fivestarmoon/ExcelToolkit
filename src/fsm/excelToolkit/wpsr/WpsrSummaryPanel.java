package fsm.excelToolkit.wpsr;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JButton;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellButton;
import fsm.excelToolkit.hmi.table.TableCellLabel;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsFileModifiedListener;
import fsm.spreadsheet.SsTable;

@SuppressWarnings("serial")
public class WpsrSummaryPanel extends TableSpreadsheet
{
   public WpsrSummaryPanel()
   {
      ssReferences_ = new ArrayList<String>();
      spreadsheets_ = new HashMap<String, LocalSpreadSheet>();
   }

   @Override
   public void createPanelBG()
   {  
      String[] columnTitle = new String[HmiColumns.length()];
      HmiColumns[] columns = HmiColumns.values();
      for ( int ii=0; ii<HmiColumns.length(); ii++ )
      {
         columnTitle[ii] = columns[ii].getColumnTitle();
      }
      setColumns(columnTitle, false);
      setColumnPrefferredSize(0, 100);

      // Load the default values for columns and rows if available
      Reader reader = getParameters().getReader();
      String rootDir = "";
      if ( reader.isKeyForValue("ROOTDIR") )
      {
         rootDir = reader.getStringValue("ROOTDIR", "");
      }
      if ( reader.isKeyForValue("resourceColDefault") )
      {
         SsColumns.RESOURCE.setDefaultValue(reader.getStringValue("resourceColDefault", "A"));
      }
      if ( reader.isKeyForValue("budgetColDefault") )
      {
         SsColumns.BUDGET.setDefaultValue(reader.getStringValue("budgetColDefault", "A"));
      }
      if ( reader.isKeyForValue("prevActualColDefault") )
      {
         SsColumns.PREVACTUAL.setDefaultValue(reader.getStringValue("prevActualColDefault", "A"));
      }
      if ( reader.isKeyForValue("thisActualColDefault") )
      {
         SsColumns.THISACTUAL.setDefaultValue(reader.getStringValue("thisActualColDefault", "A"));
      }
      if ( reader.isKeyForValue("etcColDefault") )
      {
         SsColumns.ETC.setDefaultValue(reader.getStringValue("etcColDefault", "A"));
      }
      int sheetOffsetDefault = 0;
      if ( reader.isKeyForValue("etcColDefault") )
      {
         sheetOffsetDefault = (int)reader.getLongValue("sheetOffsetDefault", 0);
      }
      int startRowDefault = 0;
      if ( reader.isKeyForValue("etcColDefault") )
      {
         startRowDefault = (int)reader.getLongValue("startRowDefault", 0);
      }
      int endRowDefault = 0;
      if ( reader.isKeyForValue("etcColDefault") )
      {
         endRowDefault = (int)reader.getLongValue("endRowDefault", 0);
      }

      Reader[] readers = getParameters().getReader().structArray("wpsrs");
      for ( Reader wpsrReader : readers )
      {
         LocalSpreadSheet ss = new LocalSpreadSheet(
            wpsrReader, 
            rootDir,
            sheetOffsetDefault,
            startRowDefault,
            endRowDefault);
         ssReferences_.add(ss.filename_);
         spreadsheets_.put(ss.filename_, ss);
      }
      for ( String ref : ssReferences_ )
      {
         spreadsheets_.get(ref).load();
      }
   }

   // --- PRIVATE

   private void displaySpreadSheet(String ssReference, LocalSpreadSheet sheet)
   {
      // Determine the resource names to sum against
      ArrayList<String> uniqueResource = new ArrayList<String>();
      for ( String ref : ssReferences_ )
      {
         LocalSpreadSheet ss = spreadsheets_.get(ref);
         SsTable table = ss.table_;
         if ( table == null )
         {
            continue;
         }
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            SsCell.AddUnique(cells[SsColumns.RESOURCE.getIndex()], uniqueResource);
         }         
      }

      // Create the cross spreadsheet summing resources
      Resource[] globalResources = new  Resource[uniqueResource.size()];
      for ( int ii=0; ii<globalResources.length; ii++ )
      {
         globalResources[ii] = new Resource(uniqueResource.get(ii));
      }
      Resource globalTotal = new Resource("GLOBAL TOTAL");

      // Process the valid spreadsheets
      removeAllRows();
      startAddRows();
      for ( String ref : ssReferences_ )
      {
         LocalSpreadSheet ss = spreadsheets_.get(ref);          
         SsTable table = ss.table_;
         ss.resetResource();

         // Add the control row
         TableCell[] controls = new TableCell[HmiColumns.length()];
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            controls[ci] =  new TableCellLabel("");
         }
         controls[0] =  new TableCellLabel(ss.label_);
         if ( table != null )
         {
            controls[1] =  new TableCellLabel(ss.table_.getFile().getReadSheet());
         }
         controls[2] =  new TableCellLabel(ss.status_);
         controls[2].setItalics(true);
         JButton reloadB = new JButton("Edit ...");
         reloadB.addActionListener(new ActionListener() 
         { 
            public void actionPerformed(ActionEvent e) 
            { 
               new Thread(new Runnable()
               {
                  @Override
                  public void run()
                  {
                     try
                     {
                        Desktop.getDesktop().open(new File(ss.filename_));
                     }
                     catch (IOException e1)
                     {
                        Log.severe("Could not open file with desktop", e1);
                     }
                  }
               }).start();
            } 
         });
         controls[HmiColumns.length()-1] =  new TableCellButton(reloadB);
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            controls[ci].setBlendBackgroundColor(ssColor_);
         }
         addRowOfCells(controls);

         //  Add the spreadsheet rows  if the table is valid
         if ( table == null )
         {
            continue;
         }

         // Create the summing resources for the spreadsheet
         Resource[] ssResources = new  Resource[uniqueResource.size()];
         for ( int ii=0; ii<ssResources.length; ii++ )
         {
            ssResources[ii] = new Resource(uniqueResource.get(ii));
         }

         // Sum up all the rows into spreadsheet and global resources
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            for ( Resource resource : ssResources )
            {
               resource.sumif(cells);
            }
            for ( Resource resource : globalResources )
            {
               resource.sumif(cells);
            }
            ss.getResource().sum(cells);
            globalTotal.sum(cells);
         }

         // Process each unique resource against the spreadsheet
         for ( Resource resource : ssResources )
         {
            if ( !resource.isUsed() )
            {
               continue;
            }
            addRowOfCells(resource.getTableCells());            
         }
      }

      // Add the global summing for resources
      {  
         addRowOfCells(getTitleRow("RESOURCE TOTAL", totalColor_));         
         for ( Resource resource : globalResources )
         {
            if ( !resource.isUsed() )
            {
               continue;
            }
            addRowOfCells(resource.getTableCells()); 
         }
      }      

      // Add the overall global summing
      {  
         addRowOfCells(getTitleRow("WPSR TOTAL", totalColor_));            
         for ( String ref : ssReferences_ )
         {
            LocalSpreadSheet ss = spreadsheets_.get(ref);
            if ( ss.table_ == null )
            {
               continue;
            }
            TableCell[] tableCells = ss.getResource().getTableCells();
            tableCells[0].setBlendBackgroundColor(ssColor_);
            addRowOfCells(tableCells);       
         }
      }
      TableCell[] tableCells = globalTotal.getTableCells();
      tableCells[0].setBold(true);
      for ( TableCell cell : tableCells )
      {
         cell.setBlendBackgroundColor(totalColor_);
      }
      addRowOfCells(tableCells); 

      // All rows added
      stopAddRows();

      // Display the panel
      displayPanel();
   }

   private class LocalSpreadSheet implements SsFileModifiedListener
   {      
      LocalSpreadSheet(
         Reader reader,
         String rootDir,
         int sheetOffsetDefault,
         int startRowDefault,
         int endRowDefault)
      {            
         reader.setVerbose(false);
         label_ = reader.getStringValue("label", "notset");
         filename_ = reader.getStringValue("file", "");
         filename_ = filename_.replace("ROOTDIR", rootDir);
         sheetOffset_ = (int)reader.getLongValue("sheetOffset", sheetOffsetDefault);
         startRow_ = -1 + (int)reader.getLongValue("startRow", startRowDefault);
         endRow_ = -1 + (int)reader.getLongValue("endRow", endRowDefault);
         resourceCol_= SsTable.ColumnLettersToIndex(reader.getStringValue(
            SsColumns.RESOURCE.getResourceKey(), 
            SsColumns.RESOURCE.getDefaultValue()));
         budgetCol_ = SsTable.ColumnLettersToIndex(reader.getStringValue(
            SsColumns.BUDGET.getResourceKey(), 
            SsColumns.BUDGET.getDefaultValue()));
         prevActualCol_ = SsTable.ColumnLettersToIndex(reader.getStringValue(
            SsColumns.PREVACTUAL.getResourceKey(), 
            SsColumns.PREVACTUAL.getDefaultValue()));
         thisActualCol_ = SsTable.ColumnLettersToIndex(reader.getStringValue(
            SsColumns.THISACTUAL.getResourceKey(), 
            SsColumns.THISACTUAL.getDefaultValue()));
         etcCol_ =  SsTable.ColumnLettersToIndex(reader.getStringValue(
            SsColumns.ETC.getResourceKey(), 
            SsColumns.ETC.getDefaultValue())); 
         status_ = "Loading...";
         file_ = SsFile.Create(filename_);
         file_.setFileModifiedListener(this);
         ssTotal_ = new Resource(label_);
      }

      public void resetResource()
      {
         ssTotal_ = new Resource(label_);
      }

      public Resource getResource()
      {
         return ssTotal_;
      }

      public void load()
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

      protected void readAndDisplayTable()
      { 
         SsTable table = null;
         try
         {       
            int startRow = startRow_;
            int endRow = endRow_;
            String shortName = new File(filename_).getName();
            if ( startRow == -1 || endRow == -1 )
            {
               table = new SsTable(file_, sheetOffset_);
               table.addRow(0,  512);
               table.addCol(SsTable.ColumnLettersToIndex("C"));
               table.addCol(SsTable.ColumnLettersToIndex("E"));
               table.readTable();
               for ( int row : table.getRowIterator() )
               {
                  SsCell[] cells = table.getCellsForRow(row);
                  if ( cells[0].toString().contains("ACTIVITIES") )
                  {
                     startRow = table.getRowsInSheet()[row] + 1;
                  }
                  if ( cells[1].toString().contains("TOTAL") )
                  {
                     endRow = table.getRowsInSheet()[row] - 1;
                  }
               }
               Log.info(shortName + " auto detect rows " + (startRow+1) + " to " + (endRow+1));
            }
            else
            {
               Log.info(shortName + " loading rows " + (startRow+1) + " to " + (endRow+1));
            }
            table = new SsTable(file_, sheetOffset_);
            table.addRow(startRow,  endRow);
            table.addCol(resourceCol_);
            table.addCol(budgetCol_);
            table.addCol(prevActualCol_);
            table.addCol(thisActualCol_);
            table.addCol(etcCol_);
            table.readTable();
            status_ = "";
         }
         catch (Exception e)
         {
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
               displaySpreadSheet(filename_, LocalSpreadSheet.this);
            }         
         });
      }

      @Override
      public void fileModified()
      {
         load();
      }

      private String label_;
      private String filename_;
      private int    sheetOffset_;
      private int    startRow_;
      private int    endRow_;
      private int    resourceCol_;
      private int    budgetCol_;
      private int    prevActualCol_;
      private int    thisActualCol_;
      private int    etcCol_; 
      private SsTable table_;

      private String status_;
      private SsFile file_;
      private Resource ssTotal_;
   }

   private TableCell[] getTitleRow(String title, Color color)
   {      
      // Add the control row
      TableCell[] controls = new TableCell[HmiColumns.length()];
      for ( int ci=0; ci<HmiColumns.length(); ci++ )
      {
         controls[ci] =  new TableCellLabel("");
      }
      controls[0] = new TableCellLabel(title);
      controls[0].setBold(true);
      for ( int ci=0; ci<HmiColumns.length(); ci++ )
      {
         controls[ci].setBlendBackgroundColor(color);
      }  
      return controls;
   }


   private ArrayList<String> ssReferences_;
   private HashMap<String, LocalSpreadSheet> spreadsheets_;
   private Color ssColor_ = new Color(0, 204, 102);
   private Color totalColor_ = new Color(255, 179, 102);

}

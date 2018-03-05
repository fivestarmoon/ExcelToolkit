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
import fsm.spreadsheet.SsTable;

@SuppressWarnings("serial")
public class WpsrSummaryPanel extends TableSpreadsheet
{
   public WpsrSummaryPanel()
   {
      ssReferences_ = new ArrayList<String>();
      spreadsheets_ = new HashMap<String, WpsrSpreadSheet>();
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
      String weekEndingCell = "";
      if ( reader.isKeyForValue("ROOTDIR") )
      {
         rootDir = reader.getStringValue("ROOTDIR", "");
      }
      if ( reader.isKeyForValue("weekEndingCellDefault") )
      {
         weekEndingCell = reader.getStringValue("weekEndingCellDefault", "");
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
      
      // Get the default order for resources if available
      resourceOrder_ = new String[0];
      if ( reader.isKeyForArrayOfValues("resourceOrder") )
      {
         resourceOrder_ = reader.getStringArray("resourceOrder");
      }

      // Load the information on the WPSR spreadsheets
      Reader[] readers = getParameters().getReader().structArray("wpsrs");
      for ( Reader wpsrReader : readers )
      {
         WpsrSpreadSheet ss = new WpsrSpreadSheet(
            this, wpsrReader, 
            rootDir,
            weekEndingCell,
            sheetOffsetDefault,
            startRowDefault,
            endRowDefault);
         ssReferences_.add(ss.filename_);
         spreadsheets_.put(ss.filename_, ss);
      }
      for ( String ref : ssReferences_ )
      {
         spreadsheets_.get(ref).loadBG();
      }
   }
   
   @Override
   protected void destroyPanel()
   {
      for ( String ref : ssReferences_ )
      {
         WpsrSpreadSheet ss = spreadsheets_.get(ref);
         ss.destroy();
      }
      
   }

   // --- PRIVATE

   void displaySpreadSheet(String ssReference, WpsrSpreadSheet sheet)
   {
      // Determine the resource names to sum against
      ArrayList<String> uniqueResourceTemp = new ArrayList<String>();
      for ( String ref : ssReferences_ )
      {
         WpsrSpreadSheet ss = spreadsheets_.get(ref);
         SsTable table = ss.table_;
         if ( table == null )
         {
            continue;
         }
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            SsCell.AddUnique(cells[SsColumns.RESOURCE.getIndex()], uniqueResourceTemp);
         }         
      }
      
      // Sort the resources based on order
      ArrayList<String> uniqueResource = new ArrayList<String>();
      for ( String resourceOrder : resourceOrder_ )
      {
         int index = uniqueResourceTemp.indexOf(resourceOrder);
         if ( index >= 0 )
         {
            uniqueResource.add(resourceOrder);
            uniqueResourceTemp.remove(index);
         }
      }
      uniqueResource.addAll(uniqueResourceTemp);

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
         WpsrSpreadSheet ss = spreadsheets_.get(ref);          
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
            controls[1] =  new TableCellLabel(ss.getSheetName());
         }
         controls[2] =  new TableCellLabel(ss.status_);
         controls[2].setItalics(true);
         
         JButton nextWeekB = new JButton("Prepare ...");
         nextWeekB.addActionListener(new ActionListener() 
         { 
            public void actionPerformed(ActionEvent e) 
            { 
               ss.nextWeek();
            } 
         });
         controls[HmiColumns.ETC.getIndex()] =  new TableCellButton(nextWeekB);
         
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
         controls[HmiColumns.VARIANCE.getIndex()] =  new TableCellButton(reloadB);
         
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
            WpsrSpreadSheet ss = spreadsheets_.get(ref);
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


   private String[] resourceOrder_;
   private ArrayList<String> ssReferences_;
   private HashMap<String, WpsrSpreadSheet> spreadsheets_;
   private Color ssColor_ = new Color(0, 204, 102);
   private Color totalColor_ = new Color(255, 179, 102);

}

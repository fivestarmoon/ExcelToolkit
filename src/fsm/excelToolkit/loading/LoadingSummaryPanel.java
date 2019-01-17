package fsm.excelToolkit.loading;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JButton;
import javax.swing.JTable;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellButton;
import fsm.excelToolkit.hmi.table.TableCellLabel;
import fsm.excelToolkit.hmi.table.TableCellSpreadsheet;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.excelToolkit.loading.LoadingMonth.LoadingMonthException;
import fsm.excelToolkit.loading.HmiColumns;
import fsm.spreadsheet.SsCell;
import fsm.excelToolkit.loading.Resource;

@SuppressWarnings("serial")
public class LoadingSummaryPanel extends TableSpreadsheet
{
   public LoadingSummaryPanel()
   {
      ssReferences_ = new ArrayList<String>();
      spreadsheets_ = new HashMap<String, LoadingSpreadSheet>();
   }

   @Override
   public void createPanelBG() throws TableException
   { 
      // Load the default values for columns and rows if available
      Reader reader = getParameters().getReader();
      try
      {
         startDate_ = new LoadingMonth(reader.getStringValue("startDate", ""));
      }
      catch (LoadingMonthException e)
      {
         throw new TableException("Failed to read startDate", e);
      }
      try
      {
         endDate_ = new LoadingMonth(reader.getStringValue("endDate", ""));
      }
      catch (LoadingMonthException e)
      {
         throw new TableException("Failed to read endDate", e);
      }
      int numMonths = LoadingMonth.DiffInMonths(endDate_, startDate_);
      if (  numMonths <= 0 )
      {
         throw new TableException(String.format("startDate[%s] must be before endDate[%s]", 
            startDate_.toString(),
            endDate_.toString()));
      }
      
      // Create the loading months
      months_ = new LoadingMonth[numMonths];
      for ( int ii=0; ii<numMonths; ii++ )
      {
         try
         {
            months_[ii] = new LoadingMonth(startDate_, ii);
         }         
         catch ( LoadingMonthException e )
         {
            throw new TableException("Unexpected LoadingMonth class error (software bug)");
         }
      }
      
      // Get the default list of resources
      resourceLabel_ = new String[0];
      if ( reader.isKeyForArrayOfValues("resourceLabel") )
      {
         resourceLabel_ = reader.getStringArray("resourceLabel");
      } 
      if ( resourceLabel_.length == 0 )
      {
         throw new TableException("Missing mandatory parameter resourceLabel");
      }

      // Setup the table
      String[] columnTitle = new String[HmiColumns.length(months_.length)];
      int[] columnPreferredSize = new int[columnTitle.length];
      HmiColumns[] columns = HmiColumns.values();
      for ( int ii=0; ii<columnTitle.length; ii++ )
      {
         columnPreferredSize[ii] = 50;
         if ( !HmiColumns.isMonth(ii) )
         {
            columnTitle[ii] = columns[ii].getColumnTitle();
         }
         else
         {
            try
            {
               columnTitle[ii] = new LoadingMonth(startDate_, HmiColumns.columnIndexToMonthIdx(ii)).toString();
            }            
            catch ( LoadingMonthException e )
            {
               throw new TableException("Unexpected LoadingMonth class error (software bug)");
            }
         }
      }
      setColumns(columnTitle, false);
      columnPreferredSize[HmiColumns.LABEL.getIndex()] = 130;
      columnPreferredSize[HmiColumns.RESOURCE.getIndex()] = 130;
      columnPreferredSize[HmiColumns.ROW_SUM.getIndex()] = 100;
      for ( int ii=0; ii<columnTitle.length; ii++ )
      {
         setColumnPrefferredSize(ii, columnPreferredSize[ii]);
      }
      
      // Read some optional parameters
      warningThreshold_ = reader.getLongValue("warningThreshold", 25);

      // Load the information on the WPSR spreadsheets
      Reader[] readers = getParameters().getReader().structArray("spreadsheets");
      for ( Reader ssReader : readers )
      {
         LoadingSpreadSheet ss = new LoadingSpreadSheet(
            this,
            ssReader, 
            resourceLabel_);
         spreadsheets_.put(ss.getFilename(), ss);
         ssReferences_.add(ss.getFilename());
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
         LoadingSpreadSheet ss = spreadsheets_.get(ref);
         ss.destroy();
      }
      
   }

   @Override
   protected int getAutoResizeMode()
   {
      return JTable.AUTO_RESIZE_OFF; 
   }

   // --- PRIVATE

   void displaySpreadSheet(String ssReference, LoadingSpreadSheet sheet)
   {
      // Number of months to sum
      final int NumberOfMonths = LoadingMonth.DiffInMonths(endDate_, startDate_);
      
      // Create the cross spreadsheet summing resources
      Resource[] globalResources = new  Resource[resourceLabel_.length];
      for ( int ii=0; ii<globalResources.length; ii++ )
      {
         globalResources[ii] = new Resource(resourceLabel_[ii], NumberOfMonths);
      }
      
      Resource globalTotal = new Resource("GLOBAL TOTAL", NumberOfMonths);
      
      // Process the valid spreadsheets
      removeAllRows();
      startAddRows();
      for ( String ref : ssReferences_ )
      {
         LoadingSpreadSheet ss = spreadsheets_.get(ref);  

         // Add the control row
         TableCell[] controls = new TableCell[HmiColumns.length(months_.length)];
         for ( int ci=0; ci<controls.length; ci++ )
         {
            controls[ci] =  new TableCellLabel("");
         }
         controls[0] =  new TableCellLabel(ss.getName());
         if ( ss.isTableValid() )
         {
            controls[1] =  new TableCellLabel(ss.getSheetName());
         }
         else
         {
            controls[1] =  new TableCellLabel(ss.getStatus());
         }
         controls[1].setItalics(true);
         
         JButton editB = new JButton("Edit ...");
         editB.addActionListener(new ActionListener() 
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
                        Desktop.getDesktop().open(new File(ss.getFilename()));
                     }
                     catch (IOException e1)
                     {
                        Log.severe("Could not open file with desktop", e1);
                     }
                  }
               }).start();
            } 
         });
         controls[2] =  new TableCellButton(editB);
         for ( int ci=0; ci<HmiColumns.length(months_.length); ci++ )
         {
            controls[ci].setBlendBackgroundColor(ssColor_);
         }
         addRowOfCells(controls);

         //  Add the spreadsheet rows  if the table is valid
         if ( !ss.isTableValid() )
         {
            continue;
         }
         
         // Process each resource
         for ( int ii=0; ii<resourceLabel_.length; ii++ )
         {
            String res = resourceLabel_[ii];
            ArrayList<SsCell[]> loading =  ss.getLoading(res, months_);
            for (SsCell[] cells : loading )
            {
               addRowOfCells(convertToRow(cells));
               globalResources[ii].sumif(cells);
               globalTotal.sum(cells);
            }
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
            addRowOfCells(convertToRow(resource.getCells())); 
         }
      }      

      // Add the overall global summing
//      TableCell[] tableCells = convertToRow(globalTotal.getCells());
//      tableCells[0].setBold(true);
//      for ( TableCell cell : tableCells )
//      {
//         cell.setBlendBackgroundColor(totalColor_);
//      }
//      addRowOfCells(tableCells); 

      // All rows added
      stopAddRows();

      // Display the panel
      displayPanel();
   }
   
   private TableCell[] convertToRow(SsCell[] cells)
   {
      TableCell[] tableCells = new TableCell[cells.length];
      for ( int ii=0; ii<cells.length; ii++ )
      {
         if ( ii == HmiColumns.LABEL.getIndex()
                  || ii == HmiColumns.RESOURCE.getIndex() )
         {
            tableCells[ii] = new TableCellSpreadsheet(cells[ii]); 
         }
         else
         {
            tableCells[ii] = new TableCellSpreadsheet(Round(cells[ii].getValue())); 
            if ( ii != HmiColumns.ROW_SUM.getIndex()
                     && cells[ii].getValue() > warningThreshold_ )
            {
               tableCells[ii].setBlendBackgroundColor(varianceWarning_);
            }
         }
      }
      return tableCells;
   }

   private TableCell[] getTitleRow(String title, Color color)
   {      
      // Add the control row
      TableCell[] controls = new TableCell[HmiColumns.length(months_.length)];
      for ( int ci=0; ci<controls.length; ci++ )
      {
         controls[ci] =  new TableCellLabel("");
      }
      controls[0] = new TableCellLabel(title);
      controls[0].setBold(true);
      for ( int ci=0; ci<controls.length; ci++ )
      {
         controls[ci].setBlendBackgroundColor(color);
      }  
      return controls;
   }
   
   private static double Round(double value)
   {
      return Math.round(value*100.0) / 100.0;
   }

   private LoadingMonth startDate_;
   private LoadingMonth endDate_;
   private LoadingMonth[] months_;
   private String[] resourceLabel_;
   private ArrayList<String> ssReferences_;
   private HashMap<String, LoadingSpreadSheet> spreadsheets_;
   private double warningThreshold_ = 25;
   private Color ssColor_ = new Color(0, 204, 102);
   private Color totalColor_ = new Color(255, 179, 102);
   private Color varianceWarning_ = new Color(255, 0, 0, 64);

}

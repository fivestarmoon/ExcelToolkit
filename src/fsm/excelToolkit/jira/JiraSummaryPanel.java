package fsm.excelToolkit.jira;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JButton;
import javax.swing.JOptionPane;

import fsm.common.parameters.Reader;
import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellButton;
import fsm.excelToolkit.hmi.table.TableCellLabel;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsTable;

@SuppressWarnings("serial")
public class JiraSummaryPanel extends TableSpreadsheet
{
   public JiraSummaryPanel()
   {
      wpsrsRef_ = new ArrayList<String>();
      wpsrs_ = new HashMap<String, String>();
   }

   @Override
   public void createPanelBG()
   {  
      String[] columnTitle = new String[HmiColumns.length()];
      int[] columnPreferredSize = new int[columnTitle.length];
      HmiColumns[] columns = HmiColumns.values();
      for ( int ii=0; ii<HmiColumns.length(); ii++ )
      {
         columnTitle[ii] = columns[ii].getColumnTitle();
         columnPreferredSize[ii] = 75;
      }
      columnPreferredSize[HmiColumns.RESOURCE.getIndex()] = 175;
      setColumns(columnTitle, columnPreferredSize, false);

      // Load the default values for columns and rows if available
      Reader reader = getParameters().getReader();
      String file = reader.getStringValue("file", "");
      String resourceCol = reader.getStringValue(SsColumns.RESOURCE.getResourceKey(), "A");
      String chargeCodeCol = reader.getStringValue(SsColumns.CHARGECODE.getResourceKey(), "B");
      String budgetCol = reader.getStringValue(SsColumns.BUDGET.getResourceKey(), "C");
      String actualCol = reader.getStringValue(SsColumns.ACTUAL.getResourceKey(), "D");
      String estimateCol = reader.getStringValue(SsColumns.ETC.getResourceKey(), "E");
      double timeConversion = reader.getDoubleValue("timeToDayConversion", 1.0);
      resourceOrder_ = new String[0];
      if ( reader.isKeyForArrayOfValues("resourceOrder") )
      {
         resourceOrder_ = reader.getStringArray("resourceOrder");
      }
      Reader[] readers = getParameters().getReader().structArray("chargeCodes");
      for ( Reader wpsrReader : readers )
      {
         String label = wpsrReader.getStringValue("label", "notset");
         String name = wpsrReader.getStringValue("name", "notset");
         wpsrsRef_.add(label);
         wpsrs_.put(label, name);
      }
      spreadSheet_ =  new JiraSpreadSheet(
         this,
         file,
         resourceCol,
         chargeCodeCol,
         budgetCol,
         actualCol,
         estimateCol,
         timeConversion);
      spreadSheet_.loadBG();
      
      actuals_ = null;
      if ( reader.isKeyForStruct("actuals") )
      {
         actuals_ = new ActualsSpreadSheet(this, reader.struct("actuals"), resourceOrder_);
         actuals_.loadBG();
      }
   }
   
   @Override
   protected void destroyPanel()
   {
      spreadSheet_.destroy();      
   }

   // --- PRIVATE

   void displaySpreadSheet()
   {
      // Determine the resource names to sum against
      ArrayList<String> uniqueResourceTemp = new ArrayList<String>();
      if ( spreadSheet_.getTable() != null )
      {
         for ( int row : spreadSheet_.getTable().getRowIterator() )
         {
            SsCell[] cells = spreadSheet_.getTable().getCellsForRow(row);
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

      // Determine if the is an alternate spreadsheet for actuals
      SsTable actualsTable = null;
      if ( actuals_ != null )
      {
         actualsTable = actuals_.getTable();
      }

      startAddRows();

      // Add some column headers
      TableCell[] headers = new TableCell[HmiColumns.length()];
      for ( int ci=0; ci<HmiColumns.length(); ci++ )
      {
         headers[ci] =  new TableCellLabel("");
         headers[ci].setBlendBackgroundColor(totalColor_);
      }
      headers[0] =  new TableCellLabel(spreadSheet_.getFileName());
      headers[0].setBlendBackgroundColor(totalColor_);
      headers[1] =  new TableCellLabel(spreadSheet_.getSheetName());
      headers[1].setBlendBackgroundColor(totalColor_);
      if ( spreadSheet_.getTable() == null )
      {
         headers[2] =  new TableCellLabel(spreadSheet_.getStatus());
      }
      else
      {  
         // Show button to change the sheet that is loaded for JIRA
         JButton sheetB = new JButton("Change sheet");
         sheetB.addActionListener(new ActionListener() 
         { 
            public void actionPerformed(ActionEvent e) 
            { 
               String input = (String) JOptionPane.showInputDialog(
                  null, 
                  "Choose now...",
                  "Select sheet", 
                  JOptionPane.QUESTION_MESSAGE, 
                  null,
                  spreadSheet_.getSheets(), 
                  spreadSheet_.getSheetName()); 
               if ( input != null )
               {
                  spreadSheet_.setSheetName(input);
                  spreadSheet_.loadBG();
               }
            } 
         });
         headers[2] =  new TableCellButton(sheetB);
      }
      headers[2].setItalics(true);
      headers[2].setBlendBackgroundColor(totalColor_);
      addRowOfCells(headers);
      if ( actuals_ != null )
      {
         headers = new TableCell[HmiColumns.length()];
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            headers[ci] =  new TableCellLabel("");
            headers[ci].setBlendBackgroundColor(totalColor_);
         }
         headers[0] =  new TableCellLabel(actuals_.getFileName());
         headers[0].setBlendBackgroundColor(totalColor_);
         headers[1] =  new TableCellLabel(actuals_.getSheetName());
         headers[1].setBlendBackgroundColor(totalColor_);
         if ( actuals_.getTable() == null )
         {
            headers[2] =  new TableCellLabel(actuals_.getStatus());
         }
         else
         {     
            // Show button to change the sheet that is loaded for actuals     
            JButton sheetB = new JButton("Change sheet");
            sheetB.addActionListener(new ActionListener() 
            { 
               public void actionPerformed(ActionEvent e) 
               { 
                  String input = (String) JOptionPane.showInputDialog(
                     null, 
                     "Choose now...",
                     "Select sheet", 
                     JOptionPane.QUESTION_MESSAGE, 
                     null,
                     actuals_.getSheets(), 
                     actuals_.getSheetName()); 
                  if ( input != null )
                  {
                     actuals_.setSheetName(input);
                     actuals_.loadBG();
                  }
               } 
            });
            headers[2] =  new TableCellButton(sheetB);
         }
         headers[2].setItalics(true);
         headers[2].setBlendBackgroundColor(totalColor_);
         addRowOfCells(headers);
      }

      // Process the WPSRs (i.e., charge codes)
      Resource[] wpsrTotal = new Resource[wpsrsRef_.size()];
      for ( int wpsrIndex=0; wpsrIndex<wpsrsRef_.size(); wpsrIndex++ )
      {      
         String wpsrLabel = wpsrsRef_.get(wpsrIndex);
         String wpsrName =  wpsrs_.get(wpsrLabel);
         SsTable table = spreadSheet_.getTable();
         wpsrTotal[wpsrIndex] = new Resource(wpsrName);

         // Add the control row
         TableCell[] controls = new TableCell[HmiColumns.length()];
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            controls[ci] =  new TableCellLabel("");
         }
         controls[0] =  new TableCellLabel(wpsrLabel + " " + wpsrName);         
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

         // Create the summing resources for the charge code
         Resource[] ssResources = new  Resource[uniqueResource.size()];
         for ( int ii=0; ii<ssResources.length; ii++ )
         {
            ssResources[ii] = new Resource(uniqueResource.get(ii));
         }

         // Sum up all the rows into spreadsheet and global resources
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            if ( !wpsrLabel.equals(cells[SsColumns.CHARGECODE.getIndex()].toString()) )
            {
               continue;
            }
            
            // Set actuals to zero if using the alternate actual spreadsheet
            if ( actuals_ != null )
            {
               cells[SsColumns.ACTUAL.getIndex()].update(0.0);
            }
            
            for ( Resource resource : ssResources )
            {
               resource.sumif(cells);
            }
            for ( Resource resource : globalResources )
            {
               resource.sumif(cells);
            }
            wpsrTotal[wpsrIndex].sum(cells);
            globalTotal.sum(cells);
         }

         // Process each unique resource against the spreadsheet
         for ( Resource resource : ssResources )
         {
            if ( !resource.isUsed() )
            {
               continue;
            }
            
            // Override actuals if actuals are provided
            if ( actualsTable != null )
            {
               // Get the actuals for the current charge code and resource
               double actuals = actuals_.getActuals(actualsTable, wpsrLabel, resource.getName());
               
               // Set the actuals for the resource in this charge code
               resource.setActuals(actuals);
               
               // Add the WPSR (charge code) total
               wpsrTotal[wpsrIndex].addActuals(actuals);
               
               // Add to the overall loading for this resource
               for ( Resource gRes : globalResources )
               {
                  if ( gRes.getName().equalsIgnoreCase(resource.getName()) )
                  {
                     gRes.addActuals(actuals);
                  }
               }
               
               // Add to the global sum
               globalTotal.addActuals(actuals);
            }
            
            addRowOfCells(resource.getTableCells());            
         
         } // for ( resource )
         
      } // for ( wpsr or charge code)
      
      // Find assigned time with no charge code
      Resource missingCC = new Resource("No Charge Code");
      SsTable table = spreadSheet_.getTable();
      if ( table != null )
      {
         // Add the control row
         TableCell[] controls = new TableCell[HmiColumns.length()];
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            controls[ci] =  new TableCellLabel("");
         }
         controls[0] =  new TableCellLabel(missingCC.getName());         
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            controls[ci].setBlendBackgroundColor(ssColor_);
         }
         addRowOfCells(controls);
         
         Resource[] ssResources = new  Resource[uniqueResource.size()];
         for ( int ii=0; ii<ssResources.length; ii++ )
         {
            ssResources[ii] = new Resource(uniqueResource.get(ii));
         }
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);

            boolean missing = true;
            for ( int wpsrIndex=0; wpsrIndex<wpsrsRef_.size(); wpsrIndex++ )
            {      
               String wpsrLabel = wpsrsRef_.get(wpsrIndex);
               if ( wpsrLabel.equals(cells[SsColumns.CHARGECODE.getIndex()].toString()) )
               {
                  missing = false;
                  break;
               }
            }
            if ( !missing )
            {
               continue;
            }
            
            for ( Resource resource : ssResources )
            {
               resource.sumif(cells);
            }
            for ( Resource resource : globalResources )
            {
               resource.sumif(cells);
            }
            missingCC.sum(cells);
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
         for ( int wpsrIndex=0; wpsrIndex<wpsrsRef_.size(); wpsrIndex++ )
         {
            TableCell[] tableCells = wpsrTotal[wpsrIndex].getTableCells();
            tableCells[0].setBlendBackgroundColor(ssColor_);
            addRowOfCells(tableCells);  
         }
         TableCell[] tableCells = missingCC.getTableCells();
         tableCells[0].setBlendBackgroundColor(ssColor_);
         addRowOfCells(tableCells); 
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
   private JiraSpreadSheet spreadSheet_;
   private ActualsSpreadSheet actuals_;
   private ArrayList<String> wpsrsRef_;
   private HashMap<String, String> wpsrs_;
   private Color ssColor_ = new Color(0, 204, 102);
   private Color totalColor_ = new Color(255, 179, 102);

}

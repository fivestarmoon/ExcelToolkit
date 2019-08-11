package fsm.excelToolkit.jira;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JButton;
import javax.swing.JOptionPane;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.excelToolkit.ActualsSpreadSheet;
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
      chargeCodeRef_ = new ArrayList<String>();
      chargeCode_ = new HashMap<String, String>();
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
      varianceColorRedThreshold_ = reader.getDoubleValue("varianceColorRedThreshold", 0.1);
      double timeConversion = reader.getDoubleValue("timeToDayConversion", 1.0);
      resourceOrder_ = new String[0];
      if ( reader.isKeyForArrayOfValues("resourceOrder") )
      {
         resourceOrder_ = reader.getStringArray("resourceOrder");
      }
      Reader[] readers = getParameters().getReader().structArray("chargeCodes");
      for ( Reader chargeCodeReader : readers )
      {
         String label = chargeCodeReader.getStringValue("label", "notset");
         String name = chargeCodeReader.getStringValue("name", "notset");
         chargeCodeRef_.add(label);
         chargeCode_.put(label, name);
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
         actuals_ = new ActualsSpreadSheet(
            this, 
            reader.struct("actuals"), 
            resourceOrder_,
            "jiraResource");
         actuals_.loadBG();
      }
   }

   public void displaySpreadSheet()
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

      // Determine if the is an alternate spreadsheet for actuals
      SsTable actualsTable = null;
      if ( actuals_ != null )
      {
         actualsTable = actuals_.getTable();
      }

      startAddRows();

      /////////////////////////////////////////////////////////////////////
      // Add the jira spreadsheet header row with controls
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

                  if ( actuals_.getTable() != null && actuals_.getAssumeSheetNameMatchParent() )
                  {
                     actuals_.setSheetName(input);
                     actuals_.loadBG();
                  }
               }
            } 
         });
         headers[2] =  new TableCellButton(sheetB);
      }
      headers[2].setItalics(true);
      headers[2].setBlendBackgroundColor(totalColor_);
      addRowOfCells(headers);
      
      /////////////////////////////////////////////////////////////////////
      // Add the "actual" header row with controls if required
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
         if ( actuals_.getTable() == null || actuals_.getAssumeSheetNameMatchParent() )
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

      /////////////////////////////////////////////////////////////////////
      // Process each charge code
      Resource globalTotal = new Resource("GLOBAL TOTAL");
      Resource[] chargeCodeTotal = new Resource[chargeCodeRef_.size()];
      for ( int ccIndex=0; ccIndex<chargeCodeRef_.size(); ccIndex++ )
      {      
         String ccLabel = chargeCodeRef_.get(ccIndex);
         String ccName =  chargeCode_.get(ccLabel);
         SsTable table = spreadSheet_.getTable();
         chargeCodeTotal[ccIndex] = new Resource(ccName);

         // Add the control row
         TableCell[] controls = new TableCell[HmiColumns.length()];
         for ( int ci=0; ci<HmiColumns.length(); ci++ )
         {
            controls[ci] =  new TableCellLabel("");
         }
         controls[0] =  new TableCellLabel(ccLabel + " " + ccName);         
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
         Resource[] ccResources = new  Resource[uniqueResource.size()];
         for ( int ii=0; ii<ccResources.length; ii++ )
         {
            ccResources[ii] = new Resource(uniqueResource.get(ii));
         }

         // Sum up all the rows into spreadsheet and global resources
         // - add time to resources for the charge code
         // - add time to charge code summing  and global summing (if matched!)
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            
            // Trim trailing zeros
            String wprsLabelTrim = ccLabel.trim().replaceAll("0+$", "");  
            String jiraChargeCode = cells[SsColumns.CHARGECODE.getIndex()].toString();
            String jiraChargeCodeTrim = jiraChargeCode.trim().replaceAll("0+$", "");  
            if ( !wprsLabelTrim.equalsIgnoreCase(jiraChargeCodeTrim) )
            {
               continue;
            }
            
            // Force the actuals to zero if using the alternate actual spreadsheet
            // - i.e., ingore the actuals logged in jira
            if ( actuals_ != null )
            {
               cells[SsColumns.ACTUAL.getIndex()].update(0.0);
            }

            // Add time to resources for the charge code
            boolean matchingResource = false;
            for ( Resource resource : ccResources )
            {
               if ( resource.sumif(cells) )
               {
                  matchingResource = true;
               }
            }
            
            // Add time to global resources
            for ( Resource resource : globalResources )
            {
               resource.sumif(cells);
            }
            
            // Add time to summing
            if ( matchingResource )
            {
               chargeCodeTotal[ccIndex].sum(cells);
               globalTotal.sum(cells);
            }
         }

         // Process each unique resource against the spreadsheet and process
         // the actual spreadsheet
         for ( Resource resource : ccResources )
         {
            // Resource is not working on charge code, skip
            if ( !resource.isUsed() )
            {
               continue;
            }
            
            // Override actuals if actuals are provided
            if ( actualsTable != null )
            {
               // Get the actuals for the current charge code and resource
               double actuals = actuals_.getActuals(actualsTable, ccLabel, resource.getName());
               
               // Set the actuals for the resource in this charge code
               resource.setActuals(actuals);
               
               // Add the charge code total
               chargeCodeTotal[ccIndex].addActuals(actuals);
               
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
            
            addRowOfCells(resource.getTableCells(varianceColorRedThreshold_));            
         
         } // for ( resource )
         
      } // for ( charge code)

      /////////////////////////////////////////////////////////////////////
      // Find assigned time with no charge code
      Resource missingCC = new Resource("Bad Resource or Charge Code (see log)");
      SsTable table = spreadSheet_.getTable();
      if ( table != null )
      {         
         Resource[] ccResources = new  Resource[uniqueResource.size()];
         for ( int ii=0; ii<ccResources.length; ii++ )
         {
            ccResources[ii] = new Resource(uniqueResource.get(ii));
         }
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            
            if ( cells[SsColumns.BUDGET.getIndex()].getValue() == 0.0
                     && cells[SsColumns.ETC.getIndex()].getValue() == 0.0 )
            {
               continue; // all zero, don't care
            }

            // Validate charge code
            boolean matchingChargeCode = false;
            for ( int ccIndex=0; ccIndex<chargeCodeRef_.size(); ccIndex++ )
            {      
               String ccLabel = chargeCodeRef_.get(ccIndex);
               // Trim trailing zeros
               String ccLabelTrim = ccLabel.trim().replaceAll("0+$", "");  
               String jiraChargeCode = cells[SsColumns.CHARGECODE.getIndex()].toString();
               String jiraChargeCodeTrim = jiraChargeCode.trim().replaceAll("0+$", "");  
               if ( ccLabelTrim.equalsIgnoreCase(jiraChargeCodeTrim) )
               {
                  matchingChargeCode = true;
                  break;
               }
            }
            
            // Validate the resource
            boolean matchingResource = false;
            if ( matchingChargeCode )
            {
               for ( Resource resource : ccResources )
               {
                  if ( resource.sumif(cells) )
                  {
                     matchingResource = true;
                  }
               }
            }
            
            if ( matchingChargeCode && matchingResource )
            {
               continue; // looks good!
            }
            Log.info("Bad JIRA entry (missing resource or charge code); relative row " + row
               + " Resource=" + cells[SsColumns.RESOURCE.getIndex()].toString()
               + " ChargeCode=" + cells[SsColumns.CHARGECODE.getIndex()].toString()
               + " Budget=" + cells[SsColumns.BUDGET.getIndex()].toString()
               + " ETC=" + cells[SsColumns.ETC.getIndex()].toString()
                     );
            missingCC.sum(cells);
            
         } // for row in jira
         
      } // if ( table != null )
      
      // Found some bad jira rows with no resource and/or charge code
      // - add a row with the some being all red
      if ( missingCC.isUsed() )
      {
         // Add missing charge code
         TableCell[] tableCells = missingCC.getTableCells(varianceColorRedThreshold_);
         for ( TableCell cell : tableCells )
         {
            cell.setBlendBackgroundColor(Color.red);
         }
         addRowOfCells(tableCells); 
      }

      /////////////////////////////////////////////////////////////////////
      // Add the global summing for each resource
      {  
         addRowOfCells(getTitleRow("RESOURCE TOTAL", totalColor_));         
         for ( Resource resource : globalResources )
         {
            if ( !resource.isUsed() )
            {
               continue;
            }
            addRowOfCells(resource.getTableCells(varianceColorRedThreshold_)); 
         }
      }      

      /////////////////////////////////////////////////////////////////////
      // Add the overall sum for each charge code
      {  
         addRowOfCells(getTitleRow("CHARGE CODE TOTAL", totalColor_));
         for ( int ccIndex=0; ccIndex<chargeCodeRef_.size(); ccIndex++ )
         {
            TableCell[] tableCells = chargeCodeTotal[ccIndex].getTableCells(varianceColorRedThreshold_);
            tableCells[0].setBlendBackgroundColor(ssColor_);
            addRowOfCells(tableCells);  
         }
      }
      
      /////////////////////////////////////////////////////////////////////
      // Add the overall global summing
      {
         TableCell[] tableCells = globalTotal.getTableCells(varianceColorRedThreshold_);
         tableCells[0].setBold(true);
         for ( TableCell cell : tableCells )
         {
            cell.setBlendBackgroundColor(totalColor_);
         }
         addRowOfCells(tableCells); 
      }

      // All rows added
      stopAddRows();

      // Display the panel
      displayPanel();
   }

   // --- PROTECTED
   
   @Override
   protected void destroyPanel()
   {
      spreadSheet_.destroy();      
   }

   // --- PRIVATE

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
   private ArrayList<String> chargeCodeRef_;
   private HashMap<String, String> chargeCode_;
   private Color ssColor_ = new Color(0, 204, 102);
   private Color totalColor_ = new Color(255, 179, 102);
   private double varianceColorRedThreshold_;

}

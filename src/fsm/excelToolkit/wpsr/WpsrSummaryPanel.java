package fsm.excelToolkit.wpsr;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
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
      String[] columnTitle = new String[Columns.length()];
      Columns[] columns = Columns.values();
      for ( int ii=0; ii<Columns.length(); ii++ )
      {
         columnTitle[ii] = columns[ii].getColumnTitle();
      }
      setColumns(columnTitle, true);
      setColumnPrefferredSize(0, 100);

      Reader[] readers = getParameters().getReader().structArray("wpsrs");
      for ( Reader reader : readers )
      {
         LocalSpreadSheet ss = new LocalSpreadSheet(reader);
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
            SsCell.AddUnique(cells[Columns.RESOURCE.getIndex()], uniqueResource);
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
         TableCell[] controls = new TableCell[Columns.length()];
         for ( int ci=0; ci<Columns.length(); ci++ )
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
         JButton reloadB = new JButton("Reload");
         reloadB.addActionListener(new ActionListener() 
         { 
            public void actionPerformed(ActionEvent e) 
            { 
               ss.load();
             } 
           });
         controls[Columns.length()-1] =  new TableCellButton(reloadB);
         for ( int ci=0; ci<Columns.length(); ci++ )
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
            TableCell[] tableCells = resource.getTableCells();
            addRowOfCells(tableCells);            
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
            TableCell[] tableCells = resource.getTableCells();
            addRowOfCells(tableCells); 
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
      LocalSpreadSheet(Reader reader)
      {             
         ssReference_ = reader.getStringValue("file", "");
         label_ = reader.getStringValue("label", "notset");
         filename_ = reader.getStringValue("file", "");
         sheetOffset_ = (int)reader.getLongValue("sheetOffset", 0);
         startRow_ = -1 + (int)reader.getLongValue("startRow", 0);
         endRow_ = -1 + (int)reader.getLongValue("endRow", 0);
         resourceCol_= SsTable.ColumnLettersToIndex(reader.getStringValue(
            Columns.RESOURCE.getResourceKey(), "A"));
         budgetCol_ = SsTable.ColumnLettersToIndex(reader.getStringValue(
            Columns.BUDGET.getResourceKey(), "A"));
         prevActualCol_ = SsTable.ColumnLettersToIndex(reader.getStringValue(
            Columns.PREVACTUAL.getResourceKey(), "A"));
         thisActualCol_ = SsTable.ColumnLettersToIndex(reader.getStringValue(
            Columns.THISACTUAL.getResourceKey(), "A"));
         etcCol_ =  SsTable.ColumnLettersToIndex(reader.getStringValue(
            Columns.ETC.getResourceKey(), "A")); 
         status_ = "Loading...";
         file_ = SsFile.Create(filename_);
         file_.setFileModifiedListener(this);
         ssTotal_ = new Resource(label_);
         display(null);
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
         int startRow = startRow_;
         int endRow = endRow_;
         if ( startRow == -1 || endRow == -1 )
         {
            SsTable table = new SsTable(file_, sheetOffset_);
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
            Log.info(ssReference_ + " auto detect rows " + (startRow+1) + " to " + (endRow+1));
         }
         else
         {
            Log.info(ssReference_ + " loading rows " + (startRow+1) + " to " + (endRow+1));
         }
         
         SsTable table = new SsTable(file_, sheetOffset_);
         table.addRow(startRow,  endRow);
         table.addCol(resourceCol_);
         table.addCol(budgetCol_);
         table.addCol(prevActualCol_);
         table.addCol(thisActualCol_);
         table.addCol(etcCol_);
         table.readTable();
         status_ = "";
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
               displaySpreadSheet(ssReference_, LocalSpreadSheet.this);
            }         
         });
      }
      
      @Override
      public void fileModified()
      {
         javax.swing.SwingUtilities.invokeLater(new Runnable()
         {
            @Override
            public void run()
            {
               status_ = "modified";
               displaySpreadSheet(ssReference_, LocalSpreadSheet.this);
            }         
         });
      }

      private String ssReference_;
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
      TableCell[] controls = new TableCell[Columns.length()];
      for ( int ci=0; ci<Columns.length(); ci++ )
      {
         controls[ci] =  new TableCellLabel("");
      }
      controls[0] = new TableCellLabel(title);
      controls[0].setBold(true);
      for ( int ci=0; ci<Columns.length(); ci++ )
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

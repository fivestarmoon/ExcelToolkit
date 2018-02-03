package fsm.excelToolkit.hmi;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JButton;
import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellSpreadsheet;
import fsm.excelToolkit.hmi.table.TableCellButton;
import fsm.excelToolkit.hmi.table.TableCellLabel;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsFile;
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
         String file = reader.getStringValue("file", "");
         LocalSpreadSheet ss = new LocalSpreadSheet(
            file, // reference
            reader.getStringValue("label", "notset"),
            file,
            (int)reader.getLongValue("sheetOffset", 0),
            -1 + (int)reader.getLongValue("startRow", 0),
            -1 + (int)reader.getLongValue("endRow", 0),
            SsTable.ColumnLettersToIndex(reader.getStringValue(
               Columns.RESOURCE.getResourceKey(), "A")),
            SsTable.ColumnLettersToIndex(reader.getStringValue(
               Columns.BUDGET.getResourceKey(), "A")),
            SsTable.ColumnLettersToIndex(reader.getStringValue(
               Columns.PREVACTUAL.getResourceKey(), "A")),
            SsTable.ColumnLettersToIndex(reader.getStringValue(
               Columns.THISACTUAL.getResourceKey(), "A")),
            SsTable.ColumnLettersToIndex(reader.getStringValue(
               Columns.ETC.getResourceKey(), "A")));

         ssReferences_.add(file);
         spreadsheets_.put(file, ss);
      }
      for ( String ref : ssReferences_ )
      {
         spreadsheets_.get(ref).load();
      }
   }

   // --- PRIVATE

   private void displaySpreadSheet(String ssReference, LocalSpreadSheet sheet)
   {
      //spreadsheets_.put(ssReference, sheet);
      removeAllRows();
      startAddRows();
      for ( String ref : ssReferences_ )
      {
         LocalSpreadSheet ss = spreadsheets_.get(ref);
         TableCell[] controls = new TableCell[Columns.length()];
         for ( int ci=0; ci<Columns.length(); ci++ )
         {
            controls[ci] =  new TableCellLabel("");
         }
         controls[Columns.RESOURCE.getIndex()] =  new TableCellLabel(ss.label_);
         controls[Columns.BUDGET.getIndex()] =  new TableCellLabel(ss.status_);
         JButton reloadB = new JButton("Reload");
         reloadB.addActionListener(new ActionListener() 
         { 
            public void actionPerformed(ActionEvent e) 
            { 
               ss.load();
             } 
           });
         controls[Columns.VARIANCE.getIndex()] =  new TableCellButton(reloadB);
         for ( int ci=0; ci<Columns.length(); ci++ )
         {
            controls[ci].setBlendBackgroundColor(new Color(230,242,255,64));
         }
         addRowOfCells(controls);
            
         SsTable table = ss.table_;
         if ( table == null )
         {
            continue;
         }
         Log.info(table.toString());
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            TableCell[] tableCells = new TableCell[Columns.length()];
            double variance = cells[Columns.BUDGET.getIndex()].getValue();
            for ( Columns col : Columns.values() )
            {
               int index = col.getIndex();
               if ( col.getResourceKey().length() > 0 )
               {
                  tableCells[index] = new TableCellSpreadsheet(cells[index]); 
                  if ( col == Columns.RESOURCE )
                  {
                     tableCells[index].setBlendBackgroundColor(new Color(255,255,0,64));
                  }     
                  else if ( col == Columns.BUDGET )
                  {
                     tableCells[index].setBold(true);
                  }                             
               }
               else if ( col == Columns.EAC )
               {
                  double sum = 0.0;
                  sum += cells[Columns.PREVACTUAL.getIndex()].getValue();
                  sum += cells[Columns.THISACTUAL.getIndex()].getValue();
                  sum += cells[Columns.ETC.getIndex()].getValue();
                  sum = Math.round(sum*100.0)/100.0;
                  tableCells[index] = new TableCellSpreadsheet(sum); 
                  tableCells[index].setBold(true);
                  variance -= sum;
               }
               else if ( col == Columns.VARIANCE )
               {
                  variance = Math.round(variance*100.0)/100.0;
                  tableCells[index] = new TableCellSpreadsheet(variance);
               }
               else
               {
                  tableCells[index] =  new TableCellSpreadsheet("");
               }
            }
            addRowOfCells(tableCells);
            for ( int col : table.getColIterator() )
            {
               Log.info(String.format("[%d][%d]=%s", row, col, table.getCell(row, col)));             
            }
         }
      }
      stopAddRows();

      // Display the panel
      displayPanel();

   }

   private class LocalSpreadSheet
   {      
      LocalSpreadSheet(String ssReference,
         String label,
         String filename,
         int    sheetOffset,
         int    startRow,
         int    endRow,
         int    resourceCol,
         int    budgetCol,
         int    prevActualCol,
         int    thisActualCol,
         int    etcCol)
      {         
         ssReference_ = ssReference;
         label_ = label;
         filename_ = filename;
         sheetOffset_ = sheetOffset;
         startRow_ = startRow;
         endRow_ = endRow;
         resourceCol_= resourceCol;
         budgetCol_ = budgetCol;
         prevActualCol_ = prevActualCol;
         thisActualCol_ = thisActualCol;
         etcCol_ = etcCol; 
         status_ = "Loading...";
         display(null);
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
               SsFile file = SsFile.Create(filename_);
               SsTable table = new SsTable(file, sheetOffset_);
               table.addRow(startRow_,  endRow_);
               table.addCol(resourceCol_);
               table.addCol(budgetCol_);
               table.addCol(prevActualCol_);
               table.addCol(thisActualCol_);
               table.addCol(etcCol_);
               table.readTable();
               status_ = "";
               display(table);
            }

         });
         bgThread.setDaemon(true);
         bgThread.start(); 
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
   }

   private enum Columns
   {
      RESOURCE(0, "resourceCol", "Resource"),
      BUDGET(1, "budgetCol", "Budget"),
      PREVACTUAL(2, "prevActualCol", "Prev Actual"),
      THISACTUAL(3, "thisActualCol", "Week Actual"),
      ETC(4, "etcCol", "ETC"),
      EAC(5, "", "EAC"),
      VARIANCE(6, "", "Variance");
      
      public static int length()
      {
         return VARIANCE.getIndex()+1;
      }
      public int getIndex()
      {
         return index_;
      }
      public String getResourceKey()
      {
         return resourceKey_;
      }
      public String getColumnTitle()
      {
         return columnTitle_;
      }
      
      
      Columns(int index, String resourceKey, String columnTitle)
      {
         index_ = index;
         resourceKey_ = resourceKey;
         columnTitle_ = columnTitle;
      }
      private int index_;
      private String resourceKey_;
      private String columnTitle_;
   }

   
   private ArrayList<String> ssReferences_;
   private HashMap<String, LocalSpreadSheet> spreadsheets_;

}

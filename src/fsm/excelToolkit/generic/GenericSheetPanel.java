package fsm.excelToolkit.generic;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JButton;

import csm.common.utils.FileModifiedListener;
import csm.common.utils.FileModifiedMonitor;
import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellButton;
import fsm.excelToolkit.hmi.table.TableCellLabel;
import fsm.excelToolkit.hmi.table.TableCellSpreadsheet;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

@SuppressWarnings("serial")
public class GenericSheetPanel extends TableSpreadsheet
{
   public GenericSheetPanel()
   {
      ssReferences_ = new ArrayList<String>();
      spreadsheets_ = new HashMap<String, LocalSpreadSheet>();
   }

   @Override
   public void createPanelBG()
   { 

      // Load the default values for columns and rows if available
      Reader[] readers = getParameters().getReader().structArray("sheets");
      int numColumns = 4;
      for ( Reader sheetReader : readers )
      {
         LocalSpreadSheet ss = new LocalSpreadSheet(sheetReader);
         ssReferences_.add(ss.filename_);
         spreadsheets_.put(ss.filename_, ss);
         numColumns = Math.max(numColumns, ss.getNumColumns());
      } 
      columns_ = new String[numColumns];
      for ( int ii=0; ii<columns_.length; ii++ )
      {
         columns_[ii] = Integer.toString(ii+1);
      }
      setColumns(columns_, true);
      setColumnPrefferredSize(0, 100);
      for ( String ref : ssReferences_ )
      {
         spreadsheets_.get(ref).load();
      }
   }
   
   @Override
   protected void destroyPanel()
   {
      for ( String ref : ssReferences_ )
      {
         LocalSpreadSheet ss = spreadsheets_.get(ref);
         ss.destroy();
      }
      
   }

   // --- PRIVATE

   private void displaySpreadSheet(String ssReference, LocalSpreadSheet sheet)
   {
      // Process the valid spreadsheets
      removeAllRows();
      startAddRows();
      for ( String ref : ssReferences_ )
      {
         LocalSpreadSheet ss = spreadsheets_.get(ref);          
         SsTable table = ss.table_;

         // Add the control row
         TableCell[] controls = new TableCell[columns_.length];
         for ( int ci=0; ci<columns_.length; ci++ )
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
         controls[3] =  new TableCellButton(reloadB);
         for ( int ci=0; ci<columns_.length; ci++ )
         {
            controls[ci].setBlendBackgroundColor(ssColor_);
         }
         addRowOfCells(controls);

         //  Add the spreadsheet rows  if the table is valid
         if ( table == null )
         {
            continue;
         }

         // Add the table cells
         for ( int row : table.getRowIterator() )
         {
            TableCell[] tableCells = new TableCell[columns_.length];
            for ( int ci=0; ci<columns_.length; ci++ )
            {
               tableCells[ci] =  new TableCellLabel("");
            }
            SsCell[] cells = table.getCellsForRow(row);
            for ( int ci=0; ci<cells.length; ci++ )
            {
               tableCells[ci] =  new TableCellSpreadsheet(cells[ci]);
            }
            addRowOfCells(tableCells); 
         }
      }

      // All rows added
      stopAddRows();

      // Display the panel
      displayPanel();
   }

   private class LocalSpreadSheet implements FileModifiedListener
   {      
      LocalSpreadSheet(Reader reader)
      {            
         reader.setVerbose(false);
         ssReference_ = reader.getStringValue("file", "");
         label_ = reader.getStringValue("label", "notset");
         filename_ = reader.getStringValue("file", "");   
         useSheetIndex_ = true;
         sheetIndex_ = 0;
         sheetName_ = "";
         if ( reader.isKeyForValue("sheetIndex") )
         {
            sheetIndex_ = -1 + (int)reader.getLongValue("sheetIndex", 1);
         }
         else
         {
            useSheetIndex_ = false;
            sheetName_ = reader.getStringValue("sheetName", "");
         }
         rows_ = new int[] {0};
         if ( reader.isKeyForArrayOfValues("rows") )
         {
            long[] temp = reader.getLongArray("rows");
            rows_ = new int[temp.length];
            for ( int ii=0; ii<temp.length; ii++ )
            {
               rows_[ii] = -1 + (int) temp[ii];
            }
         }
         else
         {
            int rowStart = -1 + (int)reader.getLongValue("rowStart", 1);
            int rowEnd = -1 + (int)reader.getLongValue("rowEnd", 1);
            rows_ = new int[Math.max(0, rowEnd-rowStart+1)];
            for ( int ii=0; ii<rows_.length; ii++ )
            {
               rows_[ii] = rowStart+ii;
            }
            
         }
         columns_ = new int[] {0};
         if ( reader.isKeyForArrayOfValues("columns") )
         {
            String[] temp = reader.getStringArray("columns");
            columns_ = new int[temp.length];
            for ( int ii=0; ii<temp.length; ii++ )
            {
               columns_[ii] = SsCell.ColumnLettersToIndex(temp[ii]);
            }
         }
         else
         {
            int columnStart = SsCell.ColumnLettersToIndex(reader.getStringValue("columnStart", "A"));
            int columnEnd = SsCell.ColumnLettersToIndex(reader.getStringValue("columnEnd", "A"));
            columns_ = new int[Math.max(0, columnEnd-columnStart+1)];
            for ( int ii=0; ii<columns_.length; ii++ )
            {
               columns_[ii] = columnStart+ii;
            }
            
         }
         status_ = "Loading...";
      }

      public void destroy()
      {   
         ssReference_ = null;
         if ( monitor_ != null )
         {
            monitor_.stop();
            monitor_ = null;
         }
      }

      public String getSheetName()
      {
         return sheetName_;
      }

      public int getNumColumns()
      {
         return columns_.length;
      }

      public void load()
      {
         if ( monitor_ != null )
         {
            monitor_.stop();
            monitor_ = null;
         }
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
         SsTable table = new SsTable();
         for ( int row : rows_ )
         {
            table.addRow(row);
         }
         for ( int column : columns_ )
         {
            table.addCol(column);
         }
         try ( SsFile file = SsFile.Create(filename_))
         {
            file.read();
            if ( useSheetIndex_ )
            {
               file.openSheet(sheetIndex_);
               sheetName_ = file.sheetIndexToName(sheetIndex_);
            }
            else
            {
               file.openSheet(file.sheetNameToIndex(sheetName_));
            }
            file.getTable(table);
            status_ = "";
            file.close();
         }
         catch (Exception e)
         {
            Log.severe("Failed to load sheet", e);
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
               displaySpreadSheet(ssReference_, LocalSpreadSheet.this);
               if ( table_ != null )
               {
                  monitor_ = new FileModifiedMonitor(new File(filename_), LocalSpreadSheet.this);
               }
            }         
         });
      }

      @Override
      public void fileModified()
      {
         if ( ssReference_ == null )
         {
            Log.info("file modified ignored beacuse panel destroyed!");
            return;
         }
         load();
      }

      private String    ssReference_;
      private String    label_;
      private String    filename_;
      private boolean   useSheetIndex_;
      private int       sheetIndex_;
      private String    sheetName_;
      private int[]     rows_;
      private int[]     columns_;
      private SsTable   table_;

      private String status_;
      
      private FileModifiedMonitor monitor_;
   }


   private ArrayList<String> ssReferences_;
   private HashMap<String, LocalSpreadSheet> spreadsheets_;
   private String[] columns_;
   private Color ssColor_ = new Color(0, 204, 102);

}

package fsm.excelToolkit.wpsr;

import java.awt.BorderLayout;
import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.SwingWorker;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.common.swing.DatePickerPanel;
import fsm.common.utils.FileModifiedListener;
import fsm.common.utils.FileModifiedMonitor;
import fsm.excelToolkit.ActualsSpreadSheet;
import fsm.spreadsheet.SsCell;
import fsm.spreadsheet.SsCell.Type;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

class WpsrSpreadSheet implements FileModifiedListener
{      
   /**
    * @param weekEndingCell 
    * @param infoshareColDefault 
    * 
    */
   WpsrSpreadSheet(
      WpsrSummaryPanel panel, 
      Reader reader,
      String rootDir,
      String weekEndingCell, 
      int sheetOffsetDefault,
      String infoshareColDefault, 
      int startInfoshareRowDefault,
      int endInfoshareRowDefault,
      int startRowDefault,
      int endRowDefault)
   {            
      parentPanel_ = panel;
      reader.setVerbose(false);
      label_ = reader.getStringValue("label", "notset");
      filename_ = reader.getStringValue("file", "");
      filename_ = filename_.replace("ROOTDIR", rootDir);
      weekEndingCell_ = reader.getStringValue("weekEndingCell", weekEndingCell);
      sheetOffset_ = (int)reader.getLongValue("sheetOffset", sheetOffsetDefault);
      infoshareChargeCode_ = reader.getStringValue("infoshareChargeCode", "NOTSET");
      infoshareCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue("infoshareCol", infoshareColDefault));
      startInfoshareRow_ = -1 + (int)reader.getLongValue("startInfoshareRow", startInfoshareRowDefault);
      endInfoshareRow_ = -1 + (int)reader.getLongValue("endInfoshareRow", endInfoshareRowDefault);
      startRow_ = -1 + (int)reader.getLongValue("startRow", startRowDefault);
      endRow_ = -1 + (int)reader.getLongValue("endRow", endRowDefault);
      resourceCol_= SsCell.ColumnLettersToIndex(reader.getStringValue(
         SsColumns.RESOURCE.getResourceKey(), 
         SsColumns.RESOURCE.getDefaultValue()));
      budgetCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue(
         SsColumns.BUDGET.getResourceKey(), 
         SsColumns.BUDGET.getDefaultValue()));
      prevActualCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue(
         SsColumns.PREVACTUAL.getResourceKey(), 
         SsColumns.PREVACTUAL.getDefaultValue()));
      thisActualCol_ = SsCell.ColumnLettersToIndex(reader.getStringValue(
         SsColumns.THISACTUAL.getResourceKey(), 
         SsColumns.THISACTUAL.getDefaultValue()));
      etcCol_ =  SsCell.ColumnLettersToIndex(reader.getStringValue(
         SsColumns.ETC.getResourceKey(), 
         SsColumns.ETC.getDefaultValue())); 
      status_ = "Loading...";
      readSheetName_ = "";
      ssTotal_ = new Resource(label_);
      readyForNextWeek_ = false;
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

   public void resetResource()
   {
      ssTotal_ = new Resource(label_);
   }

   public Resource getResource()
   {
      return ssTotal_;
   }

   public void loadBG()
   {
      table_ = null;
      status_ = "Loading...";
      readyForNextWeek_ = false;
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

   public void nextWeek(ActualsSpreadSheet actuals)
   {
      final int startInfoshareRow = startInfoshareRowForNextWeek_;
      final int endInfoshareRow = endInfoshareRowForNextWeek_;
      final int startRow = startRowForNextWeek_;
      final int endRow = endRowForNextWeek_;
      if ( !readyForNextWeek_ )
      {
         // still reading
         return;
      }    
      
      // Get the next Friday
      Calendar date = Calendar.getInstance();
      int dayOffset = Calendar.FRIDAY - date.get(Calendar.DAY_OF_WEEK);
      if ( dayOffset<0 )
      {
         dayOffset += 7;
      }
      date.set(Calendar.DATE, date.get(Calendar.DAY_OF_MONTH)+dayOffset);
      
      JLabel message1 = new JLabel("Please select end-of-week of day:");
      DatePickerPanel datePickerPanel = new DatePickerPanel(date);
      StringBuilder message2Str = new StringBuilder(
         "<html><br>Automatic preparation of WPSR includes:<br><ul>"
         + "<li>Duplicate the first sheet to selected day" + "</li>"
         + "<li>Sum PREV WEEK column and THIS WEEK colum into PREV WEEK column" + "</li>"
         + "<li>Clear THIS WEEK column to zero" + "</li>"
         + "<li>Set End Week cell to selected day" + "</li>");
      if ( actuals != null )
      {
         message2Str.append("<li>Update infoshare actuals" + "</li>");
      }
      message2Str.append("</ul><br>Proceed?</html>");
      JLabel message2 = new JLabel(message2Str.toString());
      
      JPanel panel = new JPanel(); 
      panel.setLayout(new BorderLayout());
      panel.add(message1, BorderLayout.PAGE_START);
      panel.add(datePickerPanel, BorderLayout.CENTER);
      panel.add(message2, BorderLayout.PAGE_END);

      int result = JOptionPane.showConfirmDialog(
         parentPanel_.getWindow(), 
         panel,
         "Prepare WPSR for end-of-week?", 
         JOptionPane.YES_NO_OPTION);
      if (result != JOptionPane.YES_OPTION) 
      {
         return;
      }
      
      final JOptionPane optionPane = new JOptionPane(
         "Preparing WPSR for the next week of reporting ...",
         JOptionPane.INFORMATION_MESSAGE,
         JOptionPane.DEFAULT_OPTION,
         null,
         new Object[]{},
         null);
      final JDialog dialog = new JDialog(
         parentPanel_.getWindow(), 
         "Modifying WPSR", 
         true);
      dialog.setContentPane(optionPane);
      dialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
      dialog.pack();
      dialog.setLocationRelativeTo(parentPanel_.getWindow());
      
      SwingWorker<Void, Void> sw = new SwingWorker<Void, Void>() 
      {
         @Override
         protected Void doInBackground() throws Exception 
         {
            long startTime = System.currentTimeMillis();
            prepForNext(
               datePickerPanel.getDate(), 
               weekEndingCell_, 
               startInfoshareRow,
               endInfoshareRow,
               startRow, 
               endRow, 
               actuals);
            long duration = Math.max(500, 2000 - (System.currentTimeMillis() - startTime));
            Thread.sleep(duration);
            return null;
         }

         @Override
         protected void done() 
         {
          //close the modal dialog
            dialog.dispose();
         }
      };

      sw.execute();
      dialog.setVisible(true); // pause

      fileModified();
      
      int n = JOptionPane.showConfirmDialog(
         parentPanel_.getWindow(),
         "Would like to open the WPSR for editing?",
         "WPSR",
         JOptionPane.YES_NO_OPTION);
      if ( n == JOptionPane.YES_OPTION )
      {

         new Thread(new Runnable()
         {
            @Override
            public void run()
            {
               try
               {
                  Desktop.getDesktop().open(new File(filename_));
               }
               catch (IOException e1)
               {
                  Log.severe("Could not open file with desktop", e1);
               }
            }
         }).start();
         
      }
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
         file.openSheet(sheetOffset_);
         readSheetName_ = file.sheetIndexToName(sheetOffset_);
         
         // Determine the rows with infoshare actuals
         int startInfoshareRow = startInfoshareRow_;
         int endInfoshareRow = endInfoshareRow_;
         String shortName = file.getFile().getName();
         if ( startInfoshareRow == -1 || endInfoshareRow == -1 )
         {
            table = new SsTable();
            table.addRow(0,  512);
            table.addCol(SsCell.ColumnLettersToIndex("B"));
            table.addCol(SsCell.ColumnLettersToIndex("C"));
            file.getTable(table);
            for ( int row : table.getRowIterator() )
            {
               SsCell[] cells = table.getCellsForRow(row);
               if ( cells[0].toString().contains("TOTAL") )
               {
                  endInfoshareRow = table.getRowsInSheet()[row] - 1;
               }
               if ( cells[1].toString().contains("BUDGET") )
               {
                  startInfoshareRow = table.getRowsInSheet()[row] + 1;
               }
            }
            Log.info(shortName + " auto detect infoshare rows " + (startInfoshareRow+1) + " to " + (endInfoshareRow+1));
         }
         else
         {
            Log.info(shortName + " loading infoshare rows " + (startInfoshareRow+1) + " to " + (endInfoshareRow+1));
         }
         
         // Determine the rows with actuals/etc/eac
         int startRow = startRow_;
         int endRow = endRow_;
         if ( startRow == -1 || endRow == -1 )
         {
            table = new SsTable();
            table.addRow(0,  512);
            table.addCol(SsCell.ColumnLettersToIndex("C"));
            table.addCol(SsCell.ColumnLettersToIndex("E"));
            file.getTable(table);
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

         startInfoshareRowForNextWeek_ = startInfoshareRow;
         endInfoshareRowForNextWeek_ = endInfoshareRow;
         startRowForNextWeek_ = startRow;
         endRowForNextWeek_ = endRow;
         table = new SsTable();
         table.addRow(startRow,  endRow);
         table.addCol(resourceCol_);
         table.addCol(budgetCol_);
         table.addCol(prevActualCol_);
         table.addCol(thisActualCol_);
         table.addCol(etcCol_);
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

   protected void prepForNext(Calendar date, 
                              String weekEndingCell, 
                              int startInfoshareRow, 
                              int endInfoshareRow, 
                              int startRow, 
                              int endRow, 
                              ActualsSpreadSheet actuals)
   { 
      try ( SsFile file = SsFile.Create(filename_) )
      {       
         if ( monitor_ != null )
         {
            monitor_.stop();
            monitor_ = null;
         }
         file.read();
         
         SsTable actualsTable = null;
         if ( actuals != null )
         {
            actualsTable = actuals.getTable();
         }
         
         // Duplicate the first sheet
         file.openSheet(0);      
         file.duplicateSheet((new SimpleDateFormat("MMM dd")).format(date.getTime()));  

         // Modify the duplicated sheet
         file.openSheet(0);
         if ( weekEndingCell.length() > 0 )
         {     
            file.setDateCell(weekEndingCell, date);            
         }   
         
         // Modify the infoshare rows
         if ( actualsTable != null )
         {
            SsTable table = new SsTable();
            table.addRow(startInfoshareRow,  endInfoshareRow);
            table.addCol(resourceCol_);
            table.addCol(infoshareCol_);
            file.getTable(table);
            for ( int row : table.getRowIterator() )
            {
               SsCell[] cells = table.getCellsForRow(row);
               double actDays = actuals.getActuals(actualsTable, infoshareChargeCode_, cells[0].toString());
               cells[1].update(actDays);
               if ( cells[0].getType() == Type.NUMERIC )
               {
                  // Handle empty resource cells that get treated as numeric
                  cells[0] = new SsCell("");
               }
            } 
            file.setTable(table);
         }
         
         // Modify the actual/etc/aac rows
         SsTable table = new SsTable();
         table.addRow(startRow,  endRow);
         table.addCol(prevActualCol_);
         table.addCol(thisActualCol_);
         file.getTable(table);
         for ( int row : table.getRowIterator() )
         {
            SsCell[] cells = table.getCellsForRow(row);
            if ( cells[0].getType() == SsCell.Type.NUMERIC )
            {
               cells[0].update(cells[0].getValue() + cells[1].getValue());
            }
            if ( cells[1].getType() == SsCell.Type.NUMERIC )
            {
               cells[1].update(0.0);
            }
         } 
         file.setTable(table);
         
         
         file.write();       
         file.close();
      }
      catch (Exception e)
      {
         Log.severe("Failed to duplicate", e);
         JOptionPane.showMessageDialog(null,
            "Failed to prep " + e.getMessage(),
            "WPSR Prep Error",
            JOptionPane.ERROR_MESSAGE);
      }
   }

   private void display(final SsTable table)
   {
      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         @Override
         public void run()
         {
            table_ = table;
            WpsrSummaryPanel panel = WpsrSpreadSheet.this.parentPanel_;
            if ( panel != null )
            {
               panel.displaySpreadSheet();
               if ( table != null )
               {
                  readyForNextWeek_ = true;
                  monitor_ = new FileModifiedMonitor(new File(filename_), WpsrSpreadSheet.this);
               }
            }
         }         
      });
   }
   
   private WpsrSummaryPanel parentPanel_;

   String label_;
   String filename_;
   String weekEndingCell_;
   private int    sheetOffset_;
   private String infoshareChargeCode_;
   private int    infoshareCol_;
   private int    startInfoshareRow_;
   private int    endInfoshareRow_;
   private int    startRow_;
   private int    endRow_;
   private int    resourceCol_;
   private int    budgetCol_;
   private int    prevActualCol_;
   private int    thisActualCol_;
   private int    etcCol_; 
   SsTable table_;

   String status_;
   private String readSheetName_;
   private Resource ssTotal_;
   private boolean readyForNextWeek_;
   private int    startInfoshareRowForNextWeek_;
   private int    endInfoshareRowForNextWeek_;
   private int    startRowForNextWeek_;
   private int    endRowForNextWeek_;
   
   private FileModifiedMonitor monitor_;
}
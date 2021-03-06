package fsm.excelToolkit.hmi.table;

import java.awt.GridLayout;
import java.util.ArrayList;

import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableCellRenderer;

import fsm.common.Log;
import fsm.common.parameters.Parameters;
import fsm.excelToolkit.hmi.Window;

@SuppressWarnings("serial")
public abstract  class TableSpreadsheet extends JPanel
{
   public static class TableException extends Exception
   {
      public TableException(String exception)
      {
         super(exception);
      }
      public TableException(String string, Exception exception)
      {
         super(string, exception);
      }
   }
   
   public TableSpreadsheet()
   {      
      super(new GridLayout(1,0));
      window_ = null;
      params_ = null;
      table_ = null;
      editable_ = false;
      columns_ = new String[0];
      rows_ = new ArrayList<TableCell[]>();

      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         @Override
         public void run()
         {
            if ( window_ != null )
            {
               add(new JLabel("Loading ..."));
               window_.showContent(TableSpreadsheet.this);
            }
         }         
      });
   }

   final public void createTable(Window window, Parameters params)
   {
      window_ = window;
      params_ = params;
      Thread bgThread = new Thread(new Runnable()
      {

         @Override
         public void run()
         {
            try
            {
               createPanelBG();
            }
            catch (TableException e)
            {
               Log.severe("Failed to create table", e);
            }
         }

      });
      bgThread.setDaemon(true);
      bgThread.start();
   }

   public void destroyTable()
   {
      destroyPanel();
      window_ = null;
   }
   
   public Window getWindow()
   {
      return window_;
   }
   
   public abstract void displaySpreadSheet();
   
   // --- PROTECTED

   protected abstract void createPanelBG() throws TableException;

   protected abstract void destroyPanel();

   final protected Parameters getParameters()
   {
      return params_;
   }

   protected int getAutoResizeMode()
   {
      return JTable.AUTO_RESIZE_LAST_COLUMN;
   }

   final protected void setColumns(String [] columns, int columnSize[], boolean editable)
   {
      editable_ = editable;

      columns_ = columns.clone();
      columnSize_ = columnSize.clone();
      rows_ = new ArrayList<TableCell[]>();
      
      buildTable();
   }

   final protected void startAddRows()
   {
      rows_.clear();
   }

   final protected void addRowOfCells(TableCell[] rowOfCells)
   {
      rows_.add(rowOfCells);
   }

   final protected void stopAddRows()
   {  
      buildTable();
   }

   final protected void displayPanel()
   {
      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         @Override
         public void run()
         {
            if ( window_ != null )
            {
               // Add the scroll pane to this panel.
               removeAll();
               add(new JScrollPane(table_));
               window_.showContent(TableSpreadsheet.this);
            }
         }         
      });
   }

   // --- PRIVATE

   private Window window_;
   private Parameters params_;

   private JTable table_;
   private boolean editable_;
   private String [] columns_;
   private int[] columnSize_;
   private ArrayList<TableCell[]> rows_;
   
   private void buildTable()
   {
      table_ = new MyTable(new MyTableModel());
      table_.setFillsViewportHeight(true);
      table_.setAutoResizeMode(getAutoResizeMode());
      for ( int col=0; col<columnSize_.length; col++ )
      {
         table_.getColumnModel().getColumn(col).setPreferredWidth(columnSize_[col]);
      }
      displayPanel();
      
   }

   @SuppressWarnings("rawtypes")
   private class MyTable extends JTable
   {
      public MyTable(MyTableModel myTableModel)
      {
         super.setModel(myTableModel);
         super.getTableHeader().setReorderingAllowed(false);
      }

      @Override
      public TableCellRenderer getCellRenderer(int row, int col) 
      {
          editingClass_ = rows_.get(row)[col].getColumnClass();
          return rows_.get(row)[col];
      }

      @Override
      public TableCellEditor getCellEditor(int row, int col) 
      {
         if ( rows_.get(row)[col] instanceof TableCellButton )
         {
            TableCellButton cell = (TableCellButton) rows_.get(row)[col];
            return cell;
         }
         return super.getCellEditor(row, col);
      }
      //  This method is also invoked by the editor when the value in the editor
      //  component is saved in the TableModel. The class was saved when the
      //  editor was invoked so the proper class can be created.

      @SuppressWarnings("unchecked")
      @Override
      public Class getColumnClass(int column) 
      {
          return editingClass_;
      }
      
      private Class editingClass_;
   }


   private class MyTableModel extends AbstractTableModel 
   {
      public int getColumnCount() 
      {
         return columns_.length;
      }

      public int getRowCount() 
      {
         return rows_.size();
      }

      public String getColumnName(int col) 
      {
         return columns_[col];
      }

      public Object getValueAt(int row, int col) 
      {
         return rows_.get(row)[col].getValue();
      }

      /*
       * Don't need to implement this method unless your table's
       * editable.
       */
      public boolean isCellEditable(int row, int col)
      {
         if ( rows_.get(row)[col] instanceof TableCellSpreadsheet && !editable_ )
         {
            return false;
         }
         return rows_.get(row)[col].isCellEditable();
      }


   }

}

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

import fsm.common.parameters.Parameters;
import fsm.excelToolkit.hmi.Window;
import fsm.spreadsheet.SsCell;

@SuppressWarnings("serial")
public abstract  class TableSpreadsheet extends JPanel
{
   public TableSpreadsheet()
   {      
      super(new GridLayout(1,0));
      window_ = null;
      params_ = null;
      table_ = null;
      editable_ = false;
      columns_ = new String[0];
      rows_ = new ArrayList<SsCell[]>();
      delayFireRows_ = false;

      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         @Override
         public void run()
         {
            if ( window_ != null )
            {
               add(new JLabel("Loading ..."));
               window_.showTable(TableSpreadsheet.this);
            }
         }         
      });
   }

   public void createTable(Window window, Parameters params)
   {
      window_ = window;
      params_ = params;
      Thread bgThread = new Thread(new Runnable()
      {

         @Override
         public void run()
         {
            createPanel();
         }

      });
      bgThread.setDaemon(true);
      bgThread.start();
   }

   public void setColumns(String [] columns, boolean editable)
   {
      editable_ = editable;

      columns_ = columns.clone();
      rows_ = new ArrayList<SsCell[]>();

      table_ = new JTable(new MyTableModel());
      table_.setFillsViewportHeight(true);
      displayPanel();
   }

   public void startAddRows()
   {
      delayFireRows_ = true;
   }

   public void addRowOfCells(SsCell[] rowOfCells)
   {
      rows_.add(rowOfCells);
      if ( !delayFireRows_ )
      {
         ((MyTableModel)table_.getModel()).fireTableRowsInserted(rows_.size()-1, rows_.size()-1);
      }
   }

   public void stopAddRows()
   {
      if ( !delayFireRows_ )
      {
         ((MyTableModel)table_.getModel()).fireTableRowsInserted(rows_.size()-1, rows_.size()-1);
      }
      delayFireRows_ = false;
   }

   public void destroyTable()
   {
      window_ = null;
   }

   // PROTECTED

   protected abstract void createPanel();

   protected Parameters getParameters()
   {
      return params_;
   }

   protected void displayPanel()
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
               window_.showTable(TableSpreadsheet.this);
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
   private ArrayList<SsCell[]> rows_;
   private boolean delayFireRows_;

   @SuppressWarnings("rawtypes")
   class MyTable extends JTable
   {
      private static final long serialVersionUID = 1L;
      private Class editingClass;

      @Override
      public TableCellRenderer getCellRenderer(int row, int column) 
      {
          editingClass = null;
          int modelColumn = convertColumnIndexToModel(column);
          if (modelColumn == 1) {
              Class rowClass = getModel().getValueAt(row, modelColumn).getClass();
              return getDefaultRenderer(rowClass);
          } 
          else 
          {
              return super.getCellRenderer(row, column);
          }
      }

      @Override
      public TableCellEditor getCellEditor(int row, int column) 
      {
          editingClass = null;
          int modelColumn = convertColumnIndexToModel(column);
          if (modelColumn == 1) {
              editingClass = getModel().getValueAt(row, modelColumn).getClass();
              return getDefaultEditor(editingClass);
          } 
          else 
          {
              return super.getCellEditor(row, column);
          }
      }
      //  This method is also invoked by the editor when the value in the editor
      //  component is saved in the TableModel. The class was saved when the
      //  editor was invoked so the proper class can be created.

      @Override
      public Class getColumnClass(int column) {
          return editingClass != null ? editingClass : super.getColumnClass(column);
      }
      
      private Class editingClass_;
   }


   class MyTableModel extends AbstractTableModel 
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
         return rows_.get(row)[col].toString();
      }

      /*
       * Don't need to implement this method unless your table's
       * editable.
       */
      public boolean isCellEditable(int row, int col)
      {
         if ( !editable_ )
         {
            return false;
         }
         return rows_.get(row)[col].getType() != SsCell.Type.READONLY;
      }


   }

}

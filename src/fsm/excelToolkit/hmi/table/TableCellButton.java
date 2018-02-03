package fsm.excelToolkit.hmi.table;

import java.awt.Component;
import java.util.EventObject;

import javax.swing.JButton;
import javax.swing.JTable;
import javax.swing.event.CellEditorListener;
import javax.swing.table.TableCellEditor;

public class TableCellButton extends TableCell implements TableCellEditor
{
   public TableCellButton(JButton component)
   {
      component_ = component;
      component_.setFocusable(false);
   }

   @Override
   public Object getValue()
   {
      return component_;
   }

   @SuppressWarnings("rawtypes")
   @Override
   public Class getColumnClass()
   {
      return component_.getClass();
   }

   @Override
   public boolean isCellEditable()
   {
      return true;
   }

   @Override
   public Component getTableCellRendererComponent(JTable table,
                                                  Object value,
                                                  boolean isSelected,
                                                  boolean hasFocus,
                                                  int row,
                                                  int column)
   {
      if ( !isSelected )
      {
         component_.setBackground(super.getBackgroundColor(component_.getBackground()));
      }
      component_.setFont(super.getFont(component_.getFont()));
      return component_;
   }

   @Override
   public Object getCellEditorValue()
   {
      return null;
   }

   @Override
   public boolean isCellEditable(EventObject anEvent)
   {
      return true;
   }

   @Override
   public boolean shouldSelectCell(EventObject anEvent)
   {
      return true;
   }

   @Override
   public boolean stopCellEditing()
   {
      return true;
   }

   @Override
   public void cancelCellEditing()
   {
   }

   @Override
   public void addCellEditorListener(CellEditorListener l)
   {      
   }

   @Override
   public void removeCellEditorListener(CellEditorListener l)
   {      
   }

   @Override
   public Component getTableCellEditorComponent(JTable table,
                                                Object value,
                                                boolean isSelected,
                                                int row,
                                                int column)
   {
      return component_;
   }

   // --- PRIVATE

   private JButton component_;

}

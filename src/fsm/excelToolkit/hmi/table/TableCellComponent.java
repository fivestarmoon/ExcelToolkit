package fsm.excelToolkit.hmi.table;

import java.awt.Component;
import java.util.EventObject;

import javax.swing.JTable;
import javax.swing.event.CellEditorListener;

public class TableCellComponent extends TableCell
{
   public TableCellComponent(Component component)
   {
      component_ = component;
   }

   @Override
   public Component getTableCellRendererComponent(JTable table,
                                                  Object value,
                                                  boolean isSelected,
                                                  boolean hasFocus,
                                                  int row,
                                                  int column)
   {
      return component_;
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

   @Override
   public Object getCellEditorValue()
   {
      return null;
   }

   @Override
   public boolean isCellEditable(EventObject anEvent)
   {
      return false;
   }

   @Override
   public boolean shouldSelectCell(EventObject anEvent)
   {
      return false;
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
   
   // --- PRIVATE
   
   private Component component_;

}

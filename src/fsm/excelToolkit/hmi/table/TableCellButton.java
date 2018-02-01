package fsm.excelToolkit.hmi.table;

import java.awt.Component;
import javax.swing.JButton;
import javax.swing.JTable;
import javax.swing.UIManager;

public class TableCellButton extends TableCell
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
      return false;
   }

   @Override
   public Component getTableCellRendererComponent(JTable table,
                                                  Object value,
                                                  boolean isSelected,
                                                  boolean hasFocus,
                                                  int row,
                                                  int column)
   {
      if (isSelected) 
      { 
         component_.setForeground(table.getSelectionForeground());
         component_.setBackground(table.getSelectionBackground());
      } 
      else 
      {
         component_.setForeground(table.getForeground());
         component_.setBackground(UIManager.getColor("Button.background"));
      }
      return component_;
   }

   // --- PRIVATE

   private JButton component_;

}

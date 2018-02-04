package fsm.excelToolkit.hmi.table;

import java.awt.Component;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

public class TableCellLabel extends TableCell
{
   public TableCellLabel(String text)
   {
      text_ = text;
      cellRenderer_ = new DefaultTableCellRenderer();
   }

   @Override
   public Object getValue()
   {
      return text_;
   }

   @SuppressWarnings("rawtypes")
   @Override
   public Class getColumnClass()
   {
      return String.class;
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
      Component c = cellRenderer_.getTableCellRendererComponent(
         table, value, isSelected, hasFocus, row, column);
      if ( !isSelected )
      {
         c.setBackground(super.getBackgroundColor(c.getBackground()));
      }
      c.setFont(super.getFont(c.getFont()));
      return c;
   }
   
   // --- PRIVATE
   
   private String text_;
   private DefaultTableCellRenderer cellRenderer_;

}

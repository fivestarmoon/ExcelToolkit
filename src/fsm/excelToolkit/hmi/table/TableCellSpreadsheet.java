package fsm.excelToolkit.hmi.table;

import java.awt.Component;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;
import fsm.spreadsheet.SsCell;

public class TableCellSpreadsheet extends TableCell
{
   
   public TableCellSpreadsheet(SsCell cell)
   {
      cell_ = cell;
      cellRenderer_ = new DefaultTableCellRenderer();
   }

   public TableCellSpreadsheet(double sum)
   {
      this(new SsCell(sum));
   }

   public TableCellSpreadsheet(String string)
   {
      this(new SsCell(string));
   }

   @Override
   public Object getValue()
   {
      if ( cell_.getType() == SsCell.Type.NUMERIC )
      {
         return new Double(cell_.getValue());
      }
      else
      {
         return cell_.toString();
      }
   }

   @SuppressWarnings("rawtypes")
   @Override
   public Class getColumnClass()
   {
      if ( cell_.getType() == SsCell.Type.NUMERIC )
      {
         return Double.class;
      }
      else
      {
         return String.class;
      }
   }

   @Override
   public boolean isCellEditable()
   {
      if ( cell_.getType() == SsCell.Type.NUMERIC )
      {
         return true;
      }
      else
      {
         return false;
      }
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
   
   private SsCell cell_;
   private DefaultTableCellRenderer cellRenderer_;

}

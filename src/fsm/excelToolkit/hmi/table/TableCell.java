package fsm.excelToolkit.hmi.table;

import javax.swing.table.TableCellRenderer;

public abstract class TableCell implements TableCellRenderer
{
   public abstract Object getValue();
   
   @SuppressWarnings("rawtypes")
   public abstract Class getColumnClass();

   public abstract boolean isCellEditable();
   

}

package fsm.excelToolkit.hmi.table;

import java.awt.Component;
import java.text.NumberFormat;
import java.util.EventObject;

import javax.swing.JFormattedTextField;
import javax.swing.JTable;
import javax.swing.event.CellEditorListener;
import javax.swing.text.NumberFormatter;

import fsm.spreadsheet.SsCell;

public class TableCellSpreadsheet extends TableCell
{
   public TableCellSpreadsheet(SsCell cell)
   {
      cell_ = cell;
      if ( cell_.getType() == SsCell.Type.NUMERIC )
      {
         NumberFormat format = NumberFormat.getInstance();
         NumberFormatter formatter = new NumberFormatter(format);
         formatter.setValueClass(Double.class);
         formatter.setMinimum(Double.MIN_VALUE);
         formatter.setMaximum(Double.MAX_VALUE);
         formatter.setAllowsInvalid(false);
         // If you want the value to be committed on each keystroke instead of focus lost
         formatter.setCommitsOnValidEdit(true);
         component_ = new JFormattedTextField(formatter);         
      }
      else
      {
         component_ = new JFormattedTextField(); 
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
      return component_.getValue();
   }

   @Override
   public boolean isCellEditable(EventObject anEvent)
   {
      return cell_.getType() != SsCell.Type.READONLY;
   }

   @Override
   public boolean shouldSelectCell(EventObject anEvent)
   {
      return true;
   }

   @Override
   public boolean stopCellEditing()
   {
      return cell_.getType() == SsCell.Type.READONLY;
   }

   @Override
   public void cancelCellEditing()
   {
      // nop
   }

   @Override
   public void addCellEditorListener(CellEditorListener l)
   {
      // TODO Auto-generated method stub
      
   }

   @Override
   public void removeCellEditorListener(CellEditorListener l)
   {
      // TODO Auto-generated method stub
      
   }
   
   // --- PRIVATE
   
   private SsCell cell_;
   private JFormattedTextField component_;

}

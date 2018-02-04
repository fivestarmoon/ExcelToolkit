package fsm.excelToolkit.wpsr;

import java.awt.Color;

import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellSpreadsheet;
import fsm.spreadsheet.SsCell;

public class Resource
{
   public Resource(String name)
   {
      name_ = name;
      sum_ = new double[Columns.length()];
      used_ = false;
   }
   
   public void sumif(SsCell[] cells)
   {
      if ( !name_.equals(cells[Columns.RESOURCE.getIndex()].toString()) )
      {
         return;
      }
      used_ = true;
      for ( Columns col : Columns.values() )
      {
         if ( !col.isSpreadsheetCol() )
         {
            continue;
         }
         int index = col.getIndex();
         sum_[index] += cells[index].getValue();
      }          
   }
   
   public boolean isUsed()
   {
      return used_;
   }
   
   public SsCell[] getSummedCells()
   {
      SsCell[] cells = new SsCell[Columns.length()];
      for ( Columns col : Columns.values() )
      {
         int ii = col.getIndex();
         if ( col == Columns.RESOURCE )
         {
            cells[ii] = new SsCell(name_);            
         }
         else
         {
            cells[ii] = new SsCell(Columns.Round(sum_[ii])); 
         }
      }
      return cells;
   }
   
   public TableCell[] getTableCells()
   {
      SsCell[] cells = getSummedCells();
      TableCell[] tableCells = new TableCell[Columns.length()];
      for ( Columns col : Columns.values() )
      {
         int index = col.getIndex();
         if ( col.isSpreadsheetCol() )
         {
            tableCells[index] = new TableCellSpreadsheet(cells[index]); 
            if ( col == Columns.RESOURCE )
            {
               tableCells[index].setBlendBackgroundColor(new Color(255,255,0,64));
            }     
            else if ( col == Columns.BUDGET )
            {
               tableCells[index].setBold(true);
            }                             
         }
         else if ( col == Columns.EAC )
         {
            tableCells[index] = new TableCellSpreadsheet(Columns.GetEacFromCols(cells)); 
            tableCells[index].setBold(true);
         }
         else if ( col == Columns.VARIANCE )
         {
            tableCells[index] = new TableCellSpreadsheet(Columns.GetVarianceFromCols(cells)); 
         }
         else
         {
            tableCells[index] =  new TableCellSpreadsheet("");
         }
      }
      return tableCells;
   }

   // --- PRIVATE
   
   private String name_;
   private double[] sum_;
   private boolean used_;
}

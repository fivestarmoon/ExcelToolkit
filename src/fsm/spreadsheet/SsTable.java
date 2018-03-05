package fsm.spreadsheet;

import java.util.ArrayList;

public class SsTable
{
   
   
   public SsTable()
   {
      rows_ = new ArrayList<Integer>();
      cols_ = new ArrayList<Integer>();
      cells_ = new SsCell[0][0];
   }
   
   public void addRow(int row)
   {
      rows_.add(new Integer(row));
      cells_ = new SsCell[0][0];
   }
   
   public void addCol(int col)
   {
      cols_.add(new Integer(col));
      cells_ = new SsCell[0][0];
   }
   
   public void addRow(int start, int end)
   {
      for ( int val = start; val<=end; val++ )
      {
         addRow(val);
      }
      cells_ = new SsCell[0][0];
   }
   
   public void addCol(int start, int end)
   {
      for ( int val = start; val<=end; val++ )
      {
         addCol(val);
      }
      cells_ = new SsCell[0][0];
   }
   
   public int[] getRowsInSheet()
   {
      int rows[] = new int[rows_.size()];
      for ( int ii=0; ii<rows.length; ii++ )
      {
         rows[ii] = rows_.get(ii);
      }
      return rows;
   }
   
   public int[] getRowIterator()
   {
      int rows[] = new int[rows_.size()];
      for ( int ii=0; ii<rows.length; ii++ )
      {
         rows[ii] = ii;
      }
      return rows;
   }
   
   public int[] getColsInSheet()
   {
      int cols[] = new int[cols_.size()];
      for ( int ii=0; ii<cols.length; ii++ )
      {
         cols[ii] = cols_.get(ii);
      }
      return cols;
   }
   
   public int[] getColIterator()
   {
      int cols[] = new int[cols_.size()];
      for ( int ii=0; ii<cols.length; ii++ )
      {
         cols[ii] = ii;
      }
      return cols;
   }
   
   public SsCell getCell(int row, int col)
   {
      if ( row < 0 || row >= rows_.size() )
      {
         return new SsCell();
      }
      if ( col < 0 || col >= cols_.size() )
      {
         return new SsCell();
      }
      return cells_[row][col];
   }
   
   public SsCell[] getCellsForRow(int row)
   {
      if ( row < 0 || row >= rows_.size() )
      {
         return new SsCell[0];
      }
      SsCell[] cells = cells_[row];
      return cells;
   }
   
   @Override
   public String toString()
   {
      StringBuilder stringbuilder = new StringBuilder();
      stringbuilder.append("rows [");
      for ( int ii : rows_ )
      {
         stringbuilder.append(ii + " ");
      }
      stringbuilder.append("] rows [");
      for ( int ii : cols_ )
      {
         stringbuilder.append(SsCell.ColumnIndexToLetters(ii) + " ");
      }
      stringbuilder.append("]");
      return stringbuilder.toString();
   }
   
   // --- PROTECTED
   
   protected ArrayList<Integer> rows_;
   protected ArrayList<Integer> cols_;
   protected SsCell[][] cells_;
}

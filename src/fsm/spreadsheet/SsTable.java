package fsm.spreadsheet;

import java.util.ArrayList;

import org.apache.poi.ss.util.CellReference;

public class SsTable
{
   public static int ColumnLettersToIndex(String colLetters)
   {
      return CellReference.convertColStringToIndex(colLetters);
   }
   public static String ColumnIndexToLetters(int colIndex)
   {
      return CellReference.convertNumToColString(colIndex);
   }
   
   
   public SsTable(SsFile file, String sheet)
   {
      file_ = file;
      sheetRefIsString_ = false;
      setSheetString(new String(sheet));
      setSheetInt(0);
      rows_ = new ArrayList<Integer>();
      cols_ = new ArrayList<Integer>();
      cells_ = new SsCell[0][0];
   }

   public SsTable(SsFile file, int sheet)
   {
      file_ = file;
      sheetRefIsString_ = false;
      setSheetString("");
      setSheetInt(sheet);
      rows_ = new ArrayList<Integer>();
      cols_ = new ArrayList<Integer>();
      cells_ = new SsCell[0][0];
   }
   
   // --- PRIVATE

   public SsFile getFile()
   {
      return file_;
   }

   public boolean isSheetRefIsString()
   {
      return sheetRefIsString_;
   }

   public String getSheetString()
   {
      return sheetString_;
   }

   public void setSheetString(String sheetString_)
   {
      this.sheetRefIsString_ = true;
      this.sheetString_ = sheetString_;
      this.sheetInt_ = 0;
      cells_ = new SsCell[0][0];
   }

   public int getSheetInt()
   {
      return sheetInt_;
   }

   public void setSheetInt(int sheetInt_)
   {
      this.sheetRefIsString_ = false;
      this.sheetString_ = "";
      this.sheetInt_ = sheetInt_;
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
   
   public void readTable()
   {
      if ( sheetRefIsString_ )
      {
         cells_ = file_.readTable(sheetString_, getRowsInSheet(), getColsInSheet());
      }
      else
      {
         cells_ = file_.readTable(sheetInt_, getRowsInSheet(), getColsInSheet());
      }
   }
   
   public void writeTable()
   {
      file_.writeTable(getRowsInSheet(), getColsInSheet(), cells_);      
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
   
   @Override
   public String toString()
   {
      StringBuilder stringbuilder = new StringBuilder();
      stringbuilder.append(file_.getFile().getName());
      stringbuilder.append("[");
      if ( sheetRefIsString_ )
      {
         stringbuilder.append(sheetString_);
      }
      else
      {
         stringbuilder.append(sheetInt_);
      }
      stringbuilder.append("] rows [");
      for ( int ii : rows_ )
      {
         stringbuilder.append(ii + " ");
      }
      stringbuilder.append("] rows [");
      for ( int ii : cols_ )
      {
         stringbuilder.append(ColumnIndexToLetters(ii) + " ");
      }
      stringbuilder.append("]");
      return stringbuilder.toString();
   }
   
   // --- PRIVATE

   private SsFile file_;
   private boolean sheetRefIsString_;
   private String sheetString_;
   private int sheetInt_;
   private ArrayList<Integer> rows_;
   private ArrayList<Integer> cols_;
   private SsCell[][] cells_;
}

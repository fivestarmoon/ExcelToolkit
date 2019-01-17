package fsm.excelToolkit.loading;

import fsm.spreadsheet.SsCell;

public class Resource
{
   public Resource(String name, int numMonths)
   {
      name_ = name;
      sum_ = new double[numMonths];
      used_ = false;
   }
   
   public void sumif(SsCell[] cells)
   {
      if ( !name_.equals(cells[0].toString()) )
      {
         return;
      }
      used_ = true;
      for (int ii=0; ii<sum_.length; ii++ )
      {
         sum_[ii] += cells[HmiColumns.getMonthIndexToColumnIdx(ii)].getValue();
      }          
   }
   
   public void sum(SsCell[] cells)
   {
      used_ = true;
      for (int ii=0; ii<sum_.length; ii++ )
      {
         sum_[ii] += cells[HmiColumns.getMonthIndexToColumnIdx(ii)].getValue();
      }          
   }
   
   public boolean isUsed()
   {
      return used_;
   }
   
   public SsCell[] getCells()
   {
      SsCell[] hmiCells = new SsCell[HmiColumns.length(sum_.length)];
      hmiCells[HmiColumns.LABEL.getIndex()] = new SsCell(name_);
      hmiCells[HmiColumns.RESOURCE.getIndex()] = new SsCell("");
      double rowSum = 0.0;
      for ( int ii=0; ii<sum_.length; ii++ )
      {
         hmiCells[HmiColumns.getMonthIndexToColumnIdx(ii)] = new SsCell(sum_[ii]);
         rowSum += sum_[ii];
      }
      hmiCells[HmiColumns.ROW_SUM.getIndex()] = new SsCell(rowSum);
      return hmiCells;
   }

   // --- PRIVATE
   
   private String name_;
   private double[] sum_;
   private boolean used_;
}

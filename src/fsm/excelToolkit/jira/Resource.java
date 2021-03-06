package fsm.excelToolkit.jira;

import java.awt.Color;

import fsm.excelToolkit.hmi.table.TableCell;
import fsm.excelToolkit.hmi.table.TableCellSpreadsheet;
import fsm.spreadsheet.SsCell;

public class Resource
{
   public Resource(String name)
   {
      name_ = name;
      sum_ = new double[SsColumns.length()];
      used_ = false;
   }
   
   public boolean sumif(SsCell[] cells)
   {
      if ( !name_.equals(cells[SsColumns.RESOURCE.getIndex()].toString()) )
      {
         return false;
      }
      used_ = true;
      for ( SsColumns col : SsColumns.values() )
      {
         int index = col.getIndex();
         sum_[index] += cells[index].getValue();
      }    
      return true;
   }
   
   public void sum(SsCell[] cells)
   {
      used_ = true;
      for ( SsColumns col : SsColumns.values() )
      {
         int index = col.getIndex();
         sum_[index] += cells[index].getValue();
      }          
   }
   
   public boolean isUsed()
   {
      return used_;
   }
   public void setActuals(double actuals)
   {
      sum_[SsColumns.ACTUAL.getIndex()] = actuals;
   }
   public void addActuals(double actuals)
   {
      sum_[SsColumns.ACTUAL.getIndex()] += actuals;
   }
   
   public TableCell[] getTableCells(double varianceColorRedThreshold)
   {
      TableCell[] tableCells = new TableCell[HmiColumns.length()];
      
      int index = HmiColumns.RESOURCE.getIndex();
      tableCells[index] = new TableCellSpreadsheet(new SsCell(name_)); 
      tableCells[index].setBlendBackgroundColor(new Color(255,255,0,64));
      
      index = HmiColumns.BUDGET.getIndex();
      double budget = sum_[SsColumns.BUDGET.getIndex()];
      tableCells[index] = new TableCellSpreadsheet(new SsCell(Round(budget))); 
      tableCells[index].setBold(true);
      
      index = HmiColumns.ACTUAL.getIndex();
      double actual = sum_[SsColumns.ACTUAL.getIndex()];
      tableCells[index] = new TableCellSpreadsheet(new SsCell(Round(actual))); 
      
      index = HmiColumns.ETC.getIndex();
      double etc = sum_[SsColumns.ETC.getIndex()];
      tableCells[index] = new TableCellSpreadsheet(new SsCell(Round(etc)));
      
      index = HmiColumns.EAC.getIndex();
      double eac = actual + etc;
      tableCells[index] = new TableCellSpreadsheet(new SsCell(Round(eac))); 
      tableCells[index].setBold(true); 
      
      index = HmiColumns.VARIANCE.getIndex();
      double variance = budget - eac;
      tableCells[index] = new TableCellSpreadsheet(new SsCell(Round(variance)));
      if ( variance < -1*varianceColorRedThreshold )
      {
         tableCells[index].setBlendBackgroundColor(varianceWarning_);
      }
      
      return tableCells;
   }
   
   public String getName()
   {
      return name_;
   }

   // --- PRIVATE
   
   private static double Round(double value)
   {
      return Math.round(value*100.0) / 100.0;
   }
   
   private String name_;
   private double[] sum_;
   private boolean used_;
   private Color varianceWarning_ = new Color(255, 0, 0, 64);
}

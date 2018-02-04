package fsm.excelToolkit.wpsr;

import fsm.spreadsheet.SsCell;

enum Columns
{
   RESOURCE(0, "resourceCol", "Resource"),
   BUDGET(1, "budgetCol", "Budget"),
   PREVACTUAL(2, "prevActualCol", "Prev Actual"),
   THISACTUAL(3, "thisActualCol", "Week Actual"),
   ETC(4, "etcCol", "ETC"),
   EAC(5, "", "EAC"),
   VARIANCE(6, "", "Variance");
   
   public static int length()
   {
      return VARIANCE.getIndex()+1;
   }
   public int getIndex()
   {
      return index_;
   }
   public boolean isSpreadsheetCol()
   {
      return resourceKey_.length() > 0;
   }
   public String getResourceKey()
   {
      return resourceKey_;
   }
   public String getColumnTitle()
   {
      return columnTitle_;
   }
   
   public static double GetEacFromCols(SsCell[] cells)
   {
      double sum = 0.0;
      sum += cells[Columns.PREVACTUAL.getIndex()].getValue();
      sum += cells[Columns.THISACTUAL.getIndex()].getValue();
      sum += cells[Columns.ETC.getIndex()].getValue();
      return Round(sum);
   }
   
   public static double GetVarianceFromCols(SsCell[] cells)
   {
      return Round(cells[Columns.BUDGET.getIndex()].getValue() - GetEacFromCols(cells));
   }
   
   public static double Round(double value)
   {
      return Math.round(value*100.0) / 100.0;
   }
   
   
   Columns(int index, String resourceKey, String columnTitle)
   {
      index_ = index;
      resourceKey_ = resourceKey;
      columnTitle_ = columnTitle;
   }
   private int index_;
   private String resourceKey_;
   private String columnTitle_;
}
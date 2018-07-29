package fsm.excelToolkit.jira;

enum HmiColumns
{
   RESOURCE(0, "Resource"),
   BUDGET(1, "Budget"),
   ACTUAL(2, "Actuals"),
   ETC(3, "ETC"),
   EAC(4, "EAC"),
   VARIANCE(5, "Variance");
   
   public static int length()
   {
      return VARIANCE.getIndex()+1;
   }
   public int getIndex()
   {
      return index_;
   }
   public String getColumnTitle()
   {
      return columnTitle_;
   }
   
   
   HmiColumns(int index, String columnTitle)
   {
      index_ = index;
      columnTitle_ = columnTitle;
   }
   private int index_;
   private String columnTitle_;
}
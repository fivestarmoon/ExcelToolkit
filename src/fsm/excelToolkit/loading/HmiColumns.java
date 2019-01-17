package fsm.excelToolkit.loading;

enum HmiColumns
{
   LABEL(0, "Label"),
   RESOURCE(1, "Resource"),
   ROW_SUM(2, "Row Sum"),
   MONTHS(3, "");
   
   public static int length(int numMonths)
   {
      return MONTHS.getIndex()+numMonths;
   }
   public static boolean isMonth(int index)
   {
      return (index >= MONTHS.index_);
   }
   public static int columnIndexToMonthIdx(int index)
   {
      return (index - MONTHS.index_);
   }
   public static int getMonthIndexToColumnIdx(int ii)
   {
      return MONTHS.index_+ ii;
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
package fsm.excelToolkit.jira;

enum SsColumns
{
   RESOURCE(0, "resourceCol"),
   CHARGECODE(1, "chargeCodeCol"),
   BUDGET(2, "budgetCol"),
   ACTUAL(3, "actualCol"),
   ETC(4, "estimateCol");
   
   public static int length()
   {
      return ETC.getIndex()+1;
   }
   public int getIndex()
   {
      return index_;
   }
   public String getResourceKey()
   {
      return resourceKey_;
   }
   
   
   SsColumns(int index, String resourceKey)
   {
      index_ = index;
      resourceKey_ = resourceKey;
   }
   private int index_;
   private String resourceKey_;
}
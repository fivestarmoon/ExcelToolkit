package fsm.excelToolkit.wpsr;

enum SsColumns
{
   RESOURCE(0, "resourceCol"),
   BUDGET(1, "budgetCol"),
   PREVACTUAL(2, "prevActualCol"),
   THISACTUAL(3, "thisActualCol"),
   ETC(4, "etcCol");
   
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
   public void setDefaultValue(String value)
   {
      defaultValue_ = value;;
   }
   public String getDefaultValue()
   {
      return defaultValue_;
   }
   
   
   SsColumns(int index, String resourceKey)
   {
      index_ = index;
      resourceKey_ = resourceKey;
   }
   private int index_;
   private String resourceKey_;
   private String defaultValue_;
}
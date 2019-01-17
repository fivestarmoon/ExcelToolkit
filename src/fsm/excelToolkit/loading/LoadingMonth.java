package fsm.excelToolkit.loading;

import java.util.StringTokenizer;

public class LoadingMonth
{
   @SuppressWarnings("serial")
   public static class LoadingMonthException extends Exception
   {
      public LoadingMonthException(String exception)
      {
         super(exception);
      }
   }
   public static int DiffInMonths(LoadingMonth endDate,
                                  LoadingMonth startDate)
   {
      return (endDate.year_ - startDate.year_)*12 
               + (endDate.month_ - startDate.month_) + 1;
   }
   
   public LoadingMonth(String string) throws LoadingMonthException
   {
      StringTokenizer tok = new StringTokenizer(string, "-");
      if ( tok.countTokens() != 2)
      {
         throw new LoadingMonthException("Date string should be YYYY-MM");
      }
      validate(Integer.parseInt(tok.nextToken()), Integer.parseInt(tok.nextToken()));
   }
   public LoadingMonth(int year, int month) throws LoadingMonthException
   { 
      validate(year, month);
   }
   public LoadingMonth(LoadingMonth startDate, int month) throws LoadingMonthException
   { 
      int year = startDate.year_ + (startDate.month_ + month - 1)/12;
      month = (startDate.month_ + month - 1)%12 + 1;
      validate(year, month);
   }
   public boolean equals(LoadingMonth startColDate, int jj)
   {
      int year = startColDate.year_ + (startColDate.month_ + jj - 1)/12;
      int month = (startColDate.month_ + jj - 1)%12 + 1;
      return (year_==year && month_==month);
   }
   
   @Override
   public String toString()
   {
      return String.format("%02d-%02d", year_%100, month_);
   }
   
   // --- PRIVATE
   
   private void validate(int year, int month) throws LoadingMonthException
   {
      year_ = year;
      if ( year_ < 2000 || 3000 < year_ )
      {
         throw new LoadingMonthException(String.format("Date string should be YYYY, value of %d seems wrong", year_));
      }
      month_ = month;
      if ( month_ < 1 || 12 < month_ )
      {
         throw new LoadingMonthException(String.format("Date string should be MM, value of %d is wrong", month_));
      }   
   }
   
   int year_;
   int month_;

}

package fsm.spreadsheet;

import java.util.ArrayList;

import org.apache.poi.ss.util.CellReference;

public class SsCell
{
   public enum Type
   {
      NUMERIC,
      STRING,
      READONLY      
   };
   
   public static int ColumnLettersToIndex(String colLetters)
   {
      return CellReference.convertColStringToIndex(colLetters);
   }
   public static String ColumnIndexToLetters(int colIndex)
   {
      return CellReference.convertNumToColString(colIndex);
   }
   
   public SsCell(double value)
   {
      double_ = value;
      string_ = Double.toString(value);
      type_ = Type.READONLY;
      modified_ = false;
   }

   public SsCell(String string)
   {
      double_ = 0.0;
      string_ = new String(string);
      type_ = Type.READONLY;
      modified_ = false;
   }
   
   public void update(String stringValue)
   {
      modified_ = true;
      string_ = stringValue;
      try
      {
         double_ = Double.parseDouble(stringValue);
      }
      catch (NumberFormatException e)
      {
         double_ = 0.0;
      }   
   }
   
   public void update(double numericValue)
   {
      modified_ = true;
      string_ = Double.toString(numericValue);
      double_ = numericValue;
   }
   
   @Override
   public String toString()
   {
      return string_;
   }
   
   public double getValue()
   {
      return double_;
   }
   
   public Type getType()
   {
      return type_;
   }
   
   public boolean isModified()
   {
      return modified_;
   }
   
   public static void AddUnique(SsCell cell, ArrayList<String> values)
   {
      if ( cell.toString().length() == 0 )
      {
         return;
      }
      for ( String value : values )
      {
         if ( cell.toString().equals(value)  )
         {
            return; // already on the list
         }
      }
      values.add(cell.toString());
   }
   
   // --- PROTECTED

   SsCell()
   {
      double_ = 0.0;
      string_ = "";
      type_ = Type.READONLY;
      modified_ = false;
   }

   SsCell(String string, double value, Type type)
   {
      string_ = string;
      double_ = value;
      type_ = type;
      modified_ = false;
   }

   void fill(String stringValue, double numericValue, Type type)
   {
      string_ = stringValue;
      double_ = numericValue;
      type_ = type;
      modified_ = false;
   }

   // --- PRIVATE
   
   private boolean modified_;
   private String string_;
   private double double_;
   private Type type_;
}

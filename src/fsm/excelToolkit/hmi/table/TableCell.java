package fsm.excelToolkit.hmi.table;

import java.awt.Color;
import java.awt.Font;

import javax.swing.table.TableCellRenderer;

public abstract class TableCell implements TableCellRenderer
{
   public abstract Object getValue();
   
   @SuppressWarnings("rawtypes")
   public abstract Class getColumnClass();

   public abstract boolean isCellEditable();
   
   public void setBlendBackgroundColor(Color c)
   {
      blendColor_ = c;
   }
   
   public void setBold(boolean bold)
   {
      bold_ = bold;
   }
   public void setItalics(boolean b)
   {
      italics_ = b;
   }
   
   
   protected Font getFont(Font componentFont)
   {
      if ( font_ == null )
      {
         if ( bold_ )
         {
            font_ = new Font(componentFont.getName(), Font.BOLD, componentFont.getSize());
         }
         else if ( italics_ )
         {
            font_ = new Font(componentFont.getName(), Font.ITALIC, componentFont.getSize());
         }
         else
         {
            font_ = componentFont;
         }
      }
      return font_;
   }
   
   protected Color getBackgroundColor(Color componentColor)
   {
      if ( backgroundColor_ == null )
      {
         if ( blendColor_ != null )
         {
            backgroundColor_ = blend(componentColor, blendColor_);
         }
         else
         {
            backgroundColor_ = componentColor;
         }
      }
      return backgroundColor_;
   }

   private static Color blend(Color c0, Color c1) 
   {
      double totalAlpha = c0.getAlpha() + c1.getAlpha();
      double weight0 = c0.getAlpha() / totalAlpha;
      double weight1 = c1.getAlpha() / totalAlpha;

      double r = weight0 * c0.getRed() + weight1 * c1.getRed();
      double g = weight0 * c0.getGreen() + weight1 * c1.getGreen();
      double b = weight0 * c0.getBlue() + weight1 * c1.getBlue();
      double a = Math.max(c0.getAlpha(), c1.getAlpha());

      return new Color((int) r, (int) g, (int) b, (int) a);
   }

   private Color blendColor_ = null;
   private Color backgroundColor_ = null;
   private Font  font_ = null;
   private boolean bold_;
   private boolean italics_;
}

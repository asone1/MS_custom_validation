package msexcel;

import org.apache.poi.sl.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.*;

public class Range
{
    private int m_start_zeile = 0;
    private int m_end_zeile = 0;
    private int m_start_spalte = 0;
    private int m_end_spalte = 0;

    private Sheet m_sheet = null;

    public Range( Sheet pSheet, int pStartZeile, int pStartSpalte, int pEndZeile, int pEndSpalte )
    {
        m_sheet = pSheet;
        m_start_zeile = pStartZeile;
        m_end_zeile = pEndZeile;
        m_start_spalte = pStartSpalte;
        m_end_spalte = pEndSpalte;

        m_it_zeile = m_start_zeile;
        m_it_spalte = m_start_spalte;
    }

    private int m_it_zeile = 0;
    private int m_it_spalte = 0;

    public Cell first()
    {
        m_it_zeile = m_start_zeile;
        m_it_spalte = m_start_spalte;

        return getCell( m_it_zeile, m_it_spalte );
    }

    public Cell next()
    {
        m_it_spalte++;

        if ( m_it_spalte > m_end_spalte )
        {
            m_it_zeile++;
            m_it_spalte = m_start_spalte;
        }

        return getCell( m_it_zeile, m_it_spalte );
    }


    private Cell getCell( int pZeile , int pSpalte )
    {
        /*
         * Prüfung: Iteratorzeile gültig?
         */
        if (( m_it_zeile>= m_start_zeile ) && ( m_it_zeile <= m_end_zeile ))
        {
            if (( m_it_spalte>= m_start_spalte ) && ( m_it_spalte <= m_end_spalte ))
            {
                Row row  = m_sheet.getRow( m_it_zeile );

                if ( row == null )
                {
                    row = m_sheet.createRow( m_it_zeile );
                }

                if ( row != null )
                {
                    Cell cell = row.getCell( m_it_spalte );

                    if ( cell == null )
                    {
                        cell = row.createCell( m_it_spalte );
                    }

                    return cell;
                }
            }
        }
        return null;
    }


    /**
     * Liefert den Wert der Variablen "m_end_spalte".
     *
     * @return m_end_spalte
     */
    public int getEndSpalte()
    {
        return m_end_spalte;
    }


    /**
     * Liefert den Wert der Variablen "m_end_zeile".
     *
     * @return m_end_zeile
     */
    public int getEndZeile()
    {
        return m_end_zeile;
    }


    /**
     * Liefert den Wert der Variablen "m_sheet".
     *
     * @return m_sheet
     */
    public Sheet getSheet()
    {
        return m_sheet;
    }


    /**
     * Liefert den Wert der Variablen "m_start_spalte".
     *
     * @return m_start_spalte
     */
    public int getStartSpalte()
    {
        return m_start_spalte;
    }


    /**
     * Liefert den Wert der Variablen "m_start_zeile".
     *
     * @return m_start_zeile
     */
    public int getStartZeile()
    {
        return m_start_zeile;
    }


    /**
     * Setzt den Wert der Variablen "m_end_spalte".
     *
     * @param pEndSpalte der zu setzende Wert
     */
    public void setEndSpalte( int pEndSpalte )
    {
        m_end_spalte = pEndSpalte;
    }


    /**
     * Setzt den Wert der Variablen "m_end_zeile".
     *
     * @param pEndZeile der zu setzende Wert
     */
    public void setEndZeile( int pEndZeile )
    {
        m_end_zeile = pEndZeile;
    }


    /**
     * Setzt den Wert der Variablen "m_sheet".
     *
     * @param pSheet der zu setzende Wert
     */
    public void setSheet( Sheet pSheet )
    {
        m_sheet = pSheet;
    }


    /**
     * Setzt den Wert der Variablen "m_start_spalte".
     *
     * @param pStartSpalte der zu setzende Wert
     */
    public void setStartSpalte( int pStartSpalte )
    {
        m_start_spalte = pStartSpalte;
    }


    /**
     * Setzt den Wert der Variablen "m_start_zeile".
     *
     * @param pStartZeile der zu setzende Wert
     */
    public void setStartZeile( int pStartZeile )
    {
        m_start_zeile = pStartZeile;
    }

    // ########################################################################################################################
    //                                                           toString
    // ########################################################################################################################

    /**
     * Erstellt die String-Repräsentation dieser Klasse
     *
     * @return pProperties eine Auflistung der Variablen und deren Inhalte
     */
    public String toString()
    {
        String log_string = "";

        /*
         * Ausgabe aller Werte aus der Ini-Datei
         */
        log_string += "\n + END_SPALTE    >" + m_end_spalte    + "<";
        log_string += "\n + END_ZEILE     >" + m_end_zeile     + "<";
        log_string += "\n + SHEET         >" + m_sheet         + "<";
        log_string += "\n + START_SPALTE  >" + m_start_spalte  + "<";
        log_string += "\n + START_ZEILE   >" + m_start_zeile   + "<";

        return log_string;
    }
}
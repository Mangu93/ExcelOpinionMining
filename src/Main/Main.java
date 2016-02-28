package lectorexcel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.StringTokenizer;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Adrian Portillo
 */
public class Main {

    public static String column_name = "ENTER THE COLUMN HERE";
    public static List<Cell> filtered_cells = null;
    public static List<String> filter_words = null;
    public static Map<String, Integer> dictionary = new HashMap<>();
    public static Map<Integer, List<Integer>> similarities = new HashMap<>();
    public static final double SIMILARITY_FILTER = 0.60;
    public static Map<String, Integer> association = new HashMap<>();
    public static Map<String, Integer> word_cloud = new HashMap<>();
    public static final int MAXIMUM_TWEET = 50000; //You can change this value if you don't want to take a lot of tweets.
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        initList();
        InputStream input = null;

        try {
            input = new FileInputStream("ENTER YOUR XLS HERE");
            Workbook wb = WorkbookFactory.create(input);
            Sheet actual_sheet = wb.getSheetAt(1);
            Integer column_no = null;
            List<Cell> first_cells = new ArrayList<>();
            Row firstRow = actual_sheet.getRow(0);

            for (Cell cell : firstRow) {
                if (cell.getStringCellValue().equals(column_name)) {
                    column_no = cell.getColumnIndex();
                }
            }
            if (column_no != null) {
                for (Row row : actual_sheet) {
                    Cell c = row.getCell(column_no);
                    if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK) {
                        //Skip, it's empty
                    } else {
                        first_cells.add(c);
                    }
                }
            } else {
                System.out.println("That column doesn't exist");
            }
            filter(first_cells);
            System.out.println("It has " + similarities.size() + " elements similars");
            BufferedWriter bf = new BufferedWriter(new FileWriter(new File("ENTER YOUR OUTPUT TEXT FILE HERE")));
            for (String s : dictionary.keySet()) {
                bf.write(index + ". " + s);
                bf.write("\n");
            }
            int max = 0;
            word_cloud = sortByValue(word_cloud);
            for (String s : word_cloud.keySet()) {
                if (max < 50) {
                    bf.write(s);
                    bf.write("\n");
                    max++;
                } else {
                    break;
                }
            }
            bf.close();
        } catch (Exception ex) {
            System.err.println(ex.getMessage());
        } finally {

        }
    }


    /**
     * Filter the cells and adds it to a map
     *
     * @param first_cells
     */
    private static void filter(List<Cell> first_cells) {
        boolean found = false;
        for (Cell actual_cell : first_cells) {
            if (dictionary.size() > MAXIMUM_TWEET) {
                break;
            }
            found = false;
            String cell_string = actual_cell.toString();
            if (not_in_list(cell_string)) {
                for (String key : dictionary.keySet()) {
                    if (similarity(cell_string, key) >= SIMILARITY_FILTER) {
                        dictionary.put(key, dictionary.get(key) + 1); //Add one to the map
                        Integer row = association.get(key);
                        List<Integer> list_ = similarities.get(row);
                        list_.add(actual_cell.getRowIndex());
                        similarities.put(row, list_);
                        found = !found;
                        System.out.println("Added similar");
                        tokenize(key);
                        break;
                    }
                }
                if (!found) {
                    dictionary.put(cell_string, 1);
                    similarities.put(actual_cell.getRowIndex(), new ArrayList<Integer>());
                    association.put(cell_string, actual_cell.getRowIndex());
                    System.out.println("Added new");
                    tokenize(cell_string);
                }
            }
        }
    }

    private static void tokenize(String tkn) {
        StringTokenizer stkn = new StringTokenizer(tkn, " ");
        while (stkn.hasMoreTokens()) {
            add_to_cloud(stkn.nextToken());
        }
    }

    private static void add_to_cloud(String s) {
        if (word_cloud.containsKey(s)) {
            word_cloud.put(s, word_cloud.get(s) + 1);
        } else {
            word_cloud.put(s, 1);
        }
    }

    /**
     * Initiates filter words lists
     */
    private static void initList() {
        filter_words = new ArrayList<>();
        filter_words.add(column_name);
        //Add the filters you need. In this case, we were filtering some words we thought appeared a lot.
        filter_words.add("Orange is the new Black");
        filter_words.add("Holanda");
        filter_words.add("Liga");
        filter_words.add("Club");
    }

    /**
     * Here comes dragons! Source:
     * https://stackoverflow.com/questions/109383/sort-a-mapkey-value-by-values-java
     *
     * @param <K>
     * @param <V>
     * @param map
     * @return
     */
    public static <K, V extends Comparable<? super V>> Map<K, V>
            sortByValue(Map<K, V> map) {
        List<Map.Entry<K, V>> list
                = new LinkedList<>(map.entrySet());
        Collections.sort(list, new Comparator<Map.Entry<K, V>>() {
            @Override
            public int compare(Map.Entry<K, V> o1, Map.Entry<K, V> o2) {
                return (o1.getValue()).compareTo(o2.getValue());
            }
        });

        Map<K, V> result = new LinkedHashMap<>();
        for (Map.Entry<K, V> entry : list) {
            result.put(entry.getKey(), entry.getValue());
        }
        return result;
    }

    /**
     * Calculates the similarity (a number within 0 and 1) between two strings.
     */
    public static double similarity(String s1, String s2) {
        String longer = s1, shorter = s2;
        if (s1.length() < s2.length()) { // longer should always have greater length
            longer = s2;
            shorter = s1;
        }
        int longerLength = longer.length();
        if (longerLength == 0) {
            return 1.0; /* both strings are zero length */

        }
        return (longerLength - StringUtils.getLevenshteinDistance(longer, shorter)) / (double) longerLength;
    }

    /**
     *
     * @param cell_string
     * @return true if the string doesn't contain a filter word, false if it has
     * it
     */
    private static boolean not_in_list(String cell_string) {
        for (String filter : filter_words) {
            if (cell_string.toLowerCase().contains(filter.toLowerCase())) {
                return false;
            }
        }
        return true;
    }

}

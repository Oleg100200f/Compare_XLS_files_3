import com.aspose.cells.*;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {

    public static void main(String[] args) throws Exception {
        // Вказати шлях до файлу Excel
        String filePath = "G:\\777\\IV112.xlsx"; // Файл_1
        String filePath2 = "G:\\777\\Irregular Cambridge.xlsx"; // Файл_2
        String filePath3 = "G:\\777\\Irregular Cambridge Marked.xlsx"; // Оригінальний файл
        String filePath3Copy = "G:\\777\\Irregular Cambridge Marked_2.xlsx"; // Копія файлу

        try {
            // Створити копію файлу
            Files.copy(Paths.get(filePath3), Paths.get(filePath3Copy), StandardCopyOption.REPLACE_EXISTING);

            // Відкрити файли
            Workbook workbook = new Workbook(filePath);
            Workbook workbook2 = new Workbook(filePath2);
            Workbook workbook3 = new Workbook(filePath3Copy); // Працюємо з копією файлу

            // Вибрати перший аркуш (worksheet)
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Worksheet worksheet2 = workbook2.getWorksheets().get(0);
            Worksheet worksheet3 = workbook3.getWorksheets().get(0);

            // Отримати доступ до всіх комірок
            Cells cells = worksheet.getCells();
            Cells cells2 = worksheet2.getCells();
            Cells cells3 = worksheet3.getCells();

            // Порівняти перші стовпчики рядків і записати однакові слова
            int count = 1;
            String[] equalsWords = new String[200];

            for (int row = 1; row < cells.getMaxDataRow() + 1; row++) {
                Cell cell = cells.get(row, 1);
                for (int row_2 = 1; row_2 < cells2.getMaxDataRow() + 1; row_2++) {
                    Cell cell2 = cells2.get(row_2, 1);
                    if (GetOnlyStr(cell).equals(GetOnlyStr(cell2))) {
                        System.out.printf("%d %s%n", count, cell.getStringValue());
                        equalsWords[count - 1] = GetOnlyStr(cell);
                        count++;
                    }
                }
            }

            System.out.print("-----------Following Items will be marked--------------------------\n");
            count = 1;

            for (int row = 1; row < cells3.getMaxDataRow() + 1; row++) {
                Cell cell3 = cells3.get(row, 1);
                for (int i = 0; i < equalsWords.length; i++) {
                    if (equalsWords[i] != null && GetOnlyStr(cell3).equals(equalsWords[i])) {
                        System.out.printf("%d Following Item has marked = %s%n", count, cell3.getStringValue());
                        cells3.get(row, 4).putValue("V");
                        count++;
                    }
                }
            }

            // Зберегти зміни у копії файлу
            workbook3.save(filePath3Copy);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String GetOnlyStr(Cell cell) {
        String str = cell.getStringValue();
        str = str.trim();
        if (str.indexOf(" ") > 0) {
            str = str.substring(0, str.indexOf(" "));
        }
        return str;
    }


}
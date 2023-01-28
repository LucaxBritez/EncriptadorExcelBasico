import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*Este programa tomara los datos de las dos primeras columnas de un excel(archivo xlsx),
encriptara los datos de la segunda, guardara estos datos en un hashmap
y volvera a guardar los datos ya encriptados dentro de otro archivo xlsx.
 La idea principal es utilizarlo en datos de tipo "Usario/contraseña" */
public class Main {
    public static void main(String[] args) {
        try{
            //Ingresar direccion del archivo xlsx que se desea encriptar en el parametro del constructor File.

            File filePathIn = new File("C:/Users/Zeac/Desktop/TrabajoFinalJavaBasico/prueba.xlsx");
            FileInputStream fileIn = new FileInputStream(filePathIn);

            XSSFWorkbook workbookIn= new XSSFWorkbook(fileIn);
            XSSFSheet sheetIn = workbookIn.getSheetAt(0);
            Iterator<Row> rowItr = sheetIn.iterator();

            HashMap<String, String> map = new HashMap <>();

            int contador = 0;

            String usuario = "";
            String password;

            while(rowItr.hasNext()){
                Row row = rowItr.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();

                    contador++;

                    //Este if ese utilizado para guardar los usuarios y sus contraseñas
                    //encriptadas dentro del HashMap map.

                    if(contador % 2 != 0){

                        usuario = cell.getStringCellValue();
                        System.out.println("Impar: " + usuario);
                    }else{

                        password = cell.getStringCellValue();
                        System.out.println("Par: " + password);
                        char [] encriptado = password.toCharArray();
                        for(int i = 0; i < encriptado.length; i++){
                            encriptado[i] = (char)(encriptado [i] + (char)5);
                        }

                        String passwordEncriptado = String.valueOf(encriptado);
                        System.out.println("Password encriptada: " + passwordEncriptado);
                        System.out.println("\n");
                        map.put(usuario,passwordEncriptado);
                    }
                }

            }
            System.out.println("Se termino de procesar el archivo \n\n");

            //En el parametro del constructor File ingresar la direccion en donde se desea guardar el nuevo archivo
            //Excel encriptado.
            File filePathOut = new File("C:/Users/Zeac/Desktop/TrabajoFinalJavaBasico/Destino4.xlsx");
            XSSFWorkbook workbookOut = new XSSFWorkbook();
            XSSFSheet sheetOut = workbookOut.createSheet("Hoja 1");

            int contadorCell = 0;
            int contadorRow  = 0;

            for(HashMap.Entry<String, String> m : map.entrySet()){
                Row rowOut = sheetOut.createRow(contadorRow);
                rowOut.createCell(contadorCell).setCellValue(m.getKey());
                contadorCell++;
                rowOut.createCell(contadorCell).setCellValue(m.getValue());
                contadorRow++;
                contadorCell--;
            }

            FileOutputStream fileOut = new FileOutputStream(filePathOut);
            workbookOut.write(fileOut);
            fileOut.close();

            System.out.println("Excel creado");

        }catch(Exception e){
            e.printStackTrace();

        }

    }
}
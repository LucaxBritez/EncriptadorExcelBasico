import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*Este programa tomara los datos de las dos primeras columnas de un excel(archivo xlsx),
encriptara los datos de la segunda, y volvera a guardar los datos ya encriptados dentro de otro archivo xlsx.
 La idea principal es utilizarlo en datos de tipo "Usario/contraseña" */
public class Main {
    public static void main(String[] args) {
        try{
            //Mensaje inicial.
            String solicitudPathInicial = "Porfavor ingrese la dirección del archivo a encriptar" +
                    ",esto incluye el nombre del mismo y su extension(xlsx).";
            System.out.println(solicitudPathInicial);
            //Solicitud de la direccion del archivo a encriptar.
            String pathInicial;
            Scanner pathInicialTipeado = new Scanner(System.in);
            pathInicial = pathInicialTipeado.nextLine();

            //Lectura del archivo a encriptar
            File filePathIn = new File(pathInicial);
            FileInputStream fileIn = new FileInputStream(filePathIn);

            //Se instancian los elementos que componen un archivo de excel y su iterador.
            XSSFWorkbook workbookIn= new XSSFWorkbook(fileIn);
            XSSFSheet sheetIn = workbookIn.getSheetAt(0);
            Iterator<Row> rowItr = sheetIn.iterator();

            HashMap<String, String> map = new HashMap <>();

            int contador = 0;

            String usuario = "";
            String password;

            //While inicial, se utiliza para recorrer las columnas.
            while(rowItr.hasNext()){
                Row row = rowItr.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                //While final utilizado para iterar sobre las celdas del workbook.
                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();

                    contador++;

                    //Este if ese utilizado para guardar los usuarios y sus contraseñas
                    //encriptadas dentro del HashMap map.

                    if(contador % 2 != 0){
                        usuario = cell.getStringCellValue();

                    }else{
                        password = cell.getStringCellValue();
                        char [] encriptado = password.toCharArray();
                        for(int i = 0; i < encriptado.length; i++){
                        encriptado[i] = (char)(encriptado [i] + (char)5);
                        }
                        String passwordEncriptado = String.valueOf(encriptado);
                        map.put(usuario,passwordEncriptado);
                    }
                }

            }
            System.out.println("-----Se termino de procesar el archivo----- \n\n");

            String pathDestino;
            String solicitudPathDestino = "Porfavor ingrese la direccion en la cual desee guardar el archivo encriptado" +
                    "y el nombre que desea que tenga el archivo encriptado junto con su extension(xlsx).";
            System.out.println(solicitudPathDestino);
            Scanner pathDestinoTipeado = new Scanner(System.in);
            pathDestino = pathDestinoTipeado.nextLine();

            File filePathOut = new File(pathDestino);
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

            System.out.println("----Excel creado----");

        }catch(Exception e){
            e.printStackTrace();

        }

    }
}
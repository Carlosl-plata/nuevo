package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;



public class Main {

    public static void main(String[] args) {

        //Configuracion inicail Selenium

        System.setProperty("webdriver.gecko.driver", "C:\\Users\\VG2G0EE\\seleniumFox\\geckodriver.exe");
        FirefoxDriver driver = new FirefoxDriver();
        driver.get("https://demoqa.com/text-box");

        //Excel

        //(1)Mediante el Excel ingresamos mediante columanas los inputs del formulario
        //(2)Agregamos una sendencia en el POM, llamada POI para conectar Selenium con un Excel.
        //(3)Guardar el archivo de Excel en la carpeta de Resourse, del mismo proyecto.
            //Click derecho /Open in/Abrimos la carpeta de Resourse/Guardamos en la ruta el Archivo excel

        //Carga de datos en Excel
            //Importar el excel al proyecto
        try (InputStream excelFile = ClassLoader.getSystemResourceAsStream("formulario.xlsx");
             Workbook workbook = new XSSFWorkbook(excelFile);
        ) {
            //Instrrucciones para leer excel

            //Leer la primera Hoja del Excel
            Sheet dataTypesSheet = workbook.getSheetAt(0);

            //Iteramos los registros del Excel
            Iterator<Row> iterador = dataTypesSheet.iterator();
            //Saltamos la primara fila del Excel
            if (iterador.hasNext()) {
                iterador.next();
            }
            //hasNexte=ir al siguiente renglon

            //Itermaos sobre las filas que si queremos ingresar en el formulario, mientras el excel tenga información

            while (iterador.hasNext()) {

                Row currentRow = iterador.next();
                //Accedemos a la información de cada columana del Excel
                Cell fullNameCell = currentRow.getCell(0);
                Cell emailCell = currentRow.getCell(1);
                Cell currentAdresCell = currentRow.getCell(2);
                Cell permaentAdresCell = currentRow.getCell(3);

                //LLenamos el formulario desde Excel a la pagina Web

                WebElement fullNameInput = driver.findElement(By.id("userName"));
                WebElement emailInput = driver.findElement(By.id("userEmail"));
                WebElement currentInput = driver.findElement(By.id("currentAddress"));
                WebElement permanentInput = driver.findElement(By.id("permanentAddress"));
                WebElement sendButtom = driver.findElement(By.id("submit"));

                //Mandar la informacion del excel al navegador

                fullNameInput.sendKeys(fullNameCell.getStringCellValue());
                emailInput.sendKeys(emailCell.getStringCellValue());
                currentInput.sendKeys(currentAdresCell.getStringCellValue());
                permanentInput.sendKeys(permaentAdresCell.getStringCellValue());

                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("window.scrollBy(0,350)");

                //Enviar informacion
                sendButtom.click();

                JavascriptExecutor js2 = (JavascriptExecutor) driver;
                js.executeScript("window.scrollBy(0,-350)");

                //Espara de ejecucion 2 Segundos
                Thread.sleep(3000);

                //Una ves llenado y enviado la información de cada renglon del Exel, debemos  borrar los campos para pasar al
                //Siguiente renglon

                fullNameInput.clear();
                emailInput.clear();
                currentInput.clear();
                permanentInput.clear();

            }




        } catch (IOException e) {

            e.printStackTrace();

        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        } finally {

        }

    }


}
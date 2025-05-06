/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.ase.polizas;

//Librerias utilizadas
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.text.SimpleDateFormat;
import java.util.Date;
import javax.swing.JOptionPane;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Clase utilizada para la creación de las pólizas de Ingresos
 * @author Rodriguez Santiago Yazmin L.
 */
public class JavaPoiUtilsIngresos {
    
    // Variables para guardar los libros de Excel
    XSSFWorkbook catalogo = null, conciliacion = null;
    // Variables para el texto de las pólizas 
    String polizaDescripcion = "", polizaConcepto = "", polizaReferencia1 = "", polizaReferencia2 = "";
    // Variable para guardar la unidad de negocio
    String UnidadNegocio = "";
    // Variable para dar formato a la fecha con dia/mes/año
    SimpleDateFormat DateFormat = new SimpleDateFormat("dd/MM/yyyy");
    // Variable para dar formato con dos decimales a los valores
    DecimalFormat df = new DecimalFormat("#.00");
    // Array para cuardar las cuentas del catálogo 
    ArrayList<Cuenta> cuentasList = new ArrayList<>();
    // Array para guardar las filas con las sumas encontradas
    ArrayList<Integer> filaSumas = new ArrayList<>();
    // Array para guardar que sumas ya se realizarón
    ArrayList<String> sumasRealizadas = new ArrayList<>();
    // Variables booleanas para tipo de póliza y si el archivo ya esta finalizado
    Boolean tipo_poliza, finalizado = false;
    // Arreglo de String con el encabezado de los titulos
    String[] cabecera = {"FECHA", "TIPO PÓLIZA", "No. DE PÓLIZA", "DESCRIPCIÓN GENERAL PÓLIZA (OPCIONAL)", "No. CTA.", "NOMBRE DE LA CUENTA (REQUERIDO)", "CONCEPTO", "REFERENCIA (OPCIONAL)", "REFERENCIA 2 (OPCIONAL)", "CARGO", "ABONO"};

    /**
     * Este método recibe un archivo File y lo trasforma a Workbook
     * Recibe un archivo tipo File
     * @param excelFile
     * @return XSSFWorkbook
     */
    public XSSFWorkbook trasformaWB(File excelFile) {
        try {
            InputStream myFile = new FileInputStream(excelFile);
            XSSFWorkbook wb = new XSSFWorkbook(myFile);
            return wb;
        } catch (Exception e) {
            // TODO: handle exception
            System.out.println(e.getMessage());
        }
        return null;
    }
    
    public void recorreSumas(){
        for (int i = 0; i < sumasRealizadas.size(); i++) {
            System.out.println("suma encontrada: " + sumasRealizadas.get(i));
        }
    }

    /**
     * INICIO DE SEGMENTO DE LECTURA DE CATALOGO
     *
     * Este método se encarga de buscar en el catálogo la columna
     * correspondiente al número de cuenta una vez encontrada llama a guardar la
     * cuenta
     */
    public void leerCatalogo() {
        Sheet sheet = catalogo.getSheetAt(0);
        Row row = sheet.getRow(3);
        String num, des, con;
        int rowcount = sheet.getLastRowNum();

        for (int iRow = 3; iRow < rowcount; iRow++) {
            if (row != null) {
                row = sheet.getRow(iRow);
                System.out.println(getCeldaSC(iRow, 0));
                if (getCeldaSC(iRow, 0) != null) {
                    con = getCeldaSC(iRow, 0);
                    num = getCeldaSC(iRow, 1);
                    des = getCeldaSC(iRow, 2);
                    if (con != null) {
                        Cuenta c = new Cuenta(con, num, des);
                        this.cuentasList.add(c);
                    }
                }
            }
        }
    }

    /**
    * Método para recorrer el Array con las cuentas del catálogo
    */
    public void leerCuentas() {
        for (int i = 0; i < this.cuentasList.size(); i++) {
            Cuenta c = cuentasList.get(i);
            //System.out.println("Concepto: " + c.getConcepto() + " Cuenta: " + c.getNumeroCuenta() + " Descripción: " + c.getDescripcion());
        }
    }

    /**
     * Este método busca por el concepto de la cuenta y regresa la Clase Cuenta
     * Recibe 
     * @param desc
     * @return
     */
    public Cuenta getCuenta(String desc) {
        for (int i = 0; i < cuentasList.size(); i++) {
            if (cuentasList.get(i).getConcepto() != null) {
                if (cuentasList.get(i).getConcepto().equals(desc)) {
                    return cuentasList.get(i);
                }
            }
        }
        return null;
    }
    /**
     * Método que regresa el valor de las celdas del catálogo 
     * tipo String recibe la fila y la celda a obtener
     * @param fila (row)
     * @param celda (celda)
     * @return String retorna un string con el valor de la celda
     */
    public String getCeldaSC(int fila, int celda) {
        Sheet sheet = catalogo.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return null;
                case Cell.CELL_TYPE_BLANK:
                    return null;
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue().trim();
            }
        }
        return null;
    }

    // FIN DE SEGMENTO DE LECTURA DE CATALOGO 
    // INICIO DE SEGMENTO DE LECTURA DE CONCILIACIÓN
    /**
     * Método que obtiene el valor de una celda de la conciliación
     * tipo String, recibe la fila y la celda que se busca obtener
     * @param fila (row)
     * @param celda (cell)
     * @return String retorna un string con el valor de la celda
     */
    public String getCeldaS(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return null;
                case Cell.CELL_TYPE_BLANK:
                    return null;
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue().trim();
            }
        }
        return null;
    }

    /**
     * Método que obtiene el valor de una celda de la conciliación
     * tipo Numerica, recibe la fila y la celda que se busca obtener     *
     * @param fila
     * @param celda
     * @return Double con el valor de la celda númerica
     */
    public double getCeldaD(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            return cell.getNumericCellValue();
        }
        return 0.0;
    }

    /**
     * Método que obtiene el valor de una celda tipo Numerica recibe la fila y
     * celda para retornar el valor double
     *
     * @param fila
     * @param celda
     * @return int con el valor de la celda numerica
     */
    public int getCeldaN(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return (int) cell.getNumericCellValue();
                case Cell.CELL_TYPE_STRING:
                    return 0;
            }
        }
        return 0;
    }

    /**
     * Método que obtiene el valor de una celda tipo Date recibe la fila y celda
     * para retornar el valor en fecha
     *
     * @param fila
     * @param celda
     * @return double con el valor de la celda numerica
     */
    public Date getCeldaF(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    return null;
                case Cell.CELL_TYPE_STRING:
                    return null;
                case Cell.CELL_TYPE_NUMERIC:
                    return cell.getDateCellValue();
            }
        }
        return null;
    }

    public void getTipo(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    System.out.println("Blank");
                    break;
                case Cell.CELL_TYPE_STRING:
                    System.out.println("String");
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    System.out.println("Boolean");
                    break;
                case Cell.CELL_TYPE_ERROR:
                    System.out.println("Error");
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    System.out.println("Formula");
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    System.out.println("Numeric");
                    break;
            }
        }
    }

    public String getUUID(int fila) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(9);
        if (cell_id != null) {
            //System.out.println("UUDI" + getCeldaS(fila, 9));
            return getCeldaS(fila, 9);
        }
        return null;
    }

    public String getIdPoliza(int fila) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(54);
        if (cell_id != null) {
            //System.out.println("ID" + getCeldaS(fila, 54));
            return getCeldaS(fila, 54);
        }
        return null;
    }

    public String getFecha(int fila) {
        if (this.conciliacion != null) {
            Sheet sheet = conciliacion.getSheetAt(0);

            Row row = sheet.getRow(fila);
            Cell cell_fecha1 = row.getCell(36);
            Cell cell_fecha2 = row.getCell(5);
            if (getCeldaF(fila, 36) != null) {
                Date fecha = getCeldaF(fila, 36);
                String curr_date = DateFormat.format(fecha);
                //System.out.println(curr_date);
                return curr_date;

            }
            if (getCeldaF(fila, 5) != null) {
                Date fecha = getCeldaF(fila, 5);
                String curr_date = DateFormat.format(fecha);
                //System.out.println(curr_date);
                return curr_date;
            }
        }
        return null;
    }

    /**
     * Método que guarda la unidad de negocio y obtiene el identificador de la
     * unidad de negocio
     *
     * @param fila
     * @param celda
     * @param libro
     * @return
     */
    public String getUN(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        String un = "";
        if (cell != null) {
            if (getCeldaS(fila, celda) != null) {
                un = getCeldaS(fila, celda);
                //System.out.println(un);
                this.UnidadNegocio = getCeldaS(fila, celda);
                int n = 6;
                return StringUtils.right(this.UnidadNegocio, n);
            }
        }
        return null;
    }

    /**
     * Método que obtiene el Número de Proceso
     *
     * @param fila
     * @param celda
     * @return
     */
    public Integer getNumProceso(int fila, int celda) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return (int) cell.getNumericCellValue();
                case Cell.CELL_TYPE_STRING:
                    return null;
                case Cell.CELL_TYPE_BLANK:
                    return null;
            }
        }
        return null;
    }

    /**
     * Método que obtiene la descripción de la poliza de la concatenación de las
     * columnas + 14 + 16 + 20
     *
     * @param fila recibe la fila (row)
     * @return String la descripción de la poliza
     */
    public String getDescripcionPoliza(int fila) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        if (row != null) {
            Cell cell_uudi = row.getCell(9);
            Cell cell_id = row.getCell(55);
            Cell cell_factura = row.getCell(65);
            String cell_string_uuid = getCeldaS(fila, 9);
            // variables para nomina
            String ingresos = "INGRESOS POR ";
            String ID = "", servicio = "", un = "", proceso = "", periodo = "", mes = "", trabajador = "", complemento = "";
            if (cell_uudi != null) {
                if (getCeldaS(fila, 52) == null) {
                    if (getCeldaS(fila, 21) != null) {
                        servicio = getCeldaS(fila, 21) + " ";
                    }
                    if (getCeldaS(fila, 55) != null) {
                        ID = cell_id.getStringCellValue().trim() + " ";
                    }
                    if (getUN(fila, 48) != null) {
                        un = "UN " + getUN(fila, 48) + " ";
                    } else {
                        un = "";
                    }
                    if (getNumProceso(fila, 50) != null) {
                        proceso = getNumProceso(fila, 50) + " ";
                    }
                    if (getCeldaS(fila, 54) != null) {
                        complemento = getCeldaS(fila, 54) + " ";
                    }
                    if (getCeldaS(fila, 51) != null) {
                        periodo = getCeldaS(fila, 51) + " ";
                    }
                    if (getCeldaS(fila, 56) != null) {
                        mes = getCeldaS(fila, 56);
                    }
                    //if (this.tipo_poliza == true) {
                    this.polizaDescripcion = ingresos + servicio + ID + un + proceso + periodo + complemento + mes;
                    //} else {
                    //    this.polizaDescripcion = nomina + cliente + ID + periodo + complemento + mes;
                    //}
                    //System.out.println("Descripción: "+this.polizaDescripcion);
                    return this.polizaDescripcion;
                } else {
                    String finiquito = "FINIQUITO ";
                    if (getCeldaS(fila, 21) != null) {
                        servicio = getCeldaS(fila, 21) + " ";
                    }
                    if (getUN(fila, 48) != null) {
                        un = "UN " + getUN(fila, 48) + " ";
                    } else {
                        un = "";
                    }
                    if (getNumProceso(fila, 50) != null) {
                        proceso = getNumProceso(fila, 50) + " ";
                    }
                    if (getCeldaS(fila, 54) != null) {
                        complemento = getCeldaS(fila, 54) + " ";
                    }
                    if (getCeldaS(fila, 51) != null) {
                        periodo = getCeldaS(fila, 51) + " ";
                    }
                    if (getCeldaS(fila, 53) != null) {
                        trabajador = getCeldaS(fila, 53) + " ";
                    }
                    if (getCeldaS(fila, 56) != null) {
                        mes = getCeldaS(fila, 56);
                    }
                    //if (this.tipo_poliza == true) {
                    this.polizaDescripcion = ingresos + servicio + ID + un + proceso + periodo + finiquito + trabajador + complemento + mes;
                    //} else {
                    //    this.polizaDescripcion = finiquito + cliente + ID + periodo + complemento + mes;
                    //}
                    //System.out.println("Descripción: "+this.polizaDescripcion);
                    return this.polizaDescripcion;
                }
            }
        }
        return null;
    }

    /**
     * Método que obtiene el concepto de la poliza concatenando lo obtenido de
     * las columnas: 13 + 21 del workbook prefactura almacenado
     *
     * @param fila recibe la fila (row)
     * @return String con el concepto
     */
    public String getConceptoPoliza(int fila) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        if (getCeldaS(fila, 48) != null) {
            this.polizaConcepto = getCeldaS(fila, 48) + " ";
        }
        //System.out.println("Concepto: "+this.polizaConcepto);
        return this.polizaConcepto;

    }

    public String getReferenciaPoliza1(int fila) {
        XSSFWorkbook wb = this.conciliacion;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        String concepto = "", proceso = "";
        if (getCeldaS(fila, 9) != null) {
            concepto = getCeldaS(fila, 9) + " ";
        }
        if (getNumProceso(fila, 50) != null) {
            proceso = getNumProceso(fila, 50) + " ";
        }
        this.polizaReferencia1 = concepto + proceso;
        //System.out.println("Referencia 1 :"+this.polizaReferencia1);
        return this.polizaReferencia1;

    }

    /**
     * Método que obtiene la referencia dos de la poliza obtenido de la columnas
     * 20
     *
     * @param fila recibe la fila (row)
     * @return String con la referencia dos
     */
    public String getReferenciaPoliza2(int fila) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(20);
        String periodo = "", texto = "", trabajador = "", mes = "";
        if (getCeldaS(fila, 52) == null) {
            texto = "";
            if (getCeldaS(fila, 51) != null) {
                periodo = getCeldaS(fila, 51) + " ";
            }
            if (getCeldaS(fila, 56) != null) {
                mes = getCeldaS(fila, 56) + " ";
            }
            this.polizaReferencia2 = periodo + mes;
            //System.out.println("Referencia 2: "+this.polizaReferencia2);
            return this.polizaReferencia2;
        } else {
            texto = "FINIQUITO ";
            if (getCeldaS(fila, 51) != null) {
                periodo = getCeldaS(fila, 51) + " ";
            }
            if (getCeldaS(fila, 53) != null) {
                trabajador = getCeldaS(fila, 53) + " ";
            }
            if (getCeldaS(fila, 56) != null) {
                mes = getCeldaS(fila, 56) + " ";
            }
            this.polizaReferencia2 = periodo + texto + trabajador + mes;
            //System.out.println("Referencia 2: "+this.polizaReferencia2);
            return this.polizaReferencia2;
        }
    }

    /**
     * Método que obtiene la referencia dos de la poliza obtenido de la columnas
     * 20
     *
     * @param fila recibe la fila (row)
     * @return String con la referencia dos
     */
    public String getReferenciaPoliza2Mov(int fila) {
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(20);
        String periodo = "", texto = "", trabajador = "", mes = "", mov = "";
        if (getCeldaS(fila, 52) == null) {
            texto = "";
            if (getCeldaS(fila, 51) != null) {
                periodo = getCeldaS(fila, 51) + " ";
            }
            if (getCeldaS(fila, 56) != null) {
                mes = getCeldaS(fila, 56) + " ";
            }
            if (getCeldaN(fila, 40) != 0) {
                mov = "*" + getCeldaN(fila, 40) + " ";
            }
            this.polizaReferencia2 = periodo + mes + mov;
            return this.polizaReferencia2;
        } else {
            texto = "FINIQUITO ";
            if (getCeldaS(fila, 51) != null) {
                periodo = getCeldaS(fila, 51) + " ";
            }
            if (getCeldaS(fila, 53) != null) {
                trabajador = getCeldaS(fila, 53) + " ";
            }
            if (getCeldaS(fila, 56) != null) {
                mes = getCeldaS(fila, 56) + " ";
            }
            if (getCeldaN(fila, 40) != 0) {
                mov = "*" + getCeldaN(fila, 40) + " ";
            }
            this.polizaReferencia2 = periodo + texto + trabajador + mes + mov;
            return this.polizaReferencia2;
        }
    }

    /**
     * Método que obtiene el cargo o abono de una fila y retorna un double
     * recibe la fila y celda a obtener
     *
     * @param fila
     * @param col
     * @return double
     */
    public double getCargosAbono(int fila, int col) {
        double valor;
        if (getCeldaN(fila, col) >= 0.00) {
            valor = getCeldaN(fila, col);
            if (valor <= -0.00) {
                valor = valor * -1;
                return valor;
            }
            return valor;
        }
        return 0.0;
    }

    public String getTextoSuma(int fila) {
        System.out.println("Entrando al método");
        Sheet sheet = conciliacion.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(65);
        System.out.println("antes de if");
        if (cell_id != null) {
            System.out.println("Encontro texto" + getCeldaS(fila, 65));
            return getCeldaS(fila, 65);
        }
        return null;
    }

    public void CrearPolizas() {
        // Datos para la generación del archivo
        String fileName = "Layout_ingresos.xlsx";
        String desktopPath = System.getProperty("user.home") + "/Desktop";
        String filePath = desktopPath + "/" + fileName;
        String hoja = "Hoja1";
        //Libro y primera hoja creada
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet hoja1 = book.createSheet(hoja);

        CellStyle style = book.createCellStyle();
        XSSFFont font = book.createFont();
        font.setBold(true);
        font.setFontHeight(8);
        font.setFontName("Arial");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_GENERAL);
        style.setFont(font);

        XSSFWorkbook wb = this.conciliacion;
        Sheet sheet = wb.getSheetAt(0);
        hoja1.autoSizeColumn(0);

        XSSFRow row = hoja1.createRow(0);
        for (int j = 0; j < cabecera.length; j++) {
            XSSFCell cell = row.createCell(j);
            cell.setCellStyle(style);
            cell.setCellValue(cabecera[j]);
        }

        int rowcount = sheet.getLastRowNum();
        int fila_nueva = 1;
        XSSFRow row_fila;
        XSSFCell cell_tipo, cell_fecha, cell_descrip, cell_concepto, cell_ref_1, cell_ref_2, cell_num_cuenta, cell_nombre, cell_valor, cell_valor2, cell_suma;
        //JOptionPane.showMessageDialog(null, "Se esta generando tu layout");
        for (int i = 2; i <= rowcount; i++) {
            Row row_conciliacion = sheet.getRow(i);
            System.out.println("Fila: " + i);
            if (row_conciliacion != null) {
                Cell cell_uuid = row_conciliacion.getCell(9);
                if (cell_uuid != null) {
                    if (getTextoSuma(i) != null) {
                        String busqueda = getTextoSuma(i);
                        System.out.println("texto busqueda" + busqueda);
                        boolean existe = sumasRealizadas.contains(busqueda);
                        
                        if (!existe) {
                            sumasRealizadas.add(getTextoSuma(i));
                            getFilasSumas(getTextoSuma(i));

                            //Creción de una nueva fila
                            row_fila = hoja1.createRow(fila_nueva);
                            // Creación de la celda fecha
                            if (getFecha(i) != null) {
                                cell_fecha = row_fila.createCell(0);
                                cell_fecha.setCellValue(getFecha(i));
                            }
                            //Tipo
                            cell_tipo = row_fila.createCell(1);
                            cell_tipo.setCellValue(getCeldaS(i, 67));
                            //Descripción
                            cell_descrip = row_fila.createCell(3);
                            cell_descrip.setCellValue(getDescripcionPoliza(i));
                            //Concepto
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            if (getCeldaS(i, 35) != null) {
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2Mov(i));
                            } else {
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_ref_1 = row_fila.createCell(7);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            if (!filaSumas.isEmpty()) {
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    if (getCeldaS(filaSumas.get(j), 35) != null && getCargosAbono(filaSumas.get(j), 41) != 0.00) { //Bancos
                                        if (getCuenta(getCeldaS(filaSumas.get(j), 35)) != null) {
                                            cell_nombre.setCellValue(getCuenta(getCeldaS(filaSumas.get(j), 35)).getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(filaSumas.get(j), 35)).getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 41)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia uno 
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia dos 
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));

                                        }
                                    }
                                }
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    if (getCeldaS(filaSumas.get(j), 48) != null) {
                                        if (getCuenta(getCeldaS(filaSumas.get(j), 48)) != null) {
                                            cell_nombre.setCellValue(getCuenta(getCeldaS(filaSumas.get(j), 48)).getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(filaSumas.get(j), 48)).getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 17)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    }
                                }
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    getTipo(filaSumas.get(j), 41);
                                    if (getCargosAbono(filaSumas.get(j), 68) != 0.00) {
                                        if (getCuenta("DIF EN COBROS CARGO") != null) {
                                            cell_nombre.setCellValue(getCuenta("DIF EN COBROS CARGO").getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS CARGO").getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 68)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    }
                                }
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    if (getCeldaS(filaSumas.get(j), 48) != null) {
                                        if (getCuenta(getCeldaS(filaSumas.get(j), 48)) != null) {
                                            cell_nombre.setCellValue(getCuenta(getCeldaS(i, 48)).getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 48)).getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor2.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 17)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    }
                                }
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    //ABONOS 
                                    if (getCargosAbono(filaSumas.get(j), 15) >= 0.000) { //Empresa
                                        if (getCuenta(getCeldaS(filaSumas.get(j), 47)) != null) {
                                            cell_nombre.setCellValue(getCuenta(getCeldaS(filaSumas.get(j), 47)).getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(filaSumas.get(j), 47)).getNumeroCuenta());
                                            cell_valor2.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 15)));
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    }
                                }
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    if (getCeldaS(filaSumas.get(j), 35) != null) {
                                        if (getCuenta("IVA TRASLADADO") != null) {
                                            cell_nombre.setCellValue(getCuenta("IVA TRASLADADO").getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta("IVA TRASLADADO").getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor2.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 16)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    } else {
                                        if (getCuenta("IVA POR TRASLADAR") != null) {
                                            cell_nombre.setCellValue(getCuenta("IVA POR TRASLADAR").getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta("IVA POR TRASLADAR").getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor2.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 16)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    }
                                }
                                for (int j = 0; j < filaSumas.size(); j++) {
                                    if (getCargosAbono(filaSumas.get(j), 69) > 0.000) {
                                        if (getCuenta("DIF EN COBROS ABONO") != null) { //IVA
                                            cell_nombre.setCellValue(getCuenta("DIF EN COBROS ABONO").getDescripcion());
                                            cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS ABONO").getNumeroCuenta());
                                            cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            cell_valor2.setCellValue(df.format(getCargosAbono(filaSumas.get(j), 69)));
                                            fila_nueva++;
                                            row_fila = hoja1.createRow(fila_nueva);
                                            cell_num_cuenta = row_fila.createCell(4);
                                            cell_nombre = row_fila.createCell(5);
                                            cell_valor = row_fila.createCell(9);
                                            cell_valor2 = row_fila.createCell(10);
                                            //Concepto 
                                            cell_concepto = row_fila.createCell(6);
                                            cell_concepto.setCellValue(getConceptoPoliza(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_1 = row_fila.createCell(7);
                                            //cell_ref_1.setCellValue(getReferenciaPoliza1(filaSumas.get(j)));
                                            //Referencia
                                            cell_ref_2 = row_fila.createCell(8);
                                            cell_ref_2.setCellValue(getReferenciaPoliza2(filaSumas.get(j)));
                                            //Sumas
                                            cell_suma = row_fila.createCell(11);
                                            cell_suma.setCellValue(getTextoSuma(i));
                                        }
                                    }
                                }
                                filaSumas.clear();
                            }
                        }
                    } else {//Creción de una nueva fila
                        row_fila = hoja1.createRow(fila_nueva);
                        // Creación de la celda fecha
                        if (getFecha(i) != null) {
                            cell_fecha = row_fila.createCell(0);
                            cell_fecha.setCellValue(getFecha(i));
                        }
                        //Tipo
                        cell_tipo = row_fila.createCell(1);
                        cell_tipo.setCellValue(getCeldaS(i, 67));
                        //Descripción
                        cell_descrip = row_fila.createCell(3);
                        cell_descrip.setCellValue(getDescripcionPoliza(i));
                        //Concepto
                        cell_concepto = row_fila.createCell(6);
                        cell_concepto.setCellValue(getConceptoPoliza(i));
                        //Referencia uno 
                        cell_ref_1 = row_fila.createCell(7);
                        cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                        //Referencia dos 
                        if (getCeldaS(i, 35) != null) {
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2Mov(i));
                        } else {
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                        cell_num_cuenta = row_fila.createCell(4);
                        cell_nombre = row_fila.createCell(5);
                        cell_valor = row_fila.createCell(9);
                        cell_valor2 = row_fila.createCell(10);
                        if (getCeldaS(i, 35) != null) { //Bancos
                            if (getCuenta(getCeldaS(i, 35)) != null) {
                                cell_nombre.setCellValue(getCuenta(getCeldaS(i, 35)).getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 35)).getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 41)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia uno 
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia dos 
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                        if (getCeldaS(i, 48) != null) {
                            if (getCuenta(getCeldaS(i, 48)) != null) {
                                cell_nombre.setCellValue(getCuenta(getCeldaS(i, 48)).getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 48)).getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 17)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                        //Abonos
                        if (getCeldaS(i, 48) != null) {
                            if (getCuenta(getCeldaS(i, 48)) != null) {
                                cell_nombre.setCellValue(getCuenta(getCeldaS(i, 48)).getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 48)).getNumeroCuenta());
                                cell_valor2.setCellValue(df.format(getCargosAbono(i, 17)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                        if (getCargosAbono(i, 68) != 0.00) {
                            if (getCuenta("DIF EN COBROS CARGO") != null) {
                                cell_nombre.setCellValue(getCuenta("DIF EN COBROS CARGO").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS CARGO").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 68)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                        //ABONOS 
                        if (getCargosAbono(i, 15) != 0.0) { //Empresa
                            if (getCuenta(getCeldaS(i, 47)) != null) {
                                cell_nombre.setCellValue(getCuenta(getCeldaS(i, 47)).getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 47)).getNumeroCuenta());
                                cell_valor2.setCellValue(df.format(getCargosAbono(i, 15)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                        if (getCeldaS(i, 35) != null) {
                            if (getCuenta("IVA TRASLADADO") != null) {
                                cell_nombre.setCellValue(getCuenta("IVA TRASLADADO").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("IVA TRASLADADO").getNumeroCuenta());
                                cell_valor2.setCellValue(df.format(getCargosAbono(i, 16)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        } else {
                            if (getCuenta("IVA POR TRASLADAR") != null) {
                                cell_nombre.setCellValue(getCuenta("IVA POR TRASLADAR").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("IVA POR TRASLADAR").getNumeroCuenta());
                                cell_valor2.setCellValue(df.format(getCargosAbono(i, 16)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                        if (getCargosAbono(i, 69) != 0.0) {
                            if (getCuenta("DIF EN COBROS ABONO") != null) { //IVA
                                cell_nombre.setCellValue(getCuenta("DIF EN COBROS ABONO").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS ABONO").getNumeroCuenta());
                                cell_valor2.setCellValue(df.format(getCargosAbono(i, 69)));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                                //Referencia
                                cell_ref_2 = row_fila.createCell(8);
                                cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                            }
                        }
                    }
                }
            }
        }
        recorreSumas();
        System.out.println("For Terminado");
        File excelFile;
        excelFile = new File(filePath); // Referenciando a la ruta y el archivo Excel a crear
        try ( FileInputStream fileIuS = new FileInputStream(excelFile)) {
            fileIuS.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        try ( FileOutputStream fileOuS = new FileOutputStream(excelFile)) {
            if (excelFile.exists()) { // Si el archivo existe lo eliminaremos
                excelFile.delete();
            }
            System.out.println("Creando archivo");
            book.write(fileOuS);
            fileOuS.flush();
            fileOuS.close();
            System.out.println("Archivo creado");
            this.finalizado = true;
            JOptionPane.showMessageDialog(null, "Se ha generado tu archivo");

        } catch (FileNotFoundException e) {
            this.finalizado = true;
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "No se puede crear el layout ya que el archivo se encuentra abierto");

        } catch (Exception e) {
            this.finalizado = true;
            e.printStackTrace();
        }
    }

    /**
     * Método que obtiene el listado de las filas de bancos con el mismo
     * identificardor y de tipo de operación 4
     *
     * @param id
     */
    public void getFilasSumas(String texto) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);

            Row row;
            Cell cell_sumas;
            int rowcount = sheet.getLastRowNum();
            String suma = null;
            for (int i = 1; i <= rowcount; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    cell_sumas = row.getCell(65);
                    if (cell_sumas != null) {
                        switch (cell_sumas.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                break;
                            case Cell.CELL_TYPE_STRING:
                                suma = cell_sumas.getStringCellValue().trim();
                                if (suma.equals(texto)) {
                                    filaSumas.add(i);
                                }
                                break;
                        }
                    }
                }
            }
        }
    }

    public void createExcel() {
        // Datos para la generación del archivo
        String fileName = "Layout_ingresos.xlsx";
        String desktopPath = System.getProperty("user.home") + "/Desktop";
        String filePath = desktopPath + "/" + fileName;
        String hoja = "Hoja1";
        //Libro y primera hoja creada
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet hoja1 = book.createSheet(hoja);

        CellStyle style = book.createCellStyle();
        XSSFFont font = book.createFont();
        font.setBold(true);
        font.setFontHeight(8);
        font.setFontName("Arial");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_GENERAL);
        style.setFont(font);

        XSSFWorkbook wb = this.conciliacion;
        Sheet sheet = wb.getSheetAt(0);
        hoja1.autoSizeColumn(0);

        XSSFRow row = hoja1.createRow(0);
        for (int j = 0; j < cabecera.length; j++) {
            XSSFCell cell = row.createCell(j);
            cell.setCellStyle(style);
            cell.setCellValue(cabecera[j]);
        }
        int rowcount = sheet.getLastRowNum();
        int fila_nueva = 1;
        String id_prefactura, suma;
        XSSFRow row_fila;
        XSSFCell cell_tipo, cell_fecha, cell_descrip, cell_concepto, cell_ref_1, cell_ref_2, cell_num_cuenta, cell_nombre, cell_valor, cell_valor2;
        JOptionPane.showMessageDialog(null, "Se esta generando tu layout");
        for (int i = 2; i <= rowcount; i++) {
            Row row_conciliacion = sheet.getRow(i);
            if (row_conciliacion != null) {
                Cell cell_sumas = row_conciliacion.getCell(65);
                Cell cell_uuid = row_conciliacion.getCell(9);
                if (cell_uuid != null) {
                    //Creción de una nueva fila
                    row_fila = hoja1.createRow(fila_nueva);
                    //obtener el id que se busca en la orden de pago de conciliación
                    id_prefactura = getIdPoliza(i);
                    //getFilasBancos();
                    // Creación de la celda fecha
                    if (getFecha(i) != null) {
                        cell_fecha = row_fila.createCell(0);
                        cell_fecha.setCellValue(getFecha(i));
                    }
                    //Tipo
                    cell_tipo = row_fila.createCell(1);
                    cell_tipo.setCellValue(getCeldaS(i, 67));
                    //Descripción
                    cell_descrip = row_fila.createCell(3);
                    cell_descrip.setCellValue(getDescripcionPoliza(i));
                    //Concepto
                    cell_concepto = row_fila.createCell(6);
                    cell_concepto.setCellValue(getConceptoPoliza(i));
                    //Referencia uno 
                    cell_ref_1 = row_fila.createCell(7);
                    cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                    //Referencia dos 
                    if (getCeldaS(i, 35) != null) {
                        cell_ref_2 = row_fila.createCell(8);
                        cell_ref_2.setCellValue(getReferenciaPoliza2Mov(i));
                    } else {
                        cell_ref_2 = row_fila.createCell(8);
                        cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                    }
                    cell_num_cuenta = row_fila.createCell(4);
                    cell_nombre = row_fila.createCell(5);
                    cell_valor = row_fila.createCell(9);
                    cell_valor2 = row_fila.createCell(10);

                    if (getCeldaS(i, 35) != null) { //Bancos
                        if (getCuenta(getCeldaS(i, 35)) != null) {
                            cell_nombre.setCellValue(getCuenta(getCeldaS(i, 35)).getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 35)).getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 41)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia uno 
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia dos 
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                    if (getCeldaS(i, 48) != null) {
                        if (getCuenta(getCeldaS(i, 48)) != null) {
                            cell_nombre.setCellValue(getCuenta(getCeldaS(i, 48)).getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 48)).getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 17)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                    if (getCargosAbono(i, 66) != 0.0) {
                        if (getCuenta("DIF EN COBROS CARGO") != null) {
                            cell_nombre.setCellValue(getCuenta("DIF EN COBROS CARGO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS CARGO").getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 66)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                    //ABONOS 
                    if (getCargosAbono(i, 15) != 0.0) { //Empresa
                        if (getCuenta(getCeldaS(i, 47)) != null) {
                            cell_nombre.setCellValue(getCuenta(getCeldaS(i, 47)).getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 47)).getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 15)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                    if (getCeldaS(i, 35) != null) {
                        if (getCuenta("IVA TRASLADADO") != null) {
                            cell_nombre.setCellValue(getCuenta("IVA TRASLADADO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("IVA TRASLADADO").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 16)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    } else {
                        if (getCuenta("IVA POR TRASLADAR") != null) {
                            cell_nombre.setCellValue(getCuenta("IVA POR TRASLADAR").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("IVA POR TRASLADAR").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 16)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                    if (getCeldaS(i, 48) != null) {
                        if (getCuenta(getCeldaS(i, 48)) != null) {
                            cell_nombre.setCellValue(getCuenta(getCeldaS(i, 48)).getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta(getCeldaS(i, 48)).getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 17)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                    if (getCargosAbono(i, 67) != 0.0) {
                        if (getCuenta("DIF EN COBROS ABONO") != null) { //IVA
                            cell_nombre.setCellValue(getCuenta("DIF EN COBROS ABONO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS ABONO").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 67)));
                            fila_nueva++;
                            row_fila = hoja1.createRow(fila_nueva);
                            cell_num_cuenta = row_fila.createCell(4);
                            cell_nombre = row_fila.createCell(5);
                            cell_valor = row_fila.createCell(9);
                            cell_valor2 = row_fila.createCell(10);
                            //Concepto 
                            cell_concepto = row_fila.createCell(6);
                            cell_concepto.setCellValue(getConceptoPoliza(i));
                            //Referencia
                            cell_ref_1 = row_fila.createCell(7);
                            cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            //Referencia
                            cell_ref_2 = row_fila.createCell(8);
                            cell_ref_2.setCellValue(getReferenciaPoliza2(i));
                        }
                    }
                }
            }
        }
        System.out.println("For terminado");

        File excelFile;
        excelFile = new File(filePath); // Referenciando a la ruta y el archivo Excel a crear
        try ( FileOutputStream fileOuS = new FileOutputStream(excelFile)) {
            if (excelFile.exists()) { // Si el archivo existe lo eliminaremos
                excelFile.delete();
            }
            System.out.println("Creando archivo");
            book.write(fileOuS);
            fileOuS.flush();
            fileOuS.close();
            JOptionPane.showMessageDialog(null, "Se ha generado tu archivo");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public XSSFWorkbook getCatalogo() {
        return catalogo;
    }

    public void setCatalogo(XSSFWorkbook catalogo) {
        this.catalogo = catalogo;
    }

    public XSSFWorkbook getConciliacion() {
        return conciliacion;
    }

    public void setConciliacion(XSSFWorkbook conciliacion) {
        this.conciliacion = conciliacion;
    }

    public ArrayList<Cuenta> getCuentasList() {
        return cuentasList;
    }

    public void setCuentasList(ArrayList<Cuenta> cuentasList) {
        this.cuentasList = cuentasList;
    }

    public Boolean getTipo_poliza() {
        return tipo_poliza;
    }

    public void setTipo_poliza(Boolean tipo_poliza) {
        this.tipo_poliza = tipo_poliza;
    }

    public Boolean getFinalizado() {
        return finalizado;
    }

    public void setFinalizado(Boolean finalizado) {
        this.finalizado = finalizado;
    }

}

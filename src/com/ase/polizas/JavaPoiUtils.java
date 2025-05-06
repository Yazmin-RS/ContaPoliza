/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.ase.polizas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import javax.swing.JOptionPane;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Rodriguez Santiago Yazmin L.
 */
public class JavaPoiUtils {

    XSSFWorkbook catalogo = null, prefactura = null, conciliacion = null;
    ArrayList<Cuenta> cuentasList = new ArrayList<>();
    ArrayList<Integer> filasBancos = new ArrayList<>();
    SimpleDateFormat DateFormat = new SimpleDateFormat("dd/MM/yyyy");
    Integer[] id;
    String UnidadNegocio = "", periodo = "", trabajador = "", mes_p = "", mes_c = "";
    Integer numProceso, hojaExcel;
    Boolean tipo_poliza, lugar_poliza, finalizado = false;
    String polizaDescripcion = "", polizaConcepto = "", polizaReferencia1 = "", polizaReferencia2 = "";
    DecimalFormat df = new DecimalFormat("#.00");
    String[] cabecera = {"FECHA", "TIPO PÓLIZA", "No. DE PÓLIZA", "DESCRIPCIÓN GENERAL PÓLIZA (OPCIONAL)", "No. CTA.", "NOMBRE DE LA CUENTA (REQUERIDO)", "CONCEPTO", "REFERENCIA (OPCIONAL)", "REFERENCIA 2 (OPCIONAL)", "CARGO", "ABONO"};

    /**
     * Este método recibe un archivo File y lo trasforma a Workbook
     *
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

    /**
     * INICIO DE SEGMENTO DE LECTURA DE CATALOGO
     *
     * Este método se encarga de buscar en el catálogo la columna
     * correspondiente al número de cuenta una vez encontrada llama a guardar la
     * cuenta
     */
    public void leerCatalogo() {
        XSSFWorkbook wb = this.catalogo;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(3);
        Cell cell_des;
        Cell cell_num;
        Cell cell_con;
        String num = "", des = "", con = "";
        int rowcount = sheet.getLastRowNum();

        for (int iRow = 3; iRow < rowcount; iRow++) {
            if (row != null) {
                cell_des = row.getCell(2);
                cell_num = row.getCell(1);
                cell_con = row.getCell(0);
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
            row = sheet.getRow(iRow);
        }
    }

    public void leerCuentas() {
        for (int i = 0; i < this.cuentasList.size(); i++) {
            Cuenta c = cuentasList.get(i);
            //System.out.println("Concepto: " + c.getConcepto() + " Cuenta: " + c.getNumeroCuenta() + " Descripción: " + c.getDescripcion());
        }
    }

    /**
     * Este método busca por descripción de la cuenta y regresa su número de
     * cuenta
     *
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

    // FIN DE SEGMENTO DE LECTURA DE CATALOGO 
    // INICIO DE SEGMENTO DE LECTURA DE PREFACTURA
    public String getIdPoliza(int fila) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(23);
        if (cell_id != null) {
            //System.out.println(getCeldaS(fila, 23));
            return getCeldaS(fila, 23);
        }
        return null;
    }

    public String getMesConciliacion(int fila) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(fila);
            Cell cell = row.getCell(30);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return null;
                    case Cell.CELL_TYPE_BLANK:
                        return null;
                    case Cell.CELL_TYPE_ERROR:
                        return null;
                    case Cell.CELL_TYPE_STRING:
                        //System.out.println("Mes conciliación:" +cell.getStringCellValue().trim());
                        //System.out.println("Fila "+ fila);
                        return cell.getStringCellValue().trim();
                }
            }
            return null;
        }
        return null;
    }

    public String getMesPrefactura(int fila) {
        if (this.prefactura != null) {
            XSSFWorkbook wb = this.prefactura;
            Sheet sheet = wb.getSheetAt(0);

            Row row = sheet.getRow(fila);
            Cell cell = row.getCell(20);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return null;
                    case Cell.CELL_TYPE_BLANK:
                        return null;
                    case Cell.CELL_TYPE_ERROR:
                        return null;
                    case Cell.CELL_TYPE_STRING:
                        return cell.getStringCellValue().trim();
                }
            }
            return null;

        }
        return null;
    }

    /**
     * Método que obtiene la descripción de la poliza de la concatenación de las
     * columnas 13 + 14 + 16 + 20
     *
     * @param fila recibe la fila (row)
     * @return String la descripción de la poliza
     */
    public String getDescripcionPoliza(int fila) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        if (row != null) {
            Cell cell_id = row.getCell(23);
            // variables para nomina
            String cliente = "", un = "", proceso = "", periodo = "", mes = "", trabajador = "", complemento = "";
            if (cell_id != null) {
                String ID = cell_id.getStringCellValue().trim() + " ";
                if (getCeldaS(fila, 17) == null) {
                    String nomina = "NOMINA ";

                    if (getCeldaS(fila, 10) != null) {
                        cliente = getCeldaS(fila, 10) + " ";
                    }
                    if (getUN(fila, 13) != null) {
                        un = "UN " + getUN(fila, 13) + " ";
                    }
                    if (getNumProceso(fila, 14) != null) {
                        proceso = getNumProceso(fila, 14) + " ";
                    }
                    if (getCeldaS(fila, 16) != null) {
                        periodo = getCeldaS(fila, 16) + " ";
                    }
                    if (getCeldaS(fila, 19) != null) {
                        complemento = getCeldaS(fila, 19) + " ";
                    }
                    if (getMesPrefactura(fila) != null) {
                        mes = getMesPrefactura(fila);
                    }
                    if (this.tipo_poliza == true) {
                        this.polizaDescripcion = nomina + ID + un + proceso + periodo + complemento + mes;
                    } else {
                        this.polizaDescripcion = nomina + cliente + ID + periodo + complemento + mes;
                    }
                    //System.out.println(this.polizaDescripcion);
                    return this.polizaDescripcion;
                } else {
                    String finiquito = "FINIQUITO ";
                    if (getCeldaS(fila, 18) != null) {
                        trabajador = getCeldaS(fila, 18) + " ";
                    }
                    if (getCeldaS(fila, 10) != null) {
                        cliente = getCeldaS(fila, 10) + " ";
                    }
                    if (getUN(fila, 13) != null) {
                        un = "UN " + getUN(fila, 13) + " ";
                    }
                    if (getNumProceso(fila, 14) != null) {
                        proceso = getNumProceso(fila, 14) + " ";
                    }
                    if (getCeldaS(fila, 16) != null) {
                        periodo = getCeldaS(fila, 16) + " ";
                    }
                    if (getCeldaS(fila, 19) != null) {
                        complemento = getCeldaS(fila, 19) + " ";
                    }
                    if (getMesPrefactura(fila) != null) {
                        mes = getMesPrefactura(fila);
                    }
                    //if (this.tipo_poliza == true) {
                        this.polizaDescripcion = finiquito + trabajador + cliente + ID + un + proceso + periodo + complemento + mes;
                    /*} else {
                        this.polizaDescripcion = finiquito + trabajador + ID + periodo + complemento + mes;
                    }
                    //System.out.println(this.polizaDescripcion);*/
                    return this.polizaDescripcion;
                }
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
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
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
                case Cell.CELL_TYPE_ERROR:
                    return null;
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
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            if (getCeldaS(fila, celda) != null) {
                this.lugar_poliza = true;
                this.UnidadNegocio = getCeldaS(fila, celda);
                int n = 6;
                return StringUtils.right(this.UnidadNegocio, n);
            } else {
                this.lugar_poliza = false;
            }
        }
        return "";
    }

    /**
     * Método que obtiene el concepto de la poliza concatenando lo obtenido de
     * las columnas: 13 + 21 del workbook prefactura almacenado
     *
     * @param fila recibe la fila (row)
     * @return String con el concepto
     */
    public String getConceptoPoliza(int fila) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(21);
        String cliente = "", concepto = "", division = "";
        if (getCeldaS(fila, 10) != null) {
            cliente = getCeldaS(fila, 10) + " ";
        }
        if (getUN(fila, 13) != null) {
            concepto = this.UnidadNegocio;
        }
        if (getCeldaS(fila, 21) != null) {
            division = " " + getCeldaS(fila, 21);
        }
        if (this.tipo_poliza == true) {
            this.polizaConcepto = concepto + division;
        } else {
            this.polizaConcepto = cliente;
        }
        return this.polizaConcepto;

    }

    /**
     * Método que obtiene la referencia uno de la poliza concatenando lo
     * obtenido de las columnas 14 + 16 si es nomina y 14 + 16 + 18 + 19 si es
     * finiquito
     *
     * @param fila recibe la fila (row)
     * @return String con la referencia
     */
    public String getReferenciaPoliza1(int fila) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell_id = row.getCell(23);
        // variables para nomina
        String id = "", un = "", proceso = "", periodicidad = "", complemento = "";
        // Variables para finiquito
        String trabajador = "";
        if (getCeldaS(fila, 17) == null) {
            if (getIdPoliza(fila) != null) {
                id = getIdPoliza(fila) + " ";
            }
            if (getNumProceso(fila, 14) != null) {
                proceso = getNumProceso(fila, 14) + " ";
            }
            if (getCeldaS(fila, 16) != null) {
                periodicidad = getCeldaS(fila, 16) + " ";
            }
            if (getCeldaS(fila, 19) != null) {
                complemento = getCeldaS(fila, 19);
            }

            if (this.tipo_poliza == true) {
                this.polizaReferencia1 = id + proceso + periodicidad + complemento;
            } else {
                this.polizaReferencia1 = id + periodicidad;
            }
            return this.polizaReferencia1;
        } else {
            if (getIdPoliza(fila) != null) {
                id = getIdPoliza(fila) + " ";
            }
            if (getNumProceso(fila, 14) != null) {
                proceso = getNumProceso(fila, 14) + " ";
            }
            if (getCeldaS(fila, 16) != null) {
                periodicidad = getCeldaS(fila, 16) + " ";
            }
            if (getCeldaS(fila, 18) != null) {
                trabajador = getCeldaS(fila, 18) + " ";
            }
            if (getCeldaS(fila, 19) != null) {
                complemento = getCeldaS(fila, 19);
            }
            if (this.tipo_poliza == true) {
                this.polizaReferencia1 = id + proceso + periodicidad + "F " + trabajador + complemento;
            } else {
                this.polizaReferencia1 = id + periodicidad + "F " + trabajador;
            }
            return this.polizaReferencia1;
        }
    }

    /**
     * Método que obtiene la referencia dos de la poliza obtenido de la columnas
     * 20
     *
     * @param fila recibe la fila (row)
     * @return String con la referencia dos
     */
    public String getReferenciaPoliza2prefactura(int fila) {
        if (this.prefactura != null) {
            if (getMesPrefactura(fila) != null) {
                this.mes_p = getMesPrefactura(fila);
                return getMesPrefactura(fila);
            }
        }
        return null;
    }

    /**
     * Método que obtiene la referencia dos de la poliza obtenido de la columnas
     * 20
     *
     * @param fila recibe la fila (row)
     * @return String con la referencia dos
     */
    public String getReferenciaPoliza2conciliacion(int fila) {
        if (this.conciliacion != null) {
            if (getMesConciliacion(fila) != null) {
                this.mes_c = getMesConciliacion(fila);
                return getMesConciliacion(fila);
            }
        }
        return null;
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
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return (int) cell.getNumericCellValue();
                case Cell.CELL_TYPE_STRING:
                    return 0;
                case Cell.CELL_TYPE_BLANK:
                    return 0;
                case Cell.CELL_TYPE_ERROR:
                    return 0;
            }
        }
        return 0;
    }

    /**
     * Método que obtiene el valor de una celda tipo Numerica recibe la fila y
     * celda para retornar el valor double
     *
     * @param fila
     * @param celda
     * @return double con el valor de la celda numerica
     */
    public double getCeldaD(int fila, int celda) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return cell.getNumericCellValue();
                case Cell.CELL_TYPE_STRING:
                    return 0.0;
                case Cell.CELL_TYPE_BLANK:
                    return 0.0;
                case Cell.CELL_TYPE_ERROR:
                    return 0.0;
            }
        }
        return 0.0;
    }

    /**
     * Método que obtiene el valor de una celda tipo String recibe la fila y la
     * celda del valor a obtener
     *
     * @param fila (row)
     * @param celda (cell)
     * @return String retorna un string con el valor de la celda
     */
    public String getCeldaS(int fila, int celda) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell = row.getCell(celda);
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return null;
                case Cell.CELL_TYPE_BLANK:
                    return null;
                case Cell.CELL_TYPE_ERROR:
                    return null;
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue().trim();
            }
        }
        return null;
    }

    public String getCeldaSC(int fila, int celda) {
        XSSFWorkbook wb = this.catalogo;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell;
        if (row != null) {
            cell = row.getCell(celda);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return null;
                    case Cell.CELL_TYPE_BLANK:
                        return null;
                    case Cell.CELL_TYPE_ERROR:
                        return null;
                    case Cell.CELL_TYPE_STRING:
                        return cell.getStringCellValue().trim();
                }
            }
        }
        return null;
    }

    public String getCeldaSB(int fila, int celda) {
        XSSFWorkbook wb = this.conciliacion;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell;
        if (row != null) {
            cell = row.getCell(celda);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return null;
                    case Cell.CELL_TYPE_BLANK:
                        return null;
                    case Cell.CELL_TYPE_ERROR:
                        return null;
                    case Cell.CELL_TYPE_STRING:
                        return cell.getStringCellValue().trim();
                }
            }
        }
        return null;
    }

    public String getCeldaSP(int fila, int celda) {
        XSSFWorkbook wb = this.prefactura;
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(fila);
        Cell cell;
        if (row != null) {
            cell = row.getCell(celda);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return null;
                    case Cell.CELL_TYPE_BLANK:
                        return null;
                    case Cell.CELL_TYPE_ERROR:
                        return null;
                    case Cell.CELL_TYPE_STRING:
                        return cell.getStringCellValue().trim();
                }
            }
        }
        return null;
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
        if (getCeldaN(fila, col) != 0.0) {
            valor = getCeldaD(fila, col);
            if (valor <= 0.00) {
                valor = valor * -1;
                return valor;
            }
            return valor;
        }
        return 0.0;
    }

    // FIN DE SEGMENTO DE LECTURA DE POLIZAS O PREFACTURA
    // INICIO DE SEGMENTO DE LECTURA DE CONCILIACION BANCARIA
    /**
     * Método que obtiene el listado de las filas de bancos con el mismo
     * identificardor y de tipo de operación 4
     *
     * @param id
     */
    public void getOrdenPago(String id) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);

            Row row;
            Cell cell_orden, cell_tipo;
            int rowcount = sheet.getLastRowNum();
            String order = null;
            for (int i = 1; i <= rowcount; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    cell_orden = row.getCell(32);
                    cell_tipo = row.getCell(39);
                    if (cell_orden != null) {
                        switch (cell_orden.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                break;
                            case Cell.CELL_TYPE_STRING:
                                if (cell_tipo != null) {
                                    switch (cell_tipo.getCellType()) {
                                        case Cell.CELL_TYPE_NUMERIC:
                                            order = cell_orden.getStringCellValue().trim();
                                            if (order.equals(id) && cell_tipo.getNumericCellValue() == 4) {
                                                filasBancos.add(i);
                                            } else if (order.equals(id) && cell_tipo.getNumericCellValue() == 5) {
                                                filasBancos.add(i);
                                            }
                                            break;
                                        case Cell.CELL_TYPE_STRING:
                                            break;
                                        case Cell.CELL_TYPE_BLANK:
                                            break;
                                        case Cell.CELL_TYPE_ERROR:
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                }
            }
        }
    }

    public String getFecha() {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);
            if (!filasBancos.isEmpty()) {
                Row row = sheet.getRow(filasBancos.get(0));
                Cell cell = row.getCell(4);
                if (cell != null) {
                    Date fecha = cell.getDateCellValue();
                    String curr_date = DateFormat.format(fecha);
                    //System.out.println(curr_date);
                    return curr_date;
                }
            }
        }
        return null;
    }

    public String getDescripcionBancos(int data) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);
            String valor;

            if (!filasBancos.isEmpty()) {
                Row row = sheet.getRow(filasBancos.get(data));
                Cell cell = row.getCell(2);
                if (cell != null) {
                    valor = cell.getStringCellValue();
                    return valor;
                }
            }
        }
        return null;
    }

    public String getMovimientoBancos(int filaBancos) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);
            String valor;

            if (!filasBancos.isEmpty()) {
                Row row = sheet.getRow(filasBancos.get(filaBancos));
                Cell cell_mes = row.getCell(30);
                Cell cell_mov = row.getCell(11);
                if (cell_mes != null & cell_mov != null) {
                    valor = getCeldaSB(filasBancos.get(filaBancos), 30) + " *" + (int) cell_mov.getNumericCellValue();
                    //System.out.println(valor);
                    return valor;
                }
            }
        }
        return null;
    }
    //obtienes el cargo a abonar

    public double getAbonoBancos(int filaBancos) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);
            double valor;

            if (!filasBancos.isEmpty()) {
                Row row = sheet.getRow(filasBancos.get(filaBancos));
                Cell cell = row.getCell(33);
                if (cell != null) {
                    valor = cell.getNumericCellValue();
                    return valor;
                }
            }
        }
        return 0.0;
    }

    //Obtienes el abono a cargar
    public double getCargoBancos(int filaBancos) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);
            double valor;

            if (!filasBancos.isEmpty()) {
                Row row = sheet.getRow(filasBancos.get(filaBancos));
                Cell cell = row.getCell(34);
                if (cell != null) {
                    valor = cell.getNumericCellValue();
                    return valor;
                }
            }
        }
        return 0.0;
    }

    public String getCargoAbonoBancos(int filaBancos) {
        if (this.conciliacion != null) {
            XSSFWorkbook wb = this.conciliacion;
            Sheet sheet = wb.getSheetAt(0);

            if (!filasBancos.isEmpty()) {
                Row row = sheet.getRow(filasBancos.get(filaBancos));
                Cell cell_cargo = row.getCell(33);
                if (cell_cargo != null) {
                    return "cargo";
                }
                Cell cell_abono = row.getCell(34);
                if (cell_abono != null) {
                    return "abono";
                }
            }
        }

        return "ninguno";
    }

    public void getFilasBancos() {
        for (int i = 0; i < filasBancos.size(); i++) {
            int fila = filasBancos.get(i);
            System.out.println("fila" + fila);
        }
    }

    // FIN DE SEGMENTO DE LECTURA DE CONCILICAIÓN BANCARIA
    // INICIO DE ESCRITURA DE EXCEL 
    public void createExcel() {
        // Datos para la generación del archivo
        String fileName = "Layout_egresos.xlsx";
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

        XSSFWorkbook wb = this.prefactura;
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
        String id_prefactura;
        XSSFRow row_fila;
        XSSFCell cell_tipo, cell_fecha, cell_descrip, cell_concepto, cell_ref_1, cell_ref_2, cell_num_cuenta, cell_nombre, cell_valor, cell_valor2;
        //JOptionPane.showMessageDialog(null, "Se esta generando tu layout");
        for (int i = 2; i <= rowcount; i++) {
            Row row_prefactura = sheet.getRow(i);
            if (row_prefactura != null) {
                Cell cell_id = row_prefactura.getCell(23);
                if (cell_id != null) {
                    //Creción de una nueva fila
                    row_fila = hoja1.createRow(fila_nueva);
                    //obtener el id que se busca en la orden de pago de conciliación
                    id_prefactura = getIdPoliza(i);
                    //Se llama al método que obtiene las filas del id y del tipo 4
                    getOrdenPago(id_prefactura);
                    //getFilasBancos();
                    // Creación de la celda fecha
                    if (getFecha() != null) {
                        cell_fecha = row_fila.createCell(0);
                        cell_fecha.setCellValue(getFecha());
                    }
                    //Descripción
                    cell_descrip = row_fila.createCell(3);
                    cell_descrip.setCellValue(getDescripcionPoliza(i));
                    if (getUN(i, 13) != null && getCeldaS(i, 11) == null) {
                        cell_tipo = row_fila.createCell(1);
                        cell_tipo.setCellValue("PE");
                    } else {
                        cell_tipo = row_fila.createCell(1);
                        cell_tipo.setCellValue("EGRESOS");
                    }
                    //Concepto
                    cell_concepto = row_fila.createCell(6);
                    cell_concepto.setCellValue(getConceptoPoliza(i));
                    //Referencia uno 
                    cell_ref_1 = row_fila.createCell(7);
                    cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                    //Referencia dos 
                    cell_ref_2 = row_fila.createCell(8);
                    cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                    //Cargos 
                    cell_num_cuenta = row_fila.createCell(4);
                    cell_nombre = row_fila.createCell(5);
                    cell_valor = row_fila.createCell(9);
                    cell_valor2 = row_fila.createCell(10);
                    if (getCeldaS(i, 19) != null) {
                        if (getCeldaS(i, 19).equals("DEV INFONAVIT")) {
                            if (getCuenta("DEV INFONAVIT") != null) {
                                cell_nombre.setCellValue(getCuenta("DEV INFONAVIT").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("DEV INFONAVIT").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 35)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                    }
                    if (getCeldaS(i, 19) != null) {
                        if (getCeldaS(i, 19).equals("DEV FONACOT")) {
                            if (getCuenta("DEV FONACOT") != null) {
                                cell_nombre.setCellValue(getCuenta("DEV FONACOT").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("DEV FONACOT").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 35)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                    }
                    if (getUN(i, 13) != null && getCeldaS(i, 11) == null) {
                        if (getCargosAbono(i, 79) != 0.0) {
                            if (getCuenta("SUELDOS Y SALARIOS") != null) {
                                cell_nombre.setCellValue(getCuenta("SUELDOS Y SALARIOS").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("SUELDOS Y SALARIOS").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 79)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 80) != 0.0) {
                            if (getCuenta("VACACIONES") != null) {
                                cell_nombre.setCellValue(getCuenta("VACACIONES").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("VACACIONES").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 80)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 81) != 0.0) {
                            if (getCuenta("PRIMA VACACIONAL") != null) {
                                cell_nombre.setCellValue(getCuenta("PRIMA VACACIONAL").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("PRIMA VACACIONAL").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 81)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 82) != 0.0) {
                            if (getCuenta("AGUINALDO") != null) {

                                cell_nombre.setCellValue(getCuenta("AGUINALDO").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("AGUINALDO").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 82)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 86) != 0.0) {
                            if (getCuenta("SEP VOL DE LA REL DE TRABAJO") != null) {
                                cell_nombre.setCellValue(getCuenta("SEP VOL DE LA REL DE TRABAJO").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("SEP VOL DE LA REL DE TRABAJO").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 86)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                    } else {
                        if (getCargosAbono(i, 79) != 0.0) {
                            if (getCuenta("SUELDOS Y SALARIOS OAXACA") != null) {
                                cell_nombre.setCellValue(getCuenta("SUELDOS Y SALARIOS OAXACA").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("SUELDOS Y SALARIOS OAXACA").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 79)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 80) != 0.0) {
                            if (getCuenta("VACACIONES OAXACA") != null) {
                                cell_nombre.setCellValue(getCuenta("VACACIONES OAXACA").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("VACACIONES OAXACA").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 80)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 81) != 0.0) {
                            if (getCuenta("PRIMA VACACIONAL OAXACA") != null) {
                                cell_nombre.setCellValue(getCuenta("PRIMA VACACIONAL OAXACA").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("PRIMA VACACIONAL OAXACA").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 81)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 82) != 0.0) {
                            if (getCuenta("AGUINALDO OAXACA") != null) {
                                cell_nombre.setCellValue(getCuenta("AGUINALDO OAXACA").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("AGUINALDO OAXACA").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 82)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }
                        if (getCargosAbono(i, 86) != 0.0) {
                            if (getCuenta("SEP VOL DE LA REL DE TRABAJO OAXACA") != null) {
                                cell_nombre.setCellValue(getCuenta("SEP VOL DE LA REL DE TRABAJO OAXACA").getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta("SEP VOL DE LA REL DE TRABAJO OAXACA").getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargosAbono(i, 86)));
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
                                cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                            }
                        }

                    }
                    if (getCargosAbono(i, 71) != 0.0) {
                        if (getCuenta("INFORMATIVO SUBSIDIO") != null) {
                            cell_nombre.setCellValue(getCuenta("INFORMATIVO SUBSIDIO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("INFORMATIVO SUBSIDIO").getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 71)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 87) != 0.0) {
                        if (getCuenta("FONDO DE PENSIÓN DE SUPERVIVENCIA") != null) {
                            cell_nombre.setCellValue(getCuenta("FONDO DE PENSIÓN DE SUPERVIVENCIA").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("FONDO DE PENSIÓN DE SUPERVIVENCIA").getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 87)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCuenta("DIF EN COBROS CARGO") != null) {
                        cell_nombre.setCellValue(getCuenta("DIF EN COBROS CARGO").getDescripcion());
                        cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS CARGO").getNumeroCuenta());
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
                        cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                    }
                    if (getCargosAbono(i, 43) != 0.0) {
                        if (getCuenta("PENSION ALIMENTICIA") != null) {
                            cell_nombre.setCellValue(getCuenta("PENSION ALIMENTICIA").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("PENSION ALIMENTICIA").getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 43)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 49) != 0.0) {
                        if (getCuenta("PENSION ALIMENTICIA") != null) {
                            cell_nombre.setCellValue(getCuenta("PENSION ALIMENTICIA").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("PENSION ALIMENTICIA").getNumeroCuenta());
                            cell_valor.setCellValue(df.format(getCargosAbono(i, 49)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }

                    /*
                * Apartado de lectura de cargos de conciliación
                     */
                    if (!filasBancos.isEmpty()) {
                        for (int j = 0; j < filasBancos.size(); j++) {
                            if (getCargoAbonoBancos(j).equals("abono")) {
                                cell_nombre.setCellValue(getCuenta(getDescripcionBancos(j)).getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta(getDescripcionBancos(j)).getNumeroCuenta());
                                cell_valor.setCellValue(df.format(getCargoBancos(j)));
                                cell_ref_2.setCellValue(getMovimientoBancos(j));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                cell_ref_2 = row_fila.createCell(8);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));
                            }
                        }
                    }
                    //ABONOS 
                    if (getCargosAbono(i, 43) != 0.0) {
                        if (getCuenta("PENSION ALIMENTICIA") != null) {
                            cell_nombre.setCellValue(getCuenta("PENSION ALIMENTICIA").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("PENSION ALIMENTICIA").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 43)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 49) != 0.0) {
                        if (getCuenta("PENSION ALIMENTICIA") != null) {
                            cell_nombre.setCellValue(getCuenta("PENSION ALIMENTICIA").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("PENSION ALIMENTICIA").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 49)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 65) != 0.0) {
                        if (getCuenta("IMSS MENSUAL OBRERO") != null) {
                            cell_nombre.setCellValue(getCuenta("IMSS MENSUAL OBRERO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("IMSS MENSUAL OBRERO").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 65)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 67) != 0.0) {
                        if (getCuenta("IMSS BIMESTRAL OBRERO") != null) {
                            cell_nombre.setCellValue(getCuenta("IMSS BIMESTRAL OBRERO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("IMSS BIMESTRAL OBRERO").getNumeroCuenta());
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 69) != 0.0) {
                        if (getCuenta("RETENCION INFONAVIT") != null) {
                            cell_nombre.setCellValue(getCuenta("RETENCION INFONAVIT").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("RETENCION INFONAVIT").getNumeroCuenta());
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 74) != 0.0) {
                        if (getCuenta("ISR RETENIDO") != null) {
                            cell_nombre.setCellValue(getCuenta("ISR RETENIDO").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("ISR RETENIDO").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 74)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (getCargosAbono(i, 76) != 0.0) {
                        if (getCuenta("FONACOT") != null) {
                            cell_nombre.setCellValue(getCuenta("FONACOT").getDescripcion());
                            cell_num_cuenta.setCellValue(getCuenta("FONACOT").getNumeroCuenta());
                            cell_valor2.setCellValue(df.format(getCargosAbono(i, 76)));
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
                            cell_ref_2.setCellValue(getReferenciaPoliza2prefactura(i));
                        }
                    }
                    if (!filasBancos.isEmpty()) {
                        for (int j = 0; j < filasBancos.size(); j++) {
                            if (getCargoAbonoBancos(j).equals("cargo")) {
                                
                                cell_nombre.setCellValue(getCuenta(getDescripcionBancos(j)).getDescripcion());
                                cell_num_cuenta.setCellValue(getCuenta(getDescripcionBancos(j)).getNumeroCuenta());
                                cell_valor2.setCellValue(df.format(getAbonoBancos(j)));
                                cell_ref_2.setCellValue(getMovimientoBancos(j));
                                fila_nueva++;
                                row_fila = hoja1.createRow(fila_nueva);
                                cell_num_cuenta = row_fila.createCell(4);
                                cell_nombre = row_fila.createCell(5);
                                cell_valor = row_fila.createCell(9);
                                cell_valor2 = row_fila.createCell(10);
                                cell_ref_2 = row_fila.createCell(8);
                                //Concepto 
                                cell_concepto = row_fila.createCell(6);
                                cell_concepto.setCellValue(getConceptoPoliza(i));
                                //Referencia
                                cell_ref_1 = row_fila.createCell(7);
                                cell_ref_1.setCellValue(getReferenciaPoliza1(i));

                            }
                        }
                    }
                    if (getCuenta("DIF EN COBROS ABONO") != null) {
                        cell_nombre.setCellValue(getCuenta("DIF EN COBROS ABONO").getDescripcion());
                        cell_num_cuenta.setCellValue(getCuenta("DIF EN COBROS ABONO").getNumeroCuenta());
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
                        cell_ref_2.setCellValue(getMesPrefactura(i));
                    }
                    filasBancos.clear();
                }
            }
        }

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
            //JOptionPane.showMessageDialog(null, "Se ha generado tu archivo");

        } catch (FileNotFoundException e) {
            this.finalizado = true;
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "No se puede crear el layout ya que el archivo se encuentra abierto");

        } catch (Exception e) {
            this.finalizado = true;
            e.printStackTrace();
        }

    }

    public int getHoja() {
        int fila = Integer.parseInt(JOptionPane.showInputDialog("Ingresa el número de hoja en el que se encuentra tu conciliación"));
        System.out.println(fila - 1);
        return fila - 1;
    }

    public int getCelda() {
        int fila = Integer.parseInt(JOptionPane.showInputDialog("Ingresa el número de columna"));
        return fila;
    }

    public int getDescripcion() {
        int fila = Integer.parseInt(JOptionPane.showInputDialog("Ingresa el número de columna de la descripción"));
        return fila;
    }

    public int getTipo() {
        int fila = Integer.parseInt(JOptionPane.showInputDialog("Ingresa el número de columna del tipo de cuenta"));
        return fila;
    }

    public XSSFWorkbook getCatalogo() {
        return catalogo;
    }

    public void setCatalogo(XSSFWorkbook catalogo) {
        this.catalogo = catalogo;
    }

    public XSSFWorkbook getPrefactura() {
        return prefactura;
    }

    public void setPrefactura(XSSFWorkbook prefactura) {
        this.prefactura = prefactura;
    }

    public XSSFWorkbook getConciliacion() {
        return conciliacion;
    }

    public void setConciliación(XSSFWorkbook conciliacion) {
        this.conciliacion = conciliacion;
    }

    public ArrayList<Cuenta> getCuentasList() {
        return cuentasList;
    }

    public Integer getHojaExcel() {
        return hojaExcel;
    }

    public void setHojaExcel(Integer hojaExcel) {
        this.hojaExcel = hojaExcel;
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

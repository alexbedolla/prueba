/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package botonesalta;

import com.wsclient.ButtonConfig;
import java.io.File;
import java.io.IOException;
import com.wsclient.GETButtonTempWResponse;
import com.wsclient.GetButtonTem;
import com.wsclient.GetButtonTemResponse;
import com.wsclient.GetItemTemp;
import com.wsclient.GetItemTempResponse;
import com.wsclient.NewButtonConfig;
import com.wsclient.NewButtonConfigResponse;
import jxl.*;
import jxl.read.biff.BiffException;
/**
 *
 * @author Alex
 */
public class BotonesAlta {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, BiffException {
       Workbook workbook = Workbook.getWorkbook(new File("C:\\botones.xls")); //Pasamos el excel que vamos a leer
       BotonesAlta bot=new BotonesAlta();
       Sheet sheet = workbook.getSheet(0); //Seleccionamos la hoja que vamos a leer
       bot.proceso(sheet);
       Sheet sheet2 =workbook.getSheet(1);
       bot.proceso(sheet2);
   }
    
    private void proceso(Sheet sheet){
        String nombre;
        int tamaño = sheet.getColumns();
        int raiz = 1;        
        int idanterior =0;
        for (int fila = 1; fila < sheet.getRows(); fila++) { //recorremos las filas
            for (int columna = 1; columna < sheet.getColumns(); columna++) { //recorremos las columnas
                nombre = sheet.getCell(columna, fila).getContents(); //setear la celda leida a nombre
                
                if(nombre.length()>20)nombre = nombre.substring(0,20);
                GetButtonTem indata = new GetButtonTem();
                indata.setIDInstancia(41);
                indata.setName(nombre);
                GetButtonTemResponse button = Getbutton(indata);
                if(button.getButton().getIDBTN()>0){
                    idanterior=button.getButton().getIDBTN();
                    continue;
                }//existe el boton siguiente iteracion
                else{ //no existe el boton
                    NewButtonConfig in = new NewButtonConfig();
                    ButtonConfig boton = new ButtonConfig();
                    if(columna==raiz){
                        boton.setIDSTRRT(1);
                        boton.setIDBTNSET(1);
//                        boton.setIDBTNPRNT(fila);//depende columna;
                        boton.setTYBTN("RO");//depende columna
                        boton.setBTNWT(1);
                        boton.setNMBTN(nombre);
                        boton.setDEBTN(nombre);                        
                        boton.setIDSTRGRP(1);
                        boton.setIDITMPS("");//solo ultima columna
                    }else if(columna==(tamaño-1)){//ultima columna
                        GetItemTemp initm = new GetItemTemp();
                        String tmpid = sheet.getCell(0,fila).getContents();
                        initm.setPosy(tmpid);
                        initm.setIDInstancia(41);
                        GetItemTempResponse itmresp = Getitm(initm);
                        if(itmresp.getResultado().getResultado()==-1){
                            System.out.println("no existe el articulo: "+Integer.parseInt(sheet.getCell(0, fila).getContents())+"\n cancelando");
                            break;
                        }
                        boton.setIDSTRRT(1);
                        boton.setIDBTNSET(1);
                        boton.setIDBTNPRNT(idanterior);//depende columna;
                        boton.setTYBTN("PR");//depende columna
                        boton.setBTNWT(1);
                        boton.setNMBTN(nombre);
                        boton.setDEBTN(nombre);
                        boton.setIDITM(itmresp.getItmtmp().getIDITM());//solo ultima columna
                        boton.setIDSTRGRP(1);
                        boton.setIDITMPS(itmresp.getItmtmp().getIDITMPS());//solo ultima columna
                        boton.setIDITMPSQFR(itmresp.getItmtmp().getIDITMPSQFR());//solo ultima columna
                    } else{
                        boton.setIDSTRRT(1);
                        boton.setIDBTNSET(1);
                        boton.setIDBTNPRNT(idanterior);//depende columna;
                        boton.setTYBTN("AG");//depende columna
                        boton.setBTNWT(1);
                        boton.setNMBTN(nombre);
                        boton.setDEBTN(nombre);
//                        boton.setIDITM(fila);//solo ultima columna
                        boton.setIDSTRGRP(1);
                        boton.setIDITMPS("");//solo ultima columna
//                        boton.setIDITMPSQFR(fila);//solo ultima columna
                    }  
                    in.setOper("A");
                    in.setButtonconfig(boton);
                    in.setIDInstancia(41);
                    NewButtonConfigResponse renew = abcButton(in);
                    if(renew.getResultado().getResultado()<0){
                         System.out.println("Fallo en la alta de "+nombre+" \n cancelando");
                         break;
                    } else {
                        idanterior = renew.getResultado().getResultado();
                    }
                    
                }//existe
            }
            
        }
    }
        // TODO code application logic here
    
    private static GetButtonTemResponse Getbutton(GetButtonTem log) {
        com.wsclient.EJBWebServicev20_Service service = new com.wsclient.EJBWebServicev20_Service();
        com.wsclient.EJBWebServicev20 port = service.getEJBWebServicev20Port();
        return port.getButtonTempW(log);
    }
    
     private static GetItemTempResponse Getitm(GetItemTemp log) {
        com.wsclient.EJBWebServicev20_Service service = new com.wsclient.EJBWebServicev20_Service();
        com.wsclient.EJBWebServicev20 port = service.getEJBWebServicev20Port();
        return port.getItemTempW(log);
    }
     
      private static NewButtonConfigResponse abcButton(NewButtonConfig log) {
        com.wsclient.EJBWebServicev20_Service service = new com.wsclient.EJBWebServicev20_Service();
        com.wsclient.EJBWebServicev20 port = service.getEJBWebServicev20Port();
        return port.abcButtonW(log);
    }
}

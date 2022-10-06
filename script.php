<?php
    error_reporting(error_reporting() & ~E_NOTICE);
    ini_set('memory_limit', '500M');

    require __DIR__ . "/vendor/autoload.php";
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;   

    $env_array =getenv();
    $uid = $env_array["USERDB"];
    $pwd = $env_array["PASSDB"];
    $serverName = $env_array["SERVERNAME"];


    /*Conexion AzureDB*/
    $connectionInfo = array( "Database"=>"AnalyticModelDB", "UID"=>$uid, "PWD"=>$pwd );
    $conn = sqlsrv_connect($serverName, $connectionInfo);

    
    /**************INICIO  ARCHIVO 1**************/
    $tsql= "SELECT * FROM [dbo].[Z_FORMULATION_ARTICLES_REPORT]"; /*Query de documento de excel Formulacion de Articulos*/
    $getResults= sqlsrv_query($conn, $tsql);
    
    echo ("Reading data from table - Archivo 1" . PHP_EOL);
   

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Formulación de Articulos')
        ->setSubject('Formulación de Articulos')
        ->setDescription('Formulación de Articulos');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $firstime = 0;
    $fila = 2;

    if ($getResults == FALSE)
        echo (sqlsrv_errors());

    while ($row = sqlsrv_fetch_array($getResults, SQLSRV_FETCH_ASSOC)) {
        if($firstime == 0){
            $i = 1;
            while ($value = key($row)) {
                $sheet->setCellValueByColumnAndRow($i, 1 , key($row) );
                $i++;
                next($row);
            }
            $sheet->setCellValueByColumnAndRow(sizeof($row)+1, 1 , 'Validación' );
            $sheet->setCellValueByColumnAndRow(sizeof($row)+2, 1 , 'Validación Symbol' );
            $firstime = 1;
        }

        $sheet->setCellValueByColumnAndRow(1, $fila , $row['FORMULAID'] );
        $sheet->setCellValueByColumnAndRow(2, $fila , $row['VERSIONNAME'] );
        $sheet->setCellValueByColumnAndRow(3, $fila , $row['FORMULABATCHSIZE'] );
        $sheet->setCellValueByColumnAndRow(4, $fila , $row['PURCHASEUNITSYMBOL'] );
        $sheet->setCellValueByColumnAndRow(5, $fila , $row['ISACTIVE'] );
        $sheet->setCellValueByColumnAndRow(6, $fila , $row['ITEMNUMBER'] );
        $sheet->setCellValueByColumnAndRow(7, $fila , $row['PRODUCTNAME'] );
        $sheet->setCellValueByColumnAndRow(8, $fila , $row['PRODUCTGROUPNAME'] );
        $sheet->setCellValueByColumnAndRow(9, $fila , $row['ITEMMODELGROUPNAME'] );
        $sheet->setCellValueByColumnAndRow(10, $fila , $row['PRODUCTUNITSYMBOL'] );
        $sheet->setCellValueByColumnAndRow(11, $fila , $row['QUANTITY'] );
        $sheet->setCellValueByColumnAndRow(12, $fila , $row['FORMULANAME'] );
   
        if(strcmp($row['VERSIONNAME'],$row['FORMULANAME']) == 0){
            $sheet->setCellValueByColumnAndRow(13, $fila , $row[''] );
        }else{
            $sheet->setCellValueByColumnAndRow(13, $fila , 'Modificar Nombre Formula' );
        }
        
        if(strcmp($row['PURCHASEUNITSYMBOL'],$row['PRODUCTUNITSYMBOL']) == 0 ){
            $sheet->setCellValueByColumnAndRow(14, $fila , $row[''] );
        }else{
            $sheet->setCellValueByColumnAndRow(14, $fila , 'Variación de Simbolo');
        }
        $fila++;
    }

    foreach (range('A', 'N') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }

    $writer = new Xlsx($spread);
    $writer->save('/var/www/scriptsnode/archivos/Formulación de Articulos Automático.xlsx');

    echo ("FIN - Archivo 1" . PHP_EOL);
    sqlsrv_free_stmt($getResults);

    /**************INICIO  ARCHIVO 2**************/

    $tsql= "SELECT * FROM [dbo].[Z_S2P_SUPPLIER REPORT]"; /*Query de documento de excel Reporte Proveedores*/
    $getResults= sqlsrv_query($conn, $tsql);

    echo ("Reading data from table - Archivo 2" . PHP_EOL);

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Reporte Proveedores')
        ->setSubject('Reporte Proveedores')
        ->setDescription('Reporte Proveedores');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $firstime = 0;
    $fila = 2;

    if ($getResults == FALSE)
        echo (sqlsrv_errors());

    while ($row = sqlsrv_fetch_array($getResults, SQLSRV_FETCH_ASSOC)) {
    
        if($firstime == 0){
            $i = 1;
            while ($value = key($row)) {
                $sheet->setCellValueByColumnAndRow($i, 1 , key($row) );
                $i++;
                next($row);
            }
            $firstime = 1;
        }

        $sheet->setCellValueByColumnAndRow(1, $fila , $row['NOMBRE DEL PROVEEDOR'] );
        $sheet->setCellValueByColumnAndRow(2, $fila , $row['NIT DEL PROVEEDOR'] );
        $sheet->setCellValueByColumnAndRow(3, $fila , $row['TIPO'] );
        $sheet->setCellValueByColumnAndRow(4, $fila , $row['ID TAX'] );
        $sheet->setCellValueByColumnAndRow(5, $fila , $row['CIUDAD'] );
        $sheet->setCellValueByColumnAndRow(6, $fila , $row['DEPARTAMENTO/ESTADO'] );
        $sheet->setCellValueByColumnAndRow(7, $fila , $row['PAÍS'] );
        $sheet->setCellValueByColumnAndRow(8, $fila , $row['NÚMERO DE TELÉFONO'] );
        $sheet->setCellValueByColumnAndRow(9, $fila , $row['DIRECCIÓN'] );
        $sheet->setCellValueByColumnAndRow(10, $fila , $row['CORREO ELECTRÓNICO'] );
        $sheet->setCellValueByColumnAndRow(11, $fila , $row['BANCO'] );
        $sheet->setCellValueByColumnAndRow(12, $fila , $row['NÚMERO DE CUENTA BANCARIA'] );
        $sheet->setCellValueByColumnAndRow(13, $fila , $row['TÉRMINO DE PAGO'] );
        $sheet->setCellValueByColumnAndRow(14, $fila , $row['TIPO (NATURAL O JURÍDICO)'] );
        $sheet->setCellValueByColumnAndRow(15, $fila , $row['RESPONSABLE DE IVA'] );
        $sheet->setCellValueByColumnAndRow(16, $fila , $row['DIVISA'] );
        $sheet->setCellValueByColumnAndRow(17, $fila , $row['FORMA DE PAGO'] );
        $sheet->setCellValueByColumnAndRow(18, $fila , $row['GRUPO DE IMPUESTOS'] );
        $sheet->setCellValueByColumnAndRow(19, $fila , $row['Fecha y hora de creación']->format('d/m/Y H:i:s') );
        $sheet->setCellValueByColumnAndRow(20, $fila , $row['Fecha y hora de modificación']->format('d/m/Y H:i:s') );
        
        $fila++;
    }


    foreach (range('A', 'S') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }

    $writer = new Xlsx($spread);
    $writer->save('/var/www/scriptsnode/archivos/Reporte Proveedores Automático.xlsx');

    echo ("FIN - Archivo 2" . PHP_EOL);
    sqlsrv_free_stmt($getResults);


    /**************INICIO  ARCHIVO 3**************/
    $tsql= "SELECT * FROM [dbo].[Z_ProductionStatusReport]"; /*Query de documento de excel Reporte_Producción*/
    $getResults= sqlsrv_query($conn, $tsql);

    
    echo ("Reading data from table - Archivo 3" . PHP_EOL);

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Reporte_Producción')
        ->setSubject('Reporte_Producción')
        ->setDescription('Reporte_Producción');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $firstime = 0;
    $fila = 2;

    if ($getResults == FALSE)
    echo (sqlsrv_errors());

    while ($row = sqlsrv_fetch_array($getResults, SQLSRV_FETCH_ASSOC)) {

        
        if($firstime == 0){
            $i = 1;
            while ($value = key($row)) {
                $sheet->setCellValueByColumnAndRow($i, 1 , key($row) );
                $i++;
                next($row);
            }
            $firstime = 1;
        }

        $sheet->setCellValueByColumnAndRow(1, $fila , $row['Product Number'] );
        $sheet->setCellValueByColumnAndRow(2, $fila , $row['Batch Number'] );
        $sheet->setCellValueByColumnAndRow(3, $fila , $row['Batch Order Number'] );
        $sheet->setCellValueByColumnAndRow(4, $fila , $row['Status'] );
        $sheet->setCellValueByColumnAndRow(5, $fila , $row['Backorder Status'] );
        if ($row['Order Planned Date'] != null) {
            $sheet->setCellValueByColumnAndRow(6, $fila , $row['Order Planned Date']->format('d/m/Y') );
        }
        $sheet->setCellValueByColumnAndRow(6, $fila , $row['Order Planned Date']->format('d/m/Y') );
        $sheet->setCellValueByColumnAndRow(7, $fila , $row['Journal Line Number'] );
        $sheet->setCellValueByColumnAndRow(8, $fila , $row['Route ID'] );
        $sheet->setCellValueByColumnAndRow(9, $fila , $row['Operation Number'] );
        $sheet->setCellValueByColumnAndRow(10, $fila , $row['Operation'] );
        $sheet->setCellValueByColumnAndRow(11, $fila , $row['Operation Expected time'] );
        $sheet->setCellValueByColumnAndRow(12, $fila , $row['Registered Hours'] );

        $sheet->setCellValueByColumnAndRow(7, $fila , $row['Scheduled Production Quantity'] );
        $sheet->setCellValueByColumnAndRow(8, $fila , $row['Production Quantity'] );
        $sheet->setCellValueByColumnAndRow(9, $fila , $row['Reported Good Inventory Quantity'] );
        if ($row['Operation Start Date'] != null) {
            $sheet->setCellValueByColumnAndRow(10, $fila , $row['Operation Start Date']->format('d/m/Y') );
        }
        if ($row['Operation End Date'] != null) {
            $sheet->setCellValueByColumnAndRow(11, $fila , $row['Operation End Date']->format('d/m/Y') );
        }
        if ($row['Order Delivery Date'] != null) {
            $sheet->setCellValueByColumnAndRow(12, $fila , $row['Order Delivery Date']->format('d/m/Y') );
        }
        
        
        
        $fila++;
    }

    foreach (range('A', 'R') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }

    $writer = new Xlsx($spread);
    $writer->save('/var/www/scriptsnode/archivos/Reporte_Producción Automático.xlsx');

    echo ("FIN - Archivo 3" . PHP_EOL);
    sqlsrv_free_stmt($getResults);


    /**************INICIO  ARCHIVO 4**************/
    $tsql= "SELECT * FROM [dbo].[Z_MasterProducts]"; /*Maestro de Articulos V2*/
    $getResults= sqlsrv_query($conn, $tsql);
    echo ("Reading data from table - Archivo 4" . PHP_EOL);

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Maestro de Artículos V2')
        ->setSubject('Maestro de Artículos V2')
        ->setDescription('Maestro de Artículos V2');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $firstime = 0;
    $fila = 2;

    if ($getResults == FALSE)
    echo (sqlsrv_errors());

    while ($row = sqlsrv_fetch_array($getResults, SQLSRV_FETCH_ASSOC)) {
       
        $sheet->setCellValueByColumnAndRow(1, 1 , "ITEMNUMBER" );
        $sheet->setCellValueByColumnAndRow(2, 1 , "SEVENTHPRODUCTFILTERCODE" );		
        $sheet->setCellValueByColumnAndRow(3, 1 , "PRODUCTNAME" );   	
        $sheet->setCellValueByColumnAndRow(4, 1 , "ALTERNATIVEPRODUCTVERSIONID" );   	
        $sheet->setCellValueByColumnAndRow(5, 1 , "PRODUCTTYPE" );   	
        $sheet->setCellValueByColumnAndRow(6, 1 , "PRODUCTSUBTYPE" );   	
        $sheet->setCellValueByColumnAndRow(7, 1 , "PURCHASEUNITSYMBOL" );   	
        $sheet->setCellValueByColumnAndRow(8, 1 , "COSTCALCULATIONBOMLEVEL" );   	
        $sheet->setCellValueByColumnAndRow(9, 1 , "SALESRETAILINVENTORYAVAILABILITYLEVELPROFILE" );  	
        $sheet->setCellValueByColumnAndRow(10, 1 , "TENTHPRODUCTFILTERCODE" );  	
        $sheet->setCellValueByColumnAndRow(11, 1 , "WAREHOUSERELEASESALESUNITRESTRICTED" );   	
        $sheet->setCellValueByColumnAndRow(12, 1 , "EIGHTHPRODUCTFILTERCODE" );  	
        $sheet->setCellValueByColumnAndRow(13, 1 , "INVENTORYUNITSYMBOL" );  	
        $sheet->setCellValueByColumnAndRow(14, 1 , "SALESUNITSYMBOL" );   	
        $sheet->setCellValueByColumnAndRow(15, 1 , "PRODUCTGROUPID" );   	
        $sheet->setCellValueByColumnAndRow(16, 1 , "Group Name" );  	
        $sheet->setCellValueByColumnAndRow(17, 1 , "NINTHPRODUCTFILTERCODE" );  	
        $sheet->setCellValueByColumnAndRow(18, 1 , "DEFAULTPRODUCTVERSIONID" );  	
        $sheet->setCellValueByColumnAndRow(19, 1 , "SALESRETAILINVENTORYAVAILABILITYBUFFER" );  	
        $sheet->setCellValueByColumnAndRow(20, 1 , "ITEMMODELGROUPID" );  	
        $sheet->setCellValueByColumnAndRow(21, 1 , "Nombre" );  	
        $sheet->setCellValueByColumnAndRow(22, 1 , "SIXTHPRODUCTFILTERCODE" );  	
        $sheet->setCellValueByColumnAndRow(23, 1 , "FIFTHPRODUCTFILTERCODE" );  	
        $sheet->setCellValueByColumnAndRow(24, 1 , "Manejo de Inventario" );  	
        $sheet->setCellValueByColumnAndRow(25, 1 , "Tiene costo estándar" );  	
        $sheet->setCellValueByColumnAndRow(26, 1 , "STORAGEDIMENSIONGROUPNAME" );  	
        $sheet->setCellValueByColumnAndRow(27, 1 , "TRACKINGDIMENSIONGROUPNAME" );  	
        $sheet->setCellValueByColumnAndRow(28, 1 , "ITEMOVERUNDERDELIVERYTOLERANCEGROUPID" );  	
        $sheet->setCellValueByColumnAndRow(29, 1 , "CUSTOMSDESCRIPTION" );  	
        $sheet->setCellValueByColumnAndRow(30, 1 , "TRANSFERORDERLANDEDCOSTGROUPID" ); 	
        $sheet->setCellValueByColumnAndRow(31, 1 , "VOYAGEARRIVALGROUPID" );  	
        $sheet->setCellValueByColumnAndRow(32, 1 , "LANDEDCOSTTYPEGROUPID" ); 	
        $sheet->setCellValueByColumnAndRow(33, 1 , "COMMODITYCODEID" );  	
        $sheet->setCellValueByColumnAndRow(34, 1 , "INVENTORYRESERVATIONHIERARCHYNAME" );  	
        $sheet->setCellValueByColumnAndRow(35, 1 , "UNITCOSTQUANTITY" );  	
        $sheet->setCellValueByColumnAndRow(36, 1 , "DEFAULTLEDGERDIMENSIONDISPLAYVALUE" ); 	
        $sheet->setCellValueByColumnAndRow(37, 1 , "Cuenta de Inventario" );  
        $sheet->setCellValueByColumnAndRow(38, 1 , "Contrapartida Compra GITIIC" ); 	
        $sheet->setCellValueByColumnAndRow(39, 1 , "Contrapartida Consumo Normal" );   	
        $sheet->setCellValueByColumnAndRow(40, 1 , "Contrapartida Consumo I+D" );          														
        $sheet->setCellValueByColumnAndRow(1, $fila , $row['ITEMNUMBER'] );
        $sheet->setCellValueByColumnAndRow(2, $fila , $row['SEVENTHPRODUCTFILTERCODE'] );
        $sheet->setCellValueByColumnAndRow(3, $fila , $row['PRODUCTNAME'] );
        $sheet->setCellValueByColumnAndRow(4, $fila , $row['ALTERNATIVEPRODUCTVERSIONID'] );
        $sheet->setCellValueByColumnAndRow(5, $fila , $row['PRODUCTTYPE'] );
        $sheet->setCellValueByColumnAndRow(6, $fila , $row['PRODUCTSUBTYPE']);
        $sheet->setCellValueByColumnAndRow(7, $fila , $row['PURCHASEUNITSYMBOL'] );
        $sheet->setCellValueByColumnAndRow(8, $fila , $row['COSTCALCULATIONBOMLEVEL'] );
        $sheet->setCellValueByColumnAndRow(9, $fila , $row['SALESRETAILINVENTORYAVAILABILITYLEVELPROFILE'] );
        $sheet->setCellValueByColumnAndRow(10, $fila , $row['TENTHPRODUCTFILTERCODE'] );
        $sheet->setCellValueByColumnAndRow(11, $fila , $row['WAREHOUSERELEASESALESUNITRESTRICTED'] );
        $sheet->setCellValueByColumnAndRow(12, $fila , $row['EIGHTHPRODUCTFILTERCODE'] );
        $sheet->setCellValueByColumnAndRow(13, $fila , $row['INVENTORYUNITSYMBOL'] );
        $sheet->setCellValueByColumnAndRow(14, $fila , $row['SALESUNITSYMBOL'] );
        $sheet->setCellValueByColumnAndRow(15, $fila , $row['PRODUCTGROUPID'] );
        $sheet->setCellValueByColumnAndRow(16, $fila , $row['Group Name'] );
        $sheet->setCellValueByColumnAndRow(17, $fila , $row['NINTHPRODUCTFILTERCODE'] );
        $sheet->setCellValueByColumnAndRow(18, $fila , $row['DEFAULTPRODUCTVERSIONID'] );
        $sheet->setCellValueByColumnAndRow(19, $fila , $row['SALESRETAILINVENTORYAVAILABILITYBUFFER'] );
        $sheet->setCellValueByColumnAndRow(20, $fila , $row['ITEMMODELGROUPID'] );
        $sheet->setCellValueByColumnAndRow(21, $fila , $row['Nombre'] );
        $sheet->setCellValueByColumnAndRow(22, $fila , $row['SIXTHPRODUCTFILTERCODE'] );
        $sheet->setCellValueByColumnAndRow(23, $fila , $row['FIFTHPRODUCTFILTERCODE'] );
        $sheet->setCellValueByColumnAndRow(24, $fila , $row['Manejo de Inventario'] );
        $sheet->setCellValueByColumnAndRow(25, $fila , $row['Tiene costo estándar'] );
        $sheet->setCellValueByColumnAndRow(26, $fila , $row['STORAGEDIMENSIONGROUPNAME'] );
        $sheet->setCellValueByColumnAndRow(27, $fila , $row['TRACKINGDIMENSIONGROUPNAME'] );
        $sheet->setCellValueByColumnAndRow(28, $fila , $row['ITEMOVERUNDERDELIVERYTOLERANCEGROUPID'] );
        $sheet->setCellValueByColumnAndRow(29, $fila , $row['CUSTOMSDESCRIPTION'] );
        $sheet->setCellValueByColumnAndRow(30, $fila , $row['TRANSFERORDERLANDEDCOSTGROUPID'] );
        $sheet->setCellValueByColumnAndRow(31, $fila , $row['VOYAGEARRIVALGROUPID'] );
        $sheet->setCellValueByColumnAndRow(32, $fila , $row['LANDEDCOSTTYPEGROUPID'] );
        $sheet->setCellValueByColumnAndRow(33, $fila , $row['COMMODITYCODEID'] );
        $sheet->setCellValueByColumnAndRow(34, $fila , $row['INVENTORYRESERVATIONHIERARCHYNAME'] );
        $sheet->setCellValueByColumnAndRow(35, $fila , $row['UNITCOSTQUANTITY'] );
        $sheet->setCellValueByColumnAndRow(36, $fila , $row['DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] );
        $sheet->setCellValueByColumnAndRow(37, $fila , $row['Cuenta de Inventario'] );
        $sheet->setCellValueByColumnAndRow(38, $fila , $row['Contrapartida Compra GITIIC'] );
        $sheet->setCellValueByColumnAndRow(39, $fila , $row['Contrapartida Consumo Normal'] );

        $digitos = substr($sheet->getCellByColumnAndRow(39,$fila)->getValue(), 0, 2); 
        if($digitos == 71 || $digitos == 72 || $digitos == 73 || $digitos == 74){ 
            $output = '51' . substr($sheet->getCellByColumnAndRow(39,$fila)->getValue(), 2);
            $sheet->setCellValueByColumnAndRow(40, $fila , $output ); 
        }else{$sheet->setCellValueByColumnAndRow(40, $fila , "" ); }
        
        
        $fila++;
    }

    foreach (range('A', 'Z') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }


    $writer = new Xlsx($spread);
    $writer->save('/var/www/scriptsnode/archivos/Maestro de Artículos V2 Automático.xlsx');


    echo ("FIN - Archivo 4" . PHP_EOL);
    sqlsrv_free_stmt($getResults);


    /**************INICIO  ARCHIVO 5**************/
    $tsql= "SELECT * FROM [dbo].[Z_P2E_ProductionOrdersReport]"; /*Query de documento de excel Control Ordenes Produccion*/
    $getResults = sqlsrv_query($conn, $tsql);
    
    
    echo ("Reading data from table - Archivo 5" . PHP_EOL);

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Reporte_Control_OrdenesProduccion')
        ->setSubject('Reporte_Control_OrdenesProduccion')
        ->setDescription('Reporte_Control_OrdenesProduccion');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $firstime = 0;
    $fila = 2;

    if ($getResults == FALSE)
    echo (sqlsrv_errors());
   
    while ($row = sqlsrv_fetch_array($getResults, SQLSRV_FETCH_ASSOC)) {

        
        if($firstime == 0){
            $i = 1;
            while ($value = key($row)) {
                $sheet->setCellValueByColumnAndRow($i, 1 , key($row) );
                $i++;
                next($row);
            }
            $firstime = 1;
        }
      


        $sheet->setCellValueByColumnAndRow(1, $fila , $row['PedidodeLote'] );
        $sheet->setCellValueByColumnAndRow(2, $fila , $row['CodigodeArticulo'] );
        $sheet->setCellValueByColumnAndRow(3, $fila , $row['Lote'] );
        $sheet->setCellValueByColumnAndRow(4, $fila , $row['Cantidad_Iniciada'] );
        if ($row['Estimada'] != null) {
            $sheet->setCellValueByColumnAndRow(5, $fila , $row['Estimada']);
        }
        if ($row['Programado'] != null) {
            $sheet->setCellValueByColumnAndRow(6, $fila , $row['Programado']);
        }
        if ($row['Fecha_Liberacion'] != null) {
            $sheet->setCellValueByColumnAndRow(7, $fila , $row['Fecha_Liberacion']);
        }
        if ($row['Iniciado'] != null) {
            $sheet->setCellValueByColumnAndRow(8, $fila , $row['Iniciado']);
        }
        if ($row['NotificadocomoTerminado'] != null) {
            $sheet->setCellValueByColumnAndRow(9, $fila , $row['NotificadocomoTerminado'] );  
        }
        $sheet->setCellValueByColumnAndRow(10, $fila , $row['Cantidad_Entregada_Almacen'] );
        if ($row['Terminado'] != null) {
            $sheet->setCellValueByColumnAndRow(11, $fila , $row['Terminado']);  
        }
        
        $sheet->setCellValueByColumnAndRow(12, $fila , $row['Estado'] );
        if ($row['Min_Fecha_Prog_Lote'] != null) {
            $sheet->setCellValueByColumnAndRow(13, $fila , $row['Min_Fecha_Prog_Lote']->format('d/m/Y'));  
        }
        
        if ($row['Max_Fecha_Not_Ter_Lote'] != null) {
            $sheet->setCellValueByColumnAndRow(14, $fila , $row['Max_Fecha_Not_Ter_Lote']->format('d/m/Y'));  
        }
        
        
        $fila++;
    }

    foreach (range('A', 'N') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }

    $writer = new Xlsx($spread);
    $writer->save('/var/www/scriptsnode/archivos/Reporte_Control_OrdenesProduccion Automático.xlsx');

    echo ("FIN - Archivo 5" . PHP_EOL);
    sqlsrv_free_stmt($getResults);




    /**************INICIO  ARCHIVO 6**************/

    $tsql= "SELECT * FROM [dbo].[z_Sales_Order_Price_List_CLEC]"; /*Query de documento de excel Sales_Order_Price_List*/
    $getResults= sqlsrv_query($conn, $tsql);

    echo ("Reading data from table - Archivo 6" . PHP_EOL);

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Precios Pedido de Venta')
        ->setSubject('RPrecios Pedido de Venta')
        ->setDescription('Precios Pedido de Venta');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $firstime = 0;
    $fila = 2;

    if ($getResults == FALSE)
        echo (sqlsrv_errors());

    while ($row = sqlsrv_fetch_array($getResults, SQLSRV_FETCH_ASSOC)) {
    
        if($firstime == 0){
            $i = 1;
            while ($value = key($row)) {
                $sheet->setCellValueByColumnAndRow($i, 1 , key($row) );
                $i++;
                next($row);
            }
            $firstime = 1;
        }

        $sheet->setCellValueByColumnAndRow(1, $fila , $row['Pedido de ventas'] );
        $sheet->setCellValueByColumnAndRow(2, $fila , $row['Código de artículo'] );
        $sheet->setCellValueByColumnAndRow(3, $fila , $row['Descripción de Linea'] );
        $sheet->setCellValueByColumnAndRow(4, $fila , $row['Precio Unitario'] );
        $sheet->setCellValueByColumnAndRow(5, $fila , $row['Cantidad'] );
        $sheet->setCellValueByColumnAndRow(6, $fila , $row['Importe Neto'] );
        $sheet->setCellValueByColumnAndRow(7, $fila , $row['Divisa'] );
        $sheet->setCellValueByColumnAndRow(8, $fila , $row['Estado'] );

        $fila++;
    }


    foreach (range('A', 'H') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }

    $writer = new Xlsx($spread);
    $writer->save('/var/www/scriptsnode/archivos/Precios Pedido de Venta Automático.xlsx');

    echo ("FIN - Archivo 6" . PHP_EOL);
    sqlsrv_free_stmt($getResults);


?>
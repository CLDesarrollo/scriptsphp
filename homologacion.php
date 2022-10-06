<?php
    error_reporting(error_reporting() & ~E_NOTICE);
    ini_set('memory_limit', '300M');

    require __DIR__ . "/vendor/autoload.php";
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;  

    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("/var/www/html/homologar/CodigosERP.xlsx");
    $highestrow = $spreadsheet->getActiveSheet()->getHighestRow();
    $range = 'A2:B'.$highestrow;

    $homologacion = $spreadsheet->getActiveSheet()->rangeToArray(
        $range ,     // The worksheet range that we want to retrieve
        NULL,        // Value that should be returned for empty cells
        TRUE,        // Should formulas be calculated (the equivalent of getCalculatedValue() for each cell)
        TRUE,        // Should values be formatted (the equivalent of getFormattedValue() for each cell)
        TRUE         // Should the array be indexed by cell row and cell column
    );
   
    $mjfreeway = \PhpOffice\PhpSpreadsheet\IOFactory::load("/var/www/html/homologar/ReporteMJ.xlsx");
    $highestrow = $mjfreeway->getActiveSheet()->getHighestRow();

    $spread = new Spreadsheet();
    $spread
        ->getProperties()
        ->setCreator("ECOMEDICS SAS")
        ->setLastModifiedBy('ECOMEDICS SAS')
        ->setTitle('Homologacion')
        ->setSubject('Homologacion')
        ->setDescription('Homologacion');
    
    $sheet = $spread->getActiveSheet();
    $sheet->setTitle("Hoja1");
    $sheet->setCellValueByColumnAndRow(1, 1 , "Articulo" );
    $sheet->setCellValueByColumnAndRow(2, 1 , "Lote" );
    $sheet->setCellValueByColumnAndRow(3, 1 , "Cantidad" );
    $sheet->setCellValueByColumnAndRow(4, 1 , "Serial" );
    $sheet->setCellValueByColumnAndRow(5, 1 , "Fecha de empaquetado" );
    $sheet->setCellValueByColumnAndRow(6, 1 , "Estilo" );
    $sheet->setCellValueByColumnAndRow(7, 1 , "Configuracion" );
    $sheet->setCellValueByColumnAndRow(8, 1 , "Tamaño" );
    $sheet->setCellValueByColumnAndRow(9, 1 , "Color" );


    


    for($i = 2 ; $i <= $highestrow ; $i++){

        $productnameMJ = $mjfreeway->getActiveSheet()->getCellByColumnAndRow(5, $i)->getValue();
        $key = array_search($productnameMJ, array_column($homologacion, 'B'));
       
        
        if($key !=null){
            $sheet->setCellValueByColumnAndRow(1, $i , $homologacion[$key+2]['A'] );
        }else{
            $sheet->setCellValueByColumnAndRow(1, $i , "Producto Sin homologar" );
        }
        $sheet->setCellValueByColumnAndRow(2, $i , $mjfreeway->getActiveSheet()->getCellByColumnAndRow(3, $i)->getValue());
        $sheet->setCellValueByColumnAndRow(3, $i , $mjfreeway->getActiveSheet()->getCellByColumnAndRow(4, $i)->getValue() );
        $sheet->setCellValueByColumnAndRow(4, $i , $mjfreeway->getActiveSheet()->getCellByColumnAndRow(2, $i)->getValue() );
        $sheet->setCellValueByColumnAndRow(5, $i , $mjfreeway->getActiveSheet()->getCellByColumnAndRow(7, $i)->getFormattedValue()  );
        $sheet->setCellValueByColumnAndRow(6, $i , $mjfreeway->getActiveSheet()->getCellByColumnAndRow(6, $i)->getValue() );
    }

    foreach (range('A', 'I') as $letra) {            
        $sheet->getColumnDimension($letra)->setAutoSize(true);
    }
    $fecha = date('Y_m_d');
    $writer = new Xlsx($spread);
    $writer->save('/var/www/html/homologar/Entrada_MV_'.$fecha.'.xlsx');


?>
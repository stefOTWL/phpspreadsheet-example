<?php
    require 'Dependencies/vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Spreadsheet;

    $header = array('Dummy Company', 'Dummy Report Header', 'Username');

    $connection = mysqli_connect('localhost', 'root', 'password123', 'test');

    if(!$connection){
        die("Connection failed");
    }else{
               
        $spreadsheet = new Spreadsheet();
        $userinfo = array();
        $tableinfo = array();

        //Query to get column names
        $query = "DESC demo";
        $result= mysqli_query($connection, $query);
        if(!$result){
            die("Query Failed");
        }
        while($row=mysqli_fetch_row($result)){
            $tableinfo[] = $row[0];
        }
        
        //Query to get table data
        $query = "SELECT * FROM demo";
        $result = mysqli_query($connection, $query);
        if(!$result){
            die("Query failed");
        }
        while($row = mysqli_fetch_row($result)){
            $userinfo[] = $row;
        }
        
        //Writing Data to Excel Sheet

        //Metadata
        $spreadsheet->getProperties()
        ->setCreator($header[2])
        ->setLastModifiedBy($header[2])
        ->setTitle($header[1])
        ->setSubject($header[1])
        ->setDescription($header[1])
        ->setCategory($header[1]);

        //Column Headers
        $spreadsheet->setActiveSheetIndex(0)
        ->fromArray(
            $tableinfo,
            NULL,
            'A5'
        );

        //Table data
        $spreadsheet->setActiveSheetIndex(0)
        ->fromArray(
            $userinfo,
            NULL,
            'A6'
        );

        //Getting Highest Row and Column
        $highestRow = $spreadsheet->getActiveSheet()
        ->getHighestRow();

        $highestColumn = $spreadsheet->getActiveSheet()
        ->getHighestColumn();

        // Merge Cells A1:C1
        
        if($highestColumn=='A'||$highestColumn=='B'||$highestColumn=='C'){
            $spreadsheet->getActiveSheet()->mergeCells('A1:Z1');
            $spreadsheet->getActiveSheet()->mergeCells('A3:Z3');
        }else{
            $spreadsheet->getActiveSheet()->mergeCells('A1:'.$highestColumn.'1');
            $spreadsheet->getActiveSheet()->mergeCells('A3:'.$highestColumn.'3');
        }
        
        //Rich Text for Company Header
        $company = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
        $payable = $company->createTextRun($header[0]);
        $payable->getFont()->setBold(true)
        ->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK ));

        $spreadsheet->getActiveSheet()->setCellValue('A1', $company);

        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFD1E8C0');

        //Rich Text For Report Header
        $report = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
        $payable = $report->createTextRun($header[1]);
        $payable->getFont()->setBold(true)
        ->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK ));

        $spreadsheet->getActiveSheet()->setCellValue('A3', $report);
        $spreadsheet->getActiveSheet()->getStyle('A3')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFD1E8C0');

        
        $spreadsheet->getActiveSheet()->getStyle('A5:'.$highestColumn.'5')
        ->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FF8AB5DF');
        
        $spreadsheet->getActiveSheet()->getStyle('A5:'.$highestColumn.'5')
        ->getFont()
        ->setBold(true)
        ->getColor()
        ->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
        
        //Cell Alignment
        $spreadsheet->getActiveSheet()
        ->getStyle('A5:'.$highestColumn.''.$highestRow)
        ->getAlignment()
        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        for($i='A'; $i<=$highestColumn; $i++){
            $spreadsheet->getActiveSheet()->getColumnDimension($i)->setAutoSize(true);
        }

        //Renaming Worksheet
        $spreadsheet->getActiveSheet()->setTitle('Worksheet');

        //Header and Footer
        $spreadsheet->getActiveSheet()->getHeaderFooter()
        ->setOddHeader('&L&H'. $spreadsheet->getProperties()->getTitle() . ' &R&D');
        $spreadsheet->getActiveSheet()->getHeaderFooter()
        ->setOddFooter('&RPage &P of &N');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="Report.xlsx"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;   
    }
?>
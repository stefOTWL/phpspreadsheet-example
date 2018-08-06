<?php
    require 'C:\Users\OCEAN\vendor\autoload.php';
    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Spreadsheet;

    $connection = mysqli_connect('localhost', 'root', 'password123', 'test');

    if(!$connection){
        die("Connection failed");
    }else{
               
        $spreadsheet = new Spreadsheet();
        $companyName = "Dummy Company";
        $reportHeader= "Dummy Report Header";
        $userinfo = array();
        $tableinfo = array();
        $query2 = "DESC demo";
        $result2= mysqli_query($connection, $query2);
        if(!$result2){
            die("Query Failed");
        }
        while($row=mysqli_fetch_row($result2)){
            $tableinfo[] = $row[0];
        }

        $query = "SELECT * FROM demo";
        $result = mysqli_query($connection, $query);
        if(!$result){
            die("Query failed");
        }
        while($row = mysqli_fetch_row($result)){
            $userinfo[] = $row;
        }
        
        //Writing Data to Excel Sheet
        $spreadsheet->setActiveSheetIndex(0)
        ->fromArray(
            $tableinfo,
            NULL,
            'A5'
        );

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

        //Renaming Worksheet
        $spreadsheet->getActiveSheet()->setTitle('Simple');

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
        $payable = $company->createTextRun($companyName);
        $payable->getFont()->setBold(true)
        ->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK ));

        $spreadsheet->getActiveSheet()->setCellValue('A1', $company);

        $spreadsheet->getActiveSheet()->getStyle('A1')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFD1E8C0');

        //Rich Text For Report Header
        $report = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
        $payable = $report->createTextRun($reportHeader);
        $payable->getFont()->setBold(true)
        ->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK ));

        $spreadsheet->getActiveSheet()->setCellValue('A3', $report);
        $spreadsheet->getActiveSheet()->getStyle('A3')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFD1E8C0');

        
        $spreadsheet->getActiveSheet()->getStyle('A5:'.$highestColumn.'5')->getFill()
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
        


        //Header and Footer
        $spreadsheet->getActiveSheet()->getHeaderFooter()
        ->setOddHeader('&L&HList of Users &R&D');
        $spreadsheet->getActiveSheet()->getHeaderFooter()
        ->setOddFooter('&L&B' . $spreadsheet->getProperties()->getTitle() . '&RPage &P of &N');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="Sample.xlsx"');
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
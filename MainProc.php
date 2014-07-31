<?php

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Description of MainProc
 *
 * @author user
 */
class MainProc {

    /**
     * 
     */
    public function Output() {
        
        $Info = $this->Read();
        $this->Write($Info);
    }


    /**
     * 
     * @return array
     */
    public function Read() {
        $FILE_PATH = "C:\Users\user\Desktop\sample\sample.xlsx";
        $phpExcelObj = PHPExcel_IOFactory::createReader('Excel2007');
        $readBook = $phpExcelObj->load($FILE_PATH);
        $readBook->setActiveSheetIndex(0);
        $sheet = $readBook->getActiveSheet();
        $Info = array();
        $rowIndex = 22;
        for ($i = 2; $i < $rowIndex; $i++) {
            $j = 0;
            $Obj1 = $sheet->getCellByColumnAndRow($j, $i)->getValue(); 
            $Obj2 = $sheet->getCellByColumnAndRow($j + 1, $i)->getValue(); 
            $Obj3 = $sheet->getCellByColumnAndRow($j + 2, $i)->getValue();
            array_push($Info, array($Obj1, $Obj2, $Obj3));
            print $Obj1 . "、  " . $Obj2 . "、  " . $Obj3 . "<br />\n";
        }
        return $Info;
    }

    /**
     * 
     * @param type $Info
     */
    public function Write($Info) {
        $book1 = new PHPExcel();
        $book1->setActiveSheetIndex(0);
        $sheet1 = $book1->getActiveSheet();
        $colIndex = 1;
        $rowIdx = 1;
        $targetIndex = 1;
        $Dict = array();
        foreach ($Info as $Val) {
            if (count($Dict) == 0 || !array_key_exists($Val[1], $Dict)) {
                $Dict[$Val[1]] = $colIndex;
                $sheet1->setCellValueByColumnAndRow($colIndex, 1, $Val[1]);
                $sheet1->setCellValueByColumnAndRow($colIndex, $rowIdx + 1, $Val[2]);
                $colIndex++;
            } else {
                $keyDate = $Val[1];
                $colIndexSameDate = $Dict[$keyDate];
                $sheet1->setCellValueByColumnAndRow($colIndexSameDate, $rowIdx + 1, $Val[2]);
            }
            $sheet1->setCellValueByColumnAndRow(0, $targetIndex + 1, $Val[0]);
            $rowIdx++;
            $targetIndex++;
        }
        $writer = PHPExcel_IOFactory::createWriter($book1, "Excel2007");
        $writer->save("./output3.xlsx");
    }

}

?>

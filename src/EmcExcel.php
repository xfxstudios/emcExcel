<?php
namespace Xfxstudios\Excel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Creator;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;

class EmcExcel{

    private $file     = "";
    private $name     = "";
    private $dir      = './';
    private $style    = false;
    private $download = false;
    private $array;

    public function initTest(){
        return "Init Test emcExcel Library";
    }

    //Setea el archivo a trabajar
    public function setFile($file){
        if(file_exists($this->dir.$file)){
            $this->file = $file;
            return $this;
        }else{
            $e = new \Exception("The file <b>$file</b> does not exist in the directory $this->dir", 1);
            throw $e;
        }
    }

    //Inicializa la lectura
    private function initRead(){
        if(file_exists($this->dir.$this->file)){
            $reader      = new Reader();
            $reader->setReadDataOnly(TRUE);
            $spreadsheet = $reader->load($this->dir.$this->file);
            return $spreadsheet;
        }
        $e = new \Exception("You must upload an Excel file to process", 1);
        throw $e;
    }

    public function setName($n){
        $this->name = $n;
        return $this;
    }

    public function setData($a){
        if(!is_array($a)){
            $e = new \Exception("You must send an array", 1);
            throw $e;
        }
        $this->array = $a;
        return $this;
    }

    public function isDownload(){
        $this->download = true;
        return $this;
    }

    public function setStyle($s){
        if(!is_array($s)){
            $e = new \Exception("You must send an array with the styles", 1);
            throw $e;           
        }
        $this->style = $s;
        return $this;
    }

    public function setDir($d){
        if(!file_exists($d)){
            $e = new \Exception("The indicated directory does not exist", 1);
            throw $e;             
        }
        $this->dir = $d;
        return $this;
    }

    //Estilos por defecto de titulos
    private function defaultStyle(){
        return [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endColor' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
        ];
    }

    //Retorna la letra de la columna
    private function getNameFromNumber($num) {
        $numeric = $num % 26;
        $letter  = chr(65 + $numeric);
        $num2    = intval($num / 26);
        if ($num2 > 0) {
            return $this->getNameFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }

    //Retorna la hoja activa del documento
    public function getWorkActiveSheet(){
        if(file_exists($this->dir.$this->name)){
            return $this->initRead()->getActiveSheet();
        }
        $e = new \Exception("There is no processed file", 1);
        throw $e;
    }

    //Retorna la info como array
    public function getArrayData(){
        if($this->file=='' || !file_exists($this->dir.$this->file)){
            $e  = new \Exception("There is no file to process", 1);
            throw $e;             
        }
        $worksheet   = $this->initRead()->getActiveSheet();
        $data        = $worksheet->toArray();
        return $data;
    }

    //Retorna una key buscada
    public function getMeKey($config=[]){
        if(empty($config)){
            return 0;
        }

        if(empty($config['array'])){
            $config['array'] = $this->getArrayData()[0];
            if(!$config['array']){
                $e = new \Exception("You must send an array", 1);
                throw $e;     
            }
        }
        if(!is_null($config['value'])){
            if(!is_array($config['array']) || !$config['array']){
                $e = new \Exception("You must send an array", 1);
                throw $e; 
            }
        }
        return (is_null($config['value']))?0:array_search($config['value'],$config['array'])+($config['offset']??0);
    }
    
    //Retorna un array de las letras en u rango dado
    public function getMeLetters($first,$last){
        $tmp = [];
        for($i=$first; $i<$last; $i++){
            array_push($tmp,$this->getNameFromNumber($i));//letras de columnas de libros   
        }
        return $tmp;
    }

    public function createFile(){
        if(!$this->array){
            $e = new \Exception("You must send an array of data to generate the file", 1);
            throw $e;                
        }
        if(!is_array($this->array[0])){
            $e = new \Exception("The array must be multidimensional: [[value, value ...], [value, value ...]] ", 1);
            throw $e; 
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $l = $this->getMeLetters(0,count($this->array[0]));
        
        for($f=0; $f<count($this->array); $f++){
            if(count($this->array[$f]) > 0 ){
                for($i=0; $i<count($l); $i++){
                    $sheet->setCellValue($l[$i].($f+1), $this->array[$f][$i]??'n/a');
                }
            }
        }

        for($s=0; $s<count($l); $s++){
            $sheet->getStyle($l[$s].'1')->applyFromArray((!$this->style)?$this->defaultStyle():$this->style);
        }

        $writer = new Creator($spreadsheet);
        $filename = ($this->name!="") ? $this->name.'.xlsx' : 'default'.uniqid().'.xlsx';
        $writer->save($this->dir.$filename);
        if($this->download){
            if(file_exists($this->dir.$filename)){
                header('location:'.$this->dir.$filename);
            }
        }
    }

}

?>
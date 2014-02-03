<?php
Yii::import('zii.widgets.grid.CGridView');
/**
 * Created by JetBrains PhpStorm.
 * User: mariano
 * Date: 1/7/14
 * Time: 12:03 PM
 * To change this template use File | Settings | File Templates.
 */

class EExcelWriter extends CGridView{
    public $fileName = null;
    public $stream = false; //stream to browser
    public $title = '';

    private $workBook;
    private $activeWorksheet;
    private $currentRow = 0;
    private $currentCol = 0;
    private $headerFormat;
    private $rowFormat;
    private $columnLenghts = array();

    public function init(){
        error_reporting(E_ALL ^ E_NOTICE);
        ini_set('display_errors', 0);
        Yii::import('application.vendors.*');

        require_once('excelwriter/class.writeexcel_workbookbig.inc.php');
        require_once('excelwriter/class.writeexcel_worksheet.inc.php');

        $this->initColumns();
        $this->workBook = new writeexcel_workbookbig($this->fileName);
        $this->activeWorksheet = $this->workBook->addworksheet();

        // add some default formatting
        $this->headerFormat = $this->workBook->addformat();
        $this->headerFormat->set_bold();

        $this->rowFormat = $this->workBook->addformat();

    }

    public function renderHeader()
    {

        foreach($this->columns as $column)
        {
            if($column instanceof CButtonColumn)
                $head = $column->header;
            elseif($column->header===null && $column->name!==null)
            {
                if($column->grid->dataProvider instanceof CActiveDataProvider)
                    $head = $column->grid->dataProvider->model->getAttributeLabel($column->name);
                else
                    $head = $column->name;
            } else
                $head =trim($column->header)!=='' ? $column->header : $column->grid->blankDisplay;

            $this->activeWorksheet->write( $this->currentRow , $this->currentCol, $head, $this->headerFormat);
            $this->columnLenghts[$this->currentCol] = strlen($head);
            $this->currentCol++;
        }
        $this->currentRow++;
    }

    public function renderBody()
    {
        $batchPageSize = 100;
        $provider = $this->dataProvider;
        $pager = $provider->getPagination();
        $pager->setItemCount($provider->getTotalItemCount());

        $pager->setPageSize($batchPageSize);
        $pageNumber = $pager->getPageCount();
        $allCount = 0;
        for($page = 0; $page < $pageNumber; $page++) {
            $pager->setCurrentPage($page);
            $this->dataProvider->setPagination($pager);
            $data=$this->dataProvider->getData(TRUE);
            $n=count($data);
            $allCount += $n;
            if($n>0)
            {
                for($row=0;$row<$n;++$row)
                    $this->renderRow($row, $page* $batchPageSize +$row);
            }

        }
        return $allCount;
    }

    public function renderRow($row)
    {
        $data=$this->dataProvider->getData();
        $this->currentCol = 0;
        foreach($this->columns as $n=>$column)
        {
            if($column instanceof CLinkColumn)
            {
                if($column->labelExpression!==null)
                    $value=$column->evaluateExpression($column->labelExpression,array('data'=>$data[$row],'row'=>$row));
                else
                    $value=$column->label;
            } elseif($column instanceof CButtonColumn)
                $value = ""; //Dont know what to do with buttons
            elseif($column->value!==null)
                $value=$this->evaluateExpression($column->value ,array('data'=>$data[$row]));
            elseif($column->name!==null) {
                //$value=$data[$row][$column->name];
                $value= CHtml::value($data[$row], $column->name);
                $value=$value===null ? "" : $column->grid->getFormatter()->format($value,'raw');
            }

            if(strpos($value, "$") !== FALSE) {
                $value = str_replace("$", "", $value);
            }
            $this->activeWorksheet->write( $this->currentRow , $this->currentCol, $value, $this->rowFormat);
            if( $this->columnLenghts[$this->currentCol] < strlen($value)){
                $this->columnLenghts[$this->currentCol] = strlen($value);
            }
            $this->currentCol++;
        }
        $this->currentRow++;
    }

    public function renderFooter($row)
    {
        $a=0;
        foreach($this->columns as $n=>$column)
        {
            $a=$a+1;
            if($column->footer)
            {
                $footer =trim($column->footer)!=='' ? $column->footer : $column->grid->blankDisplay;

                $cell = $this->objPHPExcel->getActiveSheet()->setCellValue($this->columnName($a).($row+2) ,$footer, true);
                if(is_callable($this->onRenderFooterCell))
                    call_user_func_array($this->onRenderFooterCell, array($cell, $footer));
            }
        }
    }

    public function autofitColumns(){
        foreach( $this->columnLenghts as $col => $length ){
            $width = 0.9 * $length;
            $this->activeWorksheet->set_column($col, $col, $width);
        }

    }


    public function run(){

        $this->renderHeader();
        $this->renderBody();
        $this->autofitColumns();
        $this->workBook->close();
        if($this->stream){ //output to browser
            header("Content-Type: application/x-msexcel; name=\"".basename($this->fileName)."\"");
            header("Content-Disposition: inline; filename=\"".basename($this->fileName)."\"");
            $fh=fopen($this->fileName, "rb");
            fpassthru($fh);
        }
    }


}
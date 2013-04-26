<?php
Yii::import('zii.widgets.grid.CGridView');

/**
* @author Nikola Kostadinov
* @license MIT License
* @version 0.3
* @link http://yiiframework.com/extension/eexcelview/
*
* @fork 0.33ab
* @forkversion 1.1
* @author A. Bennouna
* @organization tellibus.com
* @license MIT License
* @link https://github.com/tellibus/tlbExcelView
*/

/* Usage :
  $this->widget('application.components.widgets.tlbExcelView', array(
    'id'                   => 'some-grid',
    'dataProvider'         => $model->search(),
    'grid_mode'            => $production, // Same usage as EExcelView v0.33
    //'template'           => "{summary}\n{items}\n{exportbuttons}\n{pager}",
    'title'                => 'Some title - ' . date('d-m-Y - H-i-s'),
    'creator'              => 'Your Name',
    'subject'              => mb_convert_encoding('Something important with a date in French: ' . utf8_encode(strftime('%e %B %Y')), 'ISO-8859-1', 'UTF-8'),
    'description'          => mb_convert_encoding('Etat de production généré à la demande par l\'administrateur (some text in French).', 'ISO-8859-1', 'UTF-8'),
    'lastModifiedBy'       => 'Some Name',
    'sheetTitle'           => 'Report on ' . date('m-d-Y H-i'),
    'keywords'             => '',
    'category'             => '',
    'landscapeDisplay'     => true, // Default: false
    'A4'                   => true, // Default: false - ie : Letter (PHPExcel default)
    'RTL'                  => false, // Default: false
    'pageFooterText'       => '&RThis is page no. &P of &N pages', // Default: '&RPage &P of &N'
    'automaticSum'         => true, // Default: false
    'decimalSeparator'     => ',', // Default: '.'
    'thousandsSeparator'   => '.', // Default: ','
    //'displayZeros'       => false,
    //'zeroPlaceholder'    => '-',
    'sumLabel'             => 'Column totals:', // Default: 'Totals'
    'borderColor'          => '00FF00', // Default: '000000'
    'bgColor'              => 'FFFF00', // Default: 'FFFFFF'
    'textColor'            => 'FF0000', // Default: '000000'
    'rowHeight'            => 45, // Default: 15
    'headerBorderColor'    => 'FF0000', // Default: '000000'
    'headerBgColor'        => 'CCCCCC', // Default: 'CCCCCC'
    'headerTextColor'      => '0000FF', // Default: '000000'
    'headerHeight'         => 10, // Default: 20
    'footerBorderColor'    => '0000FF', // Default: '000000'
    'footerBgColor'        => '00FFCC', // Default: 'FFFFCC'
    'footerTextColor'      => 'FF00FF', // Default: '0000FF'
    'footerHeight'         => 50, // Default: 20
    'columns'              => $grid // an array of your CGridColumns
)); */

class tlbExcelView extends CGridView
{
    //the PHPExcel object
    public $libPath = 'ext.phpexcel.Classes.PHPExcel'; //the path to the PHP excel lib
    public static $objPHPExcel = null;
    public static $activeSheet = null;

    //Document properties
    public $creator = 'Nikola Kostadinov';
    public $title = null;
    public $subject = 'Subject';
    public $description = '';
    public $category = '';
    public $lastModifiedBy = 'A. Bennouna';
    public $keywords = '';
    public $sheetTitle = '';
    public $legal = 'PHPExcel generator http://phpexcel.codeplex.com/ - EExcelView Yii extension http://yiiframework.com/extension/eexcelview/ - Adaptation by A. Bennouna http://tellibus.com';
    public $landscapeDisplay = false;
    public $A4 = false;
    public $RTL = false;
    public $pageFooterText = '&RPage &P of &N';

    //config
    public $autoWidth = true;
    public $exportType = 'Excel5';
    public $disablePaging = true;
    public $filename = null; //export FileName
    public $stream = true; //stream to browser
    public $grid_mode = 'grid'; //Whether to display grid ot export it to selected format. Possible values(grid, export)
    public $grid_mode_var = 'grid_mode'; //GET var for the grid mode

    //options
    public $automaticSum = false;
    public $sumLabel = 'Totals';
    public $decimalSeparator = '.';
    public $thousandsSeparator = ',';
    public $displayZeros = false;
    public $zeroPlaceholder = '-';
    public $border_style;
    public $borderColor = '000000';
    public $bgColor = 'FFFFFF';
    public $textColor = '000000';
    public $rowHeight = 15;
    public $headerBorderColor = '000000';
    public $headerBgColor = 'CCCCCC';
    public $headerTextColor = '000000';
    public $headerHeight = 20;
    public $footerBorderColor = '000000';
    public $footerBgColor = 'FFFFCC';
    public $footerTextColor = '0000FF';
    public $footerHeight = 20;
    public static $fill_solid;
    public static $papersize_A4;
    public static $orientation_landscape;
    public static $horizontal_center;
    public static $horizontal_right;
    public static $vertical_center;
    public static $style = array();
    public static $headerStyle = array();
    public static $footerStyle = array();
    public static $summableColumns = array();
    
    //buttons config
    public $exportButtonsCSS = 'summary';
    public $exportButtons = array('Excel2007');
    public $exportText = 'Export to: ';

    //callbacks
    public $onRenderHeaderCell = null;
    public $onRenderDataCell = null;
    public $onRenderFooterCell = null;
    
    //mime types used for streaming
    public $mimeTypes = array(
        'Excel5'	=> array(
            'Content-type'=>'application/vnd.ms-excel',
            'extension'=>'xls',
            'caption'=>'Excel(*.xls)',
        ),
        'Excel2007'	=> array(
            'Content-type'=>'application/vnd.ms-excel',
            'extension'=>'xlsx',
            'caption'=>'Excel(*.xlsx)',				
        ),
        'PDF'		=>array(
            'Content-type'=>'application/pdf',
            'extension'=>'pdf',
            'caption'=>'PDF(*.pdf)',								
        ),
        'HTML'		=>array(
            'Content-type'=>'text/html',
            'extension'=>'html',
            'caption'=>'HTML(*.html)',												
        ),
        'CSV'		=>array(
            'Content-type'=>'application/csv',			
            'extension'=>'csv',
            'caption'=>'CSV(*.csv)',												
        )
    );

    public function init()
    {
        if (isset($_GET[$this->grid_mode_var])) {
            $this->grid_mode = $_GET[$this->grid_mode_var];
        }
        if (isset($_GET['exportType'])) {
            $this->exportType = $_GET['exportType'];
        }
        $lib = Yii::getPathOfAlias($this->libPath).'.php';
        if (($this->grid_mode == 'export') && (!file_exists($lib))) {
            $this->grid_mode = 'grid';
            Yii::log("PHP Excel lib not found($lib). Export disabled !", CLogger::LEVEL_WARNING, 'EExcelview');
        }
            
        if ($this->grid_mode == 'export') {
            if (!isset($this->title)) {
                $this->title = Yii::app()->getController()->getPageTitle();
            }
            $this->initColumns();
            //parent::init();
            //Autoload fix
            spl_autoload_unregister(array('YiiBase','autoload'));             
            Yii::import($this->libPath, true);

            // Get here some PHPExcel constants in order to use them elsewhere
            self::$papersize_A4 = PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4;
            self::$orientation_landscape = PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE;
            self::$fill_solid = PHPExcel_Style_Fill::FILL_SOLID;
            if (!isset($this->border_style)) {
                $this->border_style = PHPExcel_Style_Border::BORDER_THIN;
            }
            self::$horizontal_center = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
            self::$horizontal_right = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
            self::$vertical_center = PHPExcel_Style_Alignment::VERTICAL_CENTER;

            spl_autoload_register(array('YiiBase','autoload'));  

            // Creating a workbook
            self::$objPHPExcel = new PHPExcel();
            self::$activeSheet = self::$objPHPExcel->getActiveSheet();

            // Set some basic document properties
            if ($this->landscapeDisplay) {
                self::$activeSheet->getPageSetup()->setOrientation(self::$orientation_landscape);
            }

            if ($this->A4) {
                self::$activeSheet->getPageSetup()->setPaperSize(self::$papersize_A4);
            }

            if ($this->RTL) {
                self::$activeSheet->setRightToLeft(true);
            }

            self::$objPHPExcel->getProperties()
                ->setTitle($this->title)
                ->setCreator($this->creator)
                ->setSubject($this->subject)
                ->setDescription($this->description . ' // ' . $this->legal)
                ->setCategory($this->category)
                ->setLastModifiedBy($this->lastModifiedBy)
                ->setKeywords($this->keywords);

            // Initialize styles that will be used later
            self::$style = array(
                'borders' => array(
                    'allborders' => array(
                                        'style' => $this->border_style,
                                        'color' => array('rgb' => $this->borderColor),
                                    ),
                ),
                'fill' => array(
                    'type' => self::$fill_solid,
                    'color' => array('rgb' => $this->bgColor),
                ),
                'font' => array(
                    //'bold' => false,
                    'color' => array('rgb' => $this->textColor),
                )
            );
            self::$headerStyle = array(
                'borders' => array(
                    'allborders' => array(
                                        'style' => $this->border_style,
                                        'color' => array('rgb' => $this->headerBorderColor),
                                    ),
                ),
                'fill' => array(
                    'type' => self::$fill_solid,
                    'color' => array('rgb' => $this->headerBgColor),
                ),
                'font' => array(
                    'bold' => true,
                    'color' => array('rgb' => $this->headerTextColor),
                )
            );
            self::$footerStyle = array(
                'borders' => array(
                    'allborders' => array(
                                        'style' => $this->border_style,
                                        'color' => array('rgb' => $this->footerBorderColor),
                                    ),
                ),
                'fill' => array(
                    'type' => self::$fill_solid,
                    'color' => array('rgb' => $this->footerBgColor),
                ),
                'font' => array(
                    'bold' => true,
                    'color' => array('rgb' => $this->footerTextColor),
                )
            );
        } else {
            parent::init();
        }
    }

    public function renderHeader()
    {
        $a = 0;
        foreach ($this->columns as $column) {
            $a = $a + 1;
            if ($column instanceof CButtonColumn) {
                $head = $column->header;
            } else if (($column->header === null) && ($column->name !== null)) {
                if($column->grid->dataProvider instanceof CActiveDataProvider) {
                    $head = $column->grid->dataProvider->model->getAttributeLabel($column->name);
                } else {
                    $head = $column->name;
                }
            } else {
                $head =trim($column->header)!=='' ? $column->header : $column->grid->blankDisplay;
            }

            $cell = self::$activeSheet->setCellValue($this->columnName($a) . '1', $head, true);

            if (is_callable($this->onRenderHeaderCell)) {
                call_user_func_array($this->onRenderHeaderCell, array($cell, $head));				
            }
        }

        // Format the header row
        $header = self::$activeSheet->getStyle($this->columnName(1) . '1:' . $this->columnName($a) . '1');
        $header->getAlignment()
            ->setHorizontal(self::$horizontal_center)
            ->setVertical(self::$vertical_center);
        $header->applyFromArray(self::$headerStyle);
        self::$activeSheet->getRowDimension(1)->setRowHeight($this->headerHeight);
    }

    public function renderBody()
    {
        if ($this->disablePaging) {
            //if needed disable paging to export all data
            $this->dataProvider->pagination = false;
        }
        $data = $this->dataProvider->getData();
        $n = count($data);

        if ($n > 0) {
            for ($row = 0; $row < $n; ++$row) {
                $this->renderRow($row);
            }
        }
        return $n;
    }

    public function renderRow($row)
    {
        $data = $this->dataProvider->getData();			

        $a = 0;
        foreach ($this->columns as $n => $column) {
            if ($column instanceof CLinkColumn) {
                if ($column->labelExpression !== null) {
                    $value = $column->evaluateExpression($column->labelExpression, array('data' => $data[$row], 'row' => $row));
                } else {
                    $value = $column->label;
                }
            } else if ($column instanceof CButtonColumn) {
                $value = ""; //Dont know what to do with buttons
            } else if ($column->value !== null) {
                $value = $this->evaluateExpression($column->value, array('data' => $data[$row]));
            } else if ($column->name !== null) { 
                //$value = $data[$row][$column->name];
                $value = CHtml::value($data[$row], $column->name);
                $value = $value === null ? "" : $column->grid->getFormatter()->format($value, 'raw');
            }

            $a++;

            // Check if the cell value is a number, then format it accordingly
            // May be improved notably by exposing the formats as public
            // May be usable only for French-style number formatting ?
            if (preg_match("/^[0-9]*\\" . $this->thousandsSeparator . "[0-9]*\\" . $this->decimalSeparator . "[0-9]*$/", strip_tags($value))) {
                $content = str_replace($this->decimalSeparator, '.', str_replace($this->thousandsSeparator, '', strip_tags($value)));
                $format = '#\.##0.00';
            } else if (preg_match("/^[0-9]*\\" . $this->decimalSeparator . "[0-9]*$/", strip_tags($value))) {
                $content = str_replace($this->decimalSeparator, '.', strip_tags($value));
                $format = '0.00';
            } else if (!$this->displayZeros && ((strip_tags($value) === '0') || (strip_tags($value) === $this->zeroPlaceholder))) {
                $content = $this->zeroPlaceholder;
                self::$activeSheet->getStyle($this->columnName($a) . ($row + 2))->getAlignment()->setHorizontal(self::$horizontal_right);
                $format = '0.00';
            } else {
                $content = strip_tags($value);
                $format = null;
            }

            $cell = self::$activeSheet->setCellValue($this->columnName($a) . ($row + 2), $content, true);

            // Format each cell's number - if any
            if (!is_null($format)) {
                self::$summableColumns[$a] = $a;
                self::$activeSheet->getStyle($this->columnName($a) . ($row + 2))->getNumberFormat()->setFormatCode($format);
            }

            if(is_callable($this->onRenderDataCell)) {
                call_user_func_array($this->onRenderDataCell, array($cell, $data[$row], $value));
            }
        }

        // Format the row globally
        $renderedRow = self::$activeSheet->getStyle('A' . ($row + 2) . ':' . $this->columnName($a) . ($row + 2));
        $renderedRow->getAlignment()->setVertical(self::$vertical_center);
        $renderedRow->applyFromArray(self::$style);
        self::$activeSheet->getRowDimension($row + 2)->setRowHeight($this->rowHeight);
    }

    public function renderFooter($row)
    {
        $a = 0;
        foreach ($this->columns as $n => $column) {
            $a = $a + 1;
            if ($column->footer) {
                $footer = trim($column->footer) !== '' ? $column->footer : $column->grid->blankDisplay;

                $cell = self::$activeSheet->setCellValue($this->columnName($a) . ($row + 2), $footer, true);

                if(is_callable($this->onRenderFooterCell)) {
                    call_user_func_array($this->onRenderFooterCell, array($cell, $footer));				
                }
            } else if ($this->automaticSum && in_array($a, self::$summableColumns)) {
                // We want to render automatic sums in the footer if no footer was already present in the grid
                $cell = self::$activeSheet->setCellValue($this->columnName($a) . ($row + 2), '=SUM(' . $this->columnName($a) . '2:' . $this->columnName($a) . ($row + 1) . ')', true);
                $sum = self::$activeSheet->getCell($this->columnName($a) . ($row + 2))->getCalculatedValue();
                if ($sum < 1000) {
                    $format = '0.00';                    
                } else if ($sum < 1000000) {
                    $format = '#\.##0.00';
                } else {
                    $format = '#\.###\.##0.00';
                }

                // We won't set the whole row's borders and number format, so proceed with each cell individually
                self::$activeSheet->getStyle($this->columnName($a) . ($row + 2))
                    ->applyFromArray(self::$footerStyle)
                    ->getNumberFormat()->setFormatCode($format);

                if(is_callable($this->onRenderFooterCell)) {
                    call_user_func_array($this->onRenderFooterCell, array($cell, $footer));				
                }

                // Add a label before the first summable column (supposing it's not the first…)
                if (current(self::$summableColumns) == $a) {
                    $cell = self::$activeSheet->setCellValue($this->columnName($a - 1) . ($row + 2), $this->sumLabel, true);
                    self::$activeSheet->getStyle($this->columnName($a - 1) . ($row + 2))
                        ->applyFromArray(array('font' => array('bold' => true)))
                        ->getAlignment()->setHorizontal(self::$horizontal_right);
                    if(is_callable($this->onRenderFooterCell)) {
                        call_user_func_array($this->onRenderFooterCell, array($cell, $footer));				
                    }
                }
            }
        }

        // Some global formatting for the footer in case of automatic sum
        if ($this->automaticSum) {
            self::$activeSheet->getStyle('A' . ($row + 2) . ':' . $this->columnName($a) . ($row + 2))->getAlignment()->setVertical(self::$vertical_center);
            self::$activeSheet->getRowDimension($row + 2)->setRowHeight($this->footerHeight);
        }
    }

    public function run()
    {
        if ($this->grid_mode == 'export') {
            $this->renderHeader();
            $row = $this->renderBody();
            $this->renderFooter($row);

            //set auto width
            if ($this->autoWidth) {
                foreach ($this->columns as $n => $column) {
                    $cell = self::$activeSheet->getColumnDimension($this->columnName($n + 1))->setAutoSize(true);
                }
            }

            // Set some additional properties
            self::$activeSheet
                ->setTitle($this->sheetTitle)
                ->getSheetView()->setZoomScale(50);
            self::$activeSheet->getHeaderFooter()
                ->setOddHeader('&C' . $this->sheetTitle)
                ->setOddFooter('&L&B' . self::$objPHPExcel->getProperties()->getTitle() . $this->pageFooterText);
            self::$activeSheet->getPageSetup()
                ->setPrintArea('A1:' . $this->columnName(count($this->columns)) . ($row + 2))
                ->setFitToWidth();

            //create writer for saving
            $objWriter = PHPExcel_IOFactory::createWriter(self::$objPHPExcel, $this->exportType);
            if (!$this->stream) {
                $objWriter->save($this->filename);
            } else {
                //output to browser
                if(!$this->filename) {
                    $this->filename = $this->title;
                }
                $this->cleanOutput();
                header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
                header('Pragma: public');
                header('Content-type: '.$this->mimeTypes[$this->exportType]['Content-type']);
                header('Content-Disposition: attachment; filename="' . $this->filename . '.' . $this->mimeTypes[$this->exportType]['extension'] . '"');
                header('Cache-Control: max-age=0');				
                $objWriter->save('php://output');			
                Yii::app()->end();
            }
        } else {
            parent::run();
        }
    }

    /**
    * Returns the corresponding Excel column.(Abdul Rehman from yii forum)
    * 
    * @param int $index
    * @return string
    */
    public function columnName($index)
    {
        --$index;
        if (($index >= 0) && ($index < 26)) {
            return chr(ord('A') + $index);
        } else if ($index > 25) {
            return ($this->columnName($index / 26)) . ($this->columnName($index%26 + 1));
        } else {
            throw new Exception("Invalid Column # " . ($index + 1));
        }
    }
    
    public function renderExportButtons()
    {
        foreach ($this->exportButtons as $key => $button) {
            $item = is_array($button) ? CMap::mergeArray($this->mimeTypes[$key], $button) : $this->mimeTypes[$button];
            $type = is_array($button) ? $key : $button;
            $url = parse_url(Yii::app()->request->requestUri);
            //$content[] = CHtml::link($item['caption'], '?'.$url['query'].'exportType='.$type.'&'.$this->grid_mode_var.'=export');
            if (key_exists('query', $url)) {
                $content[] = CHtml::link($item['caption'], '?' . $url['query'] . '&exportType=' . $type . '&' . $this->grid_mode_var . '=export');          
            } else {
                $content[] = CHtml::link($item['caption'], '?exportType=' . $type . '&' . $this->grid_mode_var . '=export');				
            }
        }
        if ($content) {
            echo CHtml::tag('div', array('class' => $this->exportButtonsCSS), $this->exportText.implode(', ', $content));	
        }

    }
    
    /**
    * Performs cleaning on mutliple levels.
    * 
    * From le_top @ yiiframework.com
    * 
    */
    private static function cleanOutput() 
    {
        for ($level = ob_get_level(); $level > 0; --$level) {
            @ob_end_clean();
        }
    }
}